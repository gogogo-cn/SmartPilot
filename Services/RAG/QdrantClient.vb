Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Net.Http
Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services.RAG

    ''' <summary>
    ''' Qdrant 向量数据库 HTTP API 客户端
    ''' 使用 System.Net.Http.HttpClient 进行 REST 调用
    ''' </summary>
    Public Class QdrantClient
        Implements IDisposable

        Private _httpClient As HttpClient
        Private _endpoint As String
        Private _apiKey As String
        Private _isDisposed As Boolean = False

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="endpoint">Qdrant 服务地址，如 http://localhost:6333</param>
        ''' <param name="apiKey">API Key（可选）</param>
        Public Sub New(endpoint As String, Optional apiKey As String = Nothing)
            ' 明确校验：endpoint 为空时给出可读错误，而不是把 Nothing 留到后面变成 NRE
            If String.IsNullOrWhiteSpace(endpoint) Then
                Throw New ArgumentException("Qdrant endpoint 不能为空（请检查 RAG 配置）", "endpoint")
            End If
            _endpoint = NormalizeLocalEndpoint(endpoint)
            _apiKey = apiKey

            ' Qdrant 是本地/内网服务：必须禁用系统代理。
            ' 实测（Win10 + 系统代理）：默认 UseProxy=True 时 localhost 请求被代理多跳，
            ' 不可达端口要 4 秒级才失败、可达服务也要 2 秒级才返回（直连仅毫秒级）——
            ' 直接表现为「创建/删除知识库很慢」。
            Dim handler As New HttpClientHandler()
            handler.UseProxy = False
            _httpClient = New HttpClient(handler)
            _httpClient.Timeout = TimeSpan.FromSeconds(30)
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "SmartPilot-RAG/1.0")

            If Not String.IsNullOrEmpty(apiKey) Then
                _httpClient.DefaultRequestHeaders.Add("api-key", apiKey)
            End If
        End Sub

        ''' <summary>
        ''' 规范化本地服务地址：localhost → 127.0.0.1。
        ''' 实测（2026-10-21，Win10）：.NET HttpClient 对 "localhost" 会优先尝试 IPv6 ::1，
        ''' 本机未监听 v6 时每次请求固定多花约 2 秒才 fallback IPv4
        ''' （localhost:11434 2083ms vs 127.0.0.1:11434 25ms）——直接表现为「创建/删除很慢」。
        ''' </summary>
        Private Shared Function NormalizeLocalEndpoint(endpoint As String) As String
            If String.IsNullOrWhiteSpace(endpoint) Then Return ""
            Try
                Dim uri As New Uri(endpoint.Trim())
                If uri.Host = "localhost" Then
                    Dim builder As New UriBuilder(uri)
                    builder.Host = "127.0.0.1"
                    Return builder.Uri.ToString().TrimEnd("/"c)
                End If
                Return endpoint.Trim().TrimEnd("/"c)
            Catch
                Return endpoint.Trim().TrimEnd("/"c)
            End Try
        End Function

        ''' <summary>
        ''' TCP 层端口探测（快速失败）：本机安全软件会静默丢弃对未监听端口的 SYN，
        ''' HTTP 请求要 2 秒级才超时；异步 TCP 探测可在 500ms 内判定端口死活。
        ''' </summary>
        Private Async Function TcpProbeAsync(endpoint As String, timeoutMs As Integer) As Task(Of Boolean)
            Try
                Dim uri As New Uri(endpoint)
                Using client As New System.Net.Sockets.TcpClient()
                    Dim connectTask = client.ConnectAsync(uri.Host, uri.Port)
                    Dim delayTask = Task.Delay(timeoutMs)
                    Dim completed = Await Task.WhenAny(connectTask, delayTask)
                    If completed Is delayTask Then
                        ' 超时未连上：绝不再 Await connectTask（其内部网络超时要 2 秒级），
                        ' 仅挂续任务标记异常已观察，避免未观察任务异常
                        Dim ignoreContinuation = connectTask.ContinueWith(Sub(t)
                                                                              Dim ignore = t.Exception
                                                                          End Sub, TaskContinuationOptions.OnlyOnFaulted)
                        Return False
                    End If
                    Await connectTask
                    Return client.Connected
                End Using
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[Qdrant] TCP 探测异常: {ex.Message}")
                Return False
            End Try
        End Function

#Region "集合管理"

        ''' <summary>
        ''' 创建集合（如果不存在）
        ''' 向量维度 1024（BGE），距离度量 Cosine
        ''' </summary>
        ''' <param name="collectionName">集合名称</param>
        ''' <param name="vectorSize">向量维度，默认 1024</param>
        ''' <returns>是否成功创建或已存在</returns>
        Public Async Function CreateCollectionIfNotExists(collectionName As String,
                                                          Optional vectorSize As Integer = 1024) As Task(Of Boolean)
            Try
                ' 先检查是否已存在
                Dim exists = Await CollectionExistsAsync(collectionName)
                If exists Then Return True

                Dim requestBody = New With {
                    .vectors = New With {
                        .size = vectorSize,
                        .distance = "Cosine"
                    }
                }

                Dim jsonContent = JsonConvert.SerializeObject(requestBody)
                Dim content = New StringContent(jsonContent, Encoding.UTF8, "application/json")

                Dim response = Await _httpClient.PutAsync($"{_endpoint}/collections/{collectionName}", content)

                If response.IsSuccessStatusCode Then
                    System.Diagnostics.Debug.WriteLine($"[Qdrant] 集合 '{collectionName}' 创建成功 (维度={vectorSize})")
                    Return True
                Else
                    Dim errorMsg = Await response.Content.ReadAsStringAsync()
                    System.Diagnostics.Debug.WriteLine($"[Qdrant] 创建集合失败: {response.StatusCode} - {errorMsg}")
                    Return False
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[Qdrant] CreateCollectionIfNotExists 异常: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 检查集合是否存在
        ''' </summary>
        Private Async Function CollectionExistsAsync(collectionName As String) As Task(Of Boolean)
            Try
                ' 探测类请求 3 秒快速失败（Qdrant 半死时不拖慢创建流程 30 秒）
                Using cts As New Threading.CancellationTokenSource(TimeSpan.FromSeconds(3))
                    Dim response = Await _httpClient.GetAsync($"{_endpoint}/collections/{collectionName}", cts.Token)
                    Return response.IsSuccessStatusCode
                End Using
            Catch ex As Exception
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 删除整个集合
        ''' </summary>
        Public Async Function DeleteCollectionAsync(collectionName As String) As Task(Of Boolean)
            Try
                ' wait=false：不等待 Qdrant 同步完成即返回（大集合删除时 wait=true 会明显变慢）
                Dim response = Await _httpClient.DeleteAsync($"{_endpoint}/collections/{collectionName}?wait=false")

                If response.IsSuccessStatusCode Then
                    System.Diagnostics.Debug.WriteLine($"[Qdrant] 集合 '{collectionName}' 已删除")
                    Return True
                Else
                    Dim errorMsg = Await response.Content.ReadAsStringAsync()
                    System.Diagnostics.Debug.WriteLine($"[Qdrant] 删除集合失败: {response.StatusCode} - {errorMsg}")
                    Return False
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[Qdrant] DeleteCollectionAsync 异常: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 获取集合信息
        ''' </summary>
        Public Async Function GetCollectionInfoAsync(collectionName As String) As Task(Of Object)
            Try
                Dim response = Await _httpClient.GetAsync($"{_endpoint}/collections/{collectionName}")

                If response.IsSuccessStatusCode Then
                    Dim jsonStr = Await response.Content.ReadAsStringAsync()
                    Return JsonConvert.DeserializeObject(jsonStr)
                Else
                    Return Nothing
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[Qdrant] GetCollectionInfoAsync 异常: {ex.Message}")
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 获取所有集合列表
        ''' </summary>
        Public Async Function GetCollectionsAsync() As Task(Of List(Of String))
            Try
                Dim collections As New List(Of String)()
                Dim response = Await _httpClient.GetAsync($"{_endpoint}/collections")

                If response.IsSuccessStatusCode Then
                    Dim jsonStr = Await response.Content.ReadAsStringAsync()
                    Dim jsonObj = JObject.Parse(jsonStr)

                    Dim resultArray = jsonObj("result")?("collections")
                    If resultArray IsNot Nothing Then
                        For Each col In resultArray
                            collections.Add(col("name")?.ToString())
                        Next
                    End If
                End If

                Return collections
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[Qdrant] GetCollectionsAsync 异常: {ex.Message}")
                Return New List(Of String)()
            End Try
        End Function

#End Region

#Region "点操作"

        ''' <summary>
        ''' 批量插入/更新向量点（Upsert）
        ''' </summary>
        ''' <param name="collectionName">集合名称</param>
        ''' <param name="points">点列表</param>
        ''' <returns>操作结果</returns>
        Public Async Function UpsertPointsAsync(collectionName As String,
                                                 points As List(Of QdrantPoint)) As Task(Of QdrantUpsertResult)
            ' 【安全修复】参数验证
            If String.IsNullOrWhiteSpace(collectionName) Then
                Return New QdrantUpsertResult With {
                    .Success = False,
                    .Status = "error",
                    .ErrorMessage = "集合名称不能为空"
                }
            End If

            If points Is Nothing OrElse points.Count = 0 Then
                Return New QdrantUpsertResult With {
                    .Success = False,
                    .Status = "error",
                    .ErrorMessage = "点列表不能为空"
                }
            End If

            Try
                Dim pointObjects = New List(Of Object)()

                For Each point In points
                    ' 【安全修复】验证每个点的数据
                    If point.Id = Guid.Empty Then
                        System.Diagnostics.Debug.WriteLine("[QdrantClient] 警告: 发现空的Point ID")
                    End If
                    
                    pointObjects.Add(New With {
                        .id = point.Id,
                        .vector = point.Vector,
                        .payload = point.Payload
                    })
                Next

                Dim requestBody = New With {
                    .points = pointObjects
                }

                Dim jsonContent = JsonConvert.SerializeObject(requestBody)
                Dim content = New StringContent(jsonContent, Encoding.UTF8, "application/json")

                Dim sw = Diagnostics.Stopwatch.StartNew()
                Dim response = Await _httpClient.PutAsync($"{_endpoint}/collections/{collectionName}/points", content)
                sw.Stop()

                Dim jsonStr = Await response.Content.ReadAsStringAsync()

                Dim result As New QdrantUpsertResult With {
                    .Success = response.IsSuccessStatusCode,
                    .OperationId = "",
                    .Status = If(response.IsSuccessStatusCode, "completed", "failed"),
                    .LatencyMs = sw.ElapsedMilliseconds
                }

                ' 【安全修复】使用类型安全的访问方式解析响应，避免Runtime异常
                Try
                    If Not String.IsNullOrWhiteSpace(jsonStr) Then
                        Dim jsonObj = JObject.Parse(jsonStr)

                        ' 安全访问嵌套属性，避免NullReferenceException
                        Dim resultObj = jsonObj("result")
                        If resultObj IsNot Nothing Then
                            result.OperationId = resultObj("operation_id")?.ToString()
                            
                            ' 使用类型安全的status读取方式
                            Dim statusValue = resultObj("status")?.ToString()
                            If Not String.IsNullOrWhiteSpace(statusValue) Then
                                result.Status = statusValue
                            End If
                        End If
                    End If
                Catch ex As JsonException
                    System.Diagnostics.Debug.WriteLine($"[QdrantClient] JSON解析异常: {ex.Message}")
                    ' JSON解析失败不影响整体结果，保持默认值
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"[QdrantClient] 解析响应异常: {ex.Message}")
                End Try

                If result.Success Then
                    System.Diagnostics.Debug.WriteLine($"[QdrantClient] Upsert 成功: {points.Count} 个点, 耗时 {result.LatencyMs}ms")
                Else
                    System.Diagnostics.Debug.WriteLine($"[QdrantClient] Upsert 失败: {response.StatusCode} - {jsonStr}")
                End If

                Return result
            Catch ex As JsonException
                System.Diagnostics.Debug.WriteLine($"[QdrantClient] JSON序列化/反序列化异常: {ex.Message}")
                Return New QdrantUpsertResult With {
                    .Success = False,
                    .Status = "error",
                    .ErrorMessage = $"JSON处理异常: {ex.Message}"
                }
            Catch ex As HttpRequestException
                System.Diagnostics.Debug.WriteLine($"[QdrantClient] 网络请求异常: {ex.Message}")
                Return New QdrantUpsertResult With {
                    .Success = False,
                    .Status = "error",
                    .ErrorMessage = $"网络请求异常: {ex.Message}"
                }
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[QdrantClient] 未预期异常: {ex.Message}")
                Return New QdrantUpsertResult With {
                    .Success = False,
                    .Status = "error",
                    .ErrorMessage = $"未预期异常: {ex.Message}"
                }
            End Try
        End Function

        ''' <summary>
        ''' 向量相似度搜索
        ''' </summary>
        ''' <param name="collectionName">集合名称</param>
        ''' <param name="vector">查询向量</param>
        ''' <param name="topK">返回数量</param>
        ''' <param name="scoreThreshold">分数阈值</param>
        ''' <param name="filter">过滤条件（可选）</param>
        ''' <returns>检索结果列表</returns>
        Public Async Function SearchAsync(collectionName As String,
                                           vector As Single(),
                                           Optional topK As Integer = 5,
                                           Optional scoreThreshold As Single = 0.0F,
                                           Optional filter As Object = Nothing) As Task(Of List(Of SearchResult))
            Try
                Dim results As New List(Of SearchResult)()

                Dim requestBody As New Dictionary(Of String, Object) From {
                    {"vector", vector},
                    {"limit", topK},
                    {"with_payload", True}
                }

                If scoreThreshold > 0 Then
                    requestBody.Add("score_threshold", scoreThreshold)
                End If

                If filter IsNot Nothing Then
                    requestBody.Add("filter", filter)
                End If

                Dim jsonContent = JsonConvert.SerializeObject(requestBody)
                Dim content = New StringContent(jsonContent, Encoding.UTF8, "application/json")

                Dim sw = Diagnostics.Stopwatch.StartNew()
                Dim response = Await _httpClient.PostAsync($"{_endpoint}/collections/{collectionName}/points/search", content)
                sw.Stop()

                If response.IsSuccessStatusCode Then
                    Dim jsonStr = Await response.Content.ReadAsStringAsync()
                    Dim jsonObj = JObject.Parse(jsonStr)

                    Dim resultArray = jsonObj("result")
                    If resultArray IsNot Nothing Then
                        For Each item In resultArray
                            Dim scoreVal = Convert.ToSingle(item("score"))
                            Dim payload = item("payload")

                            Dim searchResult As New SearchResult With {
                                .ChunkId = item("id")?.ToString(),
                                .Score = scoreVal,
                                .Text = payload?("text")?.ToString(),
                                .DocTitle = payload?("doc_title")?.ToString(),
                                .SourceSection = payload?("source_section")?.ToString()
                            }

                            results.Add(searchResult)
                        Next
                    End If

                    System.Diagnostics.Debug.WriteLine($"[Qdrant] 搜索完成: 返回 {results.Count} 个结果, 耗时 {sw.ElapsedMilliseconds}ms")
                Else
                    System.Diagnostics.Debug.WriteLine($"[Qdrant] 搜索失败: {response.StatusCode}")
                End If

                Return results
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[Qdrant] SearchAsync 异常: {ex.Message}")
                Return New List(Of SearchResult)()
            End Try
        End Function

        ''' <summary>
        ''' 按 ID 删除点
        ''' </summary>
        Public Async Function DeletePointsAsync(collectionName As String,
                                                pointIds As List(Of Guid)) As Task(Of Boolean)
            Try
                Dim requestBody = New With {
                    .points = pointIds
                }

                Dim jsonContent = JsonConvert.SerializeObject(requestBody)
                Dim content = New StringContent(jsonContent, Encoding.UTF8, "application/json")

                Dim response = Await _httpClient.PostAsync($"{_endpoint}/collections/{collectionName}/points/delete", content)

                If response.IsSuccessStatusCode Then
                    System.Diagnostics.Debug.WriteLine($"[Qdrant] 删除点成功: {pointIds.Count} 个")
                    Return True
                Else
                    Dim errorMsg = Await response.Content.ReadAsStringAsync()
                    System.Diagnostics.Debug.WriteLine($"[Qdrant] 删除点失败: {response.StatusCode} - {errorMsg}")
                    Return False
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[Qdrant] DeletePointsAsync 异常: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 按 Payload Filter 删除点（用于删除某个文档的所有 chunk）
        ''' </summary>
        Public Async Function DeleteByFilterAsync(collectionName As String, docId As String) As Task(Of Boolean)
            Try
                ' 【安全修复】使用 JsonConvert.SerializeObject 构建 JSON，避免 docId 含特殊字符时注入
                Dim filterObj = New With {
                    .must = New Object() {
                        New With {
                            .key = "doc_id",
                            .match = New With {
                                .value = docId
                            }
                        }
                    }
                }
                Dim filterJson = JsonConvert.SerializeObject(filterObj)

                Dim url = $"{_endpoint}/collections/{collectionName}/points/delete"
                Dim content = New StringContent($"{{""filter"": {filterJson}}}", Encoding.UTF8, "application/json")

                Using response = Await _httpClient.PostAsync(url, content)
                    If response.IsSuccessStatusCode Then
                        Return True
                    Else
                        System.Diagnostics.Debug.WriteLine($"[Qdrant] DeleteByFilter 失败: {response.StatusCode}")
                        Return False
                    End If
                End Using
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[Qdrant] DeleteByFilter 异常: {ex.Message}")
                Return False
            End Try
        End Function

#End Region

#Region "健康检查"

        ''' <summary>
        ''' 检查连接是否可用
        ''' </summary>
        Public Async Function HealthCheckAsync() As Task(Of Boolean)
            ' 先做 TCP 探测（500ms 快速失败）：端口未监听时不必等 HTTP 层秒级超时
            If Not Await TcpProbeAsync(_endpoint, 500) Then
                System.Diagnostics.Debug.WriteLine($"[Qdrant] TCP 探测失败（端口未监听或不可达）: {_endpoint}")
                Return False
            End If
            Try
                Dim sw = Diagnostics.Stopwatch.StartNew()
                ' 探测类请求 3 秒快速失败（不占用 30 秒的客户端级超时）
                Using cts As New Threading.CancellationTokenSource(TimeSpan.FromSeconds(3))
                    Dim response = Await _httpClient.GetAsync($"{_endpoint}/healthz", cts.Token)
                    sw.Stop()

                    If response.IsSuccessStatusCode Then
                        System.Diagnostics.Debug.WriteLine($"[Qdrant] 健康检查通过, 耗时 {sw.ElapsedMilliseconds}ms")
                        Return True
                    Else
                        System.Diagnostics.Debug.WriteLine($"[Qdrant] 健康检查失败: {response.StatusCode}")
                        Return False
                    End If
                End Using
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[Qdrant] HealthCheckAsync 异常: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 获取详细状态信息
        ''' </summary>
        Public Async Function GetClusterInfoAsync() As Task(Of String)
            Try
                Dim response = Await _httpClient.GetAsync($"{_endpoint}/cluster")
                If response.IsSuccessStatusCode Then
                    Return Await response.Content.ReadAsStringAsync()
                End If
                Return ""
            Catch ex As Exception
                Return $"{{""error"": ""{ex.Message}""}}"
            End Try
        End Function

#End Region

#Region "IDisposable Support"

        Protected Overloads Sub Dispose(disposing As Boolean)
            If Not _isDisposed Then
                If disposing Then
                    ' 释放托管资源
                    If _httpClient IsNot Nothing Then
                        _httpClient.Dispose()
                        _httpClient = Nothing
                    End If
                End If
                _isDisposed = True
            End If
        End Sub

        Public Overloads Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

#End Region

    End Class

    ''' <summary>
    ''' Qdrant 点数据结构
    ''' </summary>
    Public Class QdrantPoint
        Public Property Id As Guid
        Public Property Vector As Single()
        Public Property Payload As Dictionary(Of String, Object)
    End Class

    ''' <summary>
    ''' Upsert 操作结果
    ''' </summary>
    Public Class QdrantUpsertResult
        Public Property Success As Boolean
        Public Property OperationId As String
        Public Property Status As String
        Public Property ErrorMessage As String
        Public Property LatencyMs As Long
    End Class

End Namespace
