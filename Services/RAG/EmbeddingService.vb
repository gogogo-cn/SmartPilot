Imports System
Imports System.Collections.Generic
Imports System.Net.Http
Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services.RAG

    ''' <summary>
    ''' 向量化服务 - 混合模式支持本地 BGE (Ollama) 和云端 API
    ''' </summary>
    Public Class EmbeddingService
        Implements IDisposable

        Private _config As RagConfig
        Private _localHttpClient As HttpClient      ' Ollama 客户端
        Private _cloudHttpClient As HttpClient        ' 云端 API 客户端
        Private _localEndpoint As String              ' 规范化后的 Ollama 地址（localhost → 127.0.0.1）
        Private _isDisposed As Boolean = False

        Public Sub New(config As RagConfig)
            _config = config

            ' Ollama 是本地服务：必须禁用系统代理。
            ' 实测（Win10 + 系统代理）：默认 UseProxy=True 时 localhost:11434 请求 2 秒级才返回（直连毫秒级）。
            Dim localHandler As New HttpClientHandler()
            localHandler.UseProxy = False
            _localHttpClient = New HttpClient(localHandler)
            _localHttpClient.Timeout = TimeSpan.FromSeconds(60)  ' 本地模型可能较慢

            ' 云端 API 保留系统代理（外网访问在代理环境下需要）
            _cloudHttpClient = New HttpClient()
            _cloudHttpClient.Timeout = TimeSpan.FromSeconds(30)

            ' localhost → 127.0.0.1：.NET 对 "localhost" 优先尝试 IPv6 ::1，
            ' 本机未监听 v6 时每次请求固定多花约 2 秒才 fallback IPv4（实测 2083ms vs 25ms）
            _localEndpoint = NormalizeLocalEndpoint(_config.LocalBgeEndpoint)
        End Sub

        ''' <summary>
        ''' 规范化本地服务地址：localhost → 127.0.0.1（避免 .NET IPv6 优先探测的 2 秒级延迟）
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
        ''' TCP 层端口探测（快速失败）：端口未监听时 500ms 内判定，
        ''' 避免批量/逐条请求各等一次秒级网络超时。
        ''' </summary>
        Private Async Function LocalTcpProbeAsync(timeoutMs As Integer) As Task(Of Boolean)
            Try
                Dim uri As New Uri(_localEndpoint)
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
                System.Diagnostics.Debug.WriteLine($"[Embedding] Ollama TCP 探测异常: {ex.Message}")
                Return False
            End Try
        End Function

#Region "公开方法"

        ''' <summary>
        ''' 将文本列表转换为向量（批量）
        ''' 自动根据配置选择本地或云端 provider
        ''' </summary>
        ''' <param name="texts">待向量化的文本列表</param>
        ''' <param name="forceProvider">强制指定 provider（可选）</param>
        ''' <returns>向量列表，每个向量为 Single() 数组</returns>
        Public Async Function EmbedBatchAsync(texts As List(Of String),
                                               Optional forceProvider As String = "") As Task(Of List(Of Single()))
            If texts Is Nothing OrElse texts.Count = 0 Then
                Return New List(Of Single())()
            End If

            Dim provider = If(String.IsNullOrEmpty(forceProvider), _config.DefaultEmbeddingProvider, forceProvider)

            Try
                Select Case provider.ToLower()
                    Case "local_bge", "ollama"
                        Return Await EmbedViaLocalBgeAsync(texts)
                    Case "openai", "azure", "qianwen", "zhipu"
                        Return Await EmbedViaCloudApiAsync(texts)
                    Case Else
                        ' 默认尝试本地，失败则降级到云端
                        Dim result = Await EmbedViaLocalBgeAsync(texts)
                        If result.Count > 0 Then Return result
                        Return Await EmbedViaCloudApiAsync(texts)
                End Select
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[Embedding] EmbedBatchAsync 异常 ({provider}): {ex.Message}")
                Throw
            End Try
        End Function

        ''' <summary>
        ''' 单条文本向量化
        ''' </summary>
        ''' <param name="text">待向量化的文本</param>
        ''' <returns>向量数组</returns>
        Public Async Function EmbedAsync(text As String) As Task(Of Single())
            If String.IsNullOrWhiteSpace(text) Then
                Return New Single() {}
            End If

            Dim results = Await EmbedBatchAsync(New List(Of String) From {text})
            If results.Count > 0 Then
                Return results(0)
            Else
                Return New Single() {}
            End If
        End Function

        ''' <summary>
        ''' 测试连通性
        ''' </summary>
        ''' <param name="provider">要测试的 provider</param>
        ''' <returns>测试结果</returns>
        Public Async Function TestConnectionAsync(provider As String) As Task(Of ConnectionTestResult)
            Dim result As New ConnectionTestResult With {
                .Provider = provider,
                .Success = False,
                .Message = ""
            }

            Try
                Dim sw = Diagnostics.Stopwatch.StartNew()

                Select Case provider.ToLower()
                    Case "local_bge", "ollama"
                        ' 测试 Ollama 连接
                        Dim testUrl = $"{_localEndpoint}/api/tags"
                        Dim testHandler As New HttpClientHandler()
                        testHandler.UseProxy = False   ' 本地服务禁代理，否则延迟虚高
                        Using testClient As New HttpClient(testHandler)
                            testClient.Timeout = TimeSpan.FromSeconds(5)
                            Dim response = Await testClient.GetAsync(testUrl)

                            sw.Stop()
                            result.LatencyMs = sw.ElapsedMilliseconds

                            If response.IsSuccessStatusCode Then
                                result.Success = True
                                result.Message = $"Ollama 服务可用 (模型: {_config.LocalBgeModel})"
                            Else
                                result.Message = $"Ollama 返回错误: {response.StatusCode}"
                            End If
                        End Using

                    Case "cloud"
                        ' 测试云端 API（简单验证 API Key 不为空）
                        sw.Stop()
                        result.LatencyMs = sw.ElapsedMilliseconds

                        If Not String.IsNullOrEmpty(_config.CloudApiKey) AndAlso
                           Not String.IsNullOrEmpty(_config.CloudEndpoint) Then
                            result.Success = True
                            result.Message = $"云端 API 配置已设置 (Provider: {_config.CloudModel})"
                        Else
                            result.Message = "云端 API 配置不完整（缺少 Endpoint 或 ApiKey）"
                        End If

                    Case Else
                        result.Message = $"未知的 Provider: {provider}"
                End Select

            Catch ex As Exception
                result.Message = $"连接测试异常: {ex.Message}"
            End Try

            Return result
        End Function

#End Region

#Region "本地 BGE (via Ollama)"

        ''' <summary>
        ''' 调用 Ollama Embeddings API
        ''' 优先使用新版批量端点 POST /api/embed（input 数组，Ollama >= 0.3.x）
        ''' 批量失败时回退旧版单条端点 POST /api/embeddings（prompt 字段）
        ''' </summary>
        Private Async Function EmbedViaLocalBgeAsync(texts As List(Of String)) As Task(Of List(Of Single()))
            Dim results As New List(Of Single())()
            If texts Is Nothing OrElse texts.Count = 0 Then Return results

            ' ---- 首选：新版批量端点 POST /api/embed，body {"model":..., "input":[...]} ----
            ' 先 TCP 探测（500ms）：Ollama 未运行时直接返回空，
            ' 避免批量失败后再逐条回退、每条各等一次秒级网络超时
            If Not Await LocalTcpProbeAsync(500) Then
                System.Diagnostics.Debug.WriteLine($"[Embedding] Ollama 不可达（TCP 探测失败）: {_localEndpoint}")
                Return results
            End If
            Try
                Dim batchBody = New With {
                    .model = _config.LocalBgeModel,
                    .input = texts
                }
                Dim batchJson = JsonConvert.SerializeObject(batchBody)
                Dim batchContent = New StringContent(batchJson, Encoding.UTF8, "application/json")

                Dim batchResponse = Await _localHttpClient.PostAsync($"{_localEndpoint}/api/embed", batchContent)

                If batchResponse.IsSuccessStatusCode Then
                    Dim batchStr = Await batchResponse.Content.ReadAsStringAsync()
                    Dim batchObj = JObject.Parse(batchStr)

                    ' 新版响应：{"embeddings": [[...], [...]]}
                    Dim embeddings = batchObj("embeddings")
                    If embeddings IsNot Nothing AndAlso embeddings.HasValues Then
                        For Each emb In embeddings
                            Dim vector As New List(Of Single)()
                            For Each v In emb
                                vector.Add(Convert.ToSingle(v))
                            Next
                            results.Add(vector.ToArray())
                        Next

                        If results.Count = texts.Count Then
                            System.Diagnostics.Debug.WriteLine($"[Embedding] Ollama 批量向量化完成(/api/embed): {texts.Count} 条文本 → {results.Count} 个向量")
                            Return results
                        End If
                        ' 数量不一致（异常响应）：清空走逐条回退，保证与文本一一对应
                        results.Clear()
                    End If
                Else
                    Dim batchErr = Await batchResponse.Content.ReadAsStringAsync()
                    System.Diagnostics.Debug.WriteLine($"[Embedding] Ollama /api/embed 失败: {batchResponse.StatusCode} - {batchErr}，回退逐条模式")
                    ' 模型缺失是配置问题（不是网络抖动）：直接抛带指引的异常，
                    ' 让文档处理失败信息明确告诉用户「ollama pull <model>」，而非晦涩的 501/404
                    If batchErr.IndexOf("not found", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                       batchErr.IndexOf("does not support", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        Throw New InvalidOperationException(
                            $"Ollama 未安装向量模型 '{_config.LocalBgeModel}'（或该实例未启用 embeddings 能力）。" &
                            $"请先在终端执行: ollama pull {_config.LocalBgeModel}，然后重试导入。")
                    End If
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[Embedding] Ollama /api/embed 异常: {ex.Message}，回退逐条模式")
            End Try

            ' ---- 回退：旧版单条端点 POST /api/embeddings，body {"model":..., "prompt":"..."} ----
            ' 单条失败以空向量占位，保证结果与输入数量一致（调用方按序对应）
            For Each txt In texts
                Dim vector As New List(Of Single)()
                Try
                    Dim singleBody = New With {
                        .model = _config.LocalBgeModel,
                        .prompt = txt
                    }
                    Dim singleJson = JsonConvert.SerializeObject(singleBody)
                    Dim singleContent = New StringContent(singleJson, Encoding.UTF8, "application/json")

                    Dim singleResponse = Await _localHttpClient.PostAsync($"{_localEndpoint}/api/embeddings", singleContent)

                    If singleResponse.IsSuccessStatusCode Then
                        Dim singleStr = Await singleResponse.Content.ReadAsStringAsync()
                        Dim singleObj = JObject.Parse(singleStr)

                        ' 旧版响应：{"embedding": [...]}
                        Dim emb = singleObj("embedding")
                        If emb IsNot Nothing AndAlso emb.HasValues Then
                            For Each v In emb
                                vector.Add(Convert.ToSingle(v))
                            Next
                        End If
                    Else
                        Dim singleErr = Await singleResponse.Content.ReadAsStringAsync()
                        System.Diagnostics.Debug.WriteLine($"[Embedding] Ollama 单条向量化失败: {singleResponse.StatusCode} - {singleErr}")
                    End If
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"[Embedding] Ollama 单条向量化异常: {ex.Message}")
                End Try

                results.Add(vector.ToArray())
            Next

            System.Diagnostics.Debug.WriteLine($"[Embedding] Ollama 逐条向量化完成: {texts.Count} 条文本 → {results.Count} 个向量")
            Return results
        End Function

#End Region

#Region "云端 API"

        ''' <summary>
        ''' 调用 OpenAI 兼容的 Embeddings API
        ''' 支持 OpenAI / Azure OpenAI / 通义千问 / 智谱 AI 等
        ''' POST {endpoint}/v1/embeddings
        ''' Body: {"input": ["text"], "model": "xxx"}
        ''' </summary>
        Private Async Function EmbedViaCloudApiAsync(texts As List(Of String)) As Task(Of List(Of Single()))
            Dim results As New List(Of Single())()

            Try
                ' 构建请求 URL（根据 provider 类型调整）
                Dim apiUrl As String = BuildCloudApiUrl()

                ' 构建请求体
                Dim requestBody As New Dictionary(Of String, Object) From {
                    {"input", texts},
                    {"model", _config.CloudModel}
                }

                Dim jsonContent = JsonConvert.SerializeObject(requestBody)
                Dim content = New StringContent(jsonContent, Encoding.UTF8, "application/json")

                ' 设置认证头
                SetCloudAuthHeader()

                Dim response = Await _cloudHttpClient.PostAsync(apiUrl, content)

                If Not response.IsSuccessStatusCode Then
                    Dim errorMsg = Await response.Content.ReadAsStringAsync()
                    System.Diagnostics.Debug.WriteLine($"[Embedding] Cloud API 请求失败: {response.StatusCode} - {errorMsg}")
                    Return results
                End If

                Dim jsonStr = Await response.Content.ReadAsStringAsync()
                Dim jsonObj = JObject.Parse(jsonStr)

                ' 解析 OpenAI 格式的响应
                Dim dataArray = jsonObj("data")
                If dataArray IsNot Nothing Then
                    For Each item In dataArray
                        Dim embeddingArray = item("embedding")
                        If embeddingArray IsNot Nothing Then
                            Dim vector As New List(Of Single)()
                            For Each v In embeddingArray
                                vector.Add(Convert.ToSingle(v))
                            Next
                            results.Add(vector.ToArray())
                        End If
                    Next
                End If

                System.Diagnostics.Debug.WriteLine($"[Embedding] Cloud API 向量化完成: {texts.Count} 条文本 → {results.Count} 个向量")

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[Embedding] EmbedViaCloudApiAsync 异常: {ex.Message}")
            End Try

            Return results
        End Function

        ''' <summary>
        ''' 构建 Cloud API URL
        ''' </summary>
        Private Function BuildCloudApiUrl() As String
            Dim endpoint = _config.CloudEndpoint.TrimEnd("/"c)

            ' 根据不同的 provider 调整路径
            Dim providerLower = If(_config.CloudModel?.ToLower(), "")

            If providerLower.Contains("azure") Then
                ' Azure OpenAI: {endpoint}/openai/deployments/{model}/embeddings?api-version=2023-05-15
                Return $"{endpoint}/openai/deployments/{_config.CloudModel}/embeddings?api-version=2023-05-15"
            ElseIf providerLower.Contains("qianwen") OrElse providerLower.Contains("text-embedding-v") Then
                ' 通义千问: 兼容 OpenAI 格式
                Return $"{endpoint}/v1/embeddings"
            ElseIf providerLower.Contains("zhipu") OrElse providerLower.Contains("embedding-") Then
                ' 智谱 AI: 兼容 OpenAI 格式
                Return $"{endpoint}/v1/embeddings"
            Else
                ' 默认 OpenAI 格式
                Return $"{endpoint}/v1/embeddings"
            End If
        End Function

        ''' <summary>
        ''' 设置云端 API 认证头
        ''' </summary>
        Private Sub SetCloudAuthHeader()
            ' 清除旧的认证头
            If _cloudHttpClient.DefaultRequestHeaders.Contains("Authorization") Then
                _cloudHttpClient.DefaultRequestHeaders.Remove("Authorization")
            End If
            If _cloudHttpClient.DefaultRequestHeaders.Contains("api-key") Then
                _cloudHttpClient.DefaultRequestHeaders.Remove("api-key")
            End If

            ' 根据不同的 provider 设置认证方式
            Dim providerLower = If(_config.CloudModel?.ToLower(), "")

            If providerLower.Contains("azure") Then
                ' Azure 使用 api-key 头
                _cloudHttpClient.DefaultRequestHeaders.Add("api-key", _config.CloudApiKey)
            Else
                ' 其他使用 Bearer Token
                _cloudHttpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {_config.CloudApiKey}")
            End If
        End Sub

#End Region

#Region "IDisposable Support"

        Protected Overloads Sub Dispose(disposing As Boolean)
            If Not _isDisposed Then
                If disposing Then
                    If _localHttpClient IsNot Nothing Then
                        _localHttpClient.Dispose()
                        _localHttpClient = Nothing
                    End If
                    If _cloudHttpClient IsNot Nothing Then
                        _cloudHttpClient.Dispose()
                        _cloudHttpClient = Nothing
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

End Namespace
