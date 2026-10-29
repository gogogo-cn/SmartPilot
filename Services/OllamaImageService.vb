Imports System.IO
Imports System.Net.Http
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services

    ''' <summary>
    ''' Ollama 图片识别服务
    ''' 支持本地运行的 Ollama 视觉模型（如 qwen3-vl、llava 等）
    ''' 使用队列确保顺序执行，避免并发冲突
    ''' </summary>
    Public Class OllamaImageService
        Implements IDisposable

        Private ReadOnly _apiUrl As String
        Private ReadOnly _tagsUrl As String
        Private ReadOnly _llamacppModelsUrl As String
        Private _isDisposed As Boolean = False

        ' 静态队列，确保所有 Ollama 请求顺序执行
        Private Shared _requestQueue As New Queue(Of Tuple(Of String, String, String, String, Double, TaskCompletionSource(Of String)))
        Private Shared _isProcessingQueue As Boolean = False
        Private Shared _queueLock As New Object()

        Public Sub New()
            Dim config = Utils.ConfigurationManager.Instance
            _apiUrl = config.GetSetting(Utils.ConfigConstants.Ollama.API_URL, Utils.ConfigConstants.Ollama.DEFAULT_API_URL)
            _tagsUrl = config.GetSetting(Utils.ConfigConstants.Ollama.TAGS_URL, Utils.ConfigConstants.Ollama.DEFAULT_TAGS_URL)
            _llamacppModelsUrl = config.GetSetting(Utils.ConfigConstants.LlamaCpp.MODELS_URL, Utils.ConfigConstants.LlamaCpp.DEFAULT_MODELS_URL)
        End Sub

        ''' <summary>
        ''' 释放资源
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            DisposeCore(True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected Overridable Sub DisposeCore(disposing As Boolean)
            If _isDisposed Then Return
            _isDisposed = True

            If disposing Then
                ' 清理静态队列(可选)
                SyncLock _queueLock
                    _requestQueue.Clear()
                    _isProcessingQueue = False
                End SyncLock
            End If
        End Sub

        ''' <summary>
        ''' 验证 API URL 是否正确配置
        ''' 防止因配置错误导致 HTTP 405 Method Not Allowed 错误
        ''' </summary>
        ''' <returns>错误信息，如果验证通过则返回 Nothing</returns>
        Private Function ValidateApiUrl() As String
            ' 检查 API URL 是否为空
            If String.IsNullOrEmpty(_apiUrl) Then
                Return "【错误：Ollama API URL 未配置】" & vbCrLf &
                       "请检查配置文件，确保 OllamaApiUrl 已正确设置。" & vbCrLf &
                       "默认值应为: http://localhost:11434/api/chat"
            End If
            
            ' 检查是否错误地使用了 /api/tags 端点（该端点只支持 GET 请求）
            If _apiUrl.ToLower().Contains("/api/tags") Then
                Return "【错误：API URL 配置错误】" & vbCrLf &
                       "当前配置的 API URL 指向了错误的端点: " & _apiUrl & vbCrLf &
                       "图片识别需要使用 /api/chat 或 /api/generate 端点，而不是 /api/tags。" & vbCrLf &
                       "请将 OllamaApiUrl 配置为: http://localhost:11434/api/chat"
            End If
            
            ' 检查是否使用了正确的协议
            If Not _apiUrl.StartsWith("http://") AndAlso Not _apiUrl.StartsWith("https://") Then
                Return "【错误：API URL 协议错误】" & vbCrLf &
                       "API URL 必须以 http:// 或 https:// 开头。" & vbCrLf &
                       "当前值: " & _apiUrl
            End If
            
            Return Nothing
        End Function

        ''' <summary>
        ''' 使用 Ollama 视觉模型识别图片内容（支持 system + user 双提示词）- 使用队列确保顺序执行
        ''' </summary>
        ''' <param name="imagePath">图片文件路径</param>
        ''' <param name="modelName">Ollama 模型名称（如 qwen3-vl:4b）</param>
        ''' <param name="systemPrompt">系统提示词</param>
        ''' <param name="userPrompt">用户提示词</param>
        ''' <param name="temperature">温度参数（0-1，越低越确定）</param>
        Public Async Function RecognizeImageWithSystemPrompt(imagePath As String, modelName As String, systemPrompt As String, userPrompt As String, Optional temperature As Double = 0.05) As Task(Of String)
            System.Diagnostics.Debug.WriteLine($"[OllamaImageService] 收到识别请求，图片: {imagePath}")
            
            ' 验证 API URL 是否正确配置
            Dim validationError = ValidateApiUrl()
            If validationError IsNot Nothing Then
                Return validationError
            End If

            ' 创建任务完成源，用于等待队列处理结果
            Dim tcs As New TaskCompletionSource(Of String)

            ' 将请求加入队列
            SyncLock _queueLock
                _requestQueue.Enqueue(Tuple.Create(imagePath, modelName, systemPrompt, userPrompt, temperature, tcs))
                System.Diagnostics.Debug.WriteLine($"[OllamaImageService] 请求已加入队列，当前队列长度: {_requestQueue.Count}")

                ' 如果队列没有在处理，启动处理
                If Not _isProcessingQueue Then
                    _isProcessingQueue = True
                    ' 启动后台任务处理队列（fire-and-forget）
#Disable Warning BC42358  ' Fire-and-forget调用，故意不等待
                    Task.Run(AddressOf ProcessQueueAsync)
#Enable Warning BC42358
                End If
            End SyncLock

            ' 等待队列处理完成
            Return Await tcs.Task
        End Function

        ''' <summary>
        ''' 处理 Ollama 请求队列 - 确保严格的顺序执行
        ''' </summary>
        Private Async Function ProcessQueueAsync() As Task
            System.Diagnostics.Debug.WriteLine("[OllamaImageService] 队列处理启动")

            While True
                Dim currentItem As Tuple(Of String, String, String, String, Double, TaskCompletionSource(Of String)) = Nothing

                ' 从队列取出一个请求
                SyncLock _queueLock
                    If _requestQueue.Count = 0 Then
                        ' 队列为空，停止处理
                        _isProcessingQueue = False
                        System.Diagnostics.Debug.WriteLine("[OllamaImageService] 队列为空，处理停止")
                        Return
                    End If
                    currentItem = _requestQueue.Dequeue()
                    System.Diagnostics.Debug.WriteLine($"[OllamaImageService] 取出请求，剩余队列长度: {_requestQueue.Count}")
                End SyncLock

                ' 执行识别
                Try
                    Dim result = Await ExecuteRecognizeInternal(currentItem.Item1, currentItem.Item2, currentItem.Item3, currentItem.Item4, currentItem.Item5)
                    currentItem.Item6.SetResult(result)
                Catch ex As Exception
                    currentItem.Item6.SetException(ex)
                End Try

                ' 等待一段时间，确保 Ollama 完全释放资源
                System.Diagnostics.Debug.WriteLine("[OllamaImageService] 等待5秒确保Ollama资源释放...")
                Await Task.Delay(5000)
            End While
        End Function

        ''' <summary>
        ''' 实际执行 Ollama 识别的内部方法
        ''' </summary>
        Private Async Function ExecuteRecognizeInternal(imagePath As String, modelName As String, systemPrompt As String, userPrompt As String, temperature As Double) As Task(Of String)
            Try
                System.Diagnostics.Debug.WriteLine("[OllamaImageService] 开始识别图片（双提示词模式）")
                System.Diagnostics.Debug.WriteLine("图片路径: " & imagePath)
                System.Diagnostics.Debug.WriteLine("使用模型: " & modelName)
                System.Diagnostics.Debug.WriteLine("Temperature: " & temperature)

                If Not File.Exists(imagePath) Then
                    Return "【错误：图片文件不存在】"
                End If

                Dim imageBytes = File.ReadAllBytes(imagePath)
                Dim base64Image = Convert.ToBase64String(imageBytes)

                ' 判断使用哪种 API 格式：原生格式 (/api/generate) 或 OpenAI 兼容格式 (/api/chat)
                Dim isChatEndpoint As Boolean = _apiUrl.EndsWith("/chat") OrElse _apiUrl.Contains("/api/chat")
                
                Dim json As String = ""
                
                If isChatEndpoint Then
                    ' Ollama /api/chat 格式（支持视觉模型）
                    System.Diagnostics.Debug.WriteLine("使用 Ollama /api/chat 格式")
                    ' Ollama 的视觉模型使用 images 数组配合 content 字符串
                    Dim requestBody = New With {
                        .model = modelName,
                        .messages = New Object() {
                            New With {
                                .role = "system",
                                .content = systemPrompt
                            },
                            New With {
                                .role = "user",
                                .content = userPrompt,
                                .images = New String() {base64Image}
                            }
                        },
                        .stream = False,
                        .temperature = temperature
                    }
                    json = JsonConvert.SerializeObject(requestBody)
                Else
                    ' Ollama 原生 API 格式（/api/generate）
                    System.Diagnostics.Debug.WriteLine("使用 Ollama 原生格式 (/api/generate)")
                    Dim requestBody = New With {
                        .model = modelName,
                        .prompt = systemPrompt & vbCrLf & vbCrLf & userPrompt,
                        .images = New String() {base64Image},
                        .stream = False,
                        .options = New With {
                            .temperature = temperature
                        }
                    }
                    json = JsonConvert.SerializeObject(requestBody)
                End If
                
                System.Diagnostics.Debug.WriteLine("Ollama 请求长度: " & json.Length)
                ' 输出完整请求体以便调试（可能包含base64图片）
                System.Diagnostics.Debug.WriteLine("===== FULL OLLAMA REQUEST START =====")
                System.Diagnostics.Debug.WriteLine(json)
                System.Diagnostics.Debug.WriteLine("===== FULL OLLAMA REQUEST END =====")

                ' 写入持久化日志
                Try
                    Utils.ChatLogger.Instance.SetEnabled(True)
                    Utils.ChatLogger.Instance.SetLogLevel(Utils.LogLevel.Debug)
                    Utils.ChatLogger.Instance.LogDebug("Ollama", "FULL_REQUEST", json)
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("写入 ChatLogger 失败: " & ex.Message)
                End Try

                Using client As New HttpClient()
                    client.Timeout = TimeSpan.FromSeconds(60)
                    Dim content = New StringContent(json, Encoding.UTF8, "application/json")
                    System.Diagnostics.Debug.WriteLine("发送请求到 Ollama API...")

                    Dim response = Await client.PostAsync(_apiUrl, content)
                    Dim responseBody = Await response.Content.ReadAsStringAsync()

                    System.Diagnostics.Debug.WriteLine("Ollama 响应状态: " & response.StatusCode)
                    System.Diagnostics.Debug.WriteLine("Ollama 响应长度: " & responseBody.Length)

                    If response.IsSuccessStatusCode Then
                        Try
                            Dim result = JObject.Parse(responseBody)
                            
                            ' 判断返回格式：原生格式有 "response" 字段，OpenAI 兼容格式有 "message" 字段
                            If result("response") IsNot Nothing Then
                                ' Ollama 原生 API 返回格式：{ "response": "...", "done": true }
                                Dim text = result("response").ToString()
                                System.Diagnostics.Debug.WriteLine("Ollama 识别结果(原生格式): " & text.Substring(0, Math.Min(text.Length, 200)))
                                Return text
                            ElseIf result("message") IsNot Nothing AndAlso result("message")("content") IsNot Nothing Then
                                ' OpenAI 兼容格式：{ "message": { "role": "...", "content": "..." } }
                                Dim text = result("message")("content").ToString()
                                System.Diagnostics.Debug.WriteLine("Ollama 识别结果(OpenAI格式): " & text.Substring(0, Math.Min(text.Length, 200)))
                                Return text
                            Else
                                ' 未知格式，返回整个响应
                                System.Diagnostics.Debug.WriteLine("Ollama 返回未知格式，返回原始响应")
                                Return responseBody
                            End If
                        Catch ex As Exception
                            System.Diagnostics.Debug.WriteLine("解析 Ollama 响应失败: " & ex.Message)
                            Return responseBody
                        End Try
                    Else
                        System.Diagnostics.Debug.WriteLine("Ollama API 错误: " & responseBody)
                        ' 尝试解析错误信息
                        Try
                            Dim errorJson = JObject.Parse(responseBody)
                            If errorJson("error") IsNot Nothing Then
                                Return "【错误：Ollama API 调用失败 - " & response.StatusCode & " - " & errorJson("error").ToString() & "】"
                            End If
                        Catch ex As Exception  ' 
                        End Try
                        Return "【错误：Ollama API 调用失败 - " & response.StatusCode & " - " & responseBody.Substring(0, Math.Min(200, responseBody.Length)) & "】"
                    End If
                End Using

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Ollama 识别失败: " & ex.Message)
                System.Diagnostics.Debug.WriteLine("异常类型: " & ex.GetType().Name)
                System.Diagnostics.Debug.WriteLine("堆栈跟踪: " & ex.StackTrace)
                
                ' 提供更详细的错误信息
                Dim detailedError = $"【错误：Ollama 识别失败】" & vbCrLf &
                                   $"异常类型: {ex.GetType().Name}" & vbCrLf &
                                   $"错误信息: {ex.Message}" & vbCrLf &
                                   $"API地址: {_apiUrl}" & vbCrLf &
                                   $"模型: {modelName}" & vbCrLf &
                                   $"图片路径: {imagePath}" & vbCrLf &
                                   $"请检查:" & vbCrLf &
                                   $"1. Ollama服务是否启动（http://localhost:11434）" & vbCrLf &
                                   $"2. 模型'{modelName}'是否已下载" & vbCrLf &
                                   $"3. 网络连接是否正常"
                Return detailedError
            End Try
        End Function

        ''' <summary>
        ''' 使用 Ollama 视觉模型识别图片内容
        ''' </summary>
        ''' <param name="imagePath">图片文件路径</param>
        ''' <param name="modelName">Ollama 模型名称（如 qwen3-vl:4b）</param>
        ''' <param name="prompt">自定义提示词</param>
        Public Async Function RecognizeImage(imagePath As String, modelName As String, Optional prompt As String = "") As Task(Of String)
            Try
                System.Diagnostics.Debug.WriteLine("OllamaImageService: 开始识别图片")
                System.Diagnostics.Debug.WriteLine("图片路径: " & imagePath)
                System.Diagnostics.Debug.WriteLine("使用模型: " & modelName)

                ' 验证 API URL 是否正确配置
                Dim validationError = ValidateApiUrl()
                If validationError IsNot Nothing Then
                    Return validationError
                End If

                If Not File.Exists(imagePath) Then
                    Return "【错误：图片文件不存在】"
                End If

                Dim imageBytes = File.ReadAllBytes(imagePath)
                Dim base64Image = Convert.ToBase64String(imageBytes)

                Dim defaultPrompt = If(String.IsNullOrEmpty(prompt),
                    "请识别这张图片中的所有文字内容。必须保留原文的换行格式，每行文字单独一行，不要合并成一行。以纯文本格式返回，保持原有的段落和行结构。",
                    prompt)

                ' 判断使用哪种 API 格式：原生格式 (/api/generate) 或 OpenAI 兼容格式 (/api/chat)
                Dim isChatEndpoint As Boolean = _apiUrl.EndsWith("/chat") OrElse _apiUrl.Contains("/api/chat")
                
                Dim json As String = ""
                
                If isChatEndpoint Then
                    ' Ollama /api/chat 格式（支持视觉模型）
                    System.Diagnostics.Debug.WriteLine("使用 Ollama /api/chat 格式")
                    ' Ollama 的视觉模型使用 images 数组配合 content 字符串
                    Dim requestBody = New With {
                        .model = modelName,
                        .messages = New Object() {
                            New With {
                                .role = "user",
                                .content = defaultPrompt,
                                .images = New String() {base64Image}
                            }
                        },
                        .stream = False
                    }
                    json = JsonConvert.SerializeObject(requestBody)
                Else
                    ' Ollama 原生 API 格式（/api/generate）
                    System.Diagnostics.Debug.WriteLine("使用 Ollama 原生格式 (/api/generate)")
                    Dim requestBody = New With {
                        .model = modelName,
                        .prompt = defaultPrompt,
                        .images = New String() {base64Image},
                        .stream = False
                    }
                    json = JsonConvert.SerializeObject(requestBody)
                End If
                
                System.Diagnostics.Debug.WriteLine("Ollama 请求长度: " & json.Length)

                Using client As New HttpClient()
                    client.Timeout = TimeSpan.FromSeconds(60)
                    Dim content = New StringContent(json, Encoding.UTF8, "application/json")
                    System.Diagnostics.Debug.WriteLine("发送请求到 Ollama API...")

                    Dim response = Await client.PostAsync(_apiUrl, content)
                    Dim responseBody = Await response.Content.ReadAsStringAsync()

                    System.Diagnostics.Debug.WriteLine("Ollama 响应状态: " & response.StatusCode)
                    System.Diagnostics.Debug.WriteLine("Ollama 响应长度: " & responseBody.Length)

                    If response.IsSuccessStatusCode Then
                        Try
                            Dim result = JObject.Parse(responseBody)
                            
                            ' 判断返回格式：原生格式有 "response" 字段，OpenAI 兼容格式有 "message" 字段
                            If result("response") IsNot Nothing Then
                                ' Ollama 原生 API 返回格式：{ "response": "...", "done": true }
                                Dim text = result("response").ToString()
                                System.Diagnostics.Debug.WriteLine("Ollama 识别结果(原生格式): " & text.Substring(0, Math.Min(text.Length, 200)))
                                Return text
                            ElseIf result("message") IsNot Nothing AndAlso result("message")("content") IsNot Nothing Then
                                ' OpenAI 兼容格式：{ "message": { "role": "...", "content": "..." } }
                                Dim text = result("message")("content").ToString()
                                System.Diagnostics.Debug.WriteLine("Ollama 识别结果(OpenAI格式): " & text.Substring(0, Math.Min(text.Length, 200)))
                                Return text
                            Else
                                ' 未知格式，返回整个响应
                                System.Diagnostics.Debug.WriteLine("Ollama 返回未知格式，返回原始响应")
                                Return responseBody
                            End If
                        Catch ex As Exception
                            System.Diagnostics.Debug.WriteLine("解析 Ollama 响应失败: " & ex.Message)
                            Return responseBody
                        End Try
                    Else
                        System.Diagnostics.Debug.WriteLine("Ollama API 错误: " & responseBody)
                        ' 尝试解析错误信息
                        Try
                            Dim errorJson = JObject.Parse(responseBody)
                            If errorJson("error") IsNot Nothing Then
                                Return "【错误：Ollama API 调用失败 - " & response.StatusCode & " - " & errorJson("error").ToString() & "】"
                            End If
                        Catch ex As Exception  ' 
                        End Try
                        Return "【错误：Ollama API 调用失败 - " & response.StatusCode & " - " & responseBody.Substring(0, Math.Min(200, responseBody.Length)) & "】"
                    End If
                End Using

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Ollama 识别失败: " & ex.Message)
                Return "【错误：Ollama 识别失败 - " & ex.Message & "】"
            End Try
        End Function

        ''' <summary>
        ''' 检查 Ollama 服务是否可用
        ''' </summary>
        Public Async Function IsOllamaAvailable() As Task(Of Boolean)
            Try
                Using client As New HttpClient()
                    client.Timeout = TimeSpan.FromSeconds(5)
                    Dim response = Await client.GetAsync(_tagsUrl)
                    Return response.IsSuccessStatusCode
                End Using
            Catch ex As Exception  ' 
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 本地模型信息：名称 + 来源（ollama / llamacpp）
        ''' 前端据此区分：Ollama 模型走 /api/chat，llama.cpp 模型走 OpenAI 兼容 /v1/chat/completions
        ''' </summary>
        Public Class LocalModelInfo
            Public Property Name As String
            Public Property Source As String

            Public Sub New(name As String, source As String)
                Me.Name = name
                Me.Source = source
            End Sub
        End Class

        ''' <summary>
        ''' 获取所有可用本地模型（Ollama + llama.cpp 合并去重）
        ''' Ollama 探测 /api/tags，llama.cpp 探测开放 OpenAI 兼容的 /v1/models
        ''' </summary>
        Public Async Function GetAvailableModels() As Task(Of List(Of LocalModelInfo))
            Dim models As New List(Of LocalModelInfo)()
            Dim names As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            ' 来源1：Ollama /api/tags（文档结构：{ "models": [ { "name": ... } ] }）
            Try
                Using client As New HttpClient()
                    client.Timeout = TimeSpan.FromSeconds(10)
                    Dim response = Await client.GetAsync(_tagsUrl)
                    If response.IsSuccessStatusCode Then
                        Dim responseBody = Await response.Content.ReadAsStringAsync()
                        Dim result = JObject.Parse(responseBody)
                        For Each model In result("models")
                            Dim mName = model("name").ToString()
                            If names.Add(mName) Then
                                models.Add(New LocalModelInfo(mName, "ollama"))
                            End If
                        Next
                    End If
                End Using
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("获取 Ollama 模型列表失败: " & ex.Message)
            End Try

            ' 来源2：llama.cpp /v1/models（OpenAI 兼容，文档结构：{ "data": [ { "id": ... } ] }）
            ' 与 Ollama 端口不同、协议不同，需单独探测；当本机未安装/未启动时静默忽略
            Try
                Using client As New HttpClient()
                    client.Timeout = TimeSpan.FromSeconds(5)
                    Dim response = Await client.GetAsync(_llamacppModelsUrl)
                    If response.IsSuccessStatusCode Then
                        Dim responseBody = Await response.Content.ReadAsStringAsync()
                        Dim result = JObject.Parse(responseBody)
                        Dim data = TryCast(result("data"), JArray)
                        If data IsNot Nothing Then
                            For Each model In data
                                Dim mName = model("id").ToString()
                                If Not String.IsNullOrWhiteSpace(mName) AndAlso names.Add(mName) Then
                                    models.Add(New LocalModelInfo(mName, "llamacpp"))
                                End If
                            Next
                        End If
                    End If
                End Using
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("获取 llama.cpp 模型列表失败: " & ex.Message)
            End Try

            Return models
        End Function

    End Class

End Namespace