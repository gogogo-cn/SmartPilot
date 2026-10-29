Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ExcelAIVSTOAddIn.Utils

Namespace Services

    ''' <summary>
    ''' 多模态内容类型
    ''' </summary>
    Public Enum ContentType
        Text
        Image
        File
        Table
        Mixed
    End Enum

    ''' <summary>
    ''' 多模态内容项
    ''' </summary>
    Public Class MultimodalContent
        Public Property Type As ContentType
        Public Property Text As String
        Public Property ImageBase64 As String
        Public Property ImageUrl As String
        Public Property FileName As String
        Public Property FileContent As String
        Public Property FileBytes As Byte()
        Public Property MimeType As String
        Public Property Metadata As Dictionary(Of String, Object)

        Public Sub New()
            Metadata = New Dictionary(Of String, Object)()
        End Sub

        ''' <summary>
        ''' 创建文本内容
        ''' </summary>
        Public Shared Function CreateText(text As String) As MultimodalContent
            Return New MultimodalContent With {
                .Type = ContentType.Text,
                .Text = text
            }
        End Function

        ''' <summary>
        ''' 创建图片内容
        ''' </summary>
        Public Shared Function CreateImage(base64Data As String, mimeType As String) As MultimodalContent
            Return New MultimodalContent With {
                .Type = ContentType.Image,
                .ImageBase64 = base64Data,
                .MimeType = mimeType
            }
        End Function

        ''' <summary>
        ''' 创建文件内容
        ''' </summary>
        Public Shared Function CreateFile(fileName As String, content As String, mimeType As String) As MultimodalContent
            Return New MultimodalContent With {
                .Type = ContentType.File,
                .FileName = fileName,
                .FileContent = content,
                .MimeType = mimeType
            }
        End Function

        ''' <summary>
        ''' 创建图片文件内容（从二进制数据）
        ''' </summary>
        Public Shared Function CreateImageFile(fileName As String, fileBytes As Byte(), mimeType As String) As MultimodalContent
            Return New MultimodalContent With {
                .Type = ContentType.Image,
                .FileName = fileName,
                .FileBytes = fileBytes,
                .ImageBase64 = Convert.ToBase64String(fileBytes),
                .MimeType = mimeType
            }
        End Function
    End Class

    ''' <summary>
    ''' LLM 请求配置
    ''' </summary>
    Public Class LLMRequestConfig
        Public Property Provider As String = "openai"
        Public Property Model As String = "gpt-4o"
        Public Property ApiKey As String
        Public Property ApiUrl As String
        Public Property Temperature As Double = 0.7
        Public Property MaxTokens As Integer = 8192
        Public Property SystemPrompt As String
        Public Property Timeout As Integer = 240
        Public Property Stream As Boolean = True
    End Class

    ''' <summary>
    ''' LLM 响应结果
    ''' </summary>
    Public Class LLMResponse
        Public Property Success As Boolean
        Public Property Content As String
        Public Property ErrorMessage As String
        Public Property Usage As Dictionary(Of String, Integer)
        Public Property RawResponse As String
        Public Property ProcessingTime As TimeSpan

        Public Sub New()
            Usage = New Dictionary(Of String, Integer)()
        End Sub
    End Class

    ''' <summary>
    ''' 多模态 LLM Engine - 替代 VBA 代码生成，直接处理多模态内容
    ''' </summary>
    Public Class MultimodalLLMEngine
        Implements IDisposable

        Private _httpClient As HttpClient
        Private _config As LLMRequestConfig
        Private _isDisposed As Boolean
        Private ReadOnly _configLock As New Object()

        ' 使用 Lazy(Of T) 替代 DCL，确保线程安全和正确的内存可见性
        Private Shared ReadOnly _lazyInstance As New Lazy(Of MultimodalLLMEngine)(
            Function() New MultimodalLLMEngine(),
            LazyThreadSafetyMode.ExecutionAndPublication)

        Public Shared ReadOnly Property SharedInstance As MultimodalLLMEngine
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        Public Function GetCurrentConfig() As LLMRequestConfig
            SyncLock _configLock
                Return _config
            End SyncLock
        End Function

        Public Sub New()
            _httpClient = New HttpClient()
            _httpClient.Timeout = TimeSpan.FromSeconds(240)
            _isDisposed = False
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            DisposeCore(True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected Overridable Sub DisposeCore(disposing As Boolean)
            If _isDisposed Then Return

            _isDisposed = True

            If disposing Then
                If _httpClient IsNot Nothing Then
                    _httpClient.Dispose()
                    _httpClient = Nothing
                End If
            End If
        End Sub

        ''' <summary>
        ''' 初始化配置（创建深拷贝，避免外部修改影响内部状态）
        ''' </summary>
        Public Sub Initialize(config As LLMRequestConfig)
            If config Is Nothing Then
                Throw New ArgumentNullException(NameOf(config), "LLMRequestConfig 不能为 null")
            End If
            
            ' 创建深拷贝，避免外部修改影响内部状态
            Dim modelName As String = config.Model
            If String.IsNullOrEmpty(modelName) Then
                If config.Provider.ToLower() = "ollama" Then
                    modelName = Utils.ConfigConstants.Ollama.DEFAULT_MODEL
                Else
                    modelName = Utils.ConfigConstants.LLM.DEFAULT_MODEL
                End If
                System.Diagnostics.Debug.WriteLine($"[MultimodalLLMEngine] 警告: 模型名称为空，使用默认值: {modelName}")
            End If
            
            Dim newConfig = New LLMRequestConfig With {
                .Provider = config.Provider,
                .Model = modelName,
                .ApiKey = config.ApiKey,
                .ApiUrl = config.ApiUrl,
                .Temperature = config.Temperature,
                .MaxTokens = config.MaxTokens,
                .SystemPrompt = config.SystemPrompt,
                .Timeout = config.Timeout
            }

            ' 使用 SyncLock 确保 _config 赋值的线程安全
            SyncLock _configLock
                _config = newConfig
            End SyncLock
        End Sub

        ''' <summary>
        ''' 处理多模态内容请求
        ''' </summary>
        Public Async Function ProcessContentAsync(contents As List(Of MultimodalContent),
                                                   Optional userPrompt As String = "") As Task(Of LLMResponse)
            Dim logger = Utils.ChatLogger.Instance

            ' SkillManager队列已确保顺序执行，直接处理
            logger.LogInfo("LLM_ENGINE", "═══════════════════════════════════════════════════════")
            logger.LogInfo("LLM_ENGINE", "【步骤 1/5】开始处理多模态内容请求")

            ' 捕获当前配置的本地引用，确保整个请求周期使用同一配置
            ' 【修复】在 SyncLock 内读取 _config 快照，避免与 Initialize() 写入产生竞态条件
            Dim currentConfig As LLMRequestConfig
            SyncLock _configLock
                currentConfig = _config
            End SyncLock
            If currentConfig Is Nothing Then
                logger.LogError("LLM_ENGINE", "【步骤 1/5】❌ LLM Engine 未初始化配置")
                Return New LLMResponse With {
                    .Success = False,
                    .ErrorMessage = "LLM Engine 未初始化配置"
                }
            End If

            Try
                logger.LogInfo("LLM_ENGINE", $"  ├─ 提供商: {currentConfig.Provider}")
                logger.LogInfo("LLM_ENGINE", $"  ├─ 模型: {currentConfig.Model}")
                logger.LogInfo("LLM_ENGINE", $"  ├─ API URL: {currentConfig.ApiUrl}")
                logger.LogInfo("LLM_ENGINE", $"  ├─ 内容项数: {contents.Count}")
                ' 【修复】防御性 null 检查，防止 Substring 抛出 NullReferenceException
                logger.LogInfo("LLM_ENGINE", $"  └─ 用户提示: {If(String.IsNullOrEmpty(userPrompt), "(无)", userPrompt.Substring(0, Math.Min(50, userPrompt.Length)))}...")

                Dim stopwatch = System.Diagnostics.Stopwatch.StartNew()

                ' 构建请求体
                logger.LogInfo("LLM_ENGINE", "【步骤 2/5】构建请求体")
                Dim requestBody = BuildRequestBody(contents, userPrompt, currentConfig)
                Dim requestBodyString = requestBody.ToString()
                logger.LogInfo("LLM_ENGINE", $"  └─ 请求体大小: {requestBodyString.Length} 字符")

                ' 记录模型在构建请求体时的“思考”摘要（调试用）
                Try
                    Dim thoughtSummary As String = $"构建请求体思考: 内容项数={contents.Count}, 模型={currentConfig.Model}"
                    logger.LogThoughts("LLM_ENGINE", thoughtSummary)
                Catch ex As Exception
                    logger.LogDebug("LLM_ENGINE", $"记录思考摘要失败: {ex.Message}")
                End Try

                ' 发送请求
                logger.LogInfo("LLM_ENGINE", "【步骤 3/5】发送LLM请求")
                Dim response = Await Utils.RetryPolicy.ExecuteAsync(
                    Async Function()
                        Return Await SendLLMRequestAsync(requestBody, currentConfig)
                    End Function,
                    maxRetries:=2,
                    initialDelayMs:=1000)

                stopwatch.Stop()
                response.ProcessingTime = stopwatch.Elapsed

                logger.LogInfo("LLM_ENGINE", "【步骤 4/5】接收响应")
                logger.LogInfo("LLM_ENGINE", $"  ├─ 成功: {response.Success}")
                logger.LogInfo("LLM_ENGINE", $"  ├─ 处理时间: {stopwatch.Elapsed.TotalSeconds:F2} 秒")
                If response.Success Then
                    Dim contentLength As Integer = If(String.IsNullOrEmpty(response.Content), 0, response.Content.Length)
                    logger.LogInfo("LLM_ENGINE", $"  ├─ 响应长度: {contentLength} 字符")
                    ' 从 Usage 字典获取 Token 信息
                    Dim promptTokens As Integer = 0
                    Dim completionTokens As Integer = 0
                    If response.Usage IsNot Nothing Then
                        If response.Usage.ContainsKey("prompt_tokens") Then
                            promptTokens = response.Usage("prompt_tokens")
                        End If
                        If response.Usage.ContainsKey("completion_tokens") Then
                            completionTokens = response.Usage("completion_tokens")
                        End If
                    End If
                    logger.LogInfo("LLM_ENGINE", $"  └─ Token使用: 提示 {promptTokens}, 完成 {completionTokens}")
                Else
                    logger.LogError("LLM_ENGINE", $"  └─ 错误: {response.ErrorMessage}")
                End If

                logger.LogInfo("LLM_ENGINE", "【步骤 5/5】处理完成")
                logger.LogInfo("LLM_ENGINE", "═══════════════════════════════════════════════════════")

                Return response
            Catch ex As Exception
                logger.LogError("LLM_ENGINE", "【步骤 3/5】❌ 处理内容时出错", ex)
                Dim fullMessage = BuildExceptionMessage(ex)
                Return New LLMResponse With {
                    .Success = False,
                    .ErrorMessage = $"处理内容时出错: {fullMessage}"
                }
            End Try
        End Function

        ''' <summary>
        ''' 递归拼接异常及内部异常信息
        ''' </summary>
        Private Function BuildExceptionMessage(ex As Exception) As String
            If ex Is Nothing Then Return ""
            Dim sb As New System.Text.StringBuilder()
            sb.Append(ex.Message)
            Dim inner = ex.InnerException
            While inner IsNot Nothing
                sb.Append(" -> ")
                sb.Append(inner.Message)
                inner = inner.InnerException
            End While
            Return sb.ToString()
        End Function

        ''' <summary>
        ''' 处理单模态文本请求（向后兼容）
        ''' </summary>
        Public Async Function ProcessTextAsync(prompt As String) As Task(Of LLMResponse)
            Dim contents = New List(Of MultimodalContent) From {
                MultimodalContent.CreateText(prompt)
            }
            Return Await ProcessContentAsync(contents)
        End Function

        ''' <summary>
        ''' 处理图片+文本请求
        ''' </summary>
        Public Async Function ProcessImageAndTextAsync(imageBase64 As String, text As String, 
                                                       Optional mimeType As String = "image/png") As Task(Of LLMResponse)
            Dim contents = New List(Of MultimodalContent) From {
                MultimodalContent.CreateImage(imageBase64, mimeType),
                MultimodalContent.CreateText(text)
            }
            Return Await ProcessContentAsync(contents)
        End Function

        ''' <summary>
        ''' 流式处理多模态内容请求 — 逐 token 推送到 WebView2 前端
        ''' 返回完整的 LLMResponse（同时流式推送中间 chunk）
        ''' </summary>
        Public Async Function ProcessContentStreamingAsync(contents As List(Of MultimodalContent),
                                                            Optional userPrompt As String = "",
                                                            Optional streamId As String = Nothing) As Task(Of LLMResponse)
            Dim logger = Utils.ChatLogger.Instance

            Dim currentConfig As LLMRequestConfig
            SyncLock _configLock
                currentConfig = _config
            End SyncLock
            If currentConfig Is Nothing Then
                Return New LLMResponse With {.Success = False, .ErrorMessage = "LLM Engine 未初始化配置"}
            End If

            Dim pushService = StreamPushService.Instance
            Dim sid = pushService.CreateStream(streamId)
            logger.LogInfo("LLM_ENGINE", $"[Streaming] 创建流 {sid}")

            Try
                Dim stopwatch = System.Diagnostics.Stopwatch.StartNew()
                Dim requestBody = BuildRequestBody(contents, userPrompt, currentConfig)
                requestBody("stream") = True

                Dim response = Await SendStreamingRequestAsync(requestBody, currentConfig, sid, pushService)

                stopwatch.Stop()
                response.ProcessingTime = stopwatch.Elapsed

                logger.LogInfo("LLM_ENGINE", $"[Streaming] 流 {sid} 完成, 耗时 {stopwatch.Elapsed.TotalSeconds:F2}s, 成功={response.Success}")
                Return response
            Catch ex As Exception
                pushService.PushError(sid, ex.Message)
                Return New LLMResponse With {.Success = False, .ErrorMessage = ex.Message}
            Finally
                pushService.RemoveStream(sid)
            End Try
        End Function

        ''' <summary>
        ''' 发送流式 HTTP 请求，逐行解析 SSE，实时推送 chunk
        ''' </summary>
        Private Async Function SendStreamingRequestAsync(requestBody As JObject,
                                                          currentConfig As LLMRequestConfig,
                                                          streamId As String,
                                                          pushService As StreamPushService) As Task(Of LLMResponse)
            Dim logger = Utils.ChatLogger.Instance
            Dim url = If(String.IsNullOrEmpty(currentConfig.ApiUrl),
                        GetDefaultApiUrl(currentConfig.Provider),
                        currentConfig.ApiUrl)

            Dim headers = New Dictionary(Of String, String) From {{"Content-Type", "application/json"}}
            Select Case currentConfig.Provider.ToLower()
                Case "openai", "azure", "moonshot", "qianwen", "zhipu", "wenxin"
                    headers("Authorization") = $"Bearer {currentConfig.ApiKey}"
            End Select

            Dim requestBodyString = requestBody.ToString()
            Dim cts = New CancellationTokenSource(TimeSpan.FromSeconds(currentConfig.Timeout))

            Using request = New HttpRequestMessage(HttpMethod.Post, url)
                request.Content = New StringContent(requestBodyString, Encoding.UTF8, "application/json")
                For Each header In headers
                    request.Headers.TryAddWithoutValidation(header.Key, header.Value)
                Next

                Dim httpResponse = Await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cts.Token)

                If Not httpResponse.IsSuccessStatusCode Then
                    Dim errorContent = Await httpResponse.Content.ReadAsStringAsync()
                    Return New LLMResponse With {
                        .Success = False,
                        .ErrorMessage = $"API 请求失败: {httpResponse.StatusCode} - {errorContent}"
                    }
                End If

                Dim fullContent As New StringBuilder()
                Dim fullReasoning As New StringBuilder()
                Dim parser As New SseParser()
                Dim lastModel As String = ""
                Dim usage As New Dictionary(Of String, Integer)()

                Dim respStream = Await httpResponse.Content.ReadAsStreamAsync()
                Using reader = New StreamReader(respStream, Encoding.UTF8)
                    Dim line As String
                    line = Await reader.ReadLineAsync()
                    While line IsNot Nothing
                        Dim chunks = parser.ProcessChunk(line & vbLf)
                        For Each chunk In chunks
                            If Not String.IsNullOrEmpty(chunk.Content) Then
                                fullContent.Append(chunk.Content)
                            End If
                            If Not String.IsNullOrEmpty(chunk.Reasoning) Then
                                fullReasoning.Append(chunk.Reasoning)
                            End If
                            If Not String.IsNullOrEmpty(chunk.Model) Then
                                lastModel = chunk.Model
                            End If
                            If chunk.Usage IsNot Nothing Then
                                For Each kvp In chunk.Usage
                                    usage(kvp.Key) = kvp.Value
                                Next
                            End If

                            pushService.PushChunk(streamId, chunk)

                            If chunk.IsDone OrElse chunk.IsError Then
                                If chunk.IsError Then
                                    Return New LLMResponse With {
                                        .Success = False,
                                        .ErrorMessage = chunk.ErrorMessage
                                    }
                                End If
                            End If
                        Next
                        line = Await reader.ReadLineAsync()
                    End While
                End Using

                Return New LLMResponse With {
                    .Success = True,
                    .Content = fullContent.ToString(),
                    .Usage = usage,
                    .RawResponse = fullContent.ToString()
                }
            End Using
        End Function

        ''' <summary>
        ''' 构建 LLM 请求体
        ''' </summary>
        Private Function BuildRequestBody(contents As List(Of MultimodalContent), userPrompt As String, currentConfig As LLMRequestConfig) As JObject
            ' 【兜底】contents 为空但调用方提供了 userPrompt 文本时，用它生成一个文本内容，避免空内容误报
            If (contents Is Nothing OrElse contents.Count = 0) AndAlso Not String.IsNullOrEmpty(userPrompt) Then
                If contents Is Nothing Then contents = New List(Of MultimodalContent)()
                contents.Add(MultimodalContent.CreateText(userPrompt))
            End If

            If contents Is Nothing OrElse contents.Count = 0 Then
                Throw New ArgumentException("contents 不能为空", NameOf(contents))
            End If

            ' 根据提供商选择不同的请求格式
            If currentConfig.Provider.ToLower() = "ollama" Then
                Return BuildOllamaRequestBody(contents, userPrompt, currentConfig)
            Else
                Return BuildOpenAICompatibleRequestBody(contents, userPrompt, currentConfig)
            End If
        End Function

        ''' <summary>
        ''' 构建 OpenAI 兼容格式的请求体（支持多模态数组内容）
        ''' </summary>
        Private Function BuildOpenAICompatibleRequestBody(contents As List(Of MultimodalContent), userPrompt As String, currentConfig As LLMRequestConfig) As JObject
            Dim messages As New JArray()
            
            ' 添加系统提示
            If Not String.IsNullOrEmpty(currentConfig.SystemPrompt) Then
                messages.Add(New JObject From {
                    {"role", "system"},
                    {"content", currentConfig.SystemPrompt}
                })
            End If

            ' 构建用户消息内容（数组格式）
            Dim userContent As New JArray()
            
            For Each content In contents
                Select Case content.Type
                    Case ContentType.Text
                        userContent.Add(New JObject From {
                            {"type", "text"},
                            {"text", content.Text}
                        })
                    
                    Case ContentType.Image
                        Dim imageUrl As String
                        If Not String.IsNullOrEmpty(content.ImageBase64) Then
                            imageUrl = $"data:{content.MimeType};base64,{CleanBase64Data(content.ImageBase64)}"
                        Else
                            imageUrl = content.ImageUrl
                        End If
                        
                        userContent.Add(New JObject From {
                            {"type", "image_url"},
                            {"image_url", New JObject From {
                                {"url", imageUrl}
                            }}
                        })
                    
                    Case ContentType.File
                        ' 根据 MIME 类型判断处理方式
                        If IsImageMimeType(content.MimeType) AndAlso content.FileBytes IsNot Nothing Then
                            Dim base64 = Convert.ToBase64String(content.FileBytes)
                            userContent.Add(New JObject From {
                                {"type", "image_url"},
                                {"image_url", New JObject From {
                                    {"url", $"data:{content.MimeType};base64,{base64}"}
                                }}
                            })
                        Else
                            ' 非图片文件，将内容作为文本处理
                            Dim fileText = $"[文件: {content.FileName}]{vbCrLf}{content.FileContent}"
                            userContent.Add(New JObject From {
                                {"type", "text"},
                                {"text", fileText}
                            })
                        End If
                    
                    Case ContentType.Table
                        ' 表格作为文本处理，添加标记
                        Dim tableText = $"[表格数据]{vbCrLf}{content.Text}"
                        userContent.Add(New JObject From {
                            {"type", "text"},
                            {"text", tableText}
                        })
                    
                    Case ContentType.Mixed
                        ' 混合类型，尝试解析 Metadata 中的内容
                        If content.Metadata IsNot Nothing Then
                            If content.Metadata.ContainsKey("text") Then
                                userContent.Add(New JObject From {
                                    {"type", "text"},
                                    {"text", content.Metadata("text").ToString()}
                                })
                            End If
                            If content.Metadata.ContainsKey("image_base64") Then
                                Dim mimeType = If(content.Metadata.ContainsKey("mime_type"), content.Metadata("mime_type").ToString(), "image/png")
                                userContent.Add(New JObject From {
                                    {"type", "image_url"},
                                    {"image_url", New JObject From {
                                        {"url", $"data:{mimeType};base64,{CleanBase64Data(content.Metadata("image_base64").ToString())}"}
                                    }}
                                })
                            End If
                        End If
                End Select
            Next

            ' 添加用户额外提示
            If Not String.IsNullOrEmpty(userPrompt) Then
                userContent.Add(New JObject From {
                    {"type", "text"},
                    {"text", userPrompt}
                })
            End If

            messages.Add(New JObject From {
                {"role", "user"},
                {"content", userContent}
            })

            ' 确保模型名称不为空
            Dim modelName As String = If(String.IsNullOrEmpty(currentConfig.Model), Utils.ConfigConstants.LLM.DEFAULT_MODEL, currentConfig.Model)

            ' DeepSeek V4 系列只允许 temperature=1（否则 API 返回 400 invalid temperature）
            Dim requestTemperature As Double = currentConfig.Temperature
            If modelName.ToLower().StartsWith("deepseek-v4") Then
                requestTemperature = 1.0
            End If

            ' 构建完整请求体
            Dim requestBody As New JObject From {
                {"model", modelName},
                {"messages", messages},
                {"temperature", requestTemperature},
                {"max_tokens", currentConfig.MaxTokens},
                {"stream", currentConfig.Stream}
            }

            Return requestBody
        End Function

        ''' <summary>
        ''' 构建 Ollama 格式的请求体（content 为字符串，图片放在 images 数组中）
        ''' </summary>
        Private Function BuildOllamaRequestBody(contents As List(Of MultimodalContent), userPrompt As String, currentConfig As LLMRequestConfig) As JObject
            Dim messages As New JArray()
            
            ' 添加系统提示
            If Not String.IsNullOrEmpty(currentConfig.SystemPrompt) Then
                messages.Add(New JObject From {
                    {"role", "system"},
                    {"content", currentConfig.SystemPrompt}
                })
            End If

            ' 收集所有文本内容和图片
            Dim textParts As New List(Of String)()
            Dim imageBase64List As New List(Of String)()
            
            For Each content In contents
                Select Case content.Type
                    Case ContentType.Text
                        If Not String.IsNullOrEmpty(content.Text) Then
                            textParts.Add(content.Text)
                        End If
                    
                    Case ContentType.Image
                        If Not String.IsNullOrEmpty(content.ImageBase64) Then
                            imageBase64List.Add(CleanBase64Data(content.ImageBase64))
                        End If
                    
                    Case ContentType.File
                        ' 根据 MIME 类型判断处理方式
                        If IsImageMimeType(content.MimeType) AndAlso content.FileBytes IsNot Nothing Then
                            imageBase64List.Add(Convert.ToBase64String(content.FileBytes))
                            textParts.Add($"[图片文件: {content.FileName}]")
                        Else
                            textParts.Add($"[文件: {content.FileName}]{vbCrLf}{content.FileContent}")
                        End If
                    
                    Case ContentType.Table
                        ' 表格作为文本处理
                        textParts.Add($"[表格数据]{vbCrLf}{content.Text}")
                    
                    Case ContentType.Mixed
                        ' 混合类型，尝试解析 Metadata 中的内容
                        If content.Metadata IsNot Nothing Then
                            If content.Metadata.ContainsKey("text") Then
                                textParts.Add(content.Metadata("text").ToString())
                            End If
                            If content.Metadata.ContainsKey("image_base64") Then
                                imageBase64List.Add(CleanBase64Data(content.Metadata("image_base64").ToString()))
                            End If
                        End If
                End Select
            Next

            ' 添加用户额外提示
            If Not String.IsNullOrEmpty(userPrompt) Then
                textParts.Add(userPrompt)
            End If

            ' 构建用户消息（Ollama 格式：content 为字符串，images 为 base64 数组）
            Dim userMessage As New JObject From {
                {"role", "user"},
                {"content", String.Join(vbCrLf & vbCrLf, textParts)}
            }
            
            ' 如果有图片，添加 images 数组
            If imageBase64List.Count > 0 Then
                Dim imagesArray As New JArray()
                For Each img In imageBase64List
                    imagesArray.Add(img)
                Next
                userMessage("images") = imagesArray
            End If
            
            messages.Add(userMessage)

            ' 确保模型名称不为空
            Dim modelName As String = If(String.IsNullOrEmpty(currentConfig.Model), Utils.ConfigConstants.Ollama.DEFAULT_MODEL, currentConfig.Model)
            
            ' 构建完整请求体（Ollama /api/chat 格式）
            Dim requestBody As New JObject From {
                {"model", modelName},
                {"messages", messages},
                {"stream", currentConfig.Stream}
            }
            
            ' 添加可选参数（Temperature >= 0 包含 0 值）
            If currentConfig.Temperature >= 0 Then
                Dim optionsObj As New JObject From {
                    {"temperature", currentConfig.Temperature}
                }
                
                If currentConfig.MaxTokens > 0 Then
                    optionsObj("num_predict") = currentConfig.MaxTokens
                End If
                
                requestBody("options") = optionsObj
            Else
                If currentConfig.MaxTokens > 0 Then
                    requestBody("options") = New JObject From {
                        {"num_predict", currentConfig.MaxTokens}
                    }
                End If
            End If

            Return requestBody
        End Function

        ''' <summary>
        ''' 判断是否为图片 MIME 类型
        ''' </summary>
        Private Shared Function IsImageMimeType(mimeType As String) As Boolean
            If String.IsNullOrEmpty(mimeType) Then Return False
            Dim imageTypes As String() = {"image/jpeg", "image/png", "image/gif", "image/webp", "image/bmp", "image/tiff"}
            Return imageTypes.Any(Function(t) mimeType.ToLower().StartsWith(t))
        End Function

        ''' <summary>
        ''' 清理 Base64 数据（移除 data URI 前缀）
        ''' </summary>
        Private Shared Function CleanBase64Data(base64Data As String) As String
            If String.IsNullOrEmpty(base64Data) Then Return base64Data
            ' 移除可能的空白字符（如换行符）
            Dim cleaned = base64Data.Trim()
            Dim commaIndex = cleaned.IndexOf(","c)
            If commaIndex >= 0 Then
                Return cleaned.Substring(commaIndex + 1).Trim()
            End If
            Return cleaned
        End Function

        ''' <summary>
        ''' 发送 LLM 请求
        ''' </summary>
        Private Async Function SendLLMRequestAsync(requestBody As JObject, currentConfig As LLMRequestConfig) As Task(Of LLMResponse)
            Dim logger = Utils.ChatLogger.Instance
            logger.LogInfo("LLM_ENGINE", "  ├─ 【SendLLMRequestAsync】开始")

            Dim url = If(String.IsNullOrEmpty(currentConfig.ApiUrl),
                        GetDefaultApiUrl(currentConfig.Provider),
                        currentConfig.ApiUrl)

            logger.LogInfo("LLM_ENGINE", $"  │   ├─ 请求URL: {url}")

            Dim headers = New Dictionary(Of String, String) From {
                {"Content-Type", "application/json"}
            }

            ' 根据提供商添加认证头
            Select Case currentConfig.Provider.ToLower()
                Case "openai", "azure"
                    headers("Authorization") = $"Bearer {currentConfig.ApiKey}"
                    logger.LogInfo("LLM_ENGINE", $"  │   ├─ 认证方式: Bearer Token (OpenAI/Azure)")
                Case "moonshot"
                    headers("Authorization") = $"Bearer {currentConfig.ApiKey}"
                    logger.LogInfo("LLM_ENGINE", $"  │   ├─ 认证方式: Bearer Token (Moonshot)")
                Case "qianwen"
                    headers("Authorization") = $"Bearer {currentConfig.ApiKey}"
                    logger.LogInfo("LLM_ENGINE", $"  │   ├─ 认证方式: Bearer Token (Qianwen)")
                Case "zhipu"
                    headers("Authorization") = $"Bearer {currentConfig.ApiKey}"
                    logger.LogInfo("LLM_ENGINE", $"  │   ├─ 认证方式: Bearer Token (Zhipu)")
                Case "wenxin"
                    headers("Authorization") = $"Bearer {currentConfig.ApiKey}"
                    logger.LogInfo("LLM_ENGINE", $"  │   ├─ 认证方式: Bearer Token (Wenxin)")
                Case "ollama"
                    logger.LogInfo("LLM_ENGINE", $"  │   ├─ 认证方式: 无 (Ollama本地)")
            End Select

            Dim requestBodyString = requestBody.ToString()
            Dim sanitizedBody = SanitizeRequestBody(requestBodyString)
            logger.LogInfo("LLM_ENGINE", $"  │   ├─ 请求体: {sanitizedBody.Substring(0, Math.Min(500, sanitizedBody.Length))}...")
            logger.LogInfo("LLM_ENGINE", $"  │   ├─ 模型: {currentConfig.Model}")
            logger.LogInfo("LLM_ENGINE", $"  │   ├─ 流式模式: {currentConfig.Stream}")

            Try
                Dim thoughtDetails As String = $"请求摘要: model={currentConfig.Model}, requestBodyLen={requestBodyString.Length}, stream={currentConfig.Stream}"
                logger.LogThoughts("LLM_INTERNAL", thoughtDetails)
            Catch ex As Exception
                logger.LogDebug("LLM_ENGINE", $"记录思考细节失败: {ex.Message}")
            End Try

            Dim cts = New CancellationTokenSource(TimeSpan.FromSeconds(currentConfig.Timeout))

            Using request = New HttpRequestMessage(HttpMethod.Post, url)
                request.Content = New StringContent(requestBodyString, Encoding.UTF8, "application/json")

                For Each header In headers
                    request.Headers.TryAddWithoutValidation(header.Key, header.Value)
                Next

                logger.LogInfo("LLM_ENGINE", $"  │   ├─ 发送HTTP请求...")
                Dim completionOption = If(currentConfig.Stream, HttpCompletionOption.ResponseHeadersRead, HttpCompletionOption.ResponseContentRead)
                Dim httpResponse = Await _httpClient.SendAsync(request, completionOption, cts.Token)

                logger.LogInfo("LLM_ENGINE", $"  │   ├─ HTTP状态码: {httpResponse.StatusCode}")

                If Not httpResponse.IsSuccessStatusCode Then
                    Dim errorContent = Await httpResponse.Content.ReadAsStringAsync()
                    logger.LogError("LLM_ENGINE", $"  │   ❌ API请求失败: {httpResponse.StatusCode}")
                    Return New LLMResponse With {
                        .Success = False,
                        .ErrorMessage = $"API 请求失败: {httpResponse.StatusCode} - {errorContent}",
                        .RawResponse = errorContent
                    }
                End If

                If currentConfig.Stream Then
                    logger.LogInfo("LLM_ENGINE", $"  │   ├─ 流式响应模式，使用 SseParser 解析...")
                    Dim fullContent As New StringBuilder()
                    Dim parser As New SseParser()
                    Dim usage As New Dictionary(Of String, Integer)()

                    Dim respStream = Await httpResponse.Content.ReadAsStreamAsync()
                    Using reader = New StreamReader(respStream, Encoding.UTF8)
                        Dim line As String
                        line = Await reader.ReadLineAsync()
                        While line IsNot Nothing
                            Dim chunks = parser.ProcessChunk(line & vbLf)
                            For Each chunk In chunks
                                If Not String.IsNullOrEmpty(chunk.Content) Then
                                    fullContent.Append(chunk.Content)
                                End If
                                If chunk.Usage IsNot Nothing Then
                                    For Each kvp In chunk.Usage
                                        usage(kvp.Key) = kvp.Value
                                    Next
                                End If
                            Next
                            line = Await reader.ReadLineAsync()
                        End While
                    End Using

                    logger.LogInfo("LLM_ENGINE", $"  │   ├─ 响应内容长度: {fullContent.Length} 字符")
                    Return New LLMResponse With {
                        .Success = True,
                        .Content = fullContent.ToString(),
                        .Usage = usage,
                        .RawResponse = fullContent.ToString()
                    }
                Else
                    Dim responseContent = Await httpResponse.Content.ReadAsStringAsync()
                    logger.LogInfo("LLM_ENGINE", $"  │   ├─ 响应大小: {responseContent.Length} 字符")
                    logger.LogInfo("LLM_ENGINE", $"  │   └─ 解析响应...")
                    Return ParseLLMResponse(responseContent, currentConfig)
                End If
            End Using
        End Function

        ''' <summary>
        ''' 脱敏请求体（移除敏感信息）
        ''' </summary>
        Private Shared Function SanitizeRequestBody(requestBody As String) As String
            Try
                Dim json = JObject.Parse(requestBody)
                ' 如果包含 images 数组，替换为占位符以减少日志大小
                Dim messages = json("messages")
                If messages IsNot Nothing AndAlso TypeOf messages Is JArray Then
                    For Each msg In CType(messages, JArray)
                        If msg("images") IsNot Nothing Then
                            Dim imageCount = msg("images").Count()
                            msg("images") = $"[{imageCount} images omitted]"
                        End If
                    Next
                End If
                Return json.ToString()
            Catch ex As Exception  ' 
                Return requestBody
            End Try
        End Function

        ''' <summary>
        ''' 解析 LLM 响应
        ''' </summary>
        Private Function ParseLLMResponse(responseContent As String, currentConfig As LLMRequestConfig) As LLMResponse
            Try
                Dim jsonResponse = JObject.Parse(responseContent)
                Dim result As New LLMResponse With {
                    .Success = True,
                    .RawResponse = responseContent
                }

                ' 根据不同提供商解析
                Select Case currentConfig.Provider.ToLower()
                    Case "ollama"
                        ' Ollama /api/chat 格式: { "message": { "role": "assistant", "content": "..." } }
                        If jsonResponse("message") IsNot Nothing Then
                            result.Content = jsonResponse("message")?("content")?.ToString()
                            
                            ' DeepSeek-R1 等推理模型：content 可能为空，推理内容在 reasoning_content 中
                            If String.IsNullOrEmpty(result.Content) Then
                                Dim reasoningContent = jsonResponse("message")?("reasoning_content")?.ToString()
                                If Not String.IsNullOrEmpty(reasoningContent) Then
                                    ' 将推理内容作为 content 返回（前端会进一步提取 JSON）
                                    result.Content = reasoningContent
                                End If
                            End If
                        Else
                            ' 尝试 /api/generate 格式: { "response": "..." }
                            result.Content = jsonResponse("response")?.ToString()
                            
                            ' DeepSeek-R1 模型的 reasoning_content
                            If String.IsNullOrEmpty(result.Content) Then
                                Dim reasoningContent = jsonResponse("reasoning_content")?.ToString()
                                If Not String.IsNullOrEmpty(reasoningContent) Then
                                    result.Content = reasoningContent
                                End If
                            End If
                        End If
                    Case Else
                        ' OpenAI 兼容格式：安全地访问 choices 数组，避免空引用和索引越界
                        Dim choices = jsonResponse("choices")
                        If choices IsNot Nothing AndAlso TypeOf choices Is JArray Then
                            Dim choicesArray = CType(choices, JArray)
                            If choicesArray.Count > 0 Then
                                result.Content = choicesArray(0)("message")?("content")?.ToString()
                            End If
                        End If
                End Select

                ' 【<think> 剥离】部分部署（LM Studio / llama.cpp / 部分中转）把思考以
                ' <think>...</think> 混在正文输出，不剥离会干扰 askJson 的 JSON 提取。
                ' 剥离后的正文非空白才替换；整个 content 都是思考时保留原文（与旧行为一致）。
                If Not String.IsNullOrEmpty(result.Content) Then
                    Dim thinkOut As String = Nothing
                    Dim stripped = SseParser.StripThinkBlocks(result.Content, thinkOut)
                    If Not String.IsNullOrWhiteSpace(stripped) Then
                        result.Content = stripped
                    End If
                End If

                ' 解析用量信息
                Dim usageToken = jsonResponse("usage")
                If usageToken IsNot Nothing AndAlso TypeOf usageToken Is JObject Then
                    Dim usageObj = CType(usageToken, JObject)
                    For Each prop In usageObj.Properties()
                        If Integer.TryParse(prop.Value.ToString(), 0) Then
                            result.Usage(prop.Name) = CInt(prop.Value)
                        End If
                    Next
                End If

                Return result
            Catch ex As Exception
                Return New LLMResponse With {
                    .Success = False,
                    .ErrorMessage = $"解析响应失败: {ex.Message}",
                    .RawResponse = responseContent
                }
            End Try
        End Function

        ''' <summary>
        ''' 获取默认 API URL
        ''' </summary>
        Private Function GetDefaultApiUrl(provider As String) As String
            Dim config = Utils.ConfigurationManager.Instance
            Select Case provider.ToLower()
                Case "openai"
                    Return config.GetSetting(Utils.ConfigConstants.LLM.API_URL_OPENAI, Utils.ConfigConstants.LLM.DEFAULT_OPENAI_URL)
                Case "azure"
                    Return config.GetSetting(Utils.ConfigConstants.LLM.API_URL_AZURE, Utils.ConfigConstants.LLM.DEFAULT_AZURE_URL)
                Case "qianwen"
                    Return config.GetSetting(Utils.ConfigConstants.LLM.API_URL_QIANWEN, Utils.ConfigConstants.LLM.DEFAULT_QIANWEN_URL)
                Case "zhipu"
                    Return config.GetSetting(Utils.ConfigConstants.LLM.API_URL_ZHIPU, Utils.ConfigConstants.LLM.DEFAULT_ZHIPU_URL)
                Case "wenxin"
                    Return config.GetSetting(Utils.ConfigConstants.LLM.API_URL_WENXIN, Utils.ConfigConstants.LLM.DEFAULT_WENXIN_URL)
                Case "moonshot"
                    Return config.GetSetting(Utils.ConfigConstants.LLM.API_URL_MOONSHOT, Utils.ConfigConstants.LLM.DEFAULT_MOONSHOT_URL)
                Case "volcengine"
                    Return config.GetSetting(Utils.ConfigConstants.LLM.API_URL_VOLCENGINE, Utils.ConfigConstants.LLM.DEFAULT_VOLCENGINE_URL)
                Case "ollama"
                    ' Ollama 使用自己的配置项
                    Return config.GetSetting(Utils.ConfigConstants.Ollama.API_URL, Utils.ConfigConstants.Ollama.DEFAULT_API_URL)
                Case Else
                    Return config.GetSetting(Utils.ConfigConstants.LLM.API_URL_OPENAI, Utils.ConfigConstants.LLM.DEFAULT_OPENAI_URL)
            End Select
        End Function

        ''' <summary>
        ''' 构建文心一言 URL（需要 access token）
        ''' </summary>
        Private Function BuildWenxinUrl(baseUrl As String) As String
            ' 这里简化处理，实际应该获取 access token
            Return baseUrl
        End Function

        ''' <summary>
        ''' 检查模型是否支持多模态
        ''' </summary>
        Public Shared Function IsMultimodalSupported(model As String) As Boolean
            If String.IsNullOrEmpty(model) Then Return False

            Dim multimodalModels As String() = New String() {
                "gpt-4o", "gpt-4o-mini", "gpt-4-vision-preview",
                "qwen-vl-max", "qwen-vl-plus",
                "glm-4v", "glm-4v-plus",
                "kimi-k3", "kimi-k2.6"
            }

            Return multimodalModels.Any(Function(m) model.ToLower().Contains(m.ToLower()))
        End Function

        ''' <summary>
        ''' 从配置管理器加载配置
        ''' </summary>
        Public Sub LoadFromConfiguration(configManager As Utils.ConfigurationManager)
            Dim cfg = If(configManager, Utils.ConfigurationManager.Instance)
            Dim provider = cfg.GetSetting(Utils.ConfigConstants.LLM.PROVIDER, Utils.ConfigConstants.LLM.DEFAULT_PROVIDER)

            ' 根据提供商获取正确的 API URL 和模型
            Dim apiUrl As String = ""
            Dim model As String = ""
            If provider.ToLower() = "ollama" Then
                apiUrl = cfg.GetSetting(Utils.ConfigConstants.Ollama.API_URL, Utils.ConfigConstants.Ollama.DEFAULT_API_URL)
                model = cfg.GetSetting(Utils.ConfigConstants.LLM.MODEL, "qwen3-vl:4b")
            Else
                apiUrl = cfg.GetSetting(Utils.ConfigConstants.LLM.API_URL, "")
                model = cfg.GetSetting(Utils.ConfigConstants.LLM.MODEL, Utils.ConfigConstants.LLM.DEFAULT_MODEL)
            End If

            ' 安全解析数值配置，避免无效格式导致启动失败
            Dim temperature As Double = Utils.ConfigConstants.LLM.DEFAULT_TEMPERATURE
            Dim maxTokens As Integer = Utils.ConfigConstants.LLM.DEFAULT_MAX_TOKENS
            Dim timeout As Integer = Utils.ConfigConstants.LLM.DEFAULT_TIMEOUT

            Dim tempStr = cfg.GetSetting(Utils.ConfigConstants.LLM.TEMPERATURE, Utils.ConfigConstants.LLM.DEFAULT_TEMPERATURE.ToString())
            If Not Double.TryParse(tempStr, temperature) Then
                System.Diagnostics.Debug.WriteLine($"[MultimodalLLMEngine] 警告: Temperature 配置值 '{tempStr}' 无效，使用默认值 {Utils.ConfigConstants.LLM.DEFAULT_TEMPERATURE}")
                temperature = Utils.ConfigConstants.LLM.DEFAULT_TEMPERATURE
            End If

            Dim tokensStr = cfg.GetSetting(Utils.ConfigConstants.LLM.MAX_TOKENS, Utils.ConfigConstants.LLM.DEFAULT_MAX_TOKENS.ToString())
            If Not Integer.TryParse(tokensStr, maxTokens) Then
                System.Diagnostics.Debug.WriteLine($"[MultimodalLLMEngine] 警告: MaxTokens 配置值 '{tokensStr}' 无效，使用默认值 {Utils.ConfigConstants.LLM.DEFAULT_MAX_TOKENS}")
                maxTokens = Utils.ConfigConstants.LLM.DEFAULT_MAX_TOKENS
            End If

            Dim timeoutStr = cfg.GetSetting(Utils.ConfigConstants.LLM.TIMEOUT, Utils.ConfigConstants.LLM.DEFAULT_TIMEOUT.ToString())
            If Not Integer.TryParse(timeoutStr, timeout) Then
                System.Diagnostics.Debug.WriteLine($"[MultimodalLLMEngine] 警告: Timeout 配置值 '{timeoutStr}' 无效，使用默认值 {Utils.ConfigConstants.LLM.DEFAULT_TIMEOUT}")
                timeout = Utils.ConfigConstants.LLM.DEFAULT_TIMEOUT
            End If

            Dim newConfig = New LLMRequestConfig With {
                .Provider = provider,
                .Model = model,
                .ApiKey = cfg.GetSetting(Utils.ConfigConstants.LLM.API_KEY, ""),
                .ApiUrl = apiUrl,
                .Temperature = temperature,
                .MaxTokens = maxTokens,
                .Timeout = timeout
            }

            SyncLock _configLock
                _config = newConfig
            End SyncLock
        End Sub

        ''' <summary>
        ''' 保存配置到配置管理器
        ''' </summary>
        Public Sub SaveToConfiguration(configManager As Utils.ConfigurationManager)
            If _config Is Nothing Then Return
            
            configManager.SetSetting("LLMProvider", _config.Provider)
            configManager.SetSetting("LLMModel", _config.Model)
            configManager.SetSetting("LLMApiKey", _config.ApiKey)
            configManager.SetSetting("LLMApiUrl", _config.ApiUrl)
            configManager.SetSetting("LLMTemperature", _config.Temperature.ToString())
            configManager.SetSetting("LLMMaxTokens", _config.MaxTokens.ToString())
        End Sub
    End Class

End Namespace