Imports System.IO
Imports System.Net.Http
Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services

    ''' <summary>
    ''' Moonshot AI 视觉模型类型
    ''' </summary>
    Public Enum MoonshotVisionModel
        ''' <summary>
        ''' 8K 上下文窗口的视觉模型
        ''' </summary>
        Vision8K
        ''' <summary>
        ''' 32K 上下文窗口的视觉模型（推荐）
        ''' </summary>
        Vision32K
        ''' <summary>
        ''' 128K 上下文窗口的视觉模型（适合复杂图片）
        ''' </summary>
        Vision128K
    End Enum

    ''' <summary>
    ''' Moonshot AI 图片识别服务
    ''' 支持多种视觉模型：8K/32K/128K 上下文窗口
    ''' 实现ILlmProtocol接口,支持协议适配器模式
    ''' </summary>
    Public Class MoonshotImageService
        Implements ILlmProtocol

        Private Const MOONSHOT_API_URL = "https://api.moonshot.cn/v1/chat/completions"

        ''' <summary>
        ''' 获取模型名称
        ''' moonshot-v1-*-vision-preview 已于 2026-08-31 下线，统一改用原生多模态的 kimi-k3
        ''' </summary>
        Public Shared Function GetModelName(model As MoonshotVisionModel) As String
            Return "kimi-k3"
        End Function

        ''' <summary>
        ''' 使用 Moonshot API 识别图片内容（默认使用 32K 模型）
        ''' </summary>
        Public Async Function RecognizeImage(imagePath As String, apiKey As String) As Task(Of String)
            Return Await RecognizeImage(imagePath, apiKey, MoonshotVisionModel.Vision32K)
        End Function

        ''' <summary>
        ''' 使用指定的 Moonshot 视觉模型识别图片内容
        ''' </summary>
        ''' <param name="imagePath">图片文件路径</param>
        ''' <param name="apiKey">Moonshot API Key</param>
        ''' <param name="model">视觉模型类型</param>
        Public Async Function RecognizeImage(imagePath As String, apiKey As String, model As MoonshotVisionModel) As Task(Of String)
            Try
                System.Diagnostics.Debug.WriteLine("MoonshotImageService: 开始识别图片")
                System.Diagnostics.Debug.WriteLine("图片路径: " & imagePath)
                System.Diagnostics.Debug.WriteLine("使用模型: " & GetModelName(model))

                ' 验证图片文件
                If Not File.Exists(imagePath) Then
                    Return "【错误：图片文件不存在】"
                End If

                ' 读取图片并转换为 base64
                Dim imageBytes = File.ReadAllBytes(imagePath)
                Dim base64Image = Convert.ToBase64String(imageBytes)
                Dim imageType = GetImageType(imagePath)

                System.Diagnostics.Debug.WriteLine("图片大小: " & imageBytes.Length & " bytes")
                System.Diagnostics.Debug.WriteLine("图片类型: " & imageType)

                ' 构建请求
                Dim requestBody = New With {
                    .model = GetModelName(model),
                    .messages = New Object() {
                        New With {
                            .role = "system",
                            .content = "你是一个专业的图片识别助手。请准确识别图片中的所有文字内容，并以纯文本格式返回。如果图片包含表格，请尽量保持表格的原始格式。支持中英文混合识别。"
                        },
                        New With {
                            .role = "user",
                            .content = New Object() {
                                New With {
                                    .type = "image_url",
                                    .image_url = New With {
                                        .url = "data:" & imageType & ";base64," & base64Image
                                    }
                                },
                                New With {
                                    .type = "text",
                                    .text = "请识别这张图片中的所有文字内容。重要：必须保留原文的换行格式，每行文字单独一行，不要合并成一行。以纯文本格式返回，保持原有的段落和行结构。"
                                }
                            }
                        }
                    },
                    .temperature = 0.3
                }

                Dim json = JsonConvert.SerializeObject(requestBody)
                System.Diagnostics.Debug.WriteLine("请求 JSON 长度: " & json.Length)

                ' 发送请求
                Using client As New HttpClient()
                    client.DefaultRequestHeaders.Add("Authorization", "Bearer " & apiKey)
                    client.Timeout = TimeSpan.FromSeconds(60)

                    Dim content = New StringContent(json, Encoding.UTF8, "application/json")
                    System.Diagnostics.Debug.WriteLine("发送请求到 Moonshot API...")

                    Dim response = Await client.PostAsync(MOONSHOT_API_URL, content)
                    Dim responseBody = Await response.Content.ReadAsStringAsync()

                    System.Diagnostics.Debug.WriteLine("响应状态: " & response.StatusCode)
                    System.Diagnostics.Debug.WriteLine("响应长度: " & responseBody.Length)

                    If response.IsSuccessStatusCode Then
                        Dim result = JObject.Parse(responseBody)
                        Dim text = result("choices")(0)("message")("content").ToString()
                        System.Diagnostics.Debug.WriteLine("识别结果: " & text.Substring(0, Math.Min(text.Length, 200)))
                        Return text
                    Else
                        System.Diagnostics.Debug.WriteLine("API 错误: " & responseBody)
                        Return "【错误：Moonshot API 调用失败 - " & response.StatusCode & "】"
                    End If
                End Using

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Moonshot 识别失败: " & ex.Message)
                Return "【错误：Moonshot 识别失败 - " & ex.Message & "】"
            End Try
        End Function

        ''' <summary>
        ''' 获取图片 MIME 类型
        ''' </summary>
        Private Function GetImageType(imagePath As String) As String
            Dim ext = Path.GetExtension(imagePath).ToLower()
            Select Case ext
                Case ".jpg", ".jpeg"
                    Return "image/jpeg"
                Case ".png"
                    Return "image/png"
                Case ".gif"
                    Return "image/gif"
                Case ".bmp"
                    Return "image/bmp"
                Case ".webp"
                    Return "image/webp"
                Case Else
                    Return "image/jpeg"
            End Select
        End Function

        ' ==================== ILlmProtocol 实现 ====================

        Public ReadOnly Property ProtocolName As String Implements ILlmProtocol.ProtocolName
            Get
                Return "Moonshot"
            End Get
        End Property

        Public Function IsAvailable(apiKey As String, apiUrl As String) As Boolean Implements ILlmProtocol.IsAvailable
            Return Not String.IsNullOrEmpty(apiKey)
        End Function

        Public Function BuildRequestBody(
            model As String,
            messages As List(Of Dictionary(Of String, Object)),
            Optional temperature As Double = 0.7,
            Optional maxTokens As Integer = 4096,
            Optional stream As Boolean = False) As String Implements ILlmProtocol.BuildRequestBody

            Dim requestBody = New Dictionary(Of String, Object) From {
                {"model", model},
                {"messages", messages},
                {"temperature", temperature},
                {"max_tokens", maxTokens},
                {"stream", stream}
            }

            Return JsonConvert.SerializeObject(requestBody)
        End Function

        Public Function ParseResponse(responseJson As String) As LLMResponse Implements ILlmProtocol.ParseResponse
            Try
                Dim json = JObject.Parse(responseJson)

                ' Moonshot响应格式(与OpenAI类似但可能有差异)
                Dim content = json("choices")?(0)?("message")?("content")?.ToString()

                If String.IsNullOrEmpty(content) Then
                    Return New LLMResponse With {
                        .Success = False,
                        .ErrorMessage = "Moonshot响应内容为空"
                    }
                End If

                Return New LLMResponse With {
                    .Success = True,
                    .Content = content
                }

            Catch ex As Exception
                Return New LLMResponse With {
                    .Success = False,
                    .ErrorMessage = $"解析Moonshot响应失败: {ex.Message}"
                }
            End Try
        End Function

        Public Function GetAuthHeaders(apiKey As String) As Dictionary(Of String, String) Implements ILlmProtocol.GetAuthHeaders
            Dim headers = New Dictionary(Of String, String)()
            headers("Authorization") = $"Bearer {apiKey}"
            headers("Content-Type") = "application/json"
            Return headers
        End Function

    End Class

End Namespace