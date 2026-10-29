Option Strict On
Option Explicit On
Option Infer On
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services

    ''' <summary>
    ''' OpenAI标准协议实现
    ''' 兼容: OpenAI, Ollama, DeepSeek, 智谱, 通义千问等
    ''' </summary>
    Public Class OpenAIProtocol
        Implements ILlmProtocol

        Public ReadOnly Property ProtocolName As String Implements ILlmProtocol.ProtocolName
            Get
                Return "OpenAI"
            End Get
        End Property

        Public Function IsAvailable(apiKey As String, apiUrl As String) As Boolean Implements ILlmProtocol.IsAvailable
            Return Not String.IsNullOrEmpty(apiUrl)
        End Function

        Public Function BuildRequestBody(
            model As String,
            messages As List(Of Dictionary(Of String, Object)),
            Optional temperature As Double = 0.7,
            Optional maxTokens As Integer = 4096,
            Optional stream As Boolean = False) As String Implements ILlmProtocol.BuildRequestBody

            ' DeepSeek V4 系列只允许 temperature=1（否则 API 返回 400 invalid temperature）
            Dim requestTemperature As Double = temperature
            If Not String.IsNullOrEmpty(model) AndAlso model.ToLower().StartsWith("deepseek-v4") Then
                requestTemperature = 1.0
            End If

            Dim requestBody = New Dictionary(Of String, Object) From {
                {"model", model},
                {"messages", messages},
                {"temperature", requestTemperature},
                {"max_tokens", maxTokens},
                {"stream", stream}
            }

            Return JsonConvert.SerializeObject(requestBody)
        End Function

        Public Function ParseResponse(responseJson As String) As LLMResponse Implements ILlmProtocol.ParseResponse
            Try
                Dim json = JObject.Parse(responseJson)

                ' 标准OpenAI格式
                Dim content = json("choices")?(0)?("message")?("content")?.ToString()
                
                If String.IsNullOrEmpty(content) Then
                    Return New LLMResponse With {
                        .Success = False,
                        .ErrorMessage = "响应内容为空"
                    }
                End If

                ' 提取usage信息
                Dim usage = New Dictionary(Of String, Integer)()
                If json("usage") IsNot Nothing Then
                    usage("prompt_tokens") = If(json("usage")("prompt_tokens")?.ToObject(Of Integer)(), 0)
                    usage("completion_tokens") = If(json("usage")("completion_tokens")?.ToObject(Of Integer)(), 0)
                    usage("total_tokens") = If(json("usage")("total_tokens")?.ToObject(Of Integer)(), 0)
                End If

                Return New LLMResponse With {
                    .Success = True,
                    .Content = content,
                    .Usage = usage
                }

            Catch ex As Exception
                Return New LLMResponse With {
                    .Success = False,
                    .ErrorMessage = $"解析响应失败: {ex.Message}"
                }
            End Try
        End Function

        Public Function GetAuthHeaders(apiKey As String) As Dictionary(Of String, String) Implements ILlmProtocol.GetAuthHeaders
            Dim headers = New Dictionary(Of String, String)()
            
            If Not String.IsNullOrEmpty(apiKey) Then
                headers("Authorization") = $"Bearer {apiKey}"
            End If
            
            headers("Content-Type") = "application/json"
            
            Return headers
        End Function

    End Class

End Namespace