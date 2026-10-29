Imports System.Threading.Tasks

Namespace Services

    ''' <summary>
    ''' LLM协议接口 - 统一不同LLM服务商的API调用
    ''' </summary>
    Public Interface ILlmProtocol

        ''' <summary>
        ''' 协议名称(用于配置)
        ''' </summary>
        ReadOnly Property ProtocolName As String

        ''' <summary>
        ''' 检查协议是否可用
        ''' </summary>
        Function IsAvailable(apiKey As String, apiUrl As String) As Boolean

        ''' <summary>
        ''' 构建请求体
        ''' </summary>
        Function BuildRequestBody(
            model As String,
            messages As List(Of Dictionary(Of String, Object)),
            Optional temperature As Double = 0.7,
            Optional maxTokens As Integer = 4096,
            Optional stream As Boolean = False) As String

        ''' <summary>
        ''' 解析响应
        ''' </summary>
        Function ParseResponse(responseJson As String) As LLMResponse

        ''' <summary>
        ''' 获取认证Header
        ''' </summary>
        Function GetAuthHeaders(apiKey As String) As Dictionary(Of String, String)

    End Interface

End Namespace