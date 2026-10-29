Namespace Services

    ''' <summary>
    ''' LLM协议工厂 - 根据配置创建对应的协议实现
    ''' </summary>
    Public Class LlmProtocolFactory

        Private Shared _protocols As New Dictionary(Of String, ILlmProtocol)()

        Shared Sub New()
            ' 注册内置协议
            RegisterProtocol(New OpenAIProtocol())
            RegisterProtocol(New MoonshotImageService())
        End Sub

        ''' <summary>
        ''' 注册协议实现
        ''' </summary>
        Public Shared Sub RegisterProtocol(protocol As ILlmProtocol)
            If protocol IsNot Nothing AndAlso Not String.IsNullOrEmpty(protocol.ProtocolName) Then
                _protocols(protocol.ProtocolName.ToLower()) = protocol
            End If
        End Sub

        ''' <summary>
        ''' 根据协议名称获取协议实现
        ''' </summary>
        Public Shared Function GetProtocol(protocolName As String) As ILlmProtocol
            If String.IsNullOrEmpty(protocolName) Then
                Return New OpenAIProtocol() ' 默认使用OpenAI协议
            End If

            Dim key = protocolName.ToLower()
            If _protocols.ContainsKey(key) Then
                Return _protocols(key)
            End If

            ' 未找到,返回默认
            System.Diagnostics.Debug.WriteLine($"[LlmProtocolFactory] 未找到协议: {protocolName}, 使用默认OpenAI协议")
            Return New OpenAIProtocol()
        End Function

        ''' <summary>
        ''' 获取所有已注册的协议名称
        ''' </summary>
        Public Shared Function GetRegisteredProtocols() As List(Of String)
            Return _protocols.Keys.ToList()
        End Function

        ''' <summary>
        ''' 清除所有协议(用于测试)
        ''' </summary>
        Public Shared Sub ClearProtocols()
            _protocols.Clear()
        End Sub

    End Class

End Namespace
