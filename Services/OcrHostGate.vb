Namespace Services

    ''' <summary>
    ''' 区分「独立 OCR 服务进程」和「Office 插件进程」。
    ''' 插件进程里的识别调用要转发到 OcrHost，不能在 Excel/Word/PPT 里拉起 Python。
    ''' </summary>
    Public NotInheritable Class OcrHostGate

        Private Sub New()
        End Sub

        Public Shared ReadOnly Property IsHostProcess As Boolean
            Get
                Return String.Equals(Environment.GetEnvironmentVariable("SMARTPILOT_OCR_HOST"), "1", StringComparison.Ordinal)
            End Get
        End Property

        Public Shared ReadOnly Property ShouldForward As Boolean
            Get
                Return Not IsHostProcess
            End Get
        End Property

    End Class

End Namespace
