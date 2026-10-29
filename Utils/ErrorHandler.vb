Imports System.Windows.Forms

Namespace Utils

    Public Class ErrorHandler

        Public Enum ErrorSeverity
            Info = 0
            Warning = 1
            [Error] = 2
            Critical = 3
        End Enum

        Public Shared Sub HandleException(ex As Exception, operation As String, 
                                          Optional severity As ErrorSeverity = ErrorSeverity.Error)
            LogException(ex, operation, severity)
            Dim message = BuildUserFriendlyMessage(ex, operation, severity)
            ShowErrorDialog(message, severity)
        End Sub

        Public Shared Function HandleExceptionAndReturn(ex As Exception, operation As String) As String
            LogException(ex, operation, ErrorSeverity.Error)
            Return BuildDetailedErrorMessage(ex, operation)
        End Function

        Private Shared Sub LogException(ex As Exception, operation As String, severity As ErrorSeverity)
            Try
                Dim logMessage As String = String.Format("[{0}] {1}{2}异常类型: {3}{4}异常消息: {5}{6}堆栈跟踪: {7}",
                    severity, operation, vbCrLf, ex.GetType().Name, vbCrLf, ex.Message, vbCrLf, ex.StackTrace)

                System.Diagnostics.Debug.WriteLine(logMessage)

                ' 统一委托到 LogServiceManager
                Select Case severity
                    Case ErrorSeverity.Info
                        LogServiceManager.Instance.LogInfo("ErrorHandler", logMessage)
                    Case ErrorSeverity.Warning
                        LogServiceManager.Instance.LogWarning("ErrorHandler", logMessage)
                    Case ErrorSeverity.Error, ErrorSeverity.Critical
                        LogServiceManager.Instance.LogError("ErrorHandler", logMessage, ex)
                End Select
            Catch logEx As Exception
                System.Diagnostics.Debug.WriteLine(
                    String.Format("[ErrorHandler] 日志写入失败: {0} | 原始异常: {1} - {2}",
                        logEx.Message, operation, ex.Message))
            End Try
        End Sub

        Private Shared Function BuildUserFriendlyMessage(ex As Exception, operation As String, 
                                                         severity As ErrorSeverity) As String
            Dim message As New System.Text.StringBuilder()

            message.AppendLine(String.Format("在执行 '{0}' 时发生问题", operation))
            message.AppendLine()
            message.AppendLine(GetExceptionSpecificAdvice(ex))
            message.AppendLine()
            message.AppendLine("建议操作：")
            message.AppendLine(GetGeneralAdvice(severity))

            If severity = ErrorSeverity.Critical Then
                message.AppendLine()
                message.AppendLine(String.Format("技术详情：{0}", ex.Message))
            End If

            Return message.ToString()
        End Function

        Private Shared Function BuildDetailedErrorMessage(ex As Exception, operation As String) As String
            Return String.Format("操作 '{0}' 失败：{1}", operation, ex.Message)
        End Function

        Private Shared Function GetExceptionSpecificAdvice(ex As Exception) As String
            Select Case True
                Case TypeOf ex Is IO.FileNotFoundException
                    Return "找不到所需的文件。请确保所有必要的文件都已正确安装。"
                
                Case TypeOf ex Is UnauthorizedAccessException
                    Return "权限不足。请确保您有足够的权限执行此操作，或尝试以管理员身份运行 Excel。"
                
                Case TypeOf ex Is IO.IOException
                    Return "文件操作失败。请检查文件是否被其他程序占用。"
                
                Case TypeOf ex Is NullReferenceException
                    Return "遇到了空引用错误。这可能是由于某些功能未正确初始化。"
                
                Case TypeOf ex Is InvalidOperationException
                    Return "当前状态不允许执行此操作。请确保 Excel 处于正确的工作状态。"
                
                Case TypeOf ex Is System.Runtime.InteropServices.COMException
                    Return "与 Excel 通信时出错。请确保 Excel 正在正常运行。"
                
                Case TypeOf ex Is Net.WebException
                    Return "网络连接失败。请检查您的网络连接。"
                
                Case TypeOf ex Is TimeoutException
                    Return "操作超时。请稍后重试。"
                
                Case Else
                    Return String.Format("发生错误：{0}", ex.Message)
            End Select
        End Function

        Private Shared Function GetGeneralAdvice(severity As ErrorSeverity) As String
            Dim advice As New System.Text.StringBuilder()

            Select Case severity
                Case ErrorSeverity.Info
                    advice.AppendLine("- 这是信息提示，无需特殊操作")
                
                Case ErrorSeverity.Warning
                    advice.AppendLine("- 请检查您的操作是否正确")
                    advice.AppendLine("- 如果问题持续，请尝试重新启动 Excel")
                
                Case ErrorSeverity.Error
                    advice.AppendLine("- 请检查您的操作步骤")
                    advice.AppendLine("- 尝试重新启动 Excel")
                    advice.AppendLine("- 如果问题持续，请查看日志文件获取更多信息")
                
                Case ErrorSeverity.Critical
                    advice.AppendLine("- 请立即保存您的工作")
                    advice.AppendLine("- 重新启动 Excel")
                    advice.AppendLine("- 如果问题持续，请联系技术支持")
            End Select

            Return advice.ToString()
        End Function

        Private Shared Sub ShowErrorDialog(message As String, severity As ErrorSeverity)
            ' 【修复】检查是否在UI线程上，非UI线程通过 SynchronizationContext 调度
            Dim syncCtx = System.Threading.SynchronizationContext.Current
            If syncCtx Is Nothing Then
                ' 无同步上下文（如单元测试或后台线程），直接显示（可能失败但不会死锁）
                ShowErrorDialogInternal(message, severity)
            Else
                ' 通过同步上下文确保在创建该上下文的线程（通常是UI线程）上执行
                Dim capturedMessage = message
                Dim capturedSeverity = severity
                syncCtx.Post(Sub(state)
                    ShowErrorDialogInternal(capturedMessage, capturedSeverity)
                End Sub, Nothing)
            End If
        End Sub

        ''' <summary>
        ''' 内部弹窗实现（必须在UI线程上调用）
        ''' </summary>
        Private Shared Sub ShowErrorDialogInternal(message As String, severity As ErrorSeverity)
            Dim icon As MessageBoxIcon
            Dim caption As String

            Select Case severity
                Case ErrorSeverity.Info
                    icon = MessageBoxIcon.Information
                    caption = "Excel AI 助手 - 信息"
                Case ErrorSeverity.Warning
                    icon = MessageBoxIcon.Warning
                    caption = "Excel AI 助手 - 警告"
                Case ErrorSeverity.Error
                    icon = MessageBoxIcon.Error
                    caption = "Excel AI 助手 - 错误"
                Case ErrorSeverity.Critical
                    icon = MessageBoxIcon.Stop
                    caption = "Excel AI 助手 - 严重错误"
                Case Else
                    icon = MessageBoxIcon.Error
                    caption = "Excel AI 助手"
            End Select

            MessageBox.Show(message, caption, MessageBoxButtons.OK, icon)
        End Sub

        Public Shared Function ExecuteSafely(action As System.Action, operationName As String) As Boolean
            Try
                action()
                Return True
            Catch ex As Exception
                HandleException(ex, operationName, ErrorSeverity.Error)
                Return False
            End Try
        End Function

        Public Shared Function ExecuteSafely(Of T)(func As System.Func(Of T), operationName As String, 
                                                   defaultValue As T) As T
            Try
                Return func()
            Catch ex As Exception
                HandleException(ex, operationName, ErrorSeverity.Error)
                Return defaultValue
            End Try
        End Function

        ''' <summary>
        ''' 静默记录异常(不弹窗,仅日志)
        ''' </summary>
        Public Shared Sub LogExceptionSilently(ex As Exception, operation As String, 
                                               Optional severity As ErrorSeverity = ErrorSeverity.Warning)
            Try
                LogException(ex, operation, severity)
            Catch silentEx As Exception
                ' 【修复】防止递归异常但保留可观测性（仅输出到 Debug，不再调 LogException 避免循环）
                System.Diagnostics.Debug.WriteLine(
                    String.Format("[ErrorHandler-Silent] 静默日志失败: {0} | 原操作: {1}",
                        silentEx.Message, operation))
            End Try
        End Sub

    End Class

End Namespace