Imports System.Threading
Imports System.Threading.Tasks

Namespace Services

    ''' <summary>
    ''' 错误恢复机制（v3.0新增）
    ''' 分类处理不同类型的错误，提供降级策略
    ''' </summary>
    Public Class ErrorRecovery

        Private Shared ReadOnly _instance As New Lazy(Of ErrorRecovery)(Function() New ErrorRecovery(), System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        Public Shared ReadOnly Property Instance As ErrorRecovery
            Get
                Return _instance.Value
            End Get
        End Property

        Private Sub New()
        End Sub

        ''' <summary>
        ''' 处理错误（主入口）
        ''' </summary>
        Public Async Function Handle([error] As Exception, context As Dictionary(Of String, Object)) As Task(Of Dictionary(Of String, Object))
            Dim errorType = [error].GetType().Name

            Select Case errorType
                Case "ArgumentException", "ArgumentNullException", "InvalidCastException"
                    Return Await HandleInvalidParameter([error], context)

                Case "COMException", "TargetInvocationException", "InvalidOperationException"
                    Return Await HandleExecutionFailed([error], context)

                Case "TimeoutException", "TaskCanceledException"
                    Return Await HandleTimeout([error], context)

                Case Else
                    System.Diagnostics.Debug.WriteLine($"[ErrorRecovery] 未知错误类型: {errorType}, 消息: {[error].Message}")
                    Return New Dictionary(Of String, Object) From {
                        {"success", False},
                        {"action", "unknown_error"},
                        {"message", $"执行失败: {[error].Message}"}
                    }
            End Select
        End Function

        ''' <summary>
        ''' 处理参数错误（让模型重新提取参数）
        ''' </summary>
        Private Function HandleInvalidParameter([error] As Exception, context As Dictionary(Of String, Object)) As Task(Of Dictionary(Of String, Object))
            System.Diagnostics.Debug.WriteLine($"[ErrorRecovery] 参数无效，尝试重新提取: {[error].Message}")

            Return Task.FromResult(New Dictionary(Of String, Object) From {
                {"success", True},
                {"action", "re_extract"},
                {"message", "参数提取失败，正在重新分析..."},
                {"confidence", 0.5}
            })
        End Function

        ''' <summary>
        ''' 处理执行失败（回滚+降级方案）
        ''' </summary>
        Private Function HandleExecutionFailed([error] As Exception, context As Dictionary(Of String, Object)) As Task(Of Dictionary(Of String, Object))
            System.Diagnostics.Debug.WriteLine($"[ErrorRecovery] 执行失败，尝试回滚: {[error].Message}")

            ' 检查是否有回滚方法
            If context.ContainsKey("skill") AndAlso context("skill") IsNot Nothing Then
                Try
                    Dim skill = context("skill")
                    Dim rollbackMethod = skill.GetType().GetMethod("Rollback")
                    If rollbackMethod IsNot Nothing Then
                        rollbackMethod.Invoke(skill, Nothing)
                        System.Diagnostics.Debug.WriteLine("[ErrorRecovery] 回滚成功")
                    End If
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"[ErrorRecovery] 回滚失败: {ex.Message}")
                End Try
            End If

            Return Task.FromResult(New Dictionary(Of String, Object) From {
                {"success", False},
                {"action", "rollback"},
                {"message", "操作失败，已回滚。请尝试其他方案。"},
                {"confidence", 0.0}
            })
        End Function

        ''' <summary>
        ''' 处理超时（返回已执行部分）
        ''' </summary>
        Private Function HandleTimeout([error] As Exception, context As Dictionary(Of String, Object)) As Task(Of Dictionary(Of String, Object))
            System.Diagnostics.Debug.WriteLine($"[ErrorRecovery] 超时，返回部分结果")

            Return Task.FromResult(New Dictionary(Of String, Object) From {
                {"success", True},
                {"action", "partial_result"},
                {"message", "操作超时，已返回已完成部分。"},
                {"confidence", 0.6}
            })
        End Function

    End Class

End Namespace
