Imports System.Runtime.InteropServices

Namespace Utils

    ''' <summary>
    ''' COM 对象安全释放辅助类
    ''' 用于管理 Excel Interop 等 COM 对象的生命周期，防止 Excel 进程残留
    ''' </summary>
    Public NotInheritable Class ComReleaseHelper

        ''' <summary>
        ''' 安全释放单个 COM 对象
        ''' </summary>
        Public Shared Function Release(comObj As Object) As Integer
            If comObj Is Nothing Then Return 0
            Try
                Return Marshal.ReleaseComObject(comObj)
            Catch ex As Exception  ' 
                Return 0
            End Try
        End Function

        ''' <summary>
        ''' 安全释放多个 COM 对象（按传入顺序释放）
        ''' </summary>
        Public Shared Sub ReleaseAll(ParamArray comObjects As Object())
            If comObjects IsNot Nothing Then
                For Each obj As Object In comObjects
                    If obj IsNot Nothing Then
                        Try
                            Marshal.ReleaseComObject(obj)
                        Catch ex As Exception  ' 
                        End Try
                    End If
                Next
            End If
        End Sub

        ''' <summary>
        ''' 安全释放并 FinalRelease（强制最后一次引用计数清零）
        ''' 仅在确认对象不再被任何地方使用时调用
        ''' </summary>
        Public Shared Sub FinalRelease(comObj As Object)
            If comObj Is Nothing Then Return
            Try
                While Marshal.ReleaseComObject(comObj) > 0
                End While
            Catch ex As Exception  ' 
            End Try
        End Sub

        ''' <summary>
        ''' 安全调用 COM 对象上的操作，并在完成后自动释放
        ''' </summary>
        Public Shared Sub UsingCom(Of T As Class)(comObj As T, action As Action(Of T))
            If comObj IsNot Nothing Then
                Try
                    action(comObj)
                Finally
                    Try
                        Marshal.ReleaseComObject(comObj)
                    Catch ex As Exception  ' 
                    End Try
                End Try
            End If
        End Sub

        ''' <summary>
        ''' 安全调用 COM 对象上的操作并返回结果，完成后自动释放
        ''' </summary>
        Public Shared Function UsingCom(Of T As Class, TResult As Class)(comObj As T, func As Func(Of T, TResult)) As TResult
            If comObj Is Nothing Then Return Nothing
            Try
                Return func(comObj)
            Finally
                Try
                    Marshal.ReleaseComObject(comObj)
                Catch ex As Exception  ' 
                End Try
            End Try
        End Function

        ''' <summary>
        ''' 强制垃圾回收以清理未释放的 COM 对象
        ''' 应在批量 COM 操作后调用
        ''' </summary>
        Public Shared Sub ForceGCCollect()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Sub

    End Class

End Namespace
