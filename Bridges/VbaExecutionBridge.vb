Option Explicit On
Option Infer On

Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Serialization
Imports ExcelAIVSTOAddIn.Bridges
Imports ExcelAIVSTOAddIn.Utils
Imports ExcelAIVSTOAddIn.Services

Namespace Bridges

    Public Class VbaExecutionBridge
        Implements IBridge

        Private Shared ReadOnly _lazyInstance As New Lazy(Of VbaExecutionBridge)(
            Function() New VbaExecutionBridge(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        Private _vbaLimiter As RateLimiter = RateLimiter.ForVBAExecution

        Public Shared ReadOnly Property Instance As VbaExecutionBridge
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        Private Sub New()
        End Sub

        Public ReadOnly Property BridgeName As String Implements IBridge.BridgeName
            Get
                Return "VbaExecutionBridge"
            End Get
        End Property

        Public Function ExecuteVBA(vbaCode As String, Optional skillName As String = "") As String
            Try
                If String.IsNullOrWhiteSpace(vbaCode) Then
                    Return ApiResponse.Fail("VBA 代码不能为空")
                End If

                vbaCode = CleanVBACode(vbaCode)

                If Not _vbaLimiter.TryAcquire() Then
                    Return ApiResponse.Fail("调用频率过高，请稍后再试")
                End If

                System.Diagnostics.Debug.WriteLine($"[VbaExecutionBridge] 开始执行 VBA，Skill: {skillName}, 代码长度: {vbaCode.Length}")

                Dim excelApp As Microsoft.Office.Interop.Excel.Application = Nothing
                Try
                    excelApp = Globals.ThisAddIn.Application
                    If excelApp Is Nothing Then
                        Return ApiResponse.Fail("Excel 应用程序不可用")
                    End If

                    ' 撤销快照：VBA 执行前捕获活动工作表整表快照（值/格式/数字格式/列宽）
                    Dim undoSheetName As String = Nothing
                    Dim undoSnap As String = Nothing
                    Try
                        Dim snapWb = excelApp.ActiveWorkbook
                        If snapWb IsNot Nothing Then
                            Dim snapWs = TryCast(snapWb.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                            If snapWs IsNot Nothing Then
                                undoSheetName = snapWs.Name
                                undoSnap = UndoBridge.Instance.CaptureSheetFullSnapshot(snapWs)
                                Utils.ComReleaseHelper.Release(snapWs)
                            End If
                            Utils.ComReleaseHelper.Release(snapWb)
                        End If
                    Catch snapEx As Exception
                        System.Diagnostics.Debug.WriteLine($"[VbaExecutionBridge] 撤销快照捕获失败: {snapEx.Message}")
                    End Try

                    Try
                        Dim result = Security.VBACodeSandbox.ExecuteCode(vbaCode, requireConfirmation:=True)

                        Dim resultJson = JsonConvert.SerializeObject(result)
                        Dim resultDict = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(resultJson)

                        Dim response As New Dictionary(Of String, Object) From {
                            {"success", True},
                            {"result", If(resultDict.ContainsKey("Message"), resultDict("Message").ToString(), "执行完成")},
                            {"executionMethod", If(resultDict.ContainsKey("ExecutionMethod"), resultDict("ExecutionMethod").ToString(), "Unknown")},
                            {"executedAt", If(resultDict.ContainsKey("ExecutedAt"), resultDict("ExecutedAt").ToString(), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))},
                            {"skillName", skillName}
                        }

                        If resultDict.ContainsKey("Warnings") Then
                            response("warnings") = resultDict("Warnings")
                        End If

                        Dim settings As New JsonSerializerSettings With {
                            .ContractResolver = New CamelCasePropertyNamesContractResolver(),
                            .Formatting = Formatting.Indented
                        }

                        System.Diagnostics.Debug.WriteLine($"[VbaExecutionBridge] 执行成功: {response("result")}")

                        ' 撤销登记：执行成功后记录 vba_execute 动作，Ctrl+Z 可恢复到执行前状态。
                        ' 注意：即使快照为空("{}")也照常登记——这样撤销时能给出明确提示，
                        ' 便于区分"快照未捕获(美化前状态无法还原)"与"根本没有可撤销的操作(通道问题)"。
                        Try
                            If String.IsNullOrEmpty(undoSnap) Then undoSnap = "{}"
                            System.Diagnostics.Debug.WriteLine($"[VbaExecutionBridge] 登记 vba_execute 撤销: Sheet={undoSheetName}, 快照长度={undoSnap.Length}, 空快照={(undoSnap = "{}")}")
                            UndoBridge.Instance.AddUndoAction(
                                New UndoAction("vba_execute", "", Nothing, Nothing, undoSheetName, undoSnap))
                        Catch undoEx As Exception
                            System.Diagnostics.Debug.WriteLine($"[VbaExecutionBridge] 撤销登记失败: {undoEx.Message}")
                        End Try

                        Return JsonConvert.SerializeObject(response, settings)
                    Finally
                        ' excelApp 来自 Globals.ThisAddIn.Application（VSTO 宿主全局缓存的 RCW），
                        ' 不能手动 ReleaseComObject：引用计数一旦清零，该 RCW 会被"分离"永久失效，
                        ' 导致下一次执行报 "COM 对象与其基础 RCW 分开后就不能再使用"（只能执行一次）。
                        ' 宿主拥有的 Application 对象交由 VSTO 生命周期管理，无需此处释放。
                    End Try

                Catch ex As OperationCanceledException
                    System.Diagnostics.Debug.WriteLine($"[VbaExecutionBridge] 用户取消执行: {ex.Message}")
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = "用户取消了执行",
                        .skillName = skillName
                    })
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"[VbaExecutionBridge] 执行异常: {ex.Message}")
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = ex.Message,
                        .skillName = skillName
                    })
                End Try
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[VbaExecutionBridge] 未预期异常: {ex.Message}")
                Return ApiResponse.Fail(ex.Message)
            End Try
        End Function

        Public Function DebugExecuteVBA(vbaCode As String) As String
            Try
                If String.IsNullOrWhiteSpace(vbaCode) Then
                    Return ApiResponse.Fail("VBA 代码不能为空")
                End If

                Dim excelApp As Microsoft.Office.Interop.Excel.Application = Nothing
                Try
                    excelApp = Globals.ThisAddIn.Application
                    If excelApp Is Nothing Then
                        Return ApiResponse.Fail("Excel 应用程序不可用")
                    End If

                    Try
                        Dim result = Security.VBACodeSandbox.DebugExecute(vbaCode)

                        Dim settings As New JsonSerializerSettings With {
                            .ContractResolver = New CamelCasePropertyNamesContractResolver(),
                            .Formatting = Formatting.Indented
                        }

                        Return JsonConvert.SerializeObject(result, settings)
                    Finally
                        ' excelApp 来自 Globals.ThisAddIn.Application（VSTO 宿主全局缓存的 RCW），
                        ' 不能手动 ReleaseComObject：引用计数一旦清零，该 RCW 会被"分离"永久失效，
                        ' 导致下一次执行报 "COM 对象与其基础 RCW 分开后就不能再使用"（只能执行一次）。
                        ' 宿主拥有的 Application 对象交由 VSTO 生命周期管理，无需此处释放。
                    End Try

                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"[VbaExecutionBridge.DebugExecute] 异常: {ex.Message}")
                    Return ApiResponse.Fail(ex.Message)
                End Try
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[VbaExecutionBridge.DebugExecute] 未预期异常: {ex.Message}")
                Return ApiResponse.Fail(ex.Message)
            End Try
        End Function

        Private Function CleanVBACode(vbaCode As String) As String
            If String.IsNullOrWhiteSpace(vbaCode) Then Return vbaCode

            Dim result = vbaCode.Trim()

            Dim codeBlockPattern As String = "^```(?:vba|vb)?\s*"
            result = System.Text.RegularExpressions.Regex.Replace(result, codeBlockPattern, "", System.Text.RegularExpressions.RegexOptions.IgnoreCase)

            If result.EndsWith("```") Then
                result = result.Substring(0, result.Length - 3).TrimEnd()
            End If

            Return result
        End Function

        Private _disposed As Boolean

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            System.Diagnostics.Debug.WriteLine("[VbaExecutionBridge] 已释放")
        End Sub

    End Class

End Namespace