'【注意】VBACodeSandbox 是 VBA 安全沙箱，通过 COM IDispatch 与 Excel 动态交互，
'  本质依赖后期绑定。此处单独关闭 Option Strict，而核心服务仍保持严格类型检查。
Option Strict Off
Option Explicit On
Option Infer On
Imports System.Text.RegularExpressions
Imports System.Security.Cryptography
Imports System.Text
Imports System.Linq
Imports System.Threading
Imports System.Windows.Forms
Imports System.Drawing
Imports ExcelAIVSToAddIn.Utils

Namespace Security

    ' VBACodeSandbox.Execution.vb — ExecuteCode 入口 / 安全日志 / 常见 Excel 操作执行器
    ' （类级 XML 文档只在主文件 VBACodeSandbox.vb 保留一份，避免 BC42314 分部重复）
    Partial Public Class VBACodeSandbox

        Public Shared Function ExecuteCode(vbaCode As String, Optional requireConfirmation As Boolean = True) As Object
            ' 0. 清理 Markdown 代码块标记（LLM 可能返回 ```vba ... ``` 格式）
            vbaCode = CleanMarkdownCodeBlock(vbaCode)

            ' 1. Perform security check
            Dim checkResult = CheckCode(vbaCode)

            ' 2. Log audit
            LogSecurityCheck(checkResult, vbaCode)

            ' 3. Check if security check passed
            If Not checkResult.IsSafe Then
                Dim errorMessage = "Code failed security check:" & Environment.NewLine &
                                  String.Join(Environment.NewLine, checkResult.Errors.Select(Function(e) "  - " & e))
                Throw New InvalidOperationException(errorMessage)
            End If

            ' 4. Medium and high risk operations require user confirmation
            '    前端「自动化免确认」开关开启时（AutoExecutionSettings.BypassConfirmation=True）跳过弹窗，
            '    供自动化任务自动执行；其余情况仍按原规则确认。
            If requireConfirmation AndAlso Not AutoExecutionSettings.BypassConfirmation AndAlso checkResult.RiskLevel >= RiskLevel.Medium Then
                Dim confirmMessage As String
                Dim isCommonExcelOp As Boolean = checkResult.Warnings.Any(Function(w) w.Contains("Common Excel operation detected"))

                If isCommonExcelOp Then
                    confirmMessage = "AI生成的代码需要执行以下Excel操作：" & Environment.NewLine & Environment.NewLine &
                                    String.Join(Environment.NewLine, checkResult.Warnings.
                                                Where(Function(w) w.Contains("Common Excel operation detected")).
                                                Select(Function(w) "  • " & w.Replace("Common Excel operation detected: ", ""))) &
                                    Environment.NewLine & Environment.NewLine &
                                    "是否继续执行？"
                Else
                    confirmMessage = String.Format(
                        "检测到风险级别为 {0} 的代码，是否继续执行？" & Environment.NewLine & Environment.NewLine &
                        "警告信息:" & Environment.NewLine & "{1}",
                        checkResult.RiskLevel,
                        String.Join(Environment.NewLine, checkResult.Warnings.Select(Function(w) "  - " & w))
                    )
                End If

                ' 【安全修复】不再通过堆栈检测自动跳过安全确认
                ' 所有 Medium/High 风险代码均需用户明确确认，无论调用来源为何
                ' 原逻辑：检测 CustomSkill 调用堆栈 → 自动确认（已移除，存在安全绕过风险）
                ' 如确需自定义Skill无交互执行，请使用 requireConfirmation:=False 参数显式传入

                ' 检查是否有不在白名单中的操作，用于显示"加入白名单"选项
                Dim hasUnknownOps As Boolean = checkResult.UnknownWhitelistOps IsNot Nothing AndAlso checkResult.UnknownWhitelistOps.Count > 0

                Dim dialogResult As DialogResult
                If hasUnknownOps Then
                    ' 有未命中白名单的操作 → 显示三按钮对话框（是/加入白名单并执行/否）
                    dialogResult = ShowWhitelistConfirmDialog(confirmMessage, checkResult.UnknownWhitelistOps, isCommonExcelOp)
                Else
                    ' 无未命中操作（纯高风险如文件删除）→ 普通二按钮对话框
                    dialogResult = System.Windows.Forms.MessageBox.Show(
                            confirmMessage,
                            "安全确认",
                            System.Windows.Forms.MessageBoxButtons.YesNo,
                            If(isCommonExcelOp, System.Windows.Forms.MessageBoxIcon.Question, System.Windows.Forms.MessageBoxIcon.Warning)
                        )
                End If

                If dialogResult = DialogResult.No OrElse dialogResult = DialogResult.Cancel Then
                    Throw New OperationCanceledException("用户取消了执行")
                End If

                ' 如果用户选择了"加入白名单并执行"，将未知操作加入白名单
                If dialogResult = DialogResult.Retry AndAlso hasUnknownOps Then
                    Dim addedCount As Integer = AddToWhitelist(checkResult.UnknownWhitelistOps.ToArray())
                    If addedCount > 0 Then
                        System.Diagnostics.Debug.WriteLine($"已将 {addedCount} 个操作加入安全白名单")
                    End If
                End If
            End If

            ' 5. Execute code (isolated in try-catch)
            Dim excelApp As Microsoft.Office.Interop.Excel.Application = Nothing
            Try
                excelApp = Globals.ThisAddIn.Application
                Dim resultMessage As String = ""
                Dim executionMethod As String = "Unknown"

                ' 【修复】Selection 预检查：如果 VBA 代码依赖 Selection 但当前无有效选中区域，提前返回友好提示
                Dim codeLower As String = vbaCode.ToLower()
                If codeLower.Contains("selection") OrElse codeLower.Contains("activecell") Then
                    Dim sel As Object = Nothing
                    Try
                        sel = excelApp.Selection
                    Catch : End Try
                    Dim isInvalidSelection As Boolean = False
                    Dim skipReason As String = ""
                    If sel Is Nothing Then
                        isInvalidSelection = True
                        skipReason = "当前没有选中任何内容"
                    ElseIf Not TypeOf sel Is Microsoft.Office.Interop.Excel.Range Then
                        ' Selection 可能是图表、形状等非 Range 对象
                        isInvalidSelection = True
                        skipReason = "当前选中的可能是图表、形状或其他对象"
                    Else
                        ' 选中了 Range 但内容为空（没有任何数据）
                        Dim selRange As Microsoft.Office.Interop.Excel.Range = Nothing
                        Try
                            selRange = DirectCast(sel, Microsoft.Office.Interop.Excel.Range)
                            Dim cellCount As Long = selRange.Cells.Count
                            If cellCount > 0 AndAlso cellCount <= 100000 Then
                                Dim countA As Long = 0
                                Try
                                    countA = excelApp.WorksheetFunction.CountA(selRange)
                                Catch : End Try
                                If countA = 0 Then
                                    isInvalidSelection = True
                                    skipReason = "选中的区域没有任何数据"
                                End If
                            End If
                        Finally
                            Utils.ComReleaseHelper.Release(selRange)
                        End Try
                    End If
                    If isInvalidSelection Then
                        Utils.ComReleaseHelper.Release(sel)
                        Return New With {
                            .Code = vbaCode,
                            .CodeHash = checkResult.CodeHash,
                            .Message = "✗ " & skipReason & "。请先在 Excel 中选择要操作的单元格区域，再执行此代码。",
                            .Warnings = checkResult.Warnings,
                            .ExecutionMethod = "PreCheck",
                            .ExecutedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                        }
                    End If
                    Utils.ComReleaseHelper.Release(sel)
                End If

                ' 使用统一检测方法识别常见 Excel 操作 → 路由到专用 Interop
                Dim commonOp As CommonExcelOperation = DetectCommonExcelOperation(vbaCode)

                Select Case commonOp
                    Case CommonExcelOperation.ClearSheet
                        resultMessage = ExecuteClearSheetOperation(excelApp)
                        executionMethod = "ClearSheet"
                    Case CommonExcelOperation.SplitWorkbook
                        resultMessage = ExecuteSplitWorkbookOperation(excelApp)
                        executionMethod = "SplitWorkbook"
                    Case CommonExcelOperation.MergeSheets
                        resultMessage = ExecuteMergeSheetsOperation(excelApp)
                        executionMethod = "MergeSheets"
                    Case CommonExcelOperation.RemoveDuplicates
                        resultMessage = ExecuteRemoveDuplicatesOperation(excelApp)
                        executionMethod = "RemoveDuplicates"
                    Case CommonExcelOperation.CreateChart
                        resultMessage = ExecuteCreateChartOperation(excelApp)
                        executionMethod = "CreateChart"

                    Case Else
                        ' 【简化】所有代码直接执行，取消沙箱
                        ' 沙箱（临时工作簿）带来的问题（路径错误、结果丢失、覆盖 A1）超过安全价值
                        ' 安全由前置 CheckCode + 用户确认弹窗保障
                        resultMessage = ExecuteDirectInWorkbook(excelApp, vbaCode)
                        executionMethod = "Direct"
                End Select

                ' 6. Log execution success
                LogExecutionSuccess(checkResult.CodeHash, executionMethod, resultMessage)

                Return New With {
                    .Code = vbaCode,
                    .CodeHash = checkResult.CodeHash,
                    .Message = If(resultMessage, "Execution completed"),
                    .Warnings = checkResult.Warnings,
                    .ExecutionMethod = executionMethod,
                    .ExecutedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                }

            Catch ex As OperationCanceledException
                Throw
            Catch ex As Exception
                LogExecutionFailure(checkResult.CodeHash, ex.Message)
                Throw New InvalidOperationException("Error executing VBA code: " & ex.Message, ex)
            End Try
        End Function

        ''' <summary>
        ''' Log security check
        ''' </summary>
        Private Shared Sub LogSecurityCheck(result As SecurityCheckResult, code As String)
            Try
                Dim logEntry = String.Format(
                    "[{0}] VBA Security Check - Hash:{1}, Safe:{2}, Risk:{3}, Warnings:{4}, Errors:{5}",
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    result.CodeHash.Substring(0, 8) & "...",
                    result.IsSafe,
                    result.RiskLevel,
                    result.Warnings.Count,
                    result.Errors.Count
                )
                System.Diagnostics.Debug.WriteLine(logEntry)
                ChatLogger.Log(logEntry, "Security")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[Security] 安全日志记录失败: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' Log execution success
        ''' </summary>
        Private Shared Sub LogExecutionSuccess(codeHash As String, method As String, message As String)
            Try
                Dim logEntry = String.Format(
                    "[{0}] VBA Execution Success - Hash:{1}, Method:{2}, Result:{3}",
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    codeHash.Substring(0, Math.Min(8, codeHash.Length)) & "...",
                    method,
                    If(message?.Length > 50, message.Substring(0, 50) & "...", message)
                )
                System.Diagnostics.Debug.WriteLine(logEntry)
                ChatLogger.Log(logEntry, "Security")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[Security] 执行成功日志记录失败: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' Log execution failure
        ''' </summary>
        Private Shared Sub LogExecutionFailure(codeHash As String, errorMessage As String)
            Try
                Dim logEntry = String.Format(
                    "[{0}] VBA Execution Failed - Hash:{1}, Error:{2}",
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    codeHash.Substring(0, Math.Min(8, codeHash.Length)) & "...",
                    errorMessage
                )
                System.Diagnostics.Debug.WriteLine(logEntry)
                ChatLogger.Log(logEntry, "Security")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[Security] 执行失败日志记录失败: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' Execute split workbook operation
        ''' </summary>
        Private Shared Function ExecuteSplitWorkbookOperation(excelApp As Microsoft.Office.Interop.Excel.Application) As String
            Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
            Try
                wb = excelApp.ActiveWorkbook
                If wb Is Nothing Then Return "Error: No open workbook"

                Dim folderPath As String = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "SplitWorkbooks_" & DateTime.Now.ToString("yyyyMMdd_HHmmss")
                )
                If Not System.IO.Directory.Exists(folderPath) Then System.IO.Directory.CreateDirectory(folderPath)

                Dim originalDisplayAlerts As Boolean = excelApp.DisplayAlerts
                excelApp.DisplayAlerts = False
                Dim wsCount As Integer = wb.Worksheets.Count
                Dim savedFiles As New List(Of String)()

                For i As Integer = 1 To wsCount
                    Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                    Dim newWb As Microsoft.Office.Interop.Excel.Workbook = Nothing
                    Try
                        ws = wb.Worksheets(i)
                        ws.Copy()
                        newWb = excelApp.ActiveWorkbook
                        Dim fileName As String = System.IO.Path.Combine(folderPath, ws.Name & ".xlsx")
                        newWb.SaveAs(fileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook)
                        savedFiles.Add(fileName)
                        newWb.Close(False)
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine($"[SplitWorkbook] Sheet {i} 处理失败: {ex.Message}")
                        If newWb IsNot Nothing Then
                            Try : newWb.Close(False) : Catch closeEx As Exception
                                System.Diagnostics.Debug.WriteLine($"[SplitWorkbook] 关闭新工作簿失败: {closeEx.Message}")
                            End Try
                        End If
                    Finally
                        ComReleaseHelper.Release(ws)
                        ComReleaseHelper.Release(newWb)
                    End Try
                Next
                excelApp.DisplayAlerts = originalDisplayAlerts
                Return $"Success! Saved {savedFiles.Count} worksheets to: {folderPath}"
            Catch ex As Exception
                Return $"Execution failed: {ex.Message}"
            Finally
                ComReleaseHelper.Release(wb)
                ComReleaseHelper.ForceGCCollect()
            End Try
        End Function

        ''' <summary>
        ''' Execute merge sheets operation
        ''' </summary>
        Private Shared Function ExecuteMergeSheetsOperation(excelApp As Microsoft.Office.Interop.Excel.Application) As String
            Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
            Try
                wb = excelApp.ActiveWorkbook
                If wb Is Nothing Then Return "Error: No open workbook"

                Dim targetWs As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Try
                    targetWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
                    targetWs.Name = "Merge Result"
                    Dim row As Long = 1
                    Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                    Dim usedRange As Microsoft.Office.Interop.Excel.Range = Nothing

                    Dim sheetCount As Integer = wb.Worksheets.Count
                    For i As Integer = 1 To sheetCount
                        Dim sheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                        Dim targetCell As Object = Nothing
                        Try
                            sheet = CType(wb.Worksheets(i), Microsoft.Office.Interop.Excel.Worksheet)
                            If sheet.Name <> targetWs.Name Then
                                usedRange = sheet.UsedRange
                                If usedRange IsNot Nothing Then
                                    targetCell = targetWs.Cells(row, 1)
                                    usedRange.Copy(targetCell)
                                    row += usedRange.Rows.Count + 1
                                End If
                                ComReleaseHelper.Release(usedRange)
                                usedRange = Nothing
                            End If
                        Finally
                            ComReleaseHelper.Release(targetCell)
                            ComReleaseHelper.Release(sheet)
                        End Try
                    Next
                    ComReleaseHelper.Release(targetWs)
                    Return "Success! Created worksheet 'Merge Result'"
                Catch ex As Exception
                    Return $"Execution failed: {ex.Message}"
                Finally
                    ComReleaseHelper.Release(targetWs)
                End Try
            Catch ex As Exception
                Return $"Execution failed: {ex.Message}"
            Finally
                ComReleaseHelper.Release(wb)
                ComReleaseHelper.ForceGCCollect()
            End Try
        End Function

        ''' <summary>
        ''' Execute remove duplicates operation
        ''' </summary>
        Private Shared Function ExecuteRemoveDuplicatesOperation(excelApp As Microsoft.Office.Interop.Excel.Application) As String
            Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim rangeToRemove As Microsoft.Office.Interop.Excel.Range = Nothing
            Try
                ws = excelApp.ActiveSheet
                If ws Is Nothing Then Return "Error: No active worksheet"

                Dim lastRow As Long = ws.Cells(ws.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
                rangeToRemove = ws.Range("A1:A" & lastRow.ToString())
                rangeToRemove.RemoveDuplicates(Columns:=1, Header:=Microsoft.Office.Interop.Excel.XlYesNoGuess.xlYes)
                Return "Success! Deduplication completed"
            Catch ex As Exception
                Return $"Execution failed: {ex.Message}"
            Finally
                ComReleaseHelper.Release(rangeToRemove)
                ComReleaseHelper.Release(ws)
            End Try
        End Function

        ''' <summary>
        ''' Execute create chart operation
        ''' </summary>
        Private Shared Function ExecuteCreateChartOperation(excelApp As Microsoft.Office.Interop.Excel.Application) As String
            Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim selection As Microsoft.Office.Interop.Excel.Range = Nothing
            Dim chartObj As Microsoft.Office.Interop.Excel.ChartObject = Nothing
            Try
                ws = excelApp.ActiveSheet
                If ws Is Nothing Then Return "Error: No active worksheet"

                selection = excelApp.Selection
                If selection Is Nothing Then Return "Error: Please select data area first"

                chartObj = ws.ChartObjects.Add(100, 100, 400, 300)
                chartObj.Chart.SetSourceData(Source:=selection)
                chartObj.Chart.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered
                Return "Success! Column chart created"
            Catch ex As Exception
                Return $"Execution failed: {ex.Message}"
            Finally
                ComReleaseHelper.Release(chartObj)
                ComReleaseHelper.Release(selection)
                ComReleaseHelper.Release(ws)
            End Try
        End Function

        ''' <summary>
        ''' Execute clear sheet operation - directly via Excel Interop
        ''' </summary>
        Private Shared Function ExecuteClearSheetOperation(excelApp As Microsoft.Office.Interop.Excel.Application) As String
            Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Try
                ws = excelApp.ActiveSheet
                If ws Is Nothing Then Return "Error: No active worksheet"

                Dim sheetName As String = ws.Name
                ws.Cells.ClearContents()
                ws.Cells.ClearFormats()
                excelApp.ScreenUpdating = True
                System.Windows.Forms.Application.DoEvents()
                Return $"Success! Cleared worksheet '{sheetName}'"
            Catch ex As Exception
                Return $"Clear sheet failed: {ex.Message}"
            Finally
                ComReleaseHelper.Release(ws)
            End Try
        End Function

        ' 【历史备注】执行 VBA 代码：优先 Interop 直调 → 回退 VBProject Import+Run 真执行
        ' （该方法定义在其他分部文件中）

    End Class

End Namespace
