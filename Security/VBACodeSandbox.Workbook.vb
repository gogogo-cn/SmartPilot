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

    ' VBACodeSandbox.Workbook.vb — 沙箱工作簿内直接执行 / 区间与清理操作
    ' （类级 XML 文档只在主文件 VBACodeSandbox.vb 保留一份，避免 BC42314 分部重复）
    Partial Public Class VBACodeSandbox

        Private Shared Function ExecuteDirectInWorkbook(
            excelApp As Microsoft.Office.Interop.Excel.Application,
            vbaCode As String) As String

            Dim vbComp As Object = Nothing
            Dim tempBasFile As String = ""
            Dim targetWb As Object = Nothing
            Dim vbProj As Object = Nothing
            Dim vbComps As Object = Nothing

            ' 保存原始状态
            Dim originalDisplayAlerts As Boolean = excelApp.DisplayAlerts
            Dim originalScreenUpdating As Boolean = excelApp.ScreenUpdating
            Dim originalEnableEvents As Boolean = excelApp.EnableEvents

            Try
                ' 禁用 UI 更新以提升性能并减少闪烁
                excelApp.DisplayAlerts = False
                excelApp.ScreenUpdating = False
                excelApp.EnableEvents = False

                targetWb = excelApp.ActiveWorkbook
                If targetWb Is Nothing Then
                    Return "✗ 没有打开的工作簿"
                End If

                ' 预处理 VBA 代码
                vbaCode = SanitizeVBAOnErrorResumeNext(vbaCode)
                vbaCode = ReplaceMsgBoxWithCellWrite(vbaCode)

                Dim macroName As String = ExtractMacroName(vbaCode)
                If String.IsNullOrEmpty(macroName) Then
                    macroName = "SBAIExec" & (DateTime.Now.Ticks Mod 10000).ToString()
                    vbaCode = "Sub " & macroName & "()" & vbCrLf & vbaCode & vbCrLf & "End Sub"
                End If

                Const moduleName As String = "ModuleSBAI"

                ' 如果之前遗留了同名模块，先移除，避免导入后产生 ModuleSBAI1 导致宏名冲突
                Dim existingComp As Object = Nothing
                Try
                    vbProj = targetWb.VBProject
                    If vbProj IsNot Nothing Then
                        vbComps = vbProj.VBComponents
                        If vbComps IsNot Nothing Then
                            existingComp = vbComps.Item(moduleName)
                            If existingComp IsNot Nothing Then
                                vbComps.Remove(existingComp)
                                System.Diagnostics.Debug.WriteLine($"[Direct] 已移除遗留模块 {moduleName}")
                            End If
                        End If
                    End If
                Catch ex As Exception  '  模块不存在时会抛异常，忽略
                Finally
                    Utils.ComReleaseHelper.Release(existingComp)
                End Try

                ' 写入临时 .bas 文件
                tempBasFile = IO.Path.Combine(IO.Path.GetTempPath(), "SBAI_" & Guid.NewGuid().ToString("N") & ".bas")
                IO.File.WriteAllText(tempBasFile, "Attribute VB_Name = """ & moduleName & """" & vbCrLf & vbaCode, System.Text.Encoding.Default)

                ' 导入到目标工作簿
                Try
                    If vbComps Is Nothing Then
                        If vbProj Is Nothing Then vbProj = targetWb.VBProject
                        If vbProj IsNot Nothing Then vbComps = vbProj.VBComponents
                    End If
                    If vbComps IsNot Nothing Then vbComp = vbComps.Import(tempBasFile)
                Catch : End Try

                ' 执行宏前临时恢复 DisplayAlerts，确保 VBA 运行时错误能正确抛出
                Dim savedDisplayAlerts As Boolean = excelApp.DisplayAlerts
                excelApp.DisplayAlerts = True
                Try
                    ' 使用工作簿限定名，避免多工作簿/重名宏导致的找不到宏问题
                    Dim wbName As String = targetWb.Name.ToString().Replace("'", "''")
                    Dim fullMacroName As String = "'" & wbName & "'!" & moduleName & "." & macroName
                    excelApp.Run(fullMacroName)
                Catch runEx As Exception
                    ' VBA 运行时错误（如 424 对象必需）会作为 COMException 抛出
                    Dim errorMsg As String = runEx.Message
                    If errorMsg.Contains("424") OrElse errorMsg.Contains("对象必需") OrElse errorMsg.Contains("Object required") Then
                        Return "✗ VBA 执行错误：代码中引用的对象无效。可能原因：选中的区域为空或代码逻辑有误。请重新选择有数据的区域后再试。"
                    End If
                    If errorMsg.Contains("无法运行") OrElse errorMsg.Contains("宏") OrElse errorMsg.Contains("macro") OrElse errorMsg.Contains("disabled") OrElse errorMsg.Contains("禁用") Then
                        Return "✗ 无法运行 VBA 宏：Excel 可能禁用了宏，或未信任 VBA 工程访问。请在 Excel 信任中心启用宏并允许 VBA 工程访问后重试。"
                    End If
                    Throw
                Finally
                    excelApp.DisplayAlerts = savedDisplayAlerts
                End Try

                ' 清理导入的模块
                Dim compToRemove As Object = Nothing
                Try
                    If vbComps Is Nothing Then
                        If vbProj Is Nothing Then vbProj = targetWb.VBProject
                        If vbProj IsNot Nothing Then vbComps = vbProj.VBComponents
                    End If
                    If vbComps IsNot Nothing Then
                        compToRemove = vbComps.Item(moduleName)
                        If compToRemove IsNot Nothing Then vbComps.Remove(compToRemove)
                    End If
                Catch cleanupEx As Exception
                    System.Diagnostics.Debug.WriteLine($"[Direct] 清理模块失败（非致命）: {cleanupEx.Message}")
                Finally
                    Utils.ComReleaseHelper.Release(compToRemove)
                End Try

                Return "✓ 执行成功"

            Catch ex As Exception
                Return "✗ 执行失败: " & ex.Message
            Finally
                ' 恢复 Excel 状态
                excelApp.DisplayAlerts = originalDisplayAlerts
                excelApp.ScreenUpdating = originalScreenUpdating
                excelApp.EnableEvents = originalEnableEvents

                ' 删除临时文件
                If Not String.IsNullOrEmpty(tempBasFile) AndAlso IO.File.Exists(tempBasFile) Then
                    Try : IO.File.Delete(tempBasFile) : Catch : End Try
                End If

                ' 释放 COM 对象
                Utils.ComReleaseHelper.Release(vbComps)
                Utils.ComReleaseHelper.Release(vbProj)
                ComReleaseHelper.Release(vbComp)
                ComReleaseHelper.Release(targetWb)
            End Try
        End Function

        ''' <summary>
        ''' Execute VBA code in a TEMPORARY sandbox workbook instead of the active workbook.
        ''' This isolates the execution and prevents malicious code from damaging user data.
        ''' 
        ''' Strategy:
        ''' 1. Create a hidden temp workbook (no UI interaction, no events)
        ''' 2. Copy the structure of the active sheet (headers and data) into the temp workbook
        ''' 3. Import and run VBA in the temp workbook
        ''' 4. After execution, copy results back to the original workbook
        ''' 5. Close temp workbook without saving to disk — no trace left behind
        ''' 
        ''' Timeout: 30 seconds max — prevents infinite loops
        ''' </summary>
        Private Shared Function ExecuteInSandboxWorkbook(
            excelApp As Microsoft.Office.Interop.Excel.Application,
            vbaCode As String,
            checkResult As SecurityCheckResult) As String

            Dim sandboxWb As Microsoft.Office.Interop.Excel.Workbook = Nothing
            Dim originalWb As Microsoft.Office.Interop.Excel.Workbook = Nothing
            Dim originalSheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim sandboxSheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim tempBasFile As String = ""
            Dim vbComp As Object = Nothing

            ' Track original state for restoration
            Dim originalDisplayAlerts As Boolean = excelApp.DisplayAlerts
            Dim originalScreenUpdating As Boolean = excelApp.ScreenUpdating
            Dim originalEnableEvents As Boolean = excelApp.EnableEvents
            Dim originalCalculation As Microsoft.Office.Interop.Excel.XlCalculation = excelApp.Calculation

            Try
                ' Step 0: Capture current context
                originalWb = excelApp.ActiveWorkbook
                If originalWb Is Nothing Then
                    Return "✗ 没有打开的工作簿，无法创建沙箱环境"
                End If
                originalSheet = excelApp.ActiveSheet

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
                        isInvalidSelection = True
                        skipReason = "当前选中的可能是图表、形状或其他对象"
                    Else
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
                        Return "✗ " & skipReason & "。请先在 Excel 中选择要操作的单元格区域，再执行此代码。"
                    End If
                    Utils.ComReleaseHelper.Release(sel)
                End If

                ' Disable UI updates, alerts, events during sandbox execution
                excelApp.DisplayAlerts = False
                excelApp.ScreenUpdating = False
                excelApp.EnableEvents = False
                excelApp.Calculation = Microsoft.Office.Interop.Excel.XlCalculation.xlCalculationManual

                System.Diagnostics.Debug.WriteLine($"[Sandbox] 开始创建沙箱工作簿...")

                ' Step 1: Create temp workbook (hidden, no template)
                sandboxWb = excelApp.Workbooks.Add()
                If sandboxWb Is Nothing Then
                    Return "✗ 无法创建沙箱工作簿"
                End If

                ' Step 2: Copy current sheet structure into sandbox
                ' Only copy headers and a limited data range (max 1000 rows for safety)
                Try
                    sandboxSheet = CType(sandboxWb.Sheets(1), Microsoft.Office.Interop.Excel.Worksheet)
                Catch : End Try

                If originalSheet IsNot Nothing AndAlso sandboxSheet IsNot Nothing Then
                    Dim srcRange As Microsoft.Office.Interop.Excel.Range = Nothing
                    Dim dstRange As Microsoft.Office.Interop.Excel.Range = Nothing
                    Try
                        Dim maxCopyRow As Long = Math.Min(SANDBOX_MAX_COPY_ROWS, 
                            originalSheet.Cells(originalSheet.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row)
                        Dim maxCopyCol As Integer = Math.Min(SANDBOX_MAX_COPY_COLS,
                            originalSheet.Cells(1, originalSheet.Columns.Count).End(Microsoft.Office.Interop.Excel.XlDirection.xlToLeft).Column)

                        If maxCopyRow > 0 Then
                            srcRange = CType(originalSheet.Range(originalSheet.Cells(1, 1), originalSheet.Cells(maxCopyRow, maxCopyCol)),
                                Microsoft.Office.Interop.Excel.Range)
                            dstRange = CType(sandboxSheet.Range(sandboxSheet.Cells(1, 1), sandboxSheet.Cells(maxCopyRow, maxCopyCol)),
                                Microsoft.Office.Interop.Excel.Range)
                            dstRange.Value = srcRange.Value

                            ' Copy column widths
                            For col As Integer = 1 To maxCopyCol
                                Dim srcCol As Object = Nothing
                                Dim dstCol As Object = Nothing
                                Try
                                    srcCol = originalSheet.Columns(col)
                                    dstCol = sandboxSheet.Columns(col)
                                    CType(dstCol, Microsoft.Office.Interop.Excel.Range).ColumnWidth =
                                        CType(srcCol, Microsoft.Office.Interop.Excel.Range).ColumnWidth
                                Catch : End Try
                                Utils.ComReleaseHelper.Release(srcCol)
                                Utils.ComReleaseHelper.Release(dstCol)
                            Next

                            System.Diagnostics.Debug.WriteLine($"[Sandbox] 已复制 {maxCopyRow} 行 x {maxCopyCol} 列到沙箱")
                        End If
                    Catch copyEx As Exception
                        System.Diagnostics.Debug.WriteLine($"[Sandbox] 复制工作表结构失败: {copyEx.Message}")
                        ' Continue with empty sandbox — better than failing
                    Finally
                        Utils.ComReleaseHelper.Release(srcRange)
                        Utils.ComReleaseHelper.Release(dstRange)
                    End Try
                End If

                ' Step 3: Prepare VBA code for sandbox execution
                ' 预处理：移除 On Error Resume Next（防止安全错误被静默忽略）
                vbaCode = SanitizeVBAOnErrorResumeNext(vbaCode)
                ' 预处理：将 MsgBox/InputBox 替换为单元格写入（沙箱抑制弹窗，需要单元格来传递输出）
                vbaCode = ReplaceMsgBoxWithCellWrite(vbaCode)

                Dim macroName As String = ExtractMacroName(vbaCode)
                If String.IsNullOrEmpty(macroName) Then
                    macroName = "SBAIExec" & (DateTime.Now.Ticks Mod 10000).ToString()
                    vbaCode = "Sub " & macroName & "()" & vbCrLf & vbaCode & vbCrLf & "End Sub"
                End If

                ' Step 4: Write .bas file and import to SANDBOX workbook (NOT original)
                Dim tempModuleName As String = "SBTMP" & (DateTime.Now.Ticks Mod 10000).ToString()
                tempBasFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), tempModuleName & ".bas")
                System.IO.File.WriteAllText(tempBasFile,
                    "Attribute VB_Name = """ & tempModuleName & """" & vbCrLf & vbaCode,
                    System.Text.Encoding.Default)

                System.Diagnostics.Debug.WriteLine($"[Sandbox] bas={tempBasFile}, macro={macroName}")

                ' Step 5: Import VBA to SANDBOX workbook
                Dim vbProj As Object = Nothing
                Dim vbComps As Object = Nothing
                Try
                    vbProj = sandboxWb.VBProject
                    If vbProj Is Nothing Then
                        Return "✗ 沙箱工作簿无法访问 VBProject（需要启用'信任对VBA项目对象模型的访问'）"
                    End If

                    vbComps = vbProj.VBComponents
                    Try
                        vbComp = vbComps.Import(tempBasFile)
                    Catch iex As Exception
                        Return "✗ 沙箱导入失败: " & iex.Message
                    End Try

                    If vbComp Is Nothing Then
                        Return "✗ 沙箱导入返回空——代码可能包含语法错误"
                    End If
                Finally
                    Utils.ComReleaseHelper.Release(vbComps)
                    Utils.ComReleaseHelper.Release(vbProj)
                End Try

                ' Step 6: Execute VBA in sandbox with COM-aware timeout
                ' NOTE: Excel COM is STA (Single-Threaded Apartment).
                ' We use a dedicated STA thread to isolate the COM call context,
                ' combined with EnableCancelKey to allow VBA interruption.
                Dim fullMacro As String = tempModuleName & "." & macroName
                System.Diagnostics.Debug.WriteLine($"[Sandbox] 在沙箱中执行: {fullMacro}")

                Dim execException As Exception = Nothing
                Dim executionCompleted As Boolean = False
                Dim execLock As New Object()

                ' Use dedicated STA thread (NOT ThreadPool) for COM isolation
                Dim execThread As New Thread(
                    Sub()
                        Try
                            excelApp.Run(fullMacro)
                        Catch ex As Exception
                            SyncLock execLock
                                execException = ex
                            End SyncLock
                            System.Diagnostics.Debug.WriteLine($"[Sandbox] 执行异常: {ex.Message}")
                        Finally
                            SyncLock execLock
                                executionCompleted = True
                            End SyncLock
                        End Try
                    End Sub)
                execThread.SetApartmentState(ApartmentState.STA)
                execThread.IsBackground = True
                execThread.Start()

                ' Wait with timeout
                Dim joined As Boolean = execThread.Join(SANDBOX_TIMEOUT_SECONDS * 1000)

                If Not joined Then
                    System.Diagnostics.Debug.WriteLine("[Sandbox] 执行超时 (30秒)，强制终止")

                    ' Step 1: Enable Ctrl+Break style interrupt
                    Try
                        excelApp.EnableCancelKey = Microsoft.Office.Interop.Excel.XlEnableCancelKey.xlInterrupt
                        System.Threading.Thread.Sleep(300)
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine($"[Sandbox] EnableCancelKey 失败: {ex.Message}")
                    End Try

                    ' Step 2: Close sandbox workbook — this kills any VBA running in it
                    Try
                        sandboxWb.Close(SaveChanges:=False)
                        sandboxWb = Nothing
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine($"[Sandbox] 关闭沙箱工作簿失败: {ex.Message}")
                        ' Force close by name
                        Try
                            Dim wbs As Object = Nothing
                            Try
                                wbs = excelApp.Workbooks
                                If wbs IsNot Nothing Then
                                    Dim wbCount As Integer = wbs.Count
                                    For i As Integer = 1 To wbCount
                                        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
                                        Try
                                            wb = CType(wbs(i), Microsoft.Office.Interop.Excel.Workbook)
                                            If wb IsNot sandboxWb AndAlso wb IsNot originalWb Then
                                                wb.Close(SaveChanges:=False)
                                            End If
                                        Finally
                                            Utils.ComReleaseHelper.Release(wb)
                                        End Try
                                    Next
                                End If
                            Finally
                                Utils.ComReleaseHelper.Release(wbs)
                            End Try
                        Catch : End Try
                    End Try

                    ' Step 3: Restore cancel key state
                    Try
                        excelApp.EnableCancelKey = Microsoft.Office.Interop.Excel.XlEnableCancelKey.xlDisabled
                    Catch : End Try

                    Return "✗ 沙箱执行超时 (30秒限制)。代码可能包含无限循环或长时间运行的操作，已被终止。"
                End If

                ' Check for exception
                SyncLock execLock
                    If execException IsNot Nothing Then
                        System.Diagnostics.Debug.WriteLine($"[Sandbox] 执行异常详情: {execException.Message}")
                        Dim errorMsg As String = execException.Message
                        If errorMsg.Contains("424") OrElse errorMsg.Contains("对象必需") OrElse errorMsg.Contains("Object required") Then
                            Return "✗ VBA 执行错误：代码中引用的对象无效。可能原因：选中的区域为空或代码逻辑有误。请重新选择有数据的区域后再试。"
                        End If
                        Return " 沙箱执行失败: " & errorMsg
                    End If
                End SyncLock

                ' Step 7: Read back results from sandbox to determine what happened
                Dim resultSummary As New StringBuilder()

                Try
                    ' Check if A1 was written to (likely MsgBox replacement output)
                    Dim a1Val As Object = Nothing
                    Dim a2Val As Object = Nothing
                    Dim a1Cell As Object = Nothing
                    Dim a2Cell As Object = Nothing
                    Dim usedRange As Object = Nothing
                    Try
                        a1Cell = sandboxSheet.Range("A1")
                        a1Val = a1Cell.Value
                        a2Cell = sandboxSheet.Range("A2")
                        a2Val = a2Cell.Value
                    Finally
                        ' 【修复】释放沙箱结果读取区的 COM Range 对象
                        Utils.ComReleaseHelper.Release(a1Cell)
                        Utils.ComReleaseHelper.Release(a2Cell)
                    End Try

                    Dim hasOutput As Boolean = False

                    If a1Val IsNot Nothing AndAlso a1Val.ToString().Length > 0 Then
                        hasOutput = True
                        Dim a1Text As String = a1Val.ToString()
                        ' If output contains line breaks, format nicely
                        If a1Text.Contains(vbCrLf) OrElse a1Text.Contains(vbLf) Then
                            resultSummary.AppendLine("✓ 执行成功！输出结果：")
                            resultSummary.AppendLine(a1Text)
                        Else
                            resultSummary.AppendLine("✓ 执行成功！输出: " & a1Text)
                        End If
                    End If
                    If a2Val IsNot Nothing AndAlso a2Val.ToString().Length > 0 Then
                        resultSummary.AppendLine("A2: " & a2Val.ToString())
                    End If

                    If Not hasOutput Then
                        resultSummary.AppendLine("✓ 沙箱执行成功")
                    End If

                    ' Check if any cells were modified
                    usedRange = sandboxSheet.UsedRange
                    If usedRange IsNot Nothing Then
                        Dim modifiedCells As Integer = 0
                        Try
                            modifiedCells = usedRange.Cells.Count
                        Finally
                            ' 【修复】释放 UsedRange COM 对象
                            Utils.ComReleaseHelper.Release(usedRange)
                        End Try
                        If modifiedCells > 0 Then
                            resultSummary.AppendLine(String.Format("沙箱中 {0} 个单元格被修改", modifiedCells))

                            ' 【安全修复】不再自动将沙箱结果回写到原工作簿
                        ' 原逻辑：dstRange.Value = srcRange.Value 直接写回（已移除，破坏沙箱隔离）
                        ' 沙箱执行结果仅在沙箱工作簿中可见，保护用户原始数据不被意外覆写
                        ' 如需结果，用户可手动确认后单独回写
                        If modifiedCells > 0 Then
                            resultSummary.AppendLine(String.Format("沙箱中 {0} 个单元格被修改（已隔离，未回写到原工作簿）", modifiedCells))
                            System.Diagnostics.Debug.WriteLine($"[Sandbox] 执行完成（隔离模式）: {modifiedCells} 个单元格被修改")
                        End If
                        End If
                    End If
                Catch readEx As Exception
                    resultSummary.AppendLine("读取沙箱结果时出错: " & readEx.Message)
                End Try

                Return resultSummary.ToString()

            Catch ex As OperationCanceledException
                Throw
            Catch ex As Exception
                Return "✗ 沙箱执行异常: " & ex.Message
            Finally
                ' CRITICAL: Restore Excel state and clean up sandbox
                Try
                    ' Step 8: Remove VBA module from sandbox
                    If vbComp IsNot Nothing AndAlso sandboxWb IsNot Nothing Then
                        Dim p As Object = Nothing
                        Dim comps As Object = Nothing
                        Try
                            p = sandboxWb.VBProject
                            If p IsNot Nothing Then
                                comps = p.VBComponents
                                If comps IsNot Nothing Then comps.Remove(vbComp)
                            End If
                        Catch : End Try
                        Utils.ComReleaseHelper.Release(comps)
                        Utils.ComReleaseHelper.Release(p)
                        Utils.ComReleaseHelper.Release(vbComp)
                        vbComp = Nothing
                    End If
                Catch : End Try

                Try
                    ' Step 9: Close sandbox workbook WITHOUT saving (no trace)
                    If sandboxWb IsNot Nothing Then
                        sandboxWb.Close(SaveChanges:=False)
                    End If
                Catch closeEx As Exception
                    System.Diagnostics.Debug.WriteLine($"[Sandbox] 关闭沙箱工作簿失败: {closeEx.Message}")
                    ' Force close if needed
                    Try
                        Dim wbs As Object = Nothing
                        Try
                            wbs = excelApp.Workbooks
                            If wbs IsNot Nothing AndAlso sandboxWb IsNot Nothing Then
                                Dim forceWb As Object = Nothing
                                Try
                                    forceWb = wbs(sandboxWb.Name)
                                    If forceWb IsNot Nothing Then forceWb.Close(SaveChanges:=False)
                                Finally
                                    Utils.ComReleaseHelper.Release(forceWb)
                                End Try
                            End If
                        Finally
                            Utils.ComReleaseHelper.Release(wbs)
                        End Try
                    Catch : End Try
                End Try

                Try
                    ' Step 10: Delete temp .bas file
                    If Not String.IsNullOrEmpty(tempBasFile) AndAlso System.IO.File.Exists(tempBasFile) Then
                        System.IO.File.Delete(tempBasFile)
                    End If
                Catch : End Try

                Try
                    ' Step 11: Restore original Excel state
                    excelApp.DisplayAlerts = originalDisplayAlerts
                    excelApp.ScreenUpdating = originalScreenUpdating
                    excelApp.EnableEvents = originalEnableEvents
                    excelApp.Calculation = originalCalculation

                    ' Reactivate original workbook
                    If originalWb IsNot Nothing Then
                        originalWb.Activate()
                    End If
                Catch restoreEx As Exception
                    System.Diagnostics.Debug.WriteLine($"[Sandbox] 恢复Excel状态失败: {restoreEx.Message}")
                End Try

                ' 【修复】释放 COM 对象防止泄漏
                Utils.ComReleaseHelper.Release(sandboxSheet)
                Utils.ComReleaseHelper.Release(originalSheet)
                Utils.ComReleaseHelper.Release(originalWb)

                System.Diagnostics.Debug.WriteLine("[Sandbox] 沙箱已清理，Excel状态已恢复")
            End Try
        End Function

        ''' <summary>
        ''' Execute Range/Cells value assignment operations via Interop
        ''' </summary>
        Private Shared Function ExecuteRangeOperations(excelApp As Microsoft.Office.Interop.Excel.Application, vbaCode As String) As String
            Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim activeCell As Object = Nothing
            Try
                Dim operations As New List(Of String)()
                ws = excelApp.ActiveSheet
                If ws Is Nothing Then Return "Error: No active worksheet"

                ' 1. Range("X").Value = Y (supports .Range / ws.Range prefix)
                For Each match As Match In Regex.Matches(vbaCode,
                    "(?:[a-zA-Z_]\w*\.)?\s*Range\s*\(\s*""([^""]*)""\s*\)\s*\.?\s*(?:Value|value)?\s*=\s*(.+?)(?:\r?\n|$)",
                    RegexOptions.IgnoreCase Or RegexOptions.Multiline)
                    If match.Groups.Count >= 3 Then
                        Dim rng As Object = Nothing
                        Try
                            rng = ws.Range(match.Groups(1).Value.Trim())
                            rng.Value = ParseValue(match.Groups(2).Value.Trim())
                            operations.Add($"Set {match.Groups(1).Value.Trim()}")
                        Catch : End Try
                        ComReleaseHelper.Release(rng)
                    End If
                Next

                ' 2. ActiveCell.Value = Y
                For Each match As Match In Regex.Matches(vbaCode,
                    "ActiveCell\s*\.?\s*(?:Value|value)?\s*=\s*(.+?)(?:\r?\n|$)",
                    RegexOptions.IgnoreCase Or RegexOptions.Multiline)
                    If match.Groups.Count >= 2 Then
                        Try
                            activeCell = excelApp.ActiveCell
                            activeCell.Value = ParseValue(match.Groups(1).Value.Trim())
                            operations.Add("Set ActiveCell")
                            Exit For
                        Catch : End Try
                    End If
                Next

                ' 3. Cells(r,c).Value = Y (supports .Cells / ws.Cells, literal row/col numbers)
                For Each match As Match In Regex.Matches(vbaCode,
                    "(?:[a-zA-Z_]\w*\.)?\s*Cells\s*\(\s*(\d+)\s*,\s*(\d+)\s*\)\s*\.?\s*(?:Value|value)?\s*=\s*(.+?)(?:\r?\n|$)",
                    RegexOptions.IgnoreCase Or RegexOptions.Multiline)
                    If match.Groups.Count >= 4 Then
                        Dim cell As Object = Nothing
                        Try
                            Dim r As Integer = CInt(match.Groups(1).Value)
                            Dim c As Integer = CInt(match.Groups(2).Value)
                            cell = ws.Cells(r, c)
                            cell.Value = ParseValue(match.Groups(3).Value.Trim())
                            operations.Add($"Set Cells({r},{c})")
                        Catch : End Try
                        ComReleaseHelper.Release(cell)
                    End If
                Next

                If operations.Count > 0 Then
                    Return "Execution completed: " & String.Join(", ", operations)
                End If
                Return "No executable Range/Cells operations found"
            Catch ex As Exception
                Return $"Range operation failed: {ex.Message}"
            Finally
                ComReleaseHelper.Release(activeCell)
                ComReleaseHelper.Release(ws)
            End Try
        End Function

        ''' <summary>
        ''' Execute Clear operations via Interop
        ''' </summary>
        Private Shared Function ExecuteClearOperations(excelApp As Microsoft.Office.Interop.Excel.Application, vbaCode As String) As String
            Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Try
                ws = excelApp.ActiveSheet
                If ws Is Nothing Then Return "Error: No active worksheet"
                Dim cl As String = vbaCode.ToLower()

                If cl.Contains("cells.clear") OrElse cl.Contains("usedrange.clear") Then
                    ws.Cells.ClearContents()
                    ws.Cells.ClearFormats()
                    Return "Cleared all content and formatting from sheet"
                End If

                For Each m As Match In Regex.Matches(vbaCode,
                    "(?:[a-zA-Z_]\w*\.)?\s*Range\s*\(\s*""([^""]*)""\s*\)\.ClearContents",
                    RegexOptions.IgnoreCase Or RegexOptions.Multiline)
                    If m.Success AndAlso m.Groups.Count >= 2 Then
                        Dim rng As Object = Nothing
                        Try
                            rng = ws.Range(m.Groups(1).Value.Trim())
                            rng.ClearContents()
                            Return $"Cleared range: {m.Groups(1).Value}"
                        Finally
                            ComReleaseHelper.Release(rng)
                        End Try
                    End If
                Next

                Return "No clear operations found"
            Catch ex As Exception
                Return $"Clear failed: {ex.Message}"
            Finally
                ComReleaseHelper.Release(ws)
            End Try
        End Function

        ''' <summary>
        ''' Handle InputBox - extract default value and write to ActiveCell
        ''' </summary>
        Private Shared Function ExecuteInputBoxOperation(excelApp As Microsoft.Office.Interop.Excel.Application, vbaCode As String) As String
            Dim activeCell As Object = Nothing
            Try
                Dim m As Match = Regex.Match(vbaCode, "InputBox\s*\(.+?,.+?,\s*""([^""]*)""", RegexOptions.IgnoreCase)
                If m.Success Then
                    Dim dv As String = m.Groups(1).Value
                    If Not String.IsNullOrWhiteSpace(dv) Then
                        Try
                            activeCell = excelApp.ActiveCell
                            activeCell.Value = dv
                            Return $"InputBox default written to ActiveCell: {dv}"
                        Catch ex As Exception  ' 
                            Return $"InputBox default: {dv}"
                        End Try
                    End If
                End If
                Return "InputBox detected but cannot auto-execute"
            Catch ex As Exception
                Return $"InputBox handling failed: {ex.Message}"
            Finally
                ComReleaseHelper.Release(activeCell)
            End Try
        End Function

        ''' <summary>
        ''' Parse value in VBA code
        ''' Supports: strings, numbers, dates, booleans, Nothing, and Excel built-in constants
        ''' Unsupported expressions (variables, function calls) are returned as original string with warning
        ''' </summary>
        Private Shared Function ParseValue(value As String) As Object
            value = value.Trim()

            ' Handle strings (with quotes) — also handle VBA "" escape inside
            If value.StartsWith("""") AndAlso value.EndsWith("""") Then
                Dim inner As String = value.Substring(1, value.Length - 2)
                ' Unescape VBA double-quote: "" → "
                inner = inner.Replace("""""", """")
                Return inner
            End If

            ' Handle Nothing / Null / Empty
            If value.Equals("Nothing", StringComparison.OrdinalIgnoreCase) OrElse
               value.Equals("Null", StringComparison.OrdinalIgnoreCase) OrElse
               value.Equals("Empty", StringComparison.OrdinalIgnoreCase) Then
                Return DBNull.Value
            End If

            ' Handle dates (#...)
            If value.Length >= 3 AndAlso value(0) = "#"c AndAlso value(value.Length - 1) = "#"c Then
                Dim dateStr As String = value.Substring(1, value.Length - 2)
                Dim resultDate As DateTime
                If DateTime.TryParse(dateStr, resultDate) Then
                    Return resultDate
                End If
                System.Diagnostics.Debug.WriteLine($"[ParseValue] 无法解析日期: {dateStr}")
            End If

            ' Handle numbers
            Dim resultInt As Integer
            If Integer.TryParse(value, resultInt) Then
                Return resultInt
            End If

            Dim resultLong As Long
            If Long.TryParse(value, resultLong) Then
                Return resultLong
            End If

            Dim resultDouble As Double
            If Double.TryParse(value, System.Globalization.NumberStyles.Float Or System.Globalization.NumberStyles.AllowThousands,
                System.Globalization.CultureInfo.InvariantCulture, resultDouble) Then
                Return resultDouble
            End If

            ' Handle booleans
            If value.Equals("true", StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
            If value.Equals("false", StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If

            ' Default: return as string (likely a variable reference or expression)
            System.Diagnostics.Debug.WriteLine($"[ParseValue] 无法解析为字面量，作为字符串返回: {value.Substring(0, Math.Min(50, value.Length))}")
            Return value
        End Function

        ' ============================================================
        ' 调试执行：一次性执行完整代码并记录每行执行后的变量状态
        ' 【安全修复】DebugExecute 不再绕过安全检查，所有代码均需通过 CheckCode

    End Class

End Namespace
