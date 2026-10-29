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

    ' VBACodeSandbox.Debug.vb — 调试执行 / 变量注入日志 / 读取调试日志
    ' （类级 XML 文档只在主文件 VBACodeSandbox.vb 保留一份，避免 BC42314 分部重复）
    Partial Public Class VBACodeSandbox

        ' ============================================================
        Public Shared Function DebugExecute(vbaCode As String) As Object
            Try
                Dim excelApp As Microsoft.Office.Interop.Excel.Application = Globals.ThisAddIn.Application
                If excelApp Is Nothing Then
                    Return New With {.success = False, .error = "Excel 应用程序不可用"}
                End If

                ' 【安全】先进行安全检查，不安全的代码拒绝调试执行
                Dim checkResult = CheckCode(vbaCode)
                If Not checkResult.IsSafe Then
                    Return New With {
                        .success = False,
                        .error = "安全检查未通过，拒绝调试执行",
                        .riskLevel = checkResult.RiskLevel.ToString(),
                        .errors = checkResult.Errors,
                        .warnings = checkResult.Warnings
                    }
                End If

                ' 对于 High/Medium 风险代码（IsSafe=True 但有警告），要求用户确认
                If checkResult.RiskLevel >= RiskLevel.Medium Then
                    Dim confirmMsg = String.Format(
                        "调试执行请求 — 代码风险级别: {0}{1}警告数: {2}{3}是否继续？",
                        checkResult.RiskLevel.ToString(), vbCrLf,
                        checkResult.Warnings.Count, vbCrLf &
                        "注意：调试模式将在当前工作簿中直接执行代码。")
                    Dim confirmResult = System.Windows.Forms.MessageBox.Show(
                        confirmMsg,
                        "调试执行安全确认",
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        System.Windows.Forms.MessageBoxIcon.Warning)
                    If confirmResult <> System.Windows.Forms.DialogResult.Yes Then
                        Return New With {.success = False, .error = "用户取消了调试执行"}
                    End If
                End If

                ' 1. 提取变量名
                Dim varNames = ExtractVariableNames(vbaCode)
                System.Diagnostics.Debug.WriteLine($"[DebugExecute] 发现变量: {String.Join(", ", varNames)}")

                ' 2. 改写代码：注入调试日志
                Dim modifiedCode = InjectDebugLogging(vbaCode, varNames)
                System.Diagnostics.Debug.WriteLine($"[DebugExecute] 改写后代码长度: {modifiedCode.Length}")

                ' 3. 执行改写后的代码
                Dim execResult = ExecuteViaRealVBAEngine(excelApp, modifiedCode)
                System.Diagnostics.Debug.WriteLine($"[DebugExecute] 执行结果: {execResult}")

                If execResult IsNot Nothing AndAlso execResult.StartsWith("✗") Then
                    Return New With {.success = False, .error = "VBA 执行失败: " & execResult}
                End If

                ' 4. 读取调试日志
                Dim log = ReadDebugLog(excelApp)

                ' 5. 清理隐藏工作表
                Try
                    excelApp.ActiveWorkbook.Worksheets("__AIDebugLog").Delete()
                Catch : End Try

                Return New With {
                    .success = True,
                    .log = log,
                    .executionMethod = "DebugEngine",
                    .executedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    .rawResult = execResult
                }

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[DebugExecute] 异常: {ex.Message}")
                Return New With {.success = False, .error = ex.Message}
            End Try
        End Function

        ' 从代码中提取 Dim 声明的变量名
        Private Shared Function ExtractVariableNames(code As String) As List(Of String)
            Dim result As New List(Of String)()
            Dim lines = code.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None)
            For Each line In lines
                Dim m = Regex.Match(line, "^\s*Dim\s+(\w+)", RegexOptions.IgnoreCase)
                If m.Success AndAlso Not result.Contains(m.Groups(1).Value) Then
                    result.Add(m.Groups(1).Value)
                End If
            Next
            Return result
        End Function

        ' 判断一行是否应该记录日志
        Private Shared Function ShouldLogLine(line As String) As Boolean
            Dim t = line.Trim()
            If String.IsNullOrEmpty(t) Then Return False
            If t.StartsWith("'") Then Return False
            If Regex.IsMatch(t, "^(Dim|Const|Sub|Function|End Sub|End Function|Exit Sub|Exit Function|On Error|Option |Declare |Type |End Type)", RegexOptions.IgnoreCase) Then Return False
            Return True
        End Function

        ' 在代码中注入调试日志语句
        Private Shared Function InjectDebugLogging(code As String, varNames As List(Of String)) As String
            Dim lines = code.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None)
            Dim sb As New StringBuilder()
            Dim insideSub As Boolean = False
            Dim baseIndent As String = ""

            For i = 0 To lines.Length - 1
                Dim line = lines(i)
                Dim trimmed = line.Trim()

                ' 检测 Sub/Function 开始（支持 Public/Private/Friend/Static 修饰符）
                If Regex.IsMatch(trimmed, "^(Public|Private|Friend|Static)?\s*(Sub|Function)\s+\w+", RegexOptions.IgnoreCase) Then
                    sb.AppendLine(line)
                    insideSub = True
                    baseIndent = "    "
                    ' 插入调试初始化
                    sb.AppendLine(baseIndent & "On Error GoTo __DebugError")
                    sb.AppendLine(baseIndent & "Dim __DebugSheet As Object, __DebugRow As Long, __DebugVars As String")
                    sb.AppendLine(baseIndent & "Application.DisplayAlerts = False")
                    sb.AppendLine(baseIndent & "On Error Resume Next")
                    sb.AppendLine(baseIndent & "ActiveWorkbook.Worksheets(""__AIDebugLog"").Delete")
                    sb.AppendLine(baseIndent & "Set __DebugSheet = ActiveWorkbook.Worksheets.Add")
                    sb.AppendLine(baseIndent & "__DebugSheet.Name = ""__AIDebugLog""")
                    sb.AppendLine(baseIndent & "__DebugSheet.Visible = 2")
                    sb.AppendLine(baseIndent & "Application.DisplayAlerts = True")
                    sb.AppendLine(baseIndent & "__DebugRow = 1")
                    sb.AppendLine(baseIndent & "On Error GoTo __DebugError")
                    Continue For
                End If

                ' 检测 End Sub/Function：先插入错误处理标签，再写 End Sub
                If Regex.IsMatch(trimmed, "^End\s+(Sub|Function)", RegexOptions.IgnoreCase) Then
                    sb.AppendLine(baseIndent & "__DebugError:")
                    sb.AppendLine(baseIndent & "__DebugSheet.Cells(__DebugRow, 1).Value = ""ERROR""")
                    sb.AppendLine(baseIndent & "__DebugSheet.Cells(__DebugRow, 2).Value = Err.Number")
                    sb.AppendLine(baseIndent & "__DebugSheet.Cells(__DebugRow, 3).Value = Err.Description")
                    sb.AppendLine(baseIndent & "__DebugRow = __DebugRow + 1")
                    sb.AppendLine(baseIndent & "Resume Next")
                    sb.AppendLine(line)
                    insideSub = False
                    Continue For
                End If

                ' 默认：追加原始行
                sb.AppendLine(line)

                ' 在可执行行后插入日志
                If insideSub AndAlso ShouldLogLine(trimmed) Then
                    Dim escaped = trimmed.Replace("""", """""")
                    sb.AppendLine(baseIndent & $"__DebugSheet.Cells(__DebugRow, 1).Value = {i + 1}")
                    sb.AppendLine(baseIndent & $"__DebugSheet.Cells(__DebugRow, 2).Value = ""{escaped}""")
                    If varNames.Count > 0 Then
                        sb.AppendLine(baseIndent & "__DebugVars = """"")
                        sb.AppendLine(baseIndent & "On Error Resume Next")
                        For Each vn In varNames
                            sb.AppendLine(baseIndent & $"__DebugVars = __DebugVars & ""{vn}="" & {vn} & ""|""")
                        Next
                        sb.AppendLine(baseIndent & "On Error GoTo __DebugError")
                        sb.AppendLine(baseIndent & "__DebugSheet.Cells(__DebugRow, 3).Value = __DebugVars")
                    End If
                    sb.AppendLine(baseIndent & "__DebugRow = __DebugRow + 1")
                End If
            Next

            ' 如果代码没有 Sub/Function，自动包装
            If Not insideSub Then
                Dim wrapped As New StringBuilder()
                wrapped.AppendLine("Sub __AIDebugMain()")
                wrapped.AppendLine("    On Error GoTo __DebugError")
                wrapped.AppendLine("    Dim __DebugSheet As Object, __DebugRow As Long, __DebugVars As String")
                wrapped.AppendLine("    Application.DisplayAlerts = False")
                wrapped.AppendLine("    On Error Resume Next")
                wrapped.AppendLine("    ActiveWorkbook.Worksheets(""__AIDebugLog"").Delete")
                wrapped.AppendLine("    Set __DebugSheet = ActiveWorkbook.Worksheets.Add")
                wrapped.AppendLine("    __DebugSheet.Name = ""__AIDebugLog""")
                wrapped.AppendLine("    __DebugSheet.Visible = 2")
                wrapped.AppendLine("    Application.DisplayAlerts = True")
                wrapped.AppendLine("    __DebugRow = 1")
                wrapped.AppendLine("    On Error GoTo __DebugError")

                For Each ln In lines
                    wrapped.AppendLine("    " & ln)
                    If ShouldLogLine(ln.Trim()) Then
                        Dim escaped = ln.Trim().Replace("""", """""")
                        wrapped.AppendLine("    " & $"__DebugSheet.Cells(__DebugRow, 1).Value = 0")
                        wrapped.AppendLine("    " & $"__DebugSheet.Cells(__DebugRow, 2).Value = ""{escaped}""")
                        If varNames.Count > 0 Then
                            wrapped.AppendLine("    " & "__DebugVars = """"")
                            wrapped.AppendLine("    " & "On Error Resume Next")
                            For Each vn In varNames
                                wrapped.AppendLine("    " & $"__DebugVars = __DebugVars & ""{vn}="" & {vn} & ""|""")
                            Next
                            wrapped.AppendLine("    " & "On Error GoTo __DebugError")
                            wrapped.AppendLine("    " & "__DebugSheet.Cells(__DebugRow, 3).Value = __DebugVars")
                        End If
                        wrapped.AppendLine("    " & "__DebugRow = __DebugRow + 1")
                    End If
                Next

                wrapped.AppendLine("__DebugError:")
                wrapped.AppendLine("    __DebugSheet.Cells(__DebugRow, 1).Value = ""ERROR""")
                wrapped.AppendLine("    __DebugSheet.Cells(__DebugRow, 2).Value = Err.Number")
                wrapped.AppendLine("    __DebugSheet.Cells(__DebugRow, 3).Value = Err.Description")
                wrapped.AppendLine("    __DebugRow = __DebugRow + 1")
                wrapped.AppendLine("    Resume Next")
                wrapped.AppendLine("End Sub")
                Return wrapped.ToString()
            End If

            Return sb.ToString()
        End Function

        ' 从 Excel 读取调试日志工作表
        Private Shared Function ReadDebugLog(excelApp As Microsoft.Office.Interop.Excel.Application) As List(Of Object)
            Dim result As New List(Of Object)()
            Dim wb As Object = Nothing
            Dim ws As Object = Nothing
            Try
                wb = excelApp.ActiveWorkbook
                If wb Is Nothing Then Return result

                Dim worksheets As Object = Nothing
                Try
                    worksheets = wb.Worksheets
                    If worksheets IsNot Nothing Then
                        Dim sheetCount As Integer = worksheets.Count
                        For i As Integer = 1 To sheetCount
                            Dim sheet As Object = Nothing
                            Try
                                sheet = worksheets(i)
                                If sheet IsNot Nothing AndAlso sheet.Name = "__AIDebugLog" Then
                                    ws = sheet
                                    sheet = Nothing
                                    Exit For
                                End If
                            Finally
                                ComReleaseHelper.Release(sheet)
                            End Try
                        Next
                    End If
                Finally
                    ComReleaseHelper.Release(worksheets)
                End Try

                If ws Is Nothing Then Return result

                Dim lastRow As Integer = 0
                Try
                    lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row
                Catch ex As Exception  ' 
                    lastRow = 0
                End Try

                If lastRow <= 0 Then Return result

                For r = 1 To lastRow
                    Dim cell1 As Object = Nothing
                    Dim cell2 As Object = Nothing
                    Dim cell3 As Object = Nothing
                    Try
                        cell1 = ws.Cells(r, 1)
                        cell2 = ws.Cells(r, 2)
                        cell3 = ws.Cells(r, 3)

                        Dim lineNumObj = cell1.Value
                        Dim codeObj = cell2.Value
                        Dim varsObj = cell3.Value

                        Dim lineNum As Integer = 0
                        If lineNumObj IsNot Nothing Then
                            Integer.TryParse(lineNumObj.ToString(), lineNum)
                        End If

                        Dim code As String = If(codeObj IsNot Nothing, codeObj.ToString(), "")
                        Dim vars As String = If(varsObj IsNot Nothing, varsObj.ToString(), "")

                        result.Add(New With {
                            .lineNumber = lineNum,
                            .code = code,
                            .variables = ParseDebugVariableString(vars)
                        })
                    Finally
                        ComReleaseHelper.Release(cell3)
                        ComReleaseHelper.Release(cell2)
                        ComReleaseHelper.Release(cell1)
                    End Try
                Next

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[ReadDebugLog] 异常: {ex.Message}")
            Finally
                ComReleaseHelper.Release(ws)
                ComReleaseHelper.Release(wb)
            End Try
            Return result
        End Function

        ' 解析 "var1=value1|var2=value2|" 格式的变量字符串
        Private Shared Function ParseDebugVariableString(varString As String) As Dictionary(Of String, String)
            Dim result As New Dictionary(Of String, String)()
            If String.IsNullOrEmpty(varString) Then Return result

            Dim pairs = varString.Split("|"c)
            For Each pair In pairs
                If String.IsNullOrWhiteSpace(pair) Then Continue For
                Dim eqIdx = pair.IndexOf("=")
                If eqIdx > 0 Then
                    Dim name = pair.Substring(0, eqIdx).Trim()
                    Dim value = pair.Substring(eqIdx + 1).Trim()
                    result(name) = value
                End If
            Next
            Return result
        End Function

    End Class

End Namespace
