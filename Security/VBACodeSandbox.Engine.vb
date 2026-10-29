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

    ' VBACodeSandbox.Engine.vb — VBA 引擎分发（Interop/RealVBAEngine）+ MsgBox/InputBox/OnError 重写
    ' （类级 XML 文档只在主文件 VBACodeSandbox.vb 保留一份，避免 BC42314 分部重复）
    Partial Public Class VBACodeSandbox

        Private Shared Function ExecuteViaVBAEngine(excelApp As Microsoft.Office.Interop.Excel.Application, vbaCode As String) As String
            Try
                ' 直接 VBProject Import+Run，保持原始代码不变
                Return ExecuteViaRealVBAEngine(excelApp, vbaCode)
            Catch ex As Exception
                Return $"✗ {ex.Message}"
            End Try
        End Function

        ''' <summary>
        ''' Interop 直接处理常见 VBA 模式，不依赖 VBProject。
        ''' 返回 "NO_MATCH" 表示未匹配到，继续走 VBAEngine。
        ''' </summary>
        Private Shared Function ExecuteViaInteropPatterns(
            excelApp As Microsoft.Office.Interop.Excel.Application,
            vbaCode As String) As String
            Dim cl As String = vbaCode.ToLower()

            ' ----- 模式1：FileSystemObject 列出文件并写入工作表 -----
            If cl.Contains("scripting.filesystemobject") OrElse cl.Contains("filesystemobject") Then
                Dim fsoResult = ExecuteFileListingInterop(excelApp, vbaCode)
                If Not String.IsNullOrEmpty(fsoResult) Then Return fsoResult
            End If

            ' ----- 模式2：简单 Clear -----
            If cl.Contains(".clear") AndAlso Not cl.Contains("fso") AndAlso Not cl.Contains("folder") Then
                Dim cr = ExecuteClearOperations(excelApp, vbaCode)
                If Not cr.Contains("No clear") Then Return cr
            End If

            ' ----- 模式3：Range/Cells 写入 -----
            If (cl.Contains("range(") OrElse cl.Contains("cells(")) AndAlso cl.Contains(".value") Then
                Return ExecuteRangeOperations(excelApp, vbaCode)
            End If

            Return "NO_MATCH"
        End Function

        ''' <summary>
        ''' 直接通过 .NET System.IO 列出文件并写入 Excel，无需 VBProject 或 VBA。
        ''' 提取 VBA 中的 folder 路径和输出 sheet 名称。
        ''' </summary>
        Private Shared Function ExecuteFileListingInterop(
            excelApp As Microsoft.Office.Interop.Excel.Application,
            vbaCode As String) As String
            Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
            Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Try
                ' 1. 提取文件夹路径：fso.GetFolder("X:\") 或 GetFolder("X:\\")
                Dim folderPath As String = Nothing
                Dim folderMatch As Match = Regex.Match(vbaCode,
                    "\.GetFolder\s*\(\s*""([^""]+)""\s*\)",
                    RegexOptions.IgnoreCase)
                If Not folderMatch.Success Then Return Nothing
                folderPath = folderMatch.Groups(1).Value.Trim()

                ' 验证路径
                If Not System.IO.Directory.Exists(folderPath) Then
                    Return $"✗ 文件夹不存在: {folderPath}"
                End If

                ' 2. 提取新建的工作表名称：ws.Name = "XXX"
                Dim sheetName As String = "文件列表"
                Dim sheetNameMatch As Match = Regex.Match(vbaCode,
                    "\.Name\s*=\s*""([^""]+)""",
                    RegexOptions.IgnoreCase)
                If sheetNameMatch.Success Then
                    sheetName = sheetNameMatch.Groups(1).Value.Trim()
                End If

                ' 3. 提取列标题（查找 .Cells(1, N).Value = "XXX" 模式）
                Dim headers As New List(Of String)()
                Dim headerMatches As MatchCollection = Regex.Matches(vbaCode,
                    "\.Cells\s*\(\s*1\s*,\s*(\d+)\s*\)\s*\.\s*(?:Value|value)\s*=\s*""([^""]+)""",
                    RegexOptions.IgnoreCase)
                For Each hm As Match In headerMatches
                    If Not headers.Contains(hm.Groups(2).Value) Then
                        headers.Add(hm.Groups(2).Value)
                    End If
                Next
                If headers.Count = 0 Then headers.Add("文件名")

                ' 4. 创建新工作表
                wb = excelApp.ActiveWorkbook
                If wb Is Nothing Then Return "✗ 没有打开的工作簿"

                Dim sheets As Object = Nothing
                Try
                    sheets = wb.Sheets
                    If sheets IsNot Nothing Then
                        Dim sheetCount As Integer = sheets.Count
                        For i As Integer = 1 To sheetCount
                            Dim s As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                            Try
                                s = CType(sheets(i), Microsoft.Office.Interop.Excel.Worksheet)
                                If s IsNot Nothing AndAlso s.Name = sheetName Then
                                    Dim originalAlerts As Boolean = excelApp.DisplayAlerts
                                    Try
                                        excelApp.DisplayAlerts = False
                                        s.Delete()
                                    Finally
                                        excelApp.DisplayAlerts = originalAlerts
                                    End Try
                                    Exit For
                                End If
                            Finally
                                ComReleaseHelper.Release(s)
                            End Try
                        Next
                    End If

                    ws = CType(wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count)), Microsoft.Office.Interop.Excel.Worksheet)
                    ws.Name = sheetName
                Catch nameEx As Exception
                    ' 名称冲突，用带时间戳的名字
                    Try
                        sheetName = sheetName & "_" & DateTime.Now.ToString("HHmmss")
                        ws.Name = sheetName
                    Catch : End Try
                Finally
                    ComReleaseHelper.Release(sheets)
                End Try

                If ws Is Nothing Then Return "✗ 无法创建工作表"

                ' 5. 写表头
                For col As Integer = 0 To headers.Count - 1
                    Dim cell As Object = Nothing
                    Try
                        cell = ws.Cells(1, col + 1)
                        cell.Value = headers(col)
                    Finally
                        ComReleaseHelper.Release(cell)
                    End Try
                Next

                ' 6. 获取文件列表
                Dim files As String() = System.IO.Directory.GetFiles(folderPath)
                Dim fileCount As Integer = files.Length
                Const MAX_FILES As Integer = 5000 ' 防止性能问题
                Dim count As Integer = Math.Min(fileCount, MAX_FILES)

                For i As Integer = 0 To count - 1
                    Dim fi As New System.IO.FileInfo(files(i))
                    Dim cell1 As Object = Nothing
                    Dim cell2 As Object = Nothing
                    Try
                        cell1 = ws.Cells(i + 2, 1)
                        cell1.Value = fi.Name
                        If headers.Count >= 2 Then
                            cell2 = ws.Cells(i + 2, 2)
                            cell2.Value = fi.FullName
                        End If
                    Finally
                        ComReleaseHelper.Release(cell2)
                        ComReleaseHelper.Release(cell1)
                    End Try
                Next

                ' 7. 自动调整列宽
                For col As Integer = 1 To headers.Count
                    Dim colRange As Object = Nothing
                    Try
                        colRange = ws.Columns(col)
                        CType(colRange, Microsoft.Office.Interop.Excel.Range).AutoFit()
                    Finally
                        ComReleaseHelper.Release(colRange)
                    End Try
                Next

                ' 8. 激活新工作表
                ws.Activate()

                Dim msg As String = $"✓ 已在工作表「{sheetName}」列出 {count} 个文件"
                If fileCount > MAX_FILES Then
                    msg &= $"（共 {fileCount} 个文件，仅显示前 {MAX_FILES} 个）"
                End If
                Return msg

            Catch ex As Exception
                Return $"✗ {ex.Message}"
            Finally
                ComReleaseHelper.Release(ws)
                ComReleaseHelper.Release(wb)
            End Try
        End Function

        ''' <summary>
        ''' 真正执行 VBA：Import 模块到 VBProject → Application.Run → 清理
        ''' 不做 On Error Resume Next 替换 —— 只在沙箱里做安全化处理，
        ''' VBAEngine 直接在当前工作簿真执行，尊重原始代码逻辑
        ''' </summary>
        Private Shared Function ExecuteViaRealVBAEngine(excelApp As Microsoft.Office.Interop.Excel.Application, vbaCode As String) As String
            Dim wb As Object = Nothing
            Dim vbComp As Object = Nothing
            Dim tempBasFile As String = ""
            Dim lastCell As Object = Nothing

            ' 保存原始状态
            Dim originalDisplayAlerts As Boolean = excelApp.DisplayAlerts
            Dim originalScreenUpdating As Boolean = excelApp.ScreenUpdating

            Try
                ' 抑制 VBA 错误对话框和屏幕更新
                excelApp.DisplayAlerts = False
                excelApp.ScreenUpdating = False

                wb = excelApp.ActiveWorkbook
                If wb Is Nothing Then Return "✗ 没有打开的工作簿"

                Dim macroName As String = ExtractMacroName(vbaCode)
                If String.IsNullOrEmpty(macroName) Then
                    macroName = "AIExec" & (DateTime.Now.Ticks Mod 10000).ToString()
                    vbaCode = "Sub " & macroName & "()" & vbCrLf & vbaCode & vbCrLf & "End Sub"
                End If
                ' VBAEngine 路径：直接执行原始代码，不做任何修改
                ' 写入 .bas
                Dim tempModuleName As String = "TMP" & (DateTime.Now.Ticks Mod 10000).ToString()
                tempBasFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), tempModuleName & ".bas")
                System.IO.File.WriteAllText(tempBasFile,
                    "Attribute VB_Name = """ & tempModuleName & """" & vbCrLf & vbaCode,
                    System.Text.Encoding.Default)

                System.Diagnostics.Debug.WriteLine($"[VBA] bas={tempBasFile}, macro={macroName}, 编码=ANSI")

                ' Step 1: VBProject
                Dim vbProj As Object = Nothing
                Dim vbComps As Object = Nothing
                Try
                    vbProj = wb.VBProject
                    If vbProj Is Nothing Then Return "✗ 无法访问 VBProject。请在 Excel 启用: 文件→选项→信任中心→信任中心设置→宏设置→勾选'信任对VBA项目对象模型的访问'"

                    ' Step 2: Import
                    vbComps = vbProj.VBComponents
                    Try
                        vbComp = vbComps.Import(tempBasFile)
                    Catch iex As Exception
                        Return "✗ 导入失败: " & iex.Message
                    End Try
                    If vbComp Is Nothing Then Return "✗ 导入返回空——代码可能包含语法错误"

                    ' Step 3: 验证
                    Dim lines As Integer = 0
                    Try : lines = vbComp.CodeModule.CountOfLines : Catch : End Try
                    System.Diagnostics.Debug.WriteLine($"[VBA] 导入成功, {lines} 行")
                Finally
                    Utils.ComReleaseHelper.Release(vbComps)
                    Utils.ComReleaseHelper.Release(vbProj)
                End Try

                ' Step 4: Run（模块限定的宏名: ModuleName.MacroName）
                Dim fullMacro As String = tempModuleName & "." & macroName
                System.Diagnostics.Debug.WriteLine($"[VBA] Run: {fullMacro}")

                ' 使用最后一格作为调试输出通道，避免干扰用户数据
                Dim originalValue As Object = Nothing
                Try
                    lastCell = excelApp.ActiveSheet.Cells(excelApp.ActiveSheet.Rows.Count, excelApp.ActiveSheet.Columns.Count)
                    originalValue = lastCell.Value
                Catch : End Try

                excelApp.ScreenUpdating = True
                System.Windows.Forms.Application.DoEvents()
                System.Threading.Thread.Sleep(100)
                excelApp.Run(fullMacro)

                System.Threading.Thread.Sleep(300)
                excelApp.ScreenUpdating = True

                ' 读取调试输出
                Dim debugOutput As Object = Nothing
                Try
                    debugOutput = lastCell.Value
                Catch : End Try

                ' 恢复原始值
                Try
                    If originalValue IsNot Nothing Then
                        lastCell.Value = originalValue
                    Else
                        lastCell.ClearContents()
                    End If
                Catch : End Try

                If debugOutput IsNot Nothing AndAlso debugOutput.ToString().Length > 0 Then
                    Return "✓ 执行成功。DebugOutput: " & debugOutput.ToString()
                End If

                Return "✓ 执行成功（未检测到输出）"

            Catch ex As Exception
                Return "✗ 异常: " & ex.Message
            Finally
                ' 恢复 Excel 状态
                excelApp.DisplayAlerts = originalDisplayAlerts
                excelApp.ScreenUpdating = originalScreenUpdating

                If Not String.IsNullOrEmpty(tempBasFile) AndAlso System.IO.File.Exists(tempBasFile) Then
                    Try : System.IO.File.Delete(tempBasFile) : Catch : End Try
                End If
                Try
                    If vbComp IsNot Nothing AndAlso wb IsNot Nothing Then
                        Dim p As Object = Nothing
                        Dim comps As Object = Nothing
                        Try
                            p = wb.VBProject
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

                Utils.ComReleaseHelper.Release(lastCell)
                Utils.ComReleaseHelper.Release(wb)
            End Try
        End Function

        ''' <summary>
        ''' 从 VBA 代码中提取宏名
        ''' </summary>
        Private Shared Function ExtractMacroName(vbaCode As String) As String
            If String.IsNullOrEmpty(vbaCode) Then Return Nothing
            ' 过程可以不写括号（Sub Main），名字也可以是中文。
            ' 旧正则要求 Sub Name( ，匹配失败后会把整段（含 Option Explicit）再包进一个 Sub，模块无法编译。
            Dim m = Regex.Match(vbaCode, "(?im)^[ \t]*(?:(?:Public|Private)[ \t]+)?(?:Sub|Function)[ \t]+([^\s(]+)")
            If m.Success Then Return m.Groups(1).Value.Trim()
            Return Nothing
        End Function

        ''' <summary>
        ''' 沙箱预处理：将 MsgBox/InputBox 替换为 Range("A1")/Range("A2") 写入
        ''' 使得沙箱执行后用户可以直观看到输出结果（而非"执行成功"后无效果）
        ''' 
        ''' 使用逐行处理，避免正则跨行匹配带来的复杂性和错误
        ''' 支持：MsgBox "text", MsgBox varName, MsgBox(expr), MsgBox(expr, ...)
        ''' </summary>
        Private Shared Function ReplaceMsgBoxWithCellWrite(vbaCode As String) As String
            If String.IsNullOrEmpty(vbaCode) Then Return vbaCode

            Dim hasMsgBox As Boolean = vbaCode.IndexOf("MsgBox", StringComparison.OrdinalIgnoreCase) >= 0
            Dim hasInputBox As Boolean = vbaCode.IndexOf("InputBox", StringComparison.OrdinalIgnoreCase) >= 0
            If Not hasMsgBox AndAlso Not hasInputBox Then Return vbaCode

            Dim lines As String() = vbaCode.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None)
            Dim resultLines As New List(Of String)(lines.Length)
            Dim modified As Boolean = False

            For Each line As String In lines
                Dim processed As String = line

                If processed.IndexOf("MsgBox", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    processed = ReplaceMsgBoxLine(processed)
                    If Not String.Equals(processed, line, StringComparison.Ordinal) Then modified = True
                End If

                If processed.IndexOf("InputBox", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    processed = ReplaceInputBoxLine(processed)
                    If Not String.Equals(processed, line, StringComparison.Ordinal) Then modified = True
                End If

                resultLines.Add(processed)
            Next

            If modified Then
                System.Diagnostics.Debug.WriteLine("[VBACodeSandbox] MsgBox/InputBox 已替换为 Debug.Print（禁止写入 A1/A2 单元格，防止破坏用户表头数据）")
            End If

            Return String.Join(vbCrLf, resultLines)
        End Function

        ''' <summary>
        ''' 单行 MsgBox 替换：将 MsgBox 调用替换为 Range("A1").Value = ...
        ''' 处理：MsgBox "text", MsgBox var, MsgBox(...)
        ''' </summary>
        Private Shared Function ReplaceMsgBoxLine(line As String) As String
            ' 1. MsgBox "literal string" [, ...] → 最常见情形
            Dim m1 As Match = Regex.Match(line,
                "MsgBox\s+(""[^""]*"")",
                RegexOptions.IgnoreCase)
            If m1.Success Then
                Dim replacement As String = 
                    "' [Sandbox: MsgBox safe] 禁止写入 A1，避免破坏用户表头（例如『表格美化完成』覆盖原数据）" & vbCrLf &
                    "On Error Resume Next: Debug.Print CStr(" & m1.Groups(1).Value & "): On Error GoTo 0"
                Return Regex.Replace(line,
                    "MsgBox\s+""[^""]*""[^']*",
                    replacement,
                    RegexOptions.IgnoreCase)
            End If

            ' 2. MsgBox variableName (without quotes, without parentheses)
            Dim m2 As Match = Regex.Match(line,
                "MsgBox\s+([A-Za-z_]\w*)\s*(?:'|$)",
                RegexOptions.IgnoreCase)
            If m2.Success Then
                Dim vName As String = m2.Groups(1).Value
                Return Regex.Replace(line,
                    "MsgBox\s+" & Regex.Escape(vName),
                    "' [Sandbox: MsgBox safe] 禁止写入 A1，避免破坏用户表头" & vbCrLf &
                    "On Error Resume Next: Debug.Print CStr(" & vName & "): On Error GoTo 0",
                    RegexOptions.IgnoreCase)
            End If

            ' 3. MsgBox(expr) or MsgBox(expr, ...) — parentheses form
            '    使用括号平衡提取第一个参数
            If line.IndexOf("MsgBox(", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               line.IndexOf("MsgBox (", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Dim firstArg As String = ExtractFirstMsgBoxArgument(line)
                If Not String.IsNullOrEmpty(firstArg) Then
                    Return Regex.Replace(line,
                        "MsgBox\s*\(.*",
                        "' [Sandbox: MsgBox safe] 禁止写入 A1，避免破坏用户表头" & vbCrLf &
                        "On Error Resume Next: Debug.Print CStr(" & firstArg & "): On Error GoTo 0",
                        RegexOptions.IgnoreCase)
                End If
            End If

            Return line
        End Function

        ''' <summary>
        ''' 从 MsgBox(expr, ...) 中提取第一个参数（支持嵌套括号和字符串）
        ''' </summary>
        Private Shared Function ExtractFirstMsgBoxArgument(line As String) As String
            ' Find "MsgBox(" or "MsgBox (" position
            Dim idx As Integer = line.IndexOf("MsgBox(", StringComparison.OrdinalIgnoreCase)
            If idx < 0 Then idx = line.IndexOf("MsgBox (", StringComparison.OrdinalIgnoreCase)
            If idx < 0 Then Return Nothing

            ' Find opening paren
            Dim parenIdx As Integer = line.IndexOf("("c, idx)
            If parenIdx < 0 Then Return Nothing

            ' Extract first arg using bracket balancing
            Dim depth As Integer = 0
            Dim inString As Boolean = False
            Dim start As Integer = parenIdx + 1
            Dim i As Integer = start

            While i < line.Length
                Dim c As Char = line(i)

                If c = """"c Then
                    ' Toggle string state (handle VBA "" escape)
                    If i + 1 < line.Length AndAlso line(i + 1) = """"c Then
                        i += 1 ' skip escaped quote
                    Else
                        inString = Not inString
                    End If
                ElseIf Not inString Then
                    If c = "("c Then
                        depth += 1
                    ElseIf c = ")"c Then
                        If depth = 0 Then
                            ' Closing paren of MsgBox — extract up to first top-level comma or this paren
                            Dim arg As String = line.Substring(start, i - start).Trim()
                            ' If arg contains commas at the top level, we take only up to first comma
                            Dim commaIdx As Integer = FindFirstTopLevelComma(arg)
                            If commaIdx > 0 Then
                                arg = arg.Substring(0, commaIdx).Trim()
                            End If
                            Return arg
                        End If
                        depth -= 1
                    ElseIf c = ","c AndAlso depth = 0 Then
                        ' Top-level comma — first argument ends here
                        Dim arg As String = line.Substring(start, i - start).Trim()
                        Return arg
                    End If
                End If
                i += 1
            End While

            Return Nothing
        End Function

        ''' <summary>
        ''' 查找字符串中第一个不在嵌套括号或引号内的逗号位置
        ''' </summary>
        Private Shared Function FindFirstTopLevelComma(s As String) As Integer
            Dim depth As Integer = 0
            Dim inString As Boolean = False
            For i As Integer = 0 To s.Length - 1
                Dim c As Char = s(i)
                If c = """"c Then
                    If i + 1 < s.Length AndAlso s(i + 1) = """"c Then
                        i += 1
                    Else
                        inString = Not inString
                    End If
                ElseIf Not inString Then
                    If c = "("c Then
                        depth += 1
                    ElseIf c = ")"c Then
                        depth -= 1
                    ElseIf c = ","c AndAlso depth = 0 Then
                        Return i
                    End If
                End If
            Next
            Return -1
        End Function

        ''' <summary>
        ''' 单行 InputBox 替换
        ''' </summary>
        Private Shared Function ReplaceInputBoxLine(line As String) As String
            ' Simple case: InputBox(prompt) → write prompt to A1 as info
            Dim m As Match = Regex.Match(line,
                "InputBox\s*\(\s*(""[^""]*"")\s*\)",
                RegexOptions.IgnoreCase)
            If m.Success Then
                Return "' [Sandbox: InputBox safe] 禁止写入 A1，避免破坏用户表头" & vbCrLf &
                       "On Error Resume Next: Debug.Print ""[InputBox] "" & CStr(" & m.Groups(1).Value & "): On Error GoTo 0 " & vbCrLf &
                       "' End Sandbox replacement"
            End If
            Return line
        End Function

        ''' <summary>
        ''' 安全化 VBA 代码：将 "On Error Resume Next" 替换为 "On Error GoTo 0"
        ''' 防止安全错误在 VBA 执行过程中被静默忽略，这是沙箱逃逸的常见手法
        ''' 同时移除 "On Error GoTo 0" 后面的任何自定义错误处理标签跳转
        ''' </summary>
        Private Shared Function SanitizeVBAOnErrorResumeNext(vbaCode As String) As String
            If String.IsNullOrEmpty(vbaCode) Then Return vbaCode

            ' 替换 On Error Resume Next（不区分大小写）为 On Error GoTo 0
            ' On Error GoTo 0 表示禁用错误处理，错误将正常抛出
            Dim result As String = Regex.Replace(
                vbaCode,
                "On\s+Error\s+Resume\s+Next",
                "On Error GoTo 0",
                RegexOptions.IgnoreCase)

            ' 同时将 On Error GoTo <label> 替换为 On Error GoTo 0（防止跳转到自定义错误处理块）
            result = Regex.Replace(
                result,
                "On\s+Error\s+GoTo\s+(?!0\b)(\w+)",
                "On Error GoTo 0",
                RegexOptions.IgnoreCase)

            If Not String.Equals(vbaCode, result, StringComparison.Ordinal) Then
                System.Diagnostics.Debug.WriteLine("[VBACodeSandbox] On Error Resume Next/GoTo 已被替换为 On Error GoTo 0")
            End If

            Return result
        End Function

        ' 【历史备注】直接在用户活跃工作簿中执行 VBA 代码（用于 Skill/工作流等非交互模式）
        ' 与沙箱模式的区别：结果直接作用于用户数据，无需回写
        ' （该方法定义在其他分部文件中）

    End Class

End Namespace
