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

    ' VBACodeSandbox.Detection.vb — 安全检测与敏感操作识别（关键词/混淆/复杂度/白名单覆盖）
    ' （类级 XML 文档只在主文件 VBACodeSandbox.vb 保留一份，避免 BC42314 分部重复）
    Partial Public Class VBACodeSandbox

        ''' <summary>
        ''' Detect common Excel operations from VBA code pattern.
        ''' Used by both security check (risk lowering) and execution routing.
        ''' </summary>
        Public Shared Function DetectCommonExcelOperation(vbaCode As String) As CommonExcelOperation
            Dim cl As String = vbaCode.ToLower()

            ' ClearSheet patterns
            If cl.Contains("clearsheet") OrElse
               (cl.Contains(".clear") AndAlso (cl.Contains("usedrange") OrElse cl.Contains("cells.clear"))) OrElse
               (cl.Contains("clearcontents") AndAlso Not cl.Contains("range(""")) Then
                Return CommonExcelOperation.ClearSheet
            End If

            ' SplitWorkbook patterns
            If cl.Contains("splitworkbook") OrElse
               (cl.Contains("sub splitworkbook") AndAlso cl.Contains("for") AndAlso cl.Contains("next")) OrElse
               (cl.Contains("for i =") AndAlso cl.Contains("sheets.count") AndAlso cl.Contains("copy") AndAlso cl.Contains("saveas")) Then
                Return CommonExcelOperation.SplitWorkbook
            End If

            ' MergeSheets patterns
            If cl.Contains("mergesheets") OrElse
               (cl.Contains("sub mergesheets") AndAlso cl.Contains("usedrange.copy")) Then
                Return CommonExcelOperation.MergeSheets
            End If

            ' RemoveDuplicates patterns
            If cl.Contains("removeduplicates") OrElse
               (cl.Contains("sub removeduplicates") AndAlso cl.Contains(".removeduplicates")) Then
                Return CommonExcelOperation.RemoveDuplicates
            End If

            ' CreateChart patterns
            If cl.Contains("createchart") OrElse
               (cl.Contains("sub createchart") AndAlso cl.Contains("chartobjects.add")) Then
                Return CommonExcelOperation.CreateChart
            End If

            Return CommonExcelOperation.None
        End Function

        ''' <summary>
        ''' 检测代码是否包含文件生成操作（PPT、PDF、文件导出等）。
        ''' 这类操作不能在沙箱中执行——沙箱工作簿是临时的，ThisWorkbook.Path 指向临时目录，
        ''' 生成的文件会保存到临时位置，沙箱关闭后用户找不到文件。
        ''' 应走直接执行路径（ExecuteDirectInWorkbook），在用户真实工作簿中执行。
        ''' </summary>
        Public Shared Function DetectFileGeneratingOperation(vbaCode As String) As Boolean
            Dim cl As String = vbaCode.ToLower()

            ' PowerPoint 生成
            If cl.Contains("createobject(""powerpoint") OrElse
               cl.Contains("createobject(""powerpoint.application") OrElse
               cl.Contains("powerpoint.application") Then
                Return True
            End If

            ' Word 生成
            If cl.Contains("createobject(""word") OrElse
               cl.Contains("createobject(""word.application") Then
                Return True
            End If

            ' 文件保存/导出操作（SaveAs、ExportAsFixedFormat 等）
            If cl.Contains(".saveas") OrElse
               cl.Contains("exportasfixedformat") OrElse
               cl.Contains(".export(") Then
                Return True
            End If

            ' Shell/文件系统操作（创建外部文件）
            If cl.Contains("createobject(""scripting.filesystemobject") OrElse
               cl.Contains("createobject(""wscript.shell") Then
                Return True
            End If

            Return False
        End Function

        ''' <summary>
        ''' Pre-execution security check - Complete code review process
        ''' </summary>
        Public Shared Function CheckCode(vbaCode As String) As SecurityCheckResult
            Dim result As New SecurityCheckResult()

            ' 1. Basic validation
            If String.IsNullOrWhiteSpace(vbaCode) Then
                result.Errors.Add("Code is empty")
                result.IsSafe = False
                result.RiskLevel = RiskLevel.Critical
                Return result
            End If

            ' 2. Code size limit
            If vbaCode.Length > MAX_CODE_LENGTH Then
                result.Errors.Add(String.Format("Code length exceeds {0} characters, execution rejected", MAX_CODE_LENGTH))
                result.RiskLevel = RiskLevel.Critical
                result.IsSafe = False
                Return result
            End If

            ' 3. Critical keyword check (reject immediately if found)
            If CheckCriticalKeywords(vbaCode, result) Then
                result.IsSafe = False
                result.RiskLevel = RiskLevel.Critical
                Return result
            End If

            ' 4. Obfuscation bypass check (string concat, Chr(), dynamic execution)
            CheckObfuscationBypasses(vbaCode, result)

            ' 5. Dangerous keyword check
            CheckDangerousKeywords(vbaCode, result)

            ' 6. Dangerous pattern check
            CheckDangerousPatterns(vbaCode, result)

            ' 7. File system operation check
            CheckFileSystemOperations(vbaCode, result)

            ' 8. Network operation check
            CheckNetworkOperations(vbaCode, result)

            ' 9. Registry operation check
            CheckRegistryOperations(vbaCode, result)

            ' 10. Code complexity check
            CheckCodeComplexity(vbaCode, result)

            ' 10.5 识别代码调用的操作，未在白名单内且未被黑名单
            '     覆盖的操作标记为 Medium 风险（执行时需用户确认），白名单命中则降低风险。
            '     设计：白名单是"允许的安全操作"集合，与黑名单互补——
            '     黑名单捕获明确危险；白名单给常见安全操作降权；两者都不匹配的操作保守视为
            '     需确认（Medium）。这从"黑名单默认放行"演进为"白名单默认信任 + 未知需确认"。
            CheckWhitelistCoverage(vbaCode, result)

            ' 11. Determine final safety
            DetermineSafety(result)

            ' 12. Generate code signature
            result.CodeHash = ComputeCodeHash(vbaCode)

            Return result
        End Function

        ''' <summary>
        ''' Check critical keywords - Reject execution immediately if found
        ''' Now also detects: string concatenation bypasses (e.g. "W" &amp; "Script.Shell"),
        ''' Chr()-based obfuscation, and dynamic code execution patterns
        ''' </summary>
        Private Shared Function CheckCriticalKeywords(vbaCode As String, result As SecurityCheckResult) As Boolean
            Dim codeLower As String = vbaCode.ToLower()
            Dim codeNoComments As String = StripComments(vbaCode).ToLower()
            Dim foundCritical As Boolean = False

            ' 1. Direct keyword matching
            For Each keyword As String In CriticalKeywords
                Dim keywordLower As String = keyword.ToLower()
                If codeLower.Contains(keywordLower) Then
                    result.Errors.Add(String.Format("Critical keyword found: {0}, execution rejected", keyword))
                    result.RiskLevel = RiskLevel.Critical
                    foundCritical = True
                End If
            Next

            ' 2. Detect Chr()-based obfuscation (e.g. Chr(87) & Chr(83) & Chr(99) = "WSc")
            If DetectChrObfuscation(vbaCode, result) Then
                foundCritical = True
            End If

            ' 3. Detect string concatenation bypass — reconstruct joined string literals
            '    e.g. "W" & "Script.Shell" or "Scr" & "ipting.File" & "SystemObject"
            If DetectStringConcatBypass(codeNoComments, result) Then
                foundCritical = True
            End If

            ' 4. Detect dynamic code execution (Application.Run with non-constant args, Evaluate with expressions)
            If DetectDynamicExecution(codeNoComments, result) Then
                foundCritical = True
            End If

            Return foundCritical
        End Function

        ''' <summary>
        ''' Strip comments from VBA code (lines starting with ' or Rem)
        ''' </summary>
        Private Shared Function StripComments(vbaCode As String) As String
            Dim lines As String() = vbaCode.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None)
            Dim result As New List(Of String)()

            For Each line As String In lines
                Dim trimmed As String = line.TrimStart()
                If trimmed.StartsWith("'") OrElse trimmed.StartsWith("Rem ", StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If
                ' Remove inline comment (but not inside string literals — best effort)
                Dim commentIdx As Integer = line.IndexOf("'")
                If commentIdx >= 0 Then
                    ' Only treat as comment if it follows a non-string context (simple heuristic)
                    ' Check if the ' is inside a string by counting quotes before it
                    Dim beforeComment As String = line.Substring(0, commentIdx)
                    Dim quoteCount As Integer = beforeComment.Count(Function(c) c = """"c)
                    If quoteCount Mod 2 = 0 Then
                        result.Add(beforeComment)
                        Continue For
                    End If
                End If
                result.Add(line)
            Next

            Return String.Join(vbCrLf, result)
        End Function

        ''' <summary>
        ''' Detect Chr()-based obfuscation used to construct dangerous strings.
        ''' E.g. Chr(67) &amp; Chr(114) &amp; Chr(101) ... = "CreateObject"
        ''' We flag any sequence of 3+ consecutive Chr() calls as suspicious.
        ''' </summary>
        Private Shared Function DetectChrObfuscation(vbaCode As String, result As SecurityCheckResult) As Boolean
            ' 仅当多个 Chr()/ChrW()/Chr$() 调用被 & 连接成连续序列（即真正用于拼接构造字符串）时才判定为混淆，
            ' 例如 Chr(87) & Chr(83) & Chr(99) & ... = "WSc"。
            ' 分散的单字符 ChrW(8226)（如 PPT 列表 bullet 符 8226 = 0x2022）不属于混淆，故不再按全文出现次数计数。
            ' VB 无重复量词插值，直接拼出正则
            Dim seqPattern As String = "(?:Chr[W\$]?\s*\([^()]*\)\s*\&\s*)\{" & CHR_OBFUSCATION_THRESHOLD.ToString() & ",}"
            Dim seqMatches As MatchCollection = Regex.Matches(vbaCode, seqPattern, RegexOptions.IgnoreCase)
            If seqMatches IsNot Nothing AndAlso seqMatches.Count > 0 Then
                result.Errors.Add("Potential Chr()-based obfuscation detected (consecutive Chr() concatenation), execution rejected")
                result.RiskLevel = RiskLevel.Critical
                Return True
            End If
            Return False
        End Function

        ''' <summary>
        ''' Reconstruct string literal concatenations and check against dangerous patterns.
        ''' E.g. "WScr" &amp; "ipt.Shell" - "WScript.Shell"
        '''      "Scr" &amp; "ipting.File" &amp; "SystemObject" - "Scripting.FileSystemObject"
        ''' </summary>
        Private Shared Function DetectStringConcatBypass(codeNoComments As String, result As SecurityCheckResult) As Boolean
            ' Find all quoted string literals separated by & operators
            ' Pattern: "string1" & "string2" & ... "stringN"
            Dim concatPattern As String = "(""[^""]*"")\s*(\&)\s*(""[^""]*"")"
            Dim concatMatches As MatchCollection = Regex.Matches(codeNoComments, concatPattern, RegexOptions.IgnoreCase)

            If concatMatches IsNot Nothing AndAlso concatMatches.Count > 0 Then
                ' Build the full reconstructed string for each match group
                Dim reconstructed As New List(Of String)()
                For Each m As Match In concatMatches
                    Dim parts As New List(Of String)()
                    ' Extract all quoted strings from the entire matched region
                    Dim strMatches As MatchCollection = Regex.Matches(m.Value, """([^""]*)""", RegexOptions.IgnoreCase)
                    For Each sm As Match In strMatches
                        parts.Add(sm.Groups(1).Value)
                    Next

                    If parts.Count >= 2 Then
                        Dim combined As String = String.Join("", parts).ToLower()
                        reconstructed.Add(combined)
                    End If
                Next

                ' Check reconstructed strings against obfuscation patterns
                For Each combined As String In reconstructed
                    For Each pattern As String In ObfuscationBypassPatterns
                        If combined.Contains(pattern) Then
                            result.Errors.Add(String.Format("String concatenation bypass detected — reconstructed: '{0}' matches dangerous pattern: {1}, execution rejected",
                                combined, pattern))
                            result.RiskLevel = RiskLevel.Critical
                            Return True
                        End If
                    Next

                    ' Also check if reconstructed string contains critical keywords
                    For Each kw As String In CriticalKeywords
                        If combined.Contains(kw.ToLower()) Then
                            result.Errors.Add(String.Format("String concatenation bypass detected — reconstructed: '{0}' contains critical keyword: {1}, execution rejected",
                                combined, kw))
                            result.RiskLevel = RiskLevel.Critical
                            Return True
                        End If
                    Next
                Next
            End If

            Return False
        End Function

        ''' <summary>
        ''' Detect dynamic code execution patterns:
        ''' - Application.Run with variable/non-constant first argument
        ''' - Application.Evaluate / ActiveSheet.Evaluate with expressions
        ''' - Eval/Execute/ExecuteGlobal calls
        ''' - ScriptControl.AddCode or .Run/.Eval
        ''' </summary>
        Private Shared Function DetectDynamicExecution(codeNoComments As String, result As SecurityCheckResult) As Boolean
            ' Check for Eval/Execute/ExecuteGlobal (these execute arbitrary strings as code)
            For Each func As String In New String() {"Eval(", "Execute(", "ExecuteGlobal("}
                If codeNoComments.Contains(func.ToLower()) Then
                    result.Errors.Add(String.Format("Dynamic code execution function detected: {0}, execution rejected", func.TrimEnd("("c)))
                    result.RiskLevel = RiskLevel.Critical
                    Return True
                End If
            Next

            ' Check for ScriptControl usage (ActiveX that can run arbitrary VBScript/JScript)
            If codeNoComments.Contains("scriptcontrol") Then
                result.Errors.Add("ScriptControl usage detected — can execute arbitrary VBScript/JScript, execution rejected")
                result.RiskLevel = RiskLevel.Critical
                Return True
            End If

            ' Check for Application.Run with non-constant first argument (parentheses after Run contain variable)
            ' Pattern: Application.Run (variable) or Application.Run expression
            Dim runMatches As MatchCollection = Regex.Matches(
                codeNoComments,
                "Application\s*\.\s*Run\s*\(\s*(?!"")([^""\)]+)\)",
                RegexOptions.IgnoreCase)
            If runMatches IsNot Nothing AndAlso runMatches.Count > 0 Then
                result.Errors.Add("Application.Run called with non-constant argument — could execute arbitrary macros, execution rejected")
                result.RiskLevel = RiskLevel.Critical
                Return True
            End If

            ' Check for Application.Evaluate with complex expressions (potential formula injection)
            Dim evalMatches As MatchCollection = Regex.Matches(
                codeNoComments,
                "(?:Application|ActiveSheet|ActiveWorkbook)\s*\.\s*Evaluate\s*\(\s*""[^""]*\+(?:cmd|shell|call|dde)[^""]*""",
                RegexOptions.IgnoreCase)
            If evalMatches IsNot Nothing AndAlso evalMatches.Count > 0 Then
                result.Errors.Add("Evaluate with potential command/Shell injection detected, execution rejected")
                result.RiskLevel = RiskLevel.Critical
                Return True
            End If

            Return False
        End Function

        ''' <summary>
        ''' Check code complexity
        ''' </summary>
        Private Shared Sub CheckCodeComplexity(vbaCode As String, result As SecurityCheckResult)
            ' Check nesting depth
            Dim maxNestingLevel As Integer = 0
            Dim currentNesting As Integer = 0

            For Each c As Char In vbaCode
                If c = "("c OrElse c = "{"c Then
                    currentNesting += 1
                    maxNestingLevel = Math.Max(maxNestingLevel, currentNesting)
                ElseIf c = ")"c OrElse c = "}"c Then
                    currentNesting -= 1
                End If
            Next

            If maxNestingLevel > MAX_NESTING_DEPTH Then
                result.Warnings.Add(String.Format("Code nesting depth too large ({0} layers), potential risk", maxNestingLevel))
                If result.RiskLevel < RiskLevel.Medium Then
                    result.RiskLevel = RiskLevel.Medium
                End If
            End If

            ' Check line count
            Dim lineCount As Integer = vbaCode.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None).Length
            If lineCount > MAX_VBA_LINE_COUNT Then
                result.Warnings.Add(String.Format("Too many code lines ({0} lines)", lineCount))
                If result.RiskLevel < RiskLevel.Low Then
                    result.RiskLevel = RiskLevel.Low
                End If
            End If

            ' Check for obfuscated code characteristics
            If CheckObfuscation(vbaCode) Then
                result.Warnings.Add("Possible code obfuscation detected")
                If result.RiskLevel < RiskLevel.High Then
                    result.RiskLevel = RiskLevel.High
                End If
            End If
        End Sub

        ''' <summary>
        ''' Check code obfuscation
        ''' </summary>
        Private Shared Function CheckObfuscation(vbaCode As String) As Boolean
            ' Check long lines (possibly obfuscated code in one line)
            Dim lines As String() = vbaCode.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None)
            For Each line As String In lines
                If line.Length > MAX_OBFUSCATED_LINE_LENGTH Then
                    Return True
                End If
            Next

            ' Check for abnormally high special character ratio
            Dim specialChars As Integer = 0
            For Each c As Char In vbaCode
                If Not Char.IsLetterOrDigit(c) AndAlso Not Char.IsWhiteSpace(c) Then
                    specialChars += 1
                End If
            Next

            Dim specialCharRatio As Double = specialChars / vbaCode.Length
            If specialCharRatio > OBFUSCATION_SPECIAL_CHAR_RATIO Then
                Return True
            End If

            Return False
        End Function

        ''' <summary>
        ''' Compute code hash (for auditing and duplicate detection)
        ''' </summary>
        Public Shared Function ComputeCodeHash(vbaCode As String) As String
            Using sha256 As SHA256 = SHA256.Create()
                Dim bytes As Byte() = Encoding.UTF8.GetBytes(vbaCode)
                Dim hash As Byte() = sha256.ComputeHash(bytes)
                Return Convert.ToBase64String(hash)
            End Using
        End Function

        Private Shared Sub CheckDangerousKeywords(vbaCode As String, result As SecurityCheckResult)
            ' Split code by lines and check line by line
            Dim lines As String() = vbaCode.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None)
            
            For Each keyword As String In DangerousKeywords
                Dim keywordLower As String = keyword.ToLower()
                
                For Each line As String In lines
                    ' Remove comments, only check code part
                    Dim codePart As String = line
                    If line.Contains("'") Then
                        codePart = line.Substring(0, line.IndexOf("'"))
                    End If
                    
                    Dim codePartLower As String = codePart.ToLower()
                    
                    ' Check if keyword is in code (as whole word)
                    If codePartLower.Contains(keywordLower) Then
                        ' Verify it's a whole word and not part of variable name
                        Dim index As Integer = codePartLower.IndexOf(keywordLower)
                        If index >= 0 Then
                            ' Check if preceding character is word boundary (space, line start, parenthesis, etc.)
                            Dim isWordBoundaryBefore As Boolean = (index = 0) OrElse 
                                Char.IsWhiteSpace(codePart(index - 1)) OrElse 
                                codePart(index - 1) = "("c OrElse 
                                codePart(index - 1) = "."c
                            
                            If isWordBoundaryBefore Then
                                result.Warnings.Add(String.Format("Potentially dangerous keyword found: {0}", keyword))
                                If result.RiskLevel < RiskLevel.High Then
                                    result.RiskLevel = RiskLevel.High
                                End If
                                Exit For ' Found once is enough, continue checking next keyword
                            End If
                        End If
                    End If
                Next
            Next
        End Sub

        Private Shared Sub CheckDangerousPatterns(vbaCode As String, result As SecurityCheckResult)
            Dim codeLower As String = vbaCode.ToLower()

            ' API declaration (Declare Function ... Lib ...)
            If codeLower.Contains("declare") AndAlso codeLower.Contains("lib") Then
                result.Warnings.Add("Potentially dangerous code pattern found: API declaration")
                If result.RiskLevel < RiskLevel.Medium Then
                    result.RiskLevel = RiskLevel.Medium
                End If
            End If

            ' COM object creation
            If vbaCode.Contains("CreateObject") Then
                result.Warnings.Add("CreateObject call found")
                If result.RiskLevel < RiskLevel.Medium Then
                    result.RiskLevel = RiskLevel.Medium
                End If
            End If

            ' GetObject — can bind to running system objects
            If vbaCode.Contains("GetObject") Then
                result.Warnings.Add("GetObject call found — can bind to running system objects")
                If result.RiskLevel < RiskLevel.Medium Then
                    result.RiskLevel = RiskLevel.Medium
                End If
            End If

            ' File output
            If codeLower.Contains("open") AndAlso codeLower.Contains("for output") Then
                result.Warnings.Add("File output operation found")
                If result.RiskLevel < RiskLevel.Medium Then
                    result.RiskLevel = RiskLevel.Medium
                End If
            End If

            ' AddressOf — function pointer, potential API callback abuse
            If vbaCode.Contains("AddressOf") Then
                result.Warnings.Add("AddressOf function pointer found")
                If result.RiskLevel < RiskLevel.Medium Then
                    result.RiskLevel = RiskLevel.Medium
                End If
            End If

            ' VBProject manipulation — can inject/delete code modules
            If codeLower.Contains("vbproject") OrElse codeLower.Contains("vbcomponents") Then
                result.Warnings.Add("VBProject manipulation detected — potential code injection")
                If result.RiskLevel < RiskLevel.High Then
                    result.RiskLevel = RiskLevel.High
                End If
            End If

            ' VBE access — can modify IDE, delete modules
            If codeLower.Contains("application.vbe") OrElse codeLower.Contains(".vbe.") Then
                result.Warnings.Add("VBE (IDE) access detected — potential IDE manipulation")
                If result.RiskLevel < RiskLevel.High Then
                    result.RiskLevel = RiskLevel.High
                End If
            End If

            ' CommandBars manipulation (legacy but can be abused)
            If codeLower.Contains("commandbars(") Then
                result.Warnings.Add("CommandBars manipulation detected")
                If result.RiskLevel < RiskLevel.Medium Then
                    result.RiskLevel = RiskLevel.Medium
                End If
            End If

            ' Scripting.Dictionary (less dangerous but can be used in data exfiltration)
            If codeLower.Contains("scripting.dictionary") Then
                result.Warnings.Add("Scripting.Dictionary usage detected")
                If result.RiskLevel < RiskLevel.Low Then
                    result.RiskLevel = RiskLevel.Low
                End If
            End If

            ' Check for common Excel operations (lower risk on these)
            CheckCommonExcelOperations(vbaCode, result)
        End Sub

        ''' <summary>
        ''' 检查常见的Excel操作，降低其风险级别
        ''' </summary>
        Private Shared Sub CheckCommonExcelOperations(vbaCode As String, result As SecurityCheckResult)
            Dim op As CommonExcelOperation = DetectCommonExcelOperation(vbaCode)

            If op = CommonExcelOperation.None Then Return

            Dim opName As String = op.ToString()
            result.Warnings.Add("Common Excel operation detected: " & opName)

            If result.RiskLevel > RiskLevel.Medium Then
                result.Warnings.Add("Risk level reduced: Common Excel operation")
                result.RiskLevel = RiskLevel.Medium
            End If
        End Sub

        Private Shared Sub CheckFileSystemOperations(vbaCode As String, result As SecurityCheckResult)
            Dim codeLower As String = vbaCode.ToLower()

            If codeLower.Contains("kill ") Then
                result.Warnings.Add("File deletion operation found (Kill)")
                If result.RiskLevel < RiskLevel.High Then
                    result.RiskLevel = RiskLevel.High
                End If
            End If

            If codeLower.Contains("rmdir ") OrElse codeLower.Contains("rmdir(") Then
                result.Warnings.Add("Directory deletion operation found (RmDir)")
                If result.RiskLevel < RiskLevel.High Then
                    result.RiskLevel = RiskLevel.High
                End If
            End If

            If codeLower.Contains("filecopy ") Then
                result.Warnings.Add("File copy operation found")
                If result.RiskLevel < RiskLevel.Medium Then
                    result.RiskLevel = RiskLevel.Medium
                End If
            End If
        End Sub

        Private Shared Sub CheckNetworkOperations(vbaCode As String, result As SecurityCheckResult)
            If vbaCode.Contains("Microsoft.XMLHTTP") Then
                result.Warnings.Add("Network request operation found: Microsoft.XMLHTTP")
                If result.RiskLevel < RiskLevel.Medium Then
                    result.RiskLevel = RiskLevel.Medium
                End If
            End If

            If vbaCode.Contains("WinHttp.WinHttpRequest") Then
                result.Warnings.Add("Network request operation found: WinHttp")
                If result.RiskLevel < RiskLevel.Medium Then
                    result.RiskLevel = RiskLevel.Medium
                End If
            End If

            If vbaCode.Contains("MSXML2.XMLHTTP") Then
                result.Warnings.Add("Network request operation found: MSXML2")
                If result.RiskLevel < RiskLevel.Medium Then
                    result.RiskLevel = RiskLevel.Medium
                End If
            End If
        End Sub

        Private Shared Sub CheckRegistryOperations(vbaCode As String, result As SecurityCheckResult)
            If vbaCode.Contains("RegWrite") Then
                result.Warnings.Add("Registry write operation found")
                If result.RiskLevel < RiskLevel.High Then
                    result.RiskLevel = RiskLevel.High
                End If
            End If

            If vbaCode.Contains("RegDelete") Then
                result.Warnings.Add("Registry delete operation found")
                If result.RiskLevel < RiskLevel.High Then
                    result.RiskLevel = RiskLevel.High
                End If
            End If

            If vbaCode.Contains("RegRead") Then
                result.Warnings.Add("Registry read operation found")
                If result.RiskLevel < RiskLevel.Medium Then
                    result.RiskLevel = RiskLevel.Medium
                End If
            End If
        End Sub

        ''' <summary>
        ''' 识别代码调用的操作，评估白名单覆盖度。
        ''' 原则：
        '''   - 代码只使用白名单安全操作（Range/Cells/Worksheets 等）→ 风险降至 Low（默认信任）
        '''   - 出现白名单外操作（如 CreateObject/GetObject 安全用法、Excel 高级 API）→ 标记 Medium
        '''     （执行时需用户确认），避免"黑名单之外全放行"的默认信任风险
        '''   - 已有黑名单命中的（文件/网络/注册表）→ 不改判（保持原风险等级）
        ''' 说明：本方法只做风险升级/降级与警告记录，不做硬性阻止；
        '''       硬性阻止仍由 Critical 关键字与 Errors 决定。
        ''' </summary>
        Private Shared Sub CheckWhitelistCoverage(vbaCode As String, result As SecurityCheckResult)
            ' 已判 Critical（有 Errors 或显式危险）→ 白名单不再降权，直接返回
            If result.Errors.Count > 0 OrElse result.RiskLevel = RiskLevel.Critical Then
                Return
            End If

            Dim codeNoComments As String = StripComments(vbaCode)
            If String.IsNullOrWhiteSpace(codeNoComments) Then Return

            ' 1. 提取代码中出现的"操作标识"（调用点/属性名）
            '    简化启发式：匹配 标识符( 形式的方法调用 + 已知的属性访问
            Dim calledOps As New List(Of String)()
            Dim callMatches As MatchCollection = Regex.Matches(codeNoComments,
                "\b([A-Za-z_][A-Za-z0-9_]*)\s*\(", RegexOptions.IgnoreCase)
            For Each m As Match In callMatches
                If m.Groups.Count > 1 Then
                    Dim opName As String = m.Groups(1).Value
                    If Not calledOps.Contains(opName) Then calledOps.Add(opName)
                End If
            Next
            ' 属性访问类操作（无括号形式）
            For Each prop In {"ActiveSheet", "ActiveWorkbook", "Selection", "UsedRange", "CurrentRegion"}
                If codeNoComments.IndexOf(prop, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    If Not calledOps.Contains(prop) Then calledOps.Add(prop)
                End If
            Next

            If calledOps.Count = 0 Then Return

            ' 2. 评估白名单覆盖：所有操作都在白名单内 → 降为 Low（默认信任）
            Dim whitelistHits As Integer = 0
            Dim unknownOps As New List(Of String)()
            For Each op In calledOps
                Dim isWhitelisted As Boolean = False
                For Each safeOp As String In AllowedSafeOperations
                    ' 白名单项如 "Range(" / "Worksheets(" / "ActiveSheet"（无括号）
                    Dim safeBase As String = safeOp.Replace("(", "").Trim()
                    If String.Compare(op, safeBase, StringComparison.OrdinalIgnoreCase) = 0 Then
                        isWhitelisted = True
                        Exit For
                    End If
                Next
                If isWhitelisted Then
                    whitelistHits += 1
                Else
                    unknownOps.Add(op)
                End If
            Next

            ' 3. 将未命中的操作存储到结果中（用于"加入白名单"功能）
            result.UnknownWhitelistOps = unknownOps

            ' 4. 风险调整
            If unknownOps.Count = 0 Then
                ' 全部操作在白名单内 → 默认信任
                ' 检查是否存在"不可降级"的严重风险（文件删除/目录删除等）
                Dim hasIrreducibleRisk As Boolean = result.Warnings.Any(Function(w)
                    Return w.Contains("File deletion") OrElse
                           w.Contains("Directory deletion") OrElse
                           w.Contains("Process termination") OrElse
                           w.Contains("Registry write") OrElse
                           w.Contains("Registry delete")
                End Function)

                If hasIrreducibleRisk Then
                    ' 存在不可降级风险（如 Kill 文件删除），保持 High 不变
                    result.Warnings.Add(String.Format(
                        "白名单校验: 检测到 {0} 个操作均在安全白名单内，但存在严重风险需确认执行",
                        calledOps.Count))
                Else
                    ' 可降级：用户已信任所有操作，降到 Low
                    If result.RiskLevel > RiskLevel.Low Then
                        result.RiskLevel = RiskLevel.Low
                    End If
                    result.Warnings.Add(String.Format(
                        "白名单校验: 检测到 {0} 个操作均在安全白名单内，默认信任",
                        calledOps.Count))
                End If
            Else
                ' 存在白名单外操作 → 保守升级到至少 Medium（执行时需用户确认）
                If result.RiskLevel < RiskLevel.Medium Then
                    result.RiskLevel = RiskLevel.Medium
                End If
                Dim unknownSummary As String = String.Join(", ", unknownOps.Take(8))
                result.Warnings.Add(String.Format(
                    "白名单校验: {0} 个操作不在安全白名单内（{1}），将需用户确认执行",
                    unknownOps.Count, unknownSummary))
            End If
        End Sub

        Private Shared Sub DetermineSafety(result As SecurityCheckResult)
            ' 【安全策略】
            ' - Errors > 0（如动态ProgID、运行时代码注入）→ 直接阻止
            ' - Critical 风险（关键字匹配）→ 直接阻止
            ' - High 风险（文件删除、注册表等）→ 不阻止，但需用户确认（在 ExecuteCode 中处理）
            ' - Medium 风险 → 不阻止，但需用户确认
            ' - Low/None → 直接执行
            If result.Errors.Count > 0 Then
                result.IsSafe = False
                Return
            End If

            ' 仅 Critical 风险无条件阻止执行
            ' 【修正】Round1 Fix#3 原将 High 也设为阻止，导致自定义 Skill 的正常 VBA 代码被误杀
            ' High 风险代码仍需用户在 MessageBox 中明确确认后才可执行
            If result.RiskLevel = RiskLevel.Critical Then
                result.IsSafe = False
                Return
            End If

            ' High/Medium/Low/None 均标记为 Safe，具体由 ExecuteCode 中的 requireConfirmation + RiskLevel 决定是否弹窗确认
            result.IsSafe = True
        End Sub

        ''' <summary>
        ''' Non-critical obfuscation bypass check — adds warnings for suspicious patterns
        ''' that don't necessarily warrant immediate rejection but should raise the risk level.
        ''' Detects: variable-based ProgID construction, indirect COM creation,
        ''' Application.Run with indirection, and namespace-wrapped dangerous calls.
        ''' </summary>
        Private Shared Sub CheckObfuscationBypasses(vbaCode As String, result As SecurityCheckResult)
            Dim codeNoComments As String = StripComments(vbaCode)

            ' 1. Detect CreateObject/GetObject whose argument is NOT a literal string (variable/regex)
            '    e.g. CreateObject(progIdVariable) or GetObject(, classVariable)
            Dim coPattern As String = "(?:CreateObject|GetObject)\s*\(\s*(?:[^""]|$)"
            Dim coMatches As MatchCollection = Regex.Matches(codeNoComments, coPattern, RegexOptions.IgnoreCase)
            If coMatches IsNot Nothing AndAlso coMatches.Count > 0 Then
                result.Warnings.Add(String.Format("CreateObject/GetObject with non-literal argument detected ({0} occurrence(s)) — ProgID is dynamically resolved", coMatches.Count))
                If result.RiskLevel < RiskLevel.Critical Then
                    result.RiskLevel = RiskLevel.Critical
                End If
                result.Errors.Add("Dynamic ProgID via CreateObject/GetObject — execution rejected")
                result.IsSafe = False
            End If

            ' 2. Detect indirect COM creation via late binding (Dim x As Object : Set x = ...)
            '    combined with suspicious namespace usage
            Dim lateBindPattern As String = "(?:Dim|Set)\s+\w+\s+As\s+(?:Object|Variant)\s*\n.*CreateObject"
            If Regex.IsMatch(codeNoComments, lateBindPattern, RegexOptions.IgnoreCase Or RegexOptions.Singleline) Then
                result.Warnings.Add("Late-bound COM object creation with Object/Variant detected")
                If result.RiskLevel < RiskLevel.Medium Then
                    result.RiskLevel = RiskLevel.Medium
                End If
            End If

            ' 3. Detect Run() or CallByName indirection
            Dim indirectCall As String = "CallByName\s*\("
            If Regex.IsMatch(codeNoComments, indirectCall, RegexOptions.IgnoreCase) Then
                result.Warnings.Add("CallByName late-bound invocation detected")
                If result.RiskLevel < RiskLevel.High Then
                    result.RiskLevel = RiskLevel.High
                End If
            End If

            ' 4. Detect SendKeys pattern (can automate UI to bypass dialogs)
            '    Already handled by DangerousKeywords but check for indirection
            If Regex.IsMatch(codeNoComments, "SendKeys\s+(?!"")(\w+)", RegexOptions.IgnoreCase) Then
                result.Warnings.Add("SendKeys with variable argument — potential UI automation bypass")
                If result.RiskLevel < RiskLevel.High Then
                    result.RiskLevel = RiskLevel.High
                End If
            End If

            ' 5. Detect access to VBA IDE extensibility or code module manipulation at runtime
            If Regex.IsMatch(codeNoComments, "\.(?:InsertLines|DeleteLines|AddFromString|AddFromFile)\s*\(", RegexOptions.IgnoreCase) Then
                result.Warnings.Add("Runtime code injection via VBA extensibility detected")
                If result.RiskLevel < RiskLevel.Critical Then
                    result.RiskLevel = RiskLevel.Critical
                End If
                result.Errors.Add("Runtime code injection (InsertLines/AddFromString) detected — execution rejected")
                result.IsSafe = False
            End If

            ' 6. Detect potential namespace resolution bypass
            '    e.g. VBA.Interaction.Shell (using fully-qualified name to avoid Shell keyword)
            If Regex.IsMatch(codeNoComments, "(?:VBA|VisualBasic)\s*\.\s*Interaction\s*\.\s*\w+", RegexOptions.IgnoreCase) Then
                result.Warnings.Add("VBA.Interaction namespace usage — potential keyword bypass")
                If result.RiskLevel < RiskLevel.High Then
                    result.RiskLevel = RiskLevel.High
                End If
            End If

            ' 7. Detect FileSystemObject usage patterns not caught by keyword check
            '    e.g. Set fso = CreateObject(...) followed by fso.DeleteFile / fso.DeleteFolder
            If Regex.IsMatch(codeNoComments, "(?:\.DeleteFile|\.DeleteFolder|\.CopyFile|\.CopyFolder|\.MoveFile|\.MoveFolder|\.CreateTextFile|\.OpenTextFile|\.OpenAsTextStream)\s*\(",
                RegexOptions.IgnoreCase) Then
                result.Warnings.Add("FileSystemObject-style file operations detected")
                If result.RiskLevel < RiskLevel.High Then
                    result.RiskLevel = RiskLevel.High
                End If
            End If
        End Sub

    End Class

End Namespace
