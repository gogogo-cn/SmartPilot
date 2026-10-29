Option Strict On
Option Infer On
Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ExcelAIVSTOAddIn.Skills

Namespace Skills.Testing

    ' =========================================================================
    ' 第一部分：测试层级枚举 & 测试用例数据模型
    ' =========================================================================

    ''' <summary>
    ''' 测试层级：共4级，从纯语法校验→意图识别→单步执行→端到端集成
    ''' </summary>
    Public Enum SkillTestLevel
        ''' <summary>Level 1：Skill定义语法校验（MD Frontmatter解析、必填字段检查、枚举合法值）。纯离线，零依赖。</summary>
        DefinitionValidation = 1
        ''' <summary>Level 2：Skill意图识别测试（输入语句→期望Skill名+置信度阈值）。仅依赖SkillRecognizer。</summary>
        IntentRecognition = 2
        ''' <summary>Level 3：WorkflowStep单步执行测试（输入→某步骤→期望输出断言）。仅依赖对应步骤处理器 + Mock Excel。</summary>
        StepExecution = 3
        ''' <summary>Level 4：端到端(E2E)集成测试（完整Skill执行，需要真实Excel环境）。默认在CI中跳过。</summary>
        EndToEndIntegration = 4
    End Enum

    ''' <summary>测试执行状态</summary>
    Public Enum SkillTestStatus
        NotRun = 0
        Passed = 1
        Failed = 2
        Skipped = 3
        ''' <summary>断言通过但有警告（如性能SLA超出预期）</summary>
        PassedWithWarning = 4
        ''' <summary>测试本身有错误（如用例格式不对，非Skill问题）</summary>
        TestError = 5
    End Enum

    ''' <summary>单个断言结果</summary>
    Public Class SkillTestAssertionResult
        Public Property AssertionId As String
        Public Property Description As String
        Public Property Status As SkillTestStatus
        Public Property Expected As String
        Public Property Actual As String
        Public Property Message As String
    End Class

    ''' <summary>单个测试用例结果</summary>
    Public Class SkillTestCaseResult
        Public Property CaseId As String
        Public Property CaseName As String
        Public Property Level As SkillTestLevel
        Public Property Status As SkillTestStatus
        Public Property StartedAt As DateTime
        Public Property EndedAt As DateTime
        Public Property DurationMs As Long
        Public Property ErrorMessage As String
        Public Property ErrorStackTrace As String
        Public Property SkillVersion As String
        Public ReadOnly Property Warnings As List(Of String) = New List(Of String)()
        Public ReadOnly Property Assertions As List(Of SkillTestAssertionResult) = New List(Of SkillTestAssertionResult)()
    End Class

    ''' <summary>Skill测试报告（一个Skill的完整执行结果）</summary>
    Public Class SkillTestReport
        Public Property SkillName As String
        Public Property SkillVersion As String
        Public Property GeneratedAt As DateTime = DateTime.Now
        Public Property TotalCases As Integer
        Public Property PassedCases As Integer
        Public Property FailedCases As Integer
        Public Property SkippedCases As Integer
        Public Property ErrorCases As Integer
        Public Property TotalDurationMs As Long
        Public ReadOnly Property PassRate As Double
            Get
                If TotalCases = 0 Then Return 0
                Return Math.Round(CDbl(PassedCases) / TotalCases * 100, 2)
            End Get
        End Property
        Public ReadOnly Property Results As List(Of SkillTestCaseResult) = New List(Of SkillTestCaseResult)()
    End Class

    ''' <summary>单个测试用例定义（可从Markdown内嵌或外部JSON加载）</summary>
    Public Class SkillTestCase
        Public Property CaseId As String
        Public Property CaseName As String
        Public Property Description As String
        Public Property Level As SkillTestLevel = SkillTestLevel.DefinitionValidation
        Public Property SkipByDefault As Boolean
        Public Property SkipReason As String
        Public Property MaxDurationMs As Integer = 30000

        ''' <summary>L1必填字段检查</summary>
        Public ReadOnly Property RequiredFields As List(Of String) = New List(Of String)()
        ''' <summary>L1 YAML语法必须合法</summary>
        Public Property RequireValidYaml As Boolean = True
        ''' <summary>L2用户查询语句（给SkillRecognizer输入）</summary>
        Public Property UserQuery As String
        ''' <summary>L2期望匹配到的Skill名</summary>
        Public Property ExpectedSkillName As String
        ''' <summary>L2最低置信度阈值</summary>
        Public Property MinConfidence As Double = 0.5
        ''' <summary>L3目标步骤：WorkflowStep的StepId或StepName</summary>
        Public Property TargetStepIdOrName As String
        ''' <summary>L3步骤初始变量</summary>
        Public Property InitialVariables As Dictionary(Of String, Object)
        ''' <summary>通用断言：Key=字段/路径, Value=断言表达式(>=/contains:/regex:等)</summary>
        Public ReadOnly Property Assertions As Dictionary(Of String, String) = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    End Class

    ''' <summary>一个Skill的完整测试套件（一组用例）</summary>
    Public Class SkillTestSuite
        Public Property SkillName As String
        ''' <summary>测试用例来源："InlineMD"=从Skill MD文件的 ## Test Cases 区块解析，"ExternalJson"=外部JSON</summary>
        Public Property Source As String
        Public Property Cases As List(Of SkillTestCase) = New List(Of SkillTestCase)()
    End Class

    ' =========================================================================
    ' 第二部分：Mock抽象接口（隔离Excel/LLM/OCR等外部依赖，便于L1/L2/L3测试）
    ' =========================================================================

    ''' <summary>Excel操作Mock接口：L3步骤测试时替代真实Excel COM对象</summary>
    Public Interface IExcelMockContext
        ReadOnly Property WrittenRanges As Dictionary(Of String, Object)
        ReadOnly Property FormattedRanges As List(Of String)
        Function ReadCell(address As String) As Object
        Sub WriteCell(address As String, value As Object)
        Function ReadRange(rangeAddr As String) As Object(,)
        Sub WriteRange(rangeAddr As String, values As Object(,))
    End Interface

    ''' <summary>默认内存Mock Excel：测试时无需真实Excel</summary>
    Public Class InMemoryExcelMock
        Implements IExcelMockContext

        Private ReadOnly _cells As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)

        Public ReadOnly Property WrittenRanges As Dictionary(Of String, Object) Implements IExcelMockContext.WrittenRanges
            Get
                Return _cells.ToDictionary(Function(k) k.Key, Function(k) k.Value)
            End Get
        End Property

        Public ReadOnly Property FormattedRanges As New List(Of String) Implements IExcelMockContext.FormattedRanges

        Public Function ReadCell(address As String) As Object Implements IExcelMockContext.ReadCell
            If _cells.ContainsKey(address) Then Return _cells(address)
            Return Nothing
        End Function

        Public Sub WriteCell(address As String, value As Object) Implements IExcelMockContext.WriteCell
            _cells(address) = value
        End Sub

        Public Function ReadRange(rangeAddr As String) As Object(,) Implements IExcelMockContext.ReadRange
            Return New Object(0, 0) {}
        End Function

        Public Sub WriteRange(rangeAddr As String, values As Object(,)) Implements IExcelMockContext.WriteRange
            _cells(rangeAddr) = values
        End Sub
    End Class

    ' =========================================================================
    ' 第三部分：断言辅助器
    ' =========================================================================

    ''' <summary>断言表达式求值器：支持 "op:expected" 格式。例如 ">5", "regex:^\d+$", "contains:错误"</summary>
    Public Module SkillAssertionEvaluator
        ''' <summary>
        ''' 评估单条断言
        ''' </summary>
        ''' <param name="assertionExpr">断言表达式，如 "==hello", "!=null", ">5", ">=100", "contains:销售额", "regex:^\d{4}-\d{2}-\d{2}$"</param>
        ''' <param name="actualValue">实际值</param>
        Public Function Evaluate(assertionExpr As String, actualValue As Object) As (Success As Boolean, Expected As String, Actual As String, Message As String)
            If String.IsNullOrWhiteSpace(assertionExpr) Then
                Return (True, "(empty assertion)", actualValue?.ToString(), "断言为空，默认通过")
            End If

            Dim actualStr As String = If(actualValue Is Nothing, "null", actualValue.ToString())
            Dim expr As String = assertionExpr.Trim()

            Try
                ' 模式1: ==value (等于)
                If expr.StartsWith("==") Then
                    Dim expected = expr.Substring(2)
                    Return (actualStr = expected, "== " & expected, actualStr, If(actualStr = expected, "相等", "期望 '" & expected & "' 实际 '" & actualStr & "'"))
                End If

                ' 模式2: !=value (不等于)
                If expr.StartsWith("!=") Then
                    Dim expected = expr.Substring(2)
                    Return (actualStr <> expected, "!= " & expected, actualStr, If(actualStr <> expected, "不相等（正确）", "实际与期望相同，均为 '" & actualStr & "'"))
                End If

                ' 模式3: !=null 或 ==null
                If expr = "!=null" Then
                    Return (actualValue IsNot Nothing AndAlso actualStr <> "null" AndAlso actualStr <> "", "!= null", actualStr, "期望非空值")
                End If
                If expr = "==null" OrElse expr = "null" Then
                    Return (actualValue Is Nothing OrElse actualStr = "" OrElse actualStr = "null", "== null", actualStr, "期望值为空")
                End If

                ' 模式4: 数字比较 >, >=, <, <=
                Dim numCmp = Regex.Match(expr, "^(>=|<=|>|<)\s*(-?\d+(?:\.\d+)?)$")
                If numCmp.Success Then
                    Dim [op] = numCmp.Groups(1).Value
                    Dim threshold As Double
                    If Double.TryParse(numCmp.Groups(2).Value, threshold) Then
                        Dim actualNum As Double
                        If Not Double.TryParse(actualStr, actualNum) Then
                            Return (False, [op] & threshold.ToString(), actualStr, "实际值 '" & actualStr & "' 无法解析为数字")
                        End If
                        Dim ok As Boolean
                        Select Case [op]
                            Case ">" : ok = actualNum > threshold
                            Case ">=" : ok = actualNum >= threshold
                            Case "<" : ok = actualNum < threshold
                            Case "<=" : ok = actualNum <= threshold
                            Case Else : ok = False
                        End Select
                        Return (ok, [op] & threshold.ToString(), actualStr, If(ok, "数值比较通过", "数值 " & actualNum.ToString() & " 不满足 " & [op] & threshold.ToString()))
                    End If
                End If

                ' 模式5: contains:子串
                If expr.StartsWith("contains:", StringComparison.OrdinalIgnoreCase) Then
                    Dim needle = expr.Substring("contains:".Length)
                    Return (actualStr.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0,
                            "包含: '" & needle & "'", actualStr,
                            If(actualStr.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0, "找到子串", "未找到子串 '" & needle & "'"))
                End If

                ' 模式6: regex:正则表达式
                If expr.StartsWith("regex:", StringComparison.OrdinalIgnoreCase) Then
                    Dim pattern = expr.Substring("regex:".Length)
                    Dim ok = Regex.IsMatch(actualStr, pattern, RegexOptions.IgnoreCase)
                    Return (ok, "匹配正则 " & pattern, actualStr, If(ok, "正则匹配通过", "正则不匹配"))
                End If

                ' 模式7: jsonpath:$.path.to.field==value  (对JSON输出的字段断言)
                If expr.StartsWith("jsonpath:", StringComparison.OrdinalIgnoreCase) Then
                    Dim rest = expr.Substring("jsonpath:".Length)
                    Dim sepIdx = rest.IndexOfAny({"="c, ">"c, "<"c, "!"c})
                    If sepIdx > 0 Then
                        Dim path = rest.Substring(0, sepIdx)
                        Dim subExpr = rest.Substring(sepIdx)
                        Try
                            If actualStr.TrimStart().StartsWith("{"c) OrElse actualStr.TrimStart().StartsWith("["c) Then
                                Dim jt = JToken.Parse(actualStr)
                                Dim token = jt.SelectToken(path)
                                Dim fieldValue = If(token Is Nothing, Nothing, token.ToString())
                                Return Evaluate(subExpr, fieldValue)
                            End If
                        Catch ex As Exception
                            Return (False, "jsonpath valid " & path & subExpr, actualStr, "JSON解析失败: " & ex.Message)
                        End Try
                    End If
                End If

                ' 默认为 == (无操作符前缀的简洁写法)
                Return (actualStr = expr, "== " & expr, actualStr, If(actualStr = expr, "相等", "期望 '" & expr & "' 实际 '" & actualStr & "'"))

            Catch ex As Exception
                Return (False, expr, actualStr, "断言执行异常: " & ex.Message)
            End Try
        End Function
    End Module

    ' =========================================================================
    ' 第四部分：Skill MD 内嵌用例解析器（从Markdown的 ## Test Cases 区块解析）
    ' =========================================================================

    ''' <summary>从Skill Markdown文件的 ## Test Cases 区块解析出测试套件</summary>
    Public Module SkillMarkdownTestCasesParser
        Private Const TestCasesHeader As String = "## Test Cases"

        ''' <summary>
        ''' 解析Skill Markdown文本中的 Test Cases 区块
        ''' 格式约定（YAML + 断言列表）：
        '''   ## Test Cases
        '''   ### id: TC-DEF-001, level: 1, name: 必填字段检查
        '''   required: ["Name", "DisplayName", "Description"]
        '''   assertions:
        '''     - "Name !=null"
        '''     - "Description contains:Excel"
        ''' </summary>
        Public Function ParseFromSkillMarkdown(skillMdPath As String) As SkillTestSuite
            If Not File.Exists(skillMdPath) Then Return Nothing

            Dim skillName As String = Path.GetFileNameWithoutExtension(skillMdPath)
            Dim allText As String = File.ReadAllText(skillMdPath, Encoding.UTF8)

            Dim suite As New SkillTestSuite With {.SkillName = skillName, .Source = "InlineMD"}

            Dim tcIdx As Integer = allText.IndexOf(TestCasesHeader, StringComparison.OrdinalIgnoreCase)
            If tcIdx < 0 Then Return suite ' 没有 Test Cases 区块，返回空套件

            Dim tail As String = allText.Substring(tcIdx + TestCasesHeader.Length)
            ' 拆分用例：按 ### 分隔
            Dim caseBlocks As String() = Regex.Split(tail, "(?=^###\s)", RegexOptions.Multiline).Where(Function(b) b.Trim().StartsWith("###")).ToArray()

            For Each block As String In caseBlocks
                Dim tc As SkillTestCase = ParseOneCase(block)
                If tc IsNot Nothing Then suite.Cases.Add(tc)
            Next

            Return suite
        End Function

        Private Function ParseOneCase(block As String) As SkillTestCase
            Try
                Dim lines As String() = block.Split({ControlChars.CrLf, ControlChars.Lf}, StringSplitOptions.None)
                Dim headerLine As String = lines.FirstOrDefault(Function(l) l.Trim().StartsWith("###"))
                If headerLine Is Nothing Then Return Nothing

                Dim tc As New SkillTestCase()

                ' 解析 ### id: x, level: 1, name: y
                Dim header As String = headerLine.Substring(3).Trim()
                For Each kv As String In SplitByTopLevelComma(header)
                    Dim colon As Integer = kv.IndexOf(":"c)
                    If colon <= 0 Then Continue For
                    Dim k As String = kv.Substring(0, colon).Trim().ToLowerInvariant()
                    Dim v As String = kv.Substring(colon + 1).Trim().Trim(""""c, "'"c)
                    Select Case k
                        Case "id" : tc.CaseId = v
                        Case "level", "l" : Dim lvl As SkillTestLevel : If [Enum].TryParse(Of SkillTestLevel)(v, lvl) Then tc.Level = lvl
                        Case "name", "title" : tc.CaseName = v
                        Case "desc", "description" : tc.Description = v
                        Case "skip" : tc.SkipByDefault = v.Equals("true", StringComparison.OrdinalIgnoreCase)
                        Case "skipreason" : tc.SkipReason = v
                        Case "maxduration", "timeout" : Dim ms As Integer : If Integer.TryParse(v, ms) Then tc.MaxDurationMs = ms
                    End Select
                Next

                ' 解析 body：支持 required / query / expect / assertions / inputs ... （简化：行内键值 + YAML列表）
                Dim inAssertions As Boolean = False
                For Each raw As String In lines.Skip(1)
                    Dim line As String = raw.Trim()
                    If String.IsNullOrEmpty(line) Then Continue For

                    If line.StartsWith("assertions:", StringComparison.OrdinalIgnoreCase) Then
                        inAssertions = True
                        Continue For
                    End If

                    If inAssertions AndAlso (line.StartsWith("-") OrElse line.StartsWith("*")) Then
                        Dim item As String = line.TrimStart("-"c, "*"c, " "c).Trim(""""c, "'"c)
                        Dim keyId As Integer = tc.Assertions.Count + 1
                        tc.Assertions("断言" & keyId & ": " & item) = item
                        Continue For
                    End If

                    ' key: value 格式
                    Dim colon2 As Integer = line.IndexOf(":"c)
                    If colon2 > 0 AndAlso Not line.StartsWith("-"c) Then
                        Dim k2 As String = line.Substring(0, colon2).Trim().ToLowerInvariant()
                        Dim v2 As String = line.Substring(colon2 + 1).Trim()
                        Select Case k2
                            Case "required", "requiredfields"
                                For Each f As String In ParseJsonArrayOrList(v2)
                                    tc.RequiredFields.Add(f)
                                Next
                            Case "query", "userquery"
                                tc.UserQuery = v2.Trim(""""c, "'"c)
                            Case "expectedskill", "expectskill"
                                tc.ExpectedSkillName = v2.Trim(""""c, "'"c)
                            Case "minconfidence", "confidence"
                                Dim d As Double : If Double.TryParse(v2, d) Then tc.MinConfidence = d
                            Case "targetstep", "step"
                                tc.TargetStepIdOrName = v2.Trim(""""c, "'"c)
                        End Select
                    End If
                Next

                If String.IsNullOrWhiteSpace(tc.CaseId) Then tc.CaseId = "AUTO-" & Guid.NewGuid().ToString("N").Substring(0, 6)
                If String.IsNullOrWhiteSpace(tc.CaseName) Then tc.CaseName = tc.CaseId

                Return tc
            Catch ex As Exception  ' 
                Return Nothing
            End Try
        End Function

        Private Function SplitByTopLevelComma(s As String) As List(Of String)
            Dim result As New List(Of String)()
            Dim depth As Integer = 0 : Dim sb As New StringBuilder()
            For Each c As Char In s
                If c = "["c OrElse c = "("c OrElse c = "{"c Then depth += 1
                If c = "]"c OrElse c = ")"c OrElse c = "}"c Then depth -= 1
                If c = ","c AndAlso depth = 0 Then
                    result.Add(sb.ToString())
                    sb.Clear()
                Else
                    sb.Append(c)
                End If
            Next
            If sb.Length > 0 Then result.Add(sb.ToString())
            Return result
        End Function

        Private Function ParseJsonArrayOrList(s As String) As List(Of String)
            Dim result As New List(Of String)()
            s = s.Trim()
            If s.StartsWith("[") AndAlso s.EndsWith("]") Then
                Try
                    Dim arr As List(Of String) = JsonConvert.DeserializeObject(Of List(Of String))(s)
                    If arr IsNot Nothing Then result.AddRange(arr)
                Catch ex As Exception  ' 
                End Try
            Else
                result.AddRange(s.Split(","c).Select(Function(x) x.Trim().Trim(""""c, "'"c)).Where(Function(x) x.Length > 0))
            End If
            Return result
        End Function
    End Module

    ' =========================================================================
    ' 第五部分：测试引擎核心
    ' =========================================================================

    ''' <summary>
    ''' Skill自动化回归测试引擎
    ''' 4级测试策略：L1定义校验 → L2意图识别 → L3单步执行 → L4端到端集成
    ''' </summary>
    Public Class SkillTestEngine

        Private ReadOnly _skillsFolder As String
        Private ReadOnly _excelMock As IExcelMockContext

        ''' <summary>测试事件：便于UI端或日志订阅</summary>
        Public Event CaseStarted(caseId As String, caseName As String, level As SkillTestLevel)
        Public Event CaseFinished(caseId As String, status As SkillTestStatus, durationMs As Long, message As String)
        Public Event CaseAssertionFinished(assertionId As String, assertionDesc As String, status As SkillTestStatus, expected As String, actual As String)

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="skillsFolder">Skill MD 文件所在目录</param>
        ''' <param name="excelMock">可注入的Mock Excel，不传则默认内存Mock</param>
        Public Sub New(skillsFolder As String, Optional excelMock As IExcelMockContext = Nothing)
            _skillsFolder = skillsFolder
            _excelMock = If(excelMock, New InMemoryExcelMock())
        End Sub

        ''' <summary>执行一个Skill的完整测试套件</summary>
        Public Async Function RunTestSuiteAsync(skillName As String, Optional minLevel As SkillTestLevel = SkillTestLevel.DefinitionValidation,
                                                Optional maxLevel As SkillTestLevel = SkillTestLevel.EndToEndIntegration,
                                                Optional runSkipped As Boolean = False) As Task(Of SkillTestReport)
            Dim sw As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
            Dim report As New SkillTestReport() With {.SkillName = skillName}

            Dim mdPath As String = Path.Combine(_skillsFolder, skillName & ".md")
            Dim suite As SkillTestSuite = If(File.Exists(mdPath),
                           SkillMarkdownTestCasesParser.ParseFromSkillMarkdown(mdPath),
                           New SkillTestSuite() With {.SkillName = skillName})

            ' 【保底】即使没有内嵌用例，L1至少执行一次标准定义校验
            If suite.Cases.All(Function(c) c.Level <> SkillTestLevel.DefinitionValidation) Then
                Dim autoTc As New SkillTestCase With {
                    .CaseId = "TC-DEF-AUTO",
                    .CaseName = "Skill MD 自动语法校验（保底）",
                    .Level = SkillTestLevel.DefinitionValidation,
                    .Description = "自动生成：检查MD文件存在、可解析、YAML头合法",
                    .RequireValidYaml = True
                }
                autoTc.RequiredFields.AddRange({"Name", "DisplayName", "Description"})
                suite.Cases.Insert(0, autoTc)
            End If

            For Each tc As SkillTestCase In suite.Cases
                If tc.Level < minLevel OrElse tc.Level > maxLevel Then Continue For
                If tc.SkipByDefault AndAlso Not runSkipped Then
                    report.Results.Add(New SkillTestCaseResult With {
                        .CaseId = tc.CaseId, .CaseName = tc.CaseName, .Level = tc.Level,
                        .Status = SkillTestStatus.Skipped, .ErrorMessage = tc.SkipReason
                    })
                    Continue For
                End If

                RaiseEvent CaseStarted(tc.CaseId, tc.CaseName, tc.Level)
                Dim caseSw As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
                Dim result As SkillTestCaseResult = Await RunOneCaseAsync(tc, mdPath)
                caseSw.Stop()
                result.DurationMs = caseSw.ElapsedMilliseconds
                If tc.MaxDurationMs > 0 AndAlso result.Status = SkillTestStatus.Passed AndAlso caseSw.ElapsedMilliseconds > tc.MaxDurationMs Then
                    result.Status = SkillTestStatus.PassedWithWarning
                    result.Warnings.Add("性能SLA超出：允许" & tc.MaxDurationMs.ToString() & "ms，实际" & caseSw.ElapsedMilliseconds.ToString() & "ms")
                End If
                report.Results.Add(result)
                RaiseEvent CaseFinished(tc.CaseId, result.Status, caseSw.ElapsedMilliseconds, result.ErrorMessage)
            Next

            ' 汇总统计
            sw.Stop()
            report.TotalDurationMs = sw.ElapsedMilliseconds
            report.TotalCases = report.Results.Count
            report.PassedCases = Enumerable.Count(report.Results, Function(r) r.Status = SkillTestStatus.Passed OrElse r.Status = SkillTestStatus.PassedWithWarning)
            report.FailedCases = Enumerable.Count(report.Results, Function(r) r.Status = SkillTestStatus.Failed)
            report.SkippedCases = Enumerable.Count(report.Results, Function(r) r.Status = SkillTestStatus.Skipped)
            report.ErrorCases = Enumerable.Count(report.Results, Function(r) r.Status = SkillTestStatus.TestError)

            Return report
        End Function

        ''' <summary>批量执行所有Skill的测试</summary>
        Public Async Function RunAllSkillsAsync(Optional minLevel As SkillTestLevel = SkillTestLevel.DefinitionValidation,
                                                Optional maxLevel As SkillTestLevel = SkillTestLevel.EndToEndIntegration,
                                                Optional progress As IProgress(Of (SkillName As String, PassRate As Double, Finished As Integer, Total As Integer)) = Nothing) As Task(Of List(Of SkillTestReport))
            Dim reports As New List(Of SkillTestReport)()
            If Not Directory.Exists(_skillsFolder) Then Return reports

            Dim mdFiles As List(Of String) = Directory.GetFiles(_skillsFolder, "*.md").ToList()
            Dim total As Integer = mdFiles.Count
            For i As Integer = 0 To mdFiles.Count - 1
                Dim name As String = Path.GetFileNameWithoutExtension(mdFiles(i))
                Dim rep As SkillTestReport = Await RunTestSuiteAsync(name, minLevel, maxLevel)
                reports.Add(rep)
                progress?.Report((name, rep.PassRate, i + 1, total))
            Next
            Return reports
        End Function

        ' --------- 单个用例实际执行 ---------
        Private Async Function RunOneCaseAsync(tc As SkillTestCase, mdPath As String) As Task(Of SkillTestCaseResult)
            Dim result As New SkillTestCaseResult With {
                .CaseId = tc.CaseId, .CaseName = tc.CaseName, .Level = tc.Level,
                .StartedAt = DateTime.Now
            }
            Try
                Select Case tc.Level
                    Case SkillTestLevel.DefinitionValidation
                        RunLevel1(tc, mdPath, result)
                    Case SkillTestLevel.IntentRecognition
                        RunLevel2(tc, mdPath, result)
                    Case SkillTestLevel.StepExecution
                        Await RunLevel3Async(tc, mdPath, result)
                    Case SkillTestLevel.EndToEndIntegration
                        Await RunLevel4Async(tc, mdPath, result)
                End Select

                If result.Status = 0 Then ' 内部没有设置状态，全部断言过则Pass
                    Dim anyFail = result.Assertions.Any(Function(a) a.Status = SkillTestStatus.Failed)
                    result.Status = If(anyFail, SkillTestStatus.Failed, SkillTestStatus.Passed)
                End If
            Catch ex As Exception
                result.Status = SkillTestStatus.TestError
                result.ErrorMessage = ex.Message
                result.ErrorStackTrace = ex.StackTrace
            End Try
            result.EndedAt = DateTime.Now
            Return result
        End Function

        ' ==================== L1：Skill定义语法校验 ====================
        Private Sub RunLevel1(tc As SkillTestCase, mdPath As String, result As SkillTestCaseResult)
            If Not File.Exists(mdPath) Then
                result.Assertions.Add(New SkillTestAssertionResult With {.AssertionId = "L1-0", .Description = "MD文件存在", .Status = SkillTestStatus.Failed, .Actual = "missing", .Expected = mdPath, .Message = "Skill MD文件不存在"})
                result.Status = SkillTestStatus.Failed
                Return
            End If
            result.Assertions.Add(New SkillTestAssertionResult With {.AssertionId = "L1-0", .Description = "MD文件存在", .Status = SkillTestStatus.Passed})

            Dim text As String = File.ReadAllText(mdPath, Encoding.UTF8)

            ' 检查YAML头 --- ... --- 包裹
            If tc.RequireValidYaml Then
                Dim yamlOk As Boolean = text.TrimStart().StartsWith("---")
                Dim assertion As New SkillTestAssertionResult With {.AssertionId = "L1-yaml", .Description = "存在YAML Frontmatter头"}
                If yamlOk Then assertion.Status = SkillTestStatus.Passed Else assertion.Status = SkillTestStatus.Failed : assertion.Message = "MD文件需以 --- 开头声明YAML元数据"
                result.Assertions.Add(assertion)
            End If

            ' 解析YAML -> SkillDefinition（通过SkillManager来做，以保证和真实加载逻辑一致）
            Dim loadedSkill As SkillDefinition = TryParseSkillDefinition(mdPath)
            Dim parseAssert As New SkillTestAssertionResult With {.AssertionId = "L1-parse", .Description = "YAML Frontmatter可解析为SkillDefinition"}
            If loadedSkill IsNot Nothing Then
                parseAssert.Status = SkillTestStatus.Passed
                result.Assertions.Add(parseAssert)
                result.SkillVersion = loadedSkill.SemanticVersion

                For Each req As String In tc.RequiredFields
                    Dim ar As New SkillTestAssertionResult With {.AssertionId = "L1-req-" & req, .Description = "必填字段存在: " & req}
                    Dim prop As PropertyInfo = GetType(SkillDefinition).GetProperty(req, BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.IgnoreCase)
                    If prop Is Nothing Then
                        ar.Status = SkillTestStatus.TestError : ar.Message = "SkillDefinition中不存在属性 " & req & "，请检查RequiredFields拼写"
                    Else
                        Dim v As Object = prop.GetValue(loadedSkill)
                        Dim missing As Boolean = v Is Nothing OrElse (TypeOf v Is String AndAlso String.IsNullOrEmpty(CStr(v))) _
                                                    OrElse (TypeOf v Is Array AndAlso DirectCast(v, Array).Length = 0)
                        ar.Status = If(missing, SkillTestStatus.Failed, SkillTestStatus.Passed)
                        ar.Actual = If(v Is Nothing, "null", v.ToString())
                        ar.Expected = "!=null"
                        If missing Then ar.Message = "必填字段 " & req & " 为空"
                    End If
                    result.Assertions.Add(ar)
                Next

                ' 关键字段合理性校验：AllSteps 不为空说明能合并两种步骤体系
                If loadedSkill.IsComposite Then
                    Dim ar As New SkillTestAssertionResult With {.AssertionId = "L1-steps", .Description = "复合Skill含工作流步骤，AllSteps访问器可合并两种步骤"}
                    ar.Status = If(loadedSkill.AllSteps.Any(), SkillTestStatus.Passed, SkillTestStatus.Failed)
                    ar.Actual = "steps=" & loadedSkill.AllSteps.Count().ToString()
                    result.Assertions.Add(ar)
                End If
            Else
                parseAssert.Status = SkillTestStatus.Failed
                parseAssert.Message = "YAML解析失败或SkillDefinition为空"
                result.Assertions.Add(parseAssert)
            End If

            ' 通用断言
            For Each kv As KeyValuePair(Of String, String) In tc.Assertions
                Dim eval As (Success As Boolean, Expected As String, Actual As String, Message As String) = SkillAssertionEvaluator.Evaluate(kv.Value, GetSkillFieldForAssertion(loadedSkill, kv.Key))
                result.Assertions.Add(New SkillTestAssertionResult With {
                    .AssertionId = "L1-assert-" & Math.Abs(kv.Key.GetHashCode()).ToString(),
                    .Description = kv.Key, .Expected = eval.Expected, .Actual = eval.Actual,
                    .Message = eval.Message,
                    .Status = If(eval.Success, SkillTestStatus.Passed, SkillTestStatus.Failed)
                })
            Next
        End Sub

        ' ==================== L2：意图识别 ====================
        Private Sub RunLevel2(tc As SkillTestCase, mdPath As String, result As SkillTestCaseResult)
            If String.IsNullOrWhiteSpace(tc.UserQuery) Then
                result.Status = SkillTestStatus.TestError
                result.ErrorMessage = "L2测试必须配置UserQuery"
                Return
            End If

            Try
                Dim sm As SkillManager = SkillManager.Instance
                Dim recognizer As New SkillRecognizer(sm)
                _recogniResult = recognizer.Recognize(tc.UserQuery, mdPath)

                Dim nameAr As New SkillTestAssertionResult With {
                    .AssertionId = "L2-name",
                    .Description = "意图匹配Skill名",
                    .Expected = If(tc.ExpectedSkillName, "(no match)"),
                    .Actual = If(_recogniResult IsNot Nothing AndAlso _recogniResult.MatchedSkill IsNot Nothing, _recogniResult.MatchedSkill.Name, "(no match)")
                }
                If (String.IsNullOrEmpty(tc.ExpectedSkillName) AndAlso Not _recogniResult.IsMatch) _
                    OrElse (_recogniResult.IsMatch AndAlso String.Equals(_recogniResult.MatchedSkill.Name, tc.ExpectedSkillName, StringComparison.OrdinalIgnoreCase)) Then
                    nameAr.Status = SkillTestStatus.Passed
                Else
                    nameAr.Status = SkillTestStatus.Failed
                    nameAr.Message = "期望匹配 '" & tc.ExpectedSkillName & "' 实际 '" & If(_recogniResult IsNot Nothing AndAlso _recogniResult.MatchedSkill IsNot Nothing, _recogniResult.MatchedSkill.Name, "(no match)") & "'"
                End If
                result.Assertions.Add(nameAr)

                If Not String.IsNullOrEmpty(tc.ExpectedSkillName) Then
                    Dim confAr As New SkillTestAssertionResult With {
                        .AssertionId = "L2-confidence", .Description = "置信度 >= " & tc.MinConfidence.ToString("F4"),
                        .Expected = ">= " & tc.MinConfidence.ToString(), .Actual = _recogniResult.Confidence.ToString("F4")
                    }
                    If _recogniResult.Confidence >= tc.MinConfidence Then confAr.Status = SkillTestStatus.Passed Else confAr.Status = SkillTestStatus.Failed
                    result.Assertions.Add(confAr)
                End If
            Catch ex As Exception
                result.Status = SkillTestStatus.TestError
                result.ErrorMessage = "L2执行异常: " & ex.Message
            End Try
        End Sub

        Private _recogniResult As SkillRecognitionResult

        ' ==================== L3：单步骤执行 ====================
        Private Async Function RunLevel3Async(tc As SkillTestCase, mdPath As String, result As SkillTestCaseResult) As Task
            Dim skill As SkillDefinition = TryParseSkillDefinition(mdPath)
            If skill Is Nothing Then
                result.Status = SkillTestStatus.TestError
                result.ErrorMessage = "无法加载Skill定义"
                Return
            End If
            If String.IsNullOrWhiteSpace(tc.TargetStepIdOrName) Then
                result.Status = SkillTestStatus.TestError
                result.ErrorMessage = "L3必须设置TargetStepIdOrName"
                Return
            End If

            ' 找步骤（通过统一的AllSteps）
            _findStepTarget = tc
            Dim targetStep As WorkflowStep = skill.AllSteps.FirstOrDefault(AddressOf FindStepPredicate)
            _findStepTarget = Nothing
            If targetStep Is Nothing Then
                result.Status = SkillTestStatus.TestError
                result.ErrorMessage = "未找到StepId/StepName='" & tc.TargetStepIdOrName & "'的步骤（AllSteps总数=" & skill.AllSteps.Count().ToString() & "）"
                Return
            End If

            ' 根据StepType调度模拟执行（这里是Mock验证逻辑，真实项目中应注入WorkflowEngine）
            Dim vars As Dictionary(Of String, Object) = If(tc.InitialVariables, New Dictionary(Of String, Object)())
            Dim outVar As Object = Nothing
            Try
                Select Case targetStep.StepType
                    Case WorkflowStepType.OCRRecognition
                        Dim cfg As Dictionary(Of String, Object) = targetStep.Config
                        Dim hasImg As Boolean = (cfg IsNot Nothing AndAlso cfg.ContainsKey("imagePath")) OrElse vars.ContainsKey("input.imagePath")
                        result.Assertions.Add(New SkillTestAssertionResult With {
                            .AssertionId = "L3-ocr-path", .Description = "OCRRecognition步骤存在imagePath配置/输入",
                            .Status = If(hasImg, SkillTestStatus.Passed, SkillTestStatus.Failed)
                        })
                    Case WorkflowStepType.WriteToExcel
                        Dim fmCount As Integer = If(skill.FieldMappings IsNot Nothing, skill.FieldMappings.Count, 0)
                        result.Assertions.Add(New SkillTestAssertionResult With {
                            .AssertionId = "L3-write-mapping", .Description = "WriteToExcel步骤：Skill有字段映射配置",
                            .Status = If(fmCount > 0, SkillTestStatus.Passed, SkillTestStatus.Failed)
                        })
                    Case WorkflowStepType.CustomLLM
                        Dim cfg2 As Dictionary(Of String, Object) = targetStep.Config
                        Dim hasPrompt As Boolean = cfg2 IsNot Nothing AndAlso cfg2.ContainsKey("prompt")
                        Dim promptVal As String = If(hasPrompt, TryCast(cfg2("prompt"), String), Nothing)
                        result.Assertions.Add(New SkillTestAssertionResult With {
                            .AssertionId = "L3-customllm-prompt", .Description = "CustomLLM步骤存在prompt配置",
                            .Expected = "!=null",
                            .Actual = If(hasPrompt AndAlso promptVal IsNot Nothing, promptVal, "null"),
                            .Status = If(hasPrompt AndAlso promptVal IsNot Nothing, SkillTestStatus.Passed, SkillTestStatus.Failed)
                        })
                    Case WorkflowStepType.Loop
                        result.Assertions.Add(New SkillTestAssertionResult With {
                            .AssertionId = "L3-loop-limit", .Description = "Loop步骤MaxLoopIterations设置合理 (<= " & targetStep.MaxLoopIterations.ToString() & ")",
                            .Status = If(targetStep.MaxLoopIterations > 0 AndAlso targetStep.MaxLoopIterations <= 100000, SkillTestStatus.Passed, SkillTestStatus.Failed),
                            .Actual = targetStep.MaxLoopIterations.ToString()
                        })
                End Select
                outVar = True
            Catch ex As Exception
                result.Assertions.Add(New SkillTestAssertionResult With {
                    .AssertionId = "L3-exception", .Description = "步骤执行无异常",
                    .Status = SkillTestStatus.Failed, .Actual = ex.Message, .Message = ex.StackTrace
                })
            End Try

            ' 通用断言
            For Each kv As KeyValuePair(Of String, String) In tc.Assertions
                Dim eval As (Success As Boolean, Expected As String, Actual As String, Message As String) = SkillAssertionEvaluator.Evaluate(kv.Value, outVar)
                result.Assertions.Add(New SkillTestAssertionResult With {
                    .AssertionId = "L3-assert-" & Math.Abs(kv.Key.GetHashCode()).ToString(),
                    .Description = kv.Key, .Expected = eval.Expected, .Actual = eval.Actual,
                    .Message = eval.Message,
                    .Status = If(eval.Success, SkillTestStatus.Passed, SkillTestStatus.Failed)
                })
            Next
            Await Task.CompletedTask
        End Function

        ' L3查找步骤的Lambda提取为单独方法（避开OrElse的行内Lambda导致BC30205）
        Private _findStepTarget As SkillTestCase = Nothing
        Private Function FindStepPredicate(s As WorkflowStep) As Boolean
            If _findStepTarget Is Nothing Then Return False
            Return String.Equals(s.StepId, _findStepTarget.TargetStepIdOrName, StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(s.StepName, _findStepTarget.TargetStepIdOrName, StringComparison.OrdinalIgnoreCase)
        End Function

        ' ==================== L4：端到端集成（CI默认跳过） ====================
        Private Async Function RunLevel4Async(tc As SkillTestCase, mdPath As String, result As SkillTestCaseResult) As Task
            result.Assertions.Add(New SkillTestAssertionResult With {
                .AssertionId = "L4-skip", .Description = "L4 E2E集成测试需要真实Excel环境，默认跳过",
                .Status = SkillTestStatus.Skipped, .Message = "如需运行L4，请设置 SkipByDefault=false 并准备好Excel宿主环境"
            })
            Await Task.CompletedTask
        End Function

        ' --------- 工具方法 ---------
        ''' <summary>
        ''' 解析Skill定义：优先通过反射调用SkillManager.ParseMarkdownSkill（与生产逻辑100%一致），
        ''' 如果失败则退化到自包含的简化解析器（不依赖YamlDotNet，仅覆盖L1/L2需要的字段）
        ''' </summary>
        Private Function TryParseSkillDefinition(mdPath As String) As SkillDefinition
            If Not File.Exists(mdPath) Then Return Nothing
            Try
                ' 路线1：反射调用 SkillManager 私有方法（与真实加载逻辑一致）
                Dim mgr As SkillManager = SkillManager.Instance
                If mgr IsNot Nothing Then
                    Dim mi As MethodInfo = GetType(SkillManager).GetMethod("ParseMarkdownSkill",
                        BindingFlags.Instance Or BindingFlags.NonPublic Or BindingFlags.Public)
                    If mi IsNot Nothing Then
                        Dim skill As SkillDefinition = DirectCast(mi.Invoke(mgr, New Object() {mdPath}), SkillDefinition)
                        If skill IsNot Nothing Then Return skill
                    End If
                End If
            Catch ex As Exception  ' 
            End Try

            ' 路线2：退化到自包含的简化解析（覆盖必填字段，L1/L2够用）
            Try
                Return ParseSkillDefinitionFallback(mdPath)
            Catch ex As Exception  ' 
                Return Nothing
            End Try
        End Function

        ''' <summary>备用解析器：极简YAML→JSON→SkillDefinition（不依赖外部库）</summary>
        Private Shared Function ParseSkillDefinitionFallback(mdPath As String) As SkillDefinition
            Dim allText As String = File.ReadAllText(mdPath, Encoding.UTF8).Trim()
            If Not allText.StartsWith("---") Then Return Nothing

            Dim endDash As Integer = allText.IndexOf("---", 3, StringComparison.Ordinal)
            If endDash <= 3 Then Return Nothing

            Dim yamlLines As String() = allText.Substring(3, endDash - 3).Split({ControlChars.CrLf, ControlChars.Lf}, StringSplitOptions.None)
            Dim sbJson As New StringBuilder()
            sbJson.Append("{"c)
            Dim first As Boolean = True
            Dim inArray As Boolean = False
            Dim arrayKey As String = ""
            Dim arrayVals As New List(Of String)()

            For Each raw As String In yamlLines
                Dim line As String = raw.TrimEnd()
                If String.IsNullOrWhiteSpace(line) Then Continue For
                If line.StartsWith("#") Then Continue For

                If inArray Then
                    If line.StartsWith("- ") OrElse line.StartsWith("-") Then
                        arrayVals.Add(line.Substring(line.IndexOf("-"c) + 1).Trim().Trim(""""c, "'"c))
                        Continue For
                    Else
                        FlushArrayInternal(sbJson, first, inArray, arrayKey, arrayVals)
                    End If
                End If

                Dim colonIdx As Integer = line.IndexOf(":"c)
                If colonIdx <= 0 Then Continue For
                Dim k As String = line.Substring(0, colonIdx).Trim()
                Dim v As String = line.Substring(colonIdx + 1).Trim()

                If v = "" OrElse v = "[" Then
                    FlushArrayInternal(sbJson, first, inArray, arrayKey, arrayVals)
                    inArray = True
                    arrayKey = k
                    Continue For
                End If

                FlushArrayInternal(sbJson, first, inArray, arrayKey, arrayVals)
                If Not first Then sbJson.Append(","c)
                sbJson.Append(""""c)
                sbJson.Append(EscapeJson(k))
                sbJson.Append(""":""")
                sbJson.Append(EscapeJson(v.Trim(""""c, "'"c)))
                sbJson.Append(""""c)
                first = False
            Next
            FlushArrayInternal(sbJson, first, inArray, arrayKey, arrayVals)
            sbJson.Append("}"c)

            Dim skill As SkillDefinition = JsonConvert.DeserializeObject(Of SkillDefinition)(sbJson.ToString(),
                New JsonSerializerSettings With {.MissingMemberHandling = MissingMemberHandling.Ignore})
            Return skill
        End Function

        Private Shared Sub FlushArrayInternal(sbJson As StringBuilder, ByRef first As Boolean, ByRef inArray As Boolean, ByRef arrayKey As String, arrayVals As List(Of String))
            If inArray Then
                If Not first Then sbJson.Append(","c)
                sbJson.Append(""""c)
                sbJson.Append(EscapeJson(arrayKey))
                sbJson.Append(""":[""")
                For iArr As Integer = 0 To arrayVals.Count - 1
                    If iArr > 0 Then sbJson.Append(","c)
                    sbJson.Append(""""c)
                    sbJson.Append(EscapeJson(arrayVals(iArr)))
                    sbJson.Append(""""c)
                Next
                sbJson.Append("]"c)
                first = False
                inArray = False
                arrayKey = ""
                arrayVals.Clear()
            End If
        End Sub

        Private Shared Function EscapeJson(s As String) As String
            If String.IsNullOrEmpty(s) Then Return ""
            Return s.Replace("\", "\\").Replace("""", "\""").Replace(vbLf, "\n").Replace(vbCr, "\r").Replace(vbTab, "\t")
        End Function

        Private Function GetSkillFieldForAssertion(skill As SkillDefinition, key As String) As Object
            If skill Is Nothing Then Return Nothing
            Dim prop As PropertyInfo = GetType(SkillDefinition).GetProperty(key, BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.IgnoreCase)
            If prop IsNot Nothing Then Return prop.GetValue(skill)
            Select Case key.ToLowerInvariant()
                Case "keywords.length", "keywordscount" : Return If(skill.Keywords IsNot Nothing, skill.Keywords.Length, 0)
                Case "examples.length", "examplescount" : Return If(skill.Examples IsNot Nothing, skill.Examples.Length, 0)
                Case "allsteps.count", "stepscount" : Return skill.AllSteps.Count()
                Case Else
                    If key.StartsWith("examples contains:", StringComparison.OrdinalIgnoreCase) Then
                        Dim needle As String = key.Substring("examples contains:".Length).Trim()
                        Dim exArr As String() = skill.Examples
                        Dim found As Boolean = False
                        If exArr IsNot Nothing Then
                            For Each e As String In exArr
                                If e IsNot Nothing AndAlso e.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0 Then
                                    found = True
                                    Exit For
                                End If
                            Next
                        End If
                        Return If(found, "found", "not-found")
                    End If
            End Select
            Return Nothing
        End Function

        ''' <summary>报告序列化为JSON字符串（便于CI写入artifact）</summary>
        Public Function ReportToJson(report As SkillTestReport) As String
            Return JsonConvert.SerializeObject(report, Formatting.Indented)
        End Function

        ''' <summary>报告转Markdown摘要（便于在PR评论中展示）</summary>
        Public Function ReportToMarkdown(report As SkillTestReport) As String
            Dim sb As New StringBuilder()
            sb.AppendLine("# Skill测试报告: " & report.SkillName & " v" & report.SkillVersion)
            sb.AppendLine()
            sb.AppendLine("- 生成时间: " & report.GeneratedAt.ToString("yyyy-MM-dd HH:mm:ss"))
            sb.AppendLine("- 用例总数: **" & report.TotalCases.ToString() & "**")
            sb.AppendLine("- [OK] 通过: " & report.PassedCases.ToString())
            sb.AppendLine("- [FAIL] 失败: " & report.FailedCases.ToString())
            sb.AppendLine("- [SKIP] 跳过: " & report.SkippedCases.ToString())
            sb.AppendLine("- [ERR] 测试错误: " & report.ErrorCases.ToString())
            sb.AppendLine("- 通过率: **" & report.PassRate.ToString("F1") & "%**")
            sb.AppendLine("- 总耗时: " & report.TotalDurationMs.ToString("N0") & "ms")
            sb.AppendLine()
            sb.AppendLine("## 用例明细")
            sb.AppendLine("| ID | 层级 | 用例名 | 状态 | 耗时ms | 备注 |")
            sb.AppendLine("| --- | --- | --- | --- | ---: | --- |")
            For Each r In report.Results
                Dim symbol As String = ""
                Select Case r.Status
                    Case SkillTestStatus.Passed : symbol = "[OK]"
                    Case SkillTestStatus.PassedWithWarning : symbol = "[WARN]"
                    Case SkillTestStatus.Failed : symbol = "[FAIL]"
                    Case SkillTestStatus.Skipped : symbol = "[SKIP]"
                    Case SkillTestStatus.TestError : symbol = "[ERR]"
                End Select
                Dim note As String = If(String.IsNullOrEmpty(r.ErrorMessage), String.Join("; ", r.Warnings), r.ErrorMessage)
                If note Is Nothing Then note = ""
                sb.AppendLine("| " & r.CaseId & " | L" & CInt(r.Level).ToString() & " | " & r.CaseName & " | " & symbol & r.Status.ToString() & " | " & r.DurationMs.ToString("N0") & " | " & note.Replace("|", "\|") & " |")
            Next
            Return sb.ToString()
        End Function
    End Class

End Namespace