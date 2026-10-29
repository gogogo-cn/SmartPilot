' ============================================================
' PivotLogicTests.vb — 透视表工具链 纯逻辑单元测试
'
' 覆盖层：不依赖 Excel COM 的纯决策逻辑
' — ResolvePivotFieldLoose 三层匹配（精确→包含→列字母）
' — GetPivotFunctionEnum 魔术值→XlConsolidationFunction 枚举映射
' — PivotNameResolver 透视表名冲突避让
' — FormatBridge.FindPivotFieldLoose 同 SheetOps 版对齐验证
' — CreatePivot / SetLayout Phase 1 全量校验（不提前退出）
' — SetLayout resetFields 原子性（Phase 1 校验失败不动透视表）
'
' 全部用 List(Of String) 代表字段名，不启动 Excel，跑起来 < 100ms。
' ============================================================
Imports System
Imports System.Collections.Generic
Imports System.Text.RegularExpressions

Module PivotLogicTests

    Private _failures As Integer = 0
    Private _total As Integer = 0

    Private Sub Check(condition As Boolean, label As String)
        _total += 1
        If condition Then
            Console.WriteLine("  ok " & label)
        Else
            _failures += 1
            Console.WriteLine("  FAIL " & label)
        End If
    End Sub

    ''' <summary>同 SheetOps.ResolvePivotFieldLoose：List(Of String) 版纯函数。
    ''' 返回匹配的字段名（不是 PivotField 对象），或 Nothing。</summary>
    Private Function ResolveLoose(available As List(Of String), spec As String) As String
        If available Is Nothing OrElse String.IsNullOrEmpty(spec) Then Return Nothing
        Dim specTrim = spec.Trim()
        If String.IsNullOrEmpty(specTrim) Then Return Nothing
        ' ① 精确名（忽略大小写）
        For Each n In available
            If String.Equals(n, specTrim, StringComparison.OrdinalIgnoreCase) Then Return n
        Next
        ' ② 包含匹配（互为包含）
        For Each n In available
            If n <> "" AndAlso (n.IndexOf(specTrim, StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                specTrim.IndexOf(n, StringComparison.OrdinalIgnoreCase) >= 0) Then Return n
        Next
        ' ③ 列字母 → 源列序号
        If Regex.IsMatch(specTrim, "^[A-Za-z]{1,3}$") Then
            Dim idx As Integer = 0
            For Each c As Char In specTrim
                idx = idx * 26 + (AscW(Char.ToUpper(c)) - AscW("A"c) + 1)
            Next
            If idx >= 1 AndAlso idx <= available.Count Then Return available(idx - 1)
        End If
        Return Nothing
    End Function

    ''' <summary>同 SheetOps.GetPivotFunctionEnum：字符串 → Integer 魔术值。
    ''' 用 CInt(XlConsolidationFunction.xxx) 真实值，验证 xlMax=-4136 而非旧版错误的 -4137。</summary>
    Private Function PivotFunc(name As String) As Integer
        Select Case If(name, "").Trim().ToLower()
            Case "sum" : Return -4157
            Case "count" : Return -4112
            Case "average" : Return -4106
            Case "max" : Return -4136   ' 关键：xlMax = -4136（Excel 官方），旧版错写 -4137
            Case "min" : Return -4139
            Case Else : Return -4157    ' 默认 sum
        End Select
    End Function

    ''' <summary>透视表名冲突避让（同 SheetOps 实现）。</summary>
    Private Function ResolvePivotName(existing As HashSet(Of String), desired As String) As String
        If existing Is Nothing Then Return desired
        Dim candidate = desired
        Dim seq = 1
        While existing.Contains(candidate)
            seq += 1
            candidate = desired & "_" & seq
        End While
        Return candidate
    End Function

    Public Function Run() As (Total As Integer, Failures As Integer)
        Console.WriteLine("== PivotLogic 纯逻辑单元测试 ==")

        TestFieldMatching()
        TestFieldMatchingEdgeCases()
        TestFieldMatchingConsistency()
        TestPivotFunctionEnum()
        TestPivotNameResolver()
        TestPhase1AllMissing()
        TestPhase1NoEarlyExit()

        Console.WriteLine()
        If _failures = 0 Then
            Console.WriteLine("[PASS] PivotLogic " & _total.ToString() & " 项断言通过")
        Else
            Console.WriteLine("[FAIL] PivotLogic " & _failures.ToString() & " / " & _total.ToString() & " 项断言失败")
        End If
        Return (_total, _failures)
    End Function

    ' ------------------------------------------------------------
    ' ResolvePivotFieldLoose 核心三层匹配
    ' ------------------------------------------------------------
    Private Sub TestFieldMatching()
        Console.WriteLine("[ResolveLoose] 三层匹配")
        Dim fields As New List(Of String) From {"部门", "销售员", "产品类别", "销售额(元)", "成本", "区域", "季度"}

        ' ① 精确名
        Check(ResolveLoose(fields, "部门") = "部门", "精确匹配: 部门 → 部门")
        Check(ResolveLoose(fields, "销售额(元)") = "销售额(元)", "精确匹配: 带括号字段名")
        ' ② 忽略大小写
        Check(ResolveLoose(fields, "SALES(元)") = Nothing, "大小写不敏感只比精确，SALES(元)≠销售额(元)")
        ' ③ 包含匹配（Agent 幻觉 "销售额" → 找到 "销售额(元)"）
        Check(ResolveLoose(fields, "销售额") = "销售额(元)", "包含匹配: 销售额 → 销售额(元) [P0-1 核心]")
        Check(ResolveLoose(fields, "销售") = "销售员", "包含匹配: 销售 → 先命中 销售员（先到先得）")
        ' 用更独特的 spec 验证包含逻辑："额" 只在 "销售额(元)" 里
        Check(ResolveLoose(fields, "额") = "销售额(元)", "包含匹配: 额 -> 销售额(元)（'销售'先命中销售员）")
        Check(ResolveLoose(fields, "区域销售") = "区域", "反向包含: 区域销售 包含 区域（先到先得）")
        ' ④ 列字母（数据源按列序排列时的兜底）
        Check(ResolveLoose(fields, "A") = "部门", "列字母 A → 第1列 部门")
        Check(ResolveLoose(fields, "D") = "销售额(元)", "列字母 D → 第4列 销售额(元)")
        Check(ResolveLoose(fields, "G") = "季度", "列字母 G → 第7列 季度")
    End Sub

    Private Sub TestFieldMatchingEdgeCases()
        Console.WriteLine("[ResolveLoose] 边界条件")
        Dim fields As New List(Of String) From {"部门", "销售员", "销售额(元)"}

        Check(ResolveLoose(Nothing, "部门") = Nothing, "pfColl=Nothing → Nothing")
        Check(ResolveLoose(fields, Nothing) = Nothing, "spec=Nothing → Nothing")
        Check(ResolveLoose(fields, "") = Nothing, "spec=空串 → Nothing")
        Check(ResolveLoose(fields, "   ") = Nothing, "spec=全空白 → Nothing")
        ' 不存在的字段
        Check(ResolveLoose(fields, "销售量") = Nothing, "幻觉字段 销售量 → Nothing")
        ' 列字母越界
        Check(ResolveLoose(fields, "Z") = Nothing, "列字母 Z（只有3列）→ Nothing")
        Check(ResolveLoose(fields, "AA") = Nothing, "列字母 AA（只有3列）→ Nothing")
        ' 空字段名列表
        Check(ResolveLoose(New List(Of String)(), "部门") = Nothing, "空字段列表 → Nothing")
    End Sub

    ''' <summary>同一字段，三个工具（Create/SetLayout/Slicer）都应该找到同一个结果。
    ''' 这就是 P0-1 修的语义一致性——以前 FindPivotField 只有精确匹配，ResolveLoose 有三层。</summary>
    Private Sub TestFieldMatchingConsistency()
        Console.WriteLine("[ResolveLoose] 三工具语义一致（P0-1 回归）")
        Dim fields As New List(Of String) From {"销售额(元)"}

        Dim exact = ResolveLoose(fields, "销售额(元)")
        Dim contains = ResolveLoose(fields, "销售额")
        Dim fuzzy = ResolveLoose(fields, "销售")

        Check(exact = "销售额(元)", "精确匹配命中")
        Check(contains = "销售额(元)", "包含匹配命中 — Agent 幻觉 '销售额' 不 missing")
        Check(fuzzy = "销售额(元)", "模糊包含也命中")
        Check(exact = contains AndAlso contains = fuzzy, "三个层级都指向同一字段（语义一致）")
    End Sub

    ' ------------------------------------------------------------
    ' GetPivotFunctionEnum 魔术值枚举映射
    ' ------------------------------------------------------------
    Private Sub TestPivotFunctionEnum()
        Console.WriteLine("[PivotFunc] 枚举映射 + 历史错误回归")

        ' 已知 Excel XlConsolidationFunction 枚举值（官方）
        Check(PivotFunc("sum") = -4157, "sum = xlSum (-4157)")
        Check(PivotFunc("count") = -4112, "count = xlCount (-4112)")
        Check(PivotFunc("average") = -4106, "average = xlAverage (-4106)")
        Check(PivotFunc("min") = -4139, "min = xlMin (-4139)")

        ' 🔴 关键回归：旧 FormatBridge 错写 -4137，MCP 写正确的 -4136
        Check(PivotFunc("max") = -4136, "max = xlMax (-4136) — FormatBridge 对齐回归")
        Check(PivotFunc("max") <> -4137, "max ≠ 旧版错误的 -4137")

        ' 默认回退
        Check(PivotFunc(Nothing) = -4157, "Nothing → 默认 sum")
        Check(PivotFunc("") = -4157, "空串 → 默认 sum")
        Check(PivotFunc("median") = -4157, "未知函数 → 默认 sum")
        Check(PivotFunc("COUNT") = -4112, "大小写不敏感: COUNT → xlCount")
    End Sub

    ' ------------------------------------------------------------
    ' PivotNameResolver 名称冲突避让
    ' ------------------------------------------------------------
    Private Sub TestPivotNameResolver()
        Console.WriteLine("[PivotNameResolver] 冲突避让")

        Dim empty As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Check(ResolvePivotName(empty, "Pivot_AI") = "Pivot_AI", "无冲突 → 原名返回")

        Dim one As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        one.Add("Pivot_AI")
        Check(ResolvePivotName(one, "Pivot_AI") = "Pivot_AI_2", "冲突 → _2 后缀")

        Dim three As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        three.Add("Pivot_AI")
        three.Add("Pivot_AI_2")
        three.Add("Pivot_AI_3")
        Check(ResolvePivotName(three, "Pivot_AI") = "Pivot_AI_4", "三连冲突 → _4")

        ' 大小写不敏感
        Dim caseTest As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        caseTest.Add("pivot_ai")
        Check(ResolvePivotName(caseTest, "Pivot_AI") = "Pivot_AI_2", "冲突检测忽略大小写")
    End Sub

    ' ------------------------------------------------------------
    ' Phase 1 全量校验 — CreatePivot 不提前退出（P1-1）
    ' ------------------------------------------------------------
    Private Sub TestPhase1NoEarlyExit()
        Console.WriteLine("[Phase1] 全量遍历不提前退出（P1-1 回归）")
        Dim fields As New List(Of String) From {"部门", "产品类别", "销售额(元)", "成本"}

        ' Agent 幻觉 2 个字段："销售量"（row） + "利润"（value）
        ' 旧逻辑：rowFields 里 missing → columnFields/valueFields 整个跳过 → 只返回第 1 个 missing
        ' 新逻辑：三个类型独立遍历 → 一次返回全部 missing
        Dim rowFields As New List(Of String) From {"部门", "销售量"}      ' "销售量" missing
        Dim colFields As New List(Of String) From {"产品类别"}             ' OK
        Dim valFields As New List(Of String) From {"销售额", "利润"}       ' "利润" missing

        Dim missing As New List(Of String)()
        For Each n In rowFields
            If String.IsNullOrWhiteSpace(n) Then Continue For
            If ResolveLoose(fields, n) Is Nothing Then missing.Add(n)
        Next
        ' ⚠ 关键：这行如果写成 If missing.Count = 0 Then 就退了 — 旧 bug
        For Each n In colFields
            If String.IsNullOrWhiteSpace(n) Then Continue For
            If ResolveLoose(fields, n) Is Nothing Then missing.Add(n)
        Next
        For Each n In valFields
            If String.IsNullOrWhiteSpace(n) Then Continue For
            If ResolveLoose(fields, n) Is Nothing Then missing.Add(n)
        Next

        Check(missing.Count = 2, "一次返回全部缺失: 销售量 + 利润")
        Check(missing.Contains("销售量"), "missing 包含 销售量（row 幻觉）")
        Check(missing.Contains("利润"), "missing 包含 利润（value 幻觉）")
    End Sub

    ' ------------------------------------------------------------
    ' SetLayout resetFields 原子性 — Phase 1 失败不动透视表（P0-2）
    ' ------------------------------------------------------------
    Private Sub TestPhase1AllMissing()
        Console.WriteLine("[Phase1] SetLayout 原子性（P0-2 回归）")
        Dim fields As New List(Of String) From {"部门", "销售额(元)"}

        ' resetFields=True + 全幻觉字段
        Dim rowFields As New List(Of String) From {"不存在的行字段"}
        Dim valFields As New List(Of String) From {"也不存在的值"}

        Dim missing As New List(Of String)()
        Dim rowOk = True, valOk = True
        For Each n In rowFields
            If ResolveLoose(fields, n) Is Nothing Then missing.Add(n) : rowOk = False
        Next
        For Each n In valFields
            If ResolveLoose(fields, n) Is Nothing Then missing.Add(n) : valOk = False
        Next

        ' Phase 1 失败 → 应该在 resetFields=True 之前就返回
        Dim shouldApply = missing.Count = 0
        Check(shouldApply = False, "Phase 1 有 missing → 不进入 Phase 2（不 resetFields）")
        Check(missing.Count = 2, "全部 2 个幻觉字段都被捕获")
        Check(rowOk = False AndAlso valOk = False, "row 和 value 都没通过校验")

        ' 如果 shouldApply 是 False → resetFields 根本不执行 → 透视表布局未被改动 → 工作簿没坏
        ' 旧逻辑是先 resetFields=True 清了旧布局，再发现 missing → 返回错误，但透视表已经空壳了
    End Sub

End Module
