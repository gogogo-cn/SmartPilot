' ============================================================
' PptDesignEngineTests.vb — 数据感知设计引擎（纯决策链）单元测试
'
' 被测对象：直接编译的元数据生产代码 Services/PptDesignEngine.vb
' — 这是生成 PPT 叙事结构 / 版式 / 图表 / 洞察的确定性规则引擎
'   （数据画像 → 版式序列），不依赖 Excel COM，可离线单测。
' — 之前该决策链仅有 JS 静态源码 sanity 检查，无行为级单元测试；
'   本套件补齐行为断言（列分类、趋势、ID 识别、方向、图表类型、
'   洞察文案、需求偏置、column hints 覆盖等）。
' ============================================================
Imports System
Imports System.Collections.Generic
Imports ExcelAIVSTOAddIn.Services
Imports Newtonsoft.Json.Linq

Module PptDesignEngineTests

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

    ''' <summary>由表头+等长行值数组构造行字典列表</summary>
    Private Function BuildRows(headers As String(), data As String()()) As List(Of Dictionary(Of String, Object))
        Dim rows As New List(Of Dictionary(Of String, Object))()
        For Each vals As String() In data
            Dim row As New Dictionary(Of String, Object)()
            For i As Integer = 0 To headers.Length - 1
                row(headers(i)) = vals(i)
            Next
            rows.Add(row)
        Next
        Return rows
    End Function

    Public Function Run() As (Total As Integer, Failures As Integer)
        Console.WriteLine("== PptDesignEngine 决策链单元测试 ==")

        TestAnalyzeDataNumericAndTrend()
        TestDateSeriesAndLineChart()
        TestIdColumnAndHints()
        TestRecommendOrientation()
        TestRecommendChartType()
        TestGenerateInsights()
        TestRequirementBias()
        TestAutoComparisonFromData()
        TestE2eSalesReport()

        Console.WriteLine()
        If _failures = 0 Then
            Console.WriteLine("[PASS] PptDesignEngine " & _total.ToString() & " 项断言通过")
        Else
            Console.WriteLine("[FAIL] PptDesignEngine " & _failures.ToString() & " / " & _total.ToString() & " 项断言失败")
        End If
        Return (_total, _failures)
    End Function

    ' ------------------------------------------------------------
    ' 列分类 + 数值统计 + 趋势 + 主指标识别
    ' ------------------------------------------------------------
    Private Sub TestAnalyzeDataNumericAndTrend()
        Console.WriteLine("[AnalyzeData] 数值列统计与趋势")
        Dim headers As String() = {"月份", "销售额", "地区"}
        Dim data As String()() = {
            New String() {"一月", "100", "华东"},
            New String() {"二月", "200", "华东"},
            New String() {"三月", "300", "华东"},
            New String() {"四月", "400", "华东"}}
        Dim rows = BuildRows(headers, data)
        Dim profile = PptDesignEngine.AnalyzeData(New List(Of String)(headers), rows)

        Dim sales As PptColumnProfile = profile.Columns.Find(Function(c) c.Name = "销售额")
        Check(sales IsNot Nothing, "识别到数值列「销售额」")
        Check(sales.Kind = "numeric", "「销售额」分类为 numeric")
        Check(sales.HasNumeric AndAlso sales.Sum = 1000, "「销售额」合计=1000（实际 " & sales.Sum.ToString() & "）")
        Check(sales.Average = 250, "「销售额」平均=250")
        Check(sales.Min = 100 AndAlso sales.Max = 400, "「销售额」区间 100~400")
        Check(sales.TrendDirection = "up", "「销售额」趋势=up")
        Check(profile.RowCount = 4, "RowCount=4")

        Dim region As PptColumnProfile = profile.Columns.Find(Function(c) c.Name = "地区")
        Check(region IsNot Nothing AndAlso region.Kind = "categorical", "「地区」分类为 categorical")
        Check(profile.MainNumericColumn IsNot Nothing AndAlso profile.MainNumericColumn.Name = "销售额", "主数值列=销售额")
    End Sub

    Private Sub TestDateSeriesAndLineChart()
        Console.WriteLine("[DesignFromData] 日期时间序列 → 折线图")
        Dim headers As String() = {"日期", "销量"}
        Dim data As String()() = {
            New String() {"2024-01-01", "10"},
            New String() {"2024-02-01", "20"},
            New String() {"2024-03-01", "30"}}
        Dim rows = BuildRows(headers, data)
        Dim slides = PptDesignEngine.DesignFromData(New List(Of String)(headers), rows, "销售报告", "测试数据")

        Check(slides.Type = JTokenType.Array AndAlso slides.Count > 0, "生成幻灯片序列（" & slides.Count.ToString() & " 页）")
        Check(slides(0)("type").ToString() = "title", "首页=title 封面")
        Dim types = slides.Select(Function(s) s.Value(Of String)("type")).ToList()
        Check(types.Contains("agenda"), "含目录页")
        Check(types.Contains("kpi"), "数值列存在 → 含 KPI 页")
        Check(types.Contains("closing"), "含结尾页")
        Dim chart = slides.FirstOrDefault(Function(s) s.Value(Of String)("type") = "chart")
        Check(chart IsNot Nothing, "时间序列 → 含图表页")
        If chart IsNot Nothing Then
            Check(chart.Value(Of String)("chartType") = "line", "时间序列图表类型=line")
            Check(chart.Value(Of String)("insight").Contains("上升"), "图表情报含趋势上升")
        End If
    End Sub

    Private Sub TestIdColumnAndHints()
        Console.WriteLine("[AnalyzeData + ApplyColumnHints] ID 识别与 column hints 覆盖")
        Dim headers As String() = {"订单ID", "销售额", "利润"}
        Dim data As String()() = {
            New String() {"A001", "100", "10"},
            New String() {"A002", "200", "20"},
            New String() {"A003", "300", "30"}}
        Dim rows = BuildRows(headers, data)
        Dim profile = PptDesignEngine.AnalyzeData(New List(Of String)(headers), rows)

        Dim idCol As PptColumnProfile = profile.Columns.Find(Function(c) c.Name = "订单ID")
        Check(idCol IsNot Nothing AndAlso idCol.IsIdLike, "「订单ID」按名称识别为 ID 列")
        Check(profile.MainNumericColumn IsNot Nothing AndAlso profile.MainNumericColumn.Name = "销售额", "主数值列默认=销售额（跳过 ID 列）")

        ' hints 覆盖主指标
        Dim hints As New JObject()
        hints("mainMetric") = "利润"
        PptDesignEngine.ApplyColumnHints(profile, hints)
        Check(profile.MainNumericColumn IsNot Nothing AndAlso profile.MainNumericColumn.Name = "利润", "hints.mainMetric=利润 覆盖主数值列")

        ' hints 空对象应幂等不抛错
        PptDesignEngine.ApplyColumnHints(profile, New JObject())
        Check(profile.MainNumericColumn IsNot Nothing AndAlso profile.MainNumericColumn.Name = "利润", "空 hints 幂等，主数值列不变")
    End Sub

    Private Sub TestRecommendOrientation()
        Console.WriteLine("[RecommendOrientation] 版式方向")
        Check(PptDesignEngine.RecommendOrientation(New List(Of String)(New String() {"a", "b", "c", "d", "e", "f"}), 5) = "landscape", "6 列宽表 → landscape")
        Check(PptDesignEngine.RecommendOrientation(New List(Of String)(New String() {"a", "b"}), 30) = "portrait", "2 列 30 行长表 → portrait")
        Check(PptDesignEngine.RecommendOrientation(New List(Of String)(New String() {"a", "b"}), 3) = "landscape", "行数少 → landscape")
        Check(PptDesignEngine.RecommendOrientation(Nothing, 3) = "landscape", "无表头且行数少 → landscape")
    End Sub

    Private Sub TestRecommendChartType()
        Console.WriteLine("[RecommendChartType] 图表类型")
        Dim dateCol As New PptColumnProfile() With {.Kind = "date"}
        Check(PptDesignEngine.RecommendChartType(Nothing, dateCol, Nothing) = "line", "时间列 → line")

        Dim catFew As New PptColumnProfile() With {.Kind = "categorical", .UniqueCount = 4}
        Check(PptDesignEngine.RecommendChartType(Nothing, catFew, Nothing) = "pie", "分类少(4) → pie")

        Dim catMany As New PptColumnProfile() With {.Kind = "categorical", .UniqueCount = 20}
        Check(PptDesignEngine.RecommendChartType(Nothing, catMany, Nothing) = "column", "分类多(20) → column")

        Check(PptDesignEngine.RecommendChartType(Nothing, Nothing, Nothing) = "column", "无 x 列 → column 兜底")
    End Sub

    Private Sub TestGenerateInsights()
        Console.WriteLine("[GenerateInsights] 洞察文案")
        Dim headers As String() = {"日期", "销量"}
        Dim data As String()() = {
            New String() {"2024-01-01", "10"},
            New String() {"2024-02-01", "20"},
            New String() {"2024-03-01", "30"}}
        Dim rows = BuildRows(headers, data)
        Dim profile As PptDataProfile = PptDesignEngine.AnalyzeData(New List(Of String)(headers), rows)
        Dim lines As List(Of String) = PptDesignEngine.GenerateInsights(profile)
        Check(lines.Count > 0, "数值列存在 → 生成洞察文案")
        Check(lines.Exists(Function(l) l.Contains("合计")), "洞察含「合计」表述")
        Check(lines.Exists(Function(l) l.Contains("上升趋势")), "洞察含「上升趋势」")

        Dim emptyProfile As New PptDataProfile()
        Dim emptyLines As List(Of String) = PptDesignEngine.GenerateInsights(emptyProfile)
        Check(emptyLines.Count = 1 AndAlso emptyLines(0).Contains("未检测到可量化"), "无数值列 → 兜底提示")
    End Sub

    Private Sub TestRequirementBias()
        Console.WriteLine("[DesignFromData] 需求偏置")
        ' 长表 + 类别 + 数值：默认应含"数据明细"；显式"不要明细"则裁剪
        Dim headers As String() = {"类别", "金额"}
        Dim data As New List(Of String())()
        For i As Integer = 1 To 15
            data.Add(New String() {("类" & (i Mod 3 + 1).ToString()), (i * 10).ToString()})
        Next
        Dim rows As New List(Of Dictionary(Of String, Object))()
        For Each vals As String() In data
            Dim row As New Dictionary(Of String, Object)()
            row("类别") = vals(0)
            row("金额") = vals(1)
            rows.Add(row)
        Next

        Dim defSlides = PptDesignEngine.DesignFromData(New List(Of String)(headers), rows, "报告", "来源")
        Dim defTypes = defSlides.Select(Function(s) s.Value(Of String)("type")).ToList()
        Check(Not defTypes.Contains("table"), "默认不含「数据明细」table 页（选择性输出）")

        ' 需求含「明细」→ 生成数据明细页
        Dim detailSlides = PptDesignEngine.DesignFromData(New List(Of String)(headers), rows, "报告", "来源", "要明细")
        Dim detailTypes = detailSlides.Select(Function(s) s.Value(Of String)("type")).ToList()
        Check(detailTypes.Contains("table"), "需求含「明细」→ 含数据明细 table 页")

        ' 「占比」→ 保留构成图表页（标题含「占比」）；此数据集无时间列，仅构成图
        Dim compSlides = PptDesignEngine.DesignFromData(New List(Of String)(headers), rows, "报告", "来源", "占比 明细")
        Dim compTypes = compSlides.Select(Function(s) s.Value(Of String)("type")).ToList()
        Dim hasCompChart = compSlides.Any(Function(s)
                                              Return s.Value(Of String)("type") = "chart" AndAlso (s.Value(Of String)("title") & "").Contains("占比")
                                          End Function)
        Check(compTypes.Contains("table"), "「占比 明细」→ 含数据明细 table 页")
        Check(hasCompChart, "「占比」→ 含构成占比图")

        ' 「只要趋势」→ SkipComposition，裁剪构成占比图；无时间列故趋势图本就不能生成
        Dim trimmedSlides = PptDesignEngine.DesignFromData(New List(Of String)(headers), rows, "报告", "来源", "只要趋势")
        Dim hasCompChartTrimmed = trimmedSlides.Any(Function(s)
                                                        Return s.Value(Of String)("type") = "chart" AndAlso (s.Value(Of String)("title") & "").Contains("占比")
                                                    End Function)
        Dim hasAnyChartTrimmed = trimmedSlides.Any(Function(s) s.Value(Of String)("type") = "chart")
        Check(Not hasCompChartTrimmed, "「只要趋势」→ 裁剪构成占比图")
        Check(Not hasAnyChartTrimmed, "「只要趋势」+ 无时间列 → 不生成任何图表页")
    End Sub

    Private Sub TestAutoComparisonFromData()
        Console.WriteLine("[DesignFromData] 对比页：仅当用户需求含「对比」时生成（选择性输出）")
        Dim headers As String() = {"商品", "销售额", "利润"}
        Dim data As String()() = {
            New String() {"A", "100", "10"},
            New String() {"B", "200", "20"},
            New String() {"C", "300", "25"},
            New String() {"A", "120", "12"},
            New String() {"B", "180", "18"},
            New String() {"C", "320", "30"},
            New String() {"A", "100", "10"},
            New String() {"B", "200", "20"},
            New String() {"C", "300", "25"},
            New String() {"A", "120", "12"},
            New String() {"B", "180", "18"},
            New String() {"C", "320", "30"}}
        Dim rows = BuildRows(headers, data)
        ' 无需求：不自动生成对比页（选择性输出）
        Dim slides = PptDesignEngine.DesignFromData(New List(Of String)(headers), rows, "报告", "来源", "")
        Dim types = slides.Select(Function(s) s.Value(Of String)("type")).ToList()
        Check(Not types.Contains("comparison"), "无对比需求 → 不自动加入对比页（选择性输出）")
        ' 需求含「对比」：生成对比页
        Dim reqSlides = PptDesignEngine.DesignFromData(New List(Of String)(headers), rows, "报告", "来源", "做一下对比分析")
        Dim reqTypes = reqSlides.Select(Function(s) s.Value(Of String)("type")).ToList()
        Check(reqTypes.Contains("comparison"), "需求含「对比」→ 加入对比页")
        Dim cmp = reqSlides.FirstOrDefault(Function(s) s.Value(Of String)("type") = "comparison")
        Check(cmp IsNot Nothing AndAlso cmp.Value(Of String)("leftTitle") = "销售额" AndAlso cmp.Value(Of String)("rightTitle") = "利润",
            "对比页左右指标=前两个数值列（" & If(cmp Is Nothing, "无", cmp.Value(Of String)("leftTitle") & "/" & cmp.Value(Of String)("rightTitle")) & "）")
    End Sub

    ''' 端到端销售报表案例：真实混合数据（日期+订单ID+地区+双数值），
    ''' 期望整链触发 KPI / 趋势 / 自动对比 / 构成 / 洞察 等版式并输出真实 slides JSON 供前端规划基线使用。
    '''
    Private Sub TestE2eSalesReport()
        Console.WriteLine("[E2E] 销售报表：真实引擎 DesignFromData 输出")
        Dim headers As String() = {"日期", "订单ID", "地区", "销售额", "利润"}
        Dim data As String()() = {
            New String() {"2024-01-15", "A001", "华东", "120", "12"},
            New String() {"2024-01-20", "A002", "华南", "150", "15"},
            New String() {"2024-02-10", "A003", "华北", "130", "13"},
            New String() {"2024-02-25", "A004", "华东", "160", "16"},
            New String() {"2024-03-05", "A005", "华南", "140", "14"},
            New String() {"2024-03-18", "A006", "华北", "170", "18"},
            New String() {"2024-04-02", "A007", "华东", "155", "15"},
            New String() {"2024-04-21", "A008", "华南", "180", "20"},
            New String() {"2024-04-30", "A009", "华北", "175", "17"},
            New String() {"2024-05-12", "A010", "华东", "190", "21"},
            New String() {"2024-05-28", "A011", "华南", "185", "19"},
            New String() {"2024-06-09", "A012", "华北", "200", "22"}}
        Dim rows = BuildRows(headers, data)
        Dim slides = PptDesignEngine.DesignFromData(New List(Of String)(headers), rows, "2024年上半年销售分析", "数据来源：SalesData 工作表", "")
        Dim types = slides.Select(Function(s) s.Value(Of String)("type")).ToList()
        Console.WriteLine(slides.ToString(Newtonsoft.Json.Formatting.None))

        Check(types.Contains("kpi"), "E2E: 数值列 → KPI 页")
        Check(types.Contains("chart"), "E2E: 时间序列 → 图表页")
        Check(Not types.Contains("comparison"), "E2E: 无对比需求 → 不自动加对比页（选择性输出）")
        Check(types.Contains("closing"), "E2E: 结尾页")
        Dim chart = slides.FirstOrDefault(Function(s) s.Value(Of String)("type") = "chart")
        Check(chart IsNot Nothing AndAlso chart.Value(Of String)("chartType") = "line", "E2E: 日期序列图表=line")
        ' 需求含「对比」→ 生成对比页
        Dim reqSlides = PptDesignEngine.DesignFromData(New List(Of String)(headers), rows, "2024年上半年销售分析", "数据来源：SalesData 工作表", "做对比分析")
        Dim cmp = reqSlides.FirstOrDefault(Function(s) s.Value(Of String)("type") = "comparison")
        Check(cmp IsNot Nothing AndAlso cmp.Value(Of String)("leftTitle") = "销售额" AndAlso cmp.Value(Of String)("rightTitle") = "利润",
            "E2E: 对比需求 → 对比页=销售额/利润")
    End Sub

End Module