' ============================================================
' PptDesignEngine - 数据感知的 PPT 设计引擎
'
' 职责：
'   - 分析选区数据（列类型/时间序列/数值/分类/聚合/趋势）→ 数据画像
'   - 根据数据画像推荐版式组合与图表类型（不再固定套用）
'   - 生成规则化洞察文案（关键发现/趋势/占比）
'   - 构建完整叙事结构（封面/目录/章节/结尾）
'   - 解析参考模板（Config/ppt_templates.json 版式注册表）驱动设计决策
'
' 设计原则：确定性规则引擎，不依赖 LLM，符合"少报错"的既有哲学。
' ============================================================
Option Explicit On
Option Infer On
Option Strict Off

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json.Linq
Imports ExcelAIVSTOAddIn.Utils

Namespace Services

    ''' <summary>列画像：类型 + 数值统计 + 分类统计 + 趋势</summary>
    Public Class PptColumnProfile
        Public Property Name As String = ""
        Public Property Kind As String = "text"   ' numeric / categorical / date / text
        Public Property HasNumeric As Boolean = False
        Public Property Sum As Double = 0
        Public Property Average As Double = 0
        Public Property Min As Double = 0
        Public Property Max As Double = 0
        Public Property UniqueCount As Integer = 0
        Public Property TopValues As List(Of KeyValuePair(Of String, Integer)) = New List(Of KeyValuePair(Of String, Integer))()
        Public Property IsTimeSeries As Boolean = False
        Public Property IsIdLike As Boolean = False
        Public Property FirstValue As Double = 0
        Public Property LastValue As Double = 0
        Public Property TrendPct As Double = 0
        Public Property TrendDirection As String = "flat"   ' up / down / flat
    End Class

    ''' <summary>数据画像：整体 + 各列分类汇总</summary>
    Public Class PptDataProfile
        Public Property Headers As List(Of String) = New List(Of String)()
        Public Property RowCount As Integer = 0
        Public Property Columns As List(Of PptColumnProfile) = New List(Of PptColumnProfile)()
        Public Property NumericColumns As List(Of PptColumnProfile) = New List(Of PptColumnProfile)()
        Public Property CategoricalColumns As List(Of PptColumnProfile) = New List(Of PptColumnProfile)()
        Public Property DateColumns As List(Of PptColumnProfile) = New List(Of PptColumnProfile)()
        Public Property TimeSeriesColumn As PptColumnProfile = Nothing
        Public Property MainCategoryColumn As PptColumnProfile = Nothing
        Public Property MainNumericColumn As PptColumnProfile = Nothing
    End Class

    ''' <summary>
    ''' 数据感知设计引擎（静态入口，供 PptGenerationService / ExcelBridge 调用）
    ''' </summary>
    Public NotInheritable Class PptDesignEngine

        Private Sub New()
        End Sub

        ' ------------------------------------------------------------
        ' 参考模板注册表：从 Config/ppt_templates.json 的 slideLayouts 加载
        ' 设计引擎据此知道"参考模板"提供哪些版式与描述，驱动版式选择（单一数据源）
        ' ------------------------------------------------------------
        Private Class LayoutMeta
            Public Property Type As String = ""
            Public Property Name As String = ""
            Public Property Description As String = ""
        End Class

        Private Shared _layoutRegistry As Dictionary(Of String, LayoutMeta)
        Private Shared _layoutLoadAttempted As Boolean
        Private Shared _layoutLoadLock As New Object()

        Private Shared Function LoadLayoutRegistry() As Dictionary(Of String, LayoutMeta)
            If _layoutLoadAttempted Then Return _layoutRegistry
            SyncLock _layoutLoadLock
                If _layoutLoadAttempted Then Return _layoutRegistry
                _layoutLoadAttempted = True
                Try
                    Dim configPath As String = Path.Combine(PathHelper.ConfigDirectory, "ppt_templates.json")
                    If File.Exists(configPath) Then
                        Dim cfg As JObject = JObject.Parse(File.ReadAllText(configPath))
                        Dim layouts As JArray = cfg.Value(Of JArray)("slideLayouts")
                        If layouts IsNot Nothing Then
                            Dim reg As New Dictionary(Of String, LayoutMeta)()
                            For Each item As JToken In layouts
                                Dim obj As JObject = TryCast(item, JObject)
                                If obj Is Nothing Then Continue For
                                Dim t As String = obj.Value(Of String)("type")
                                If String.IsNullOrWhiteSpace(t) Then Continue For
                                reg(t.ToLowerInvariant()) = New LayoutMeta() With {
                                    .Type = t,
                                    .Name = obj.Value(Of String)("name"),
                                    .Description = obj.Value(Of String)("description")
                                }
                            Next
                            _layoutRegistry = reg
                        End If
                    End If
                Catch ex As Exception  ' 
                    _layoutRegistry = Nothing
                End Try
                Return _layoutRegistry
            End SyncLock
        End Function

        ''' <summary>
        ''' 分析参考模板：返回可用版式与主题的摘要（供前端展示"参考模板能力"）。
        ''' 数据源：Config/ppt_templates.json（与后端 PptGenerationService 共用单一数据源）。
        ''' </summary>
        Public Shared Function AnalyzeReferenceTemplate() As JObject
            Dim result As New JObject()
            result("success") = True
            Dim reg As Dictionary(Of String, LayoutMeta) = LoadLayoutRegistry()
            Dim layouts As New JArray()
            If reg IsNot Nothing Then
                For Each kv As KeyValuePair(Of String, LayoutMeta) In reg
                    layouts.Add(New JObject() From {
                        { "type", kv.Value.Type },
                        { "name", kv.Value.Name },
                        { "description", kv.Value.Description }
                    })
                Next
            End If
            result("layouts") = layouts
            Return result
        End Function

        ' ------------------------------------------------------------
        ' 数据画像：列分类 + 数值统计 + 分类统计 + 趋势检测
        ' ------------------------------------------------------------
        Public Shared Function AnalyzeData(headers As List(Of String),
                                           rows As List(Of Dictionary(Of String, Object))) As PptDataProfile
            Dim profile As New PptDataProfile()
            If headers Is Nothing OrElse headers.Count = 0 Then Return profile
            profile.Headers = headers
            profile.RowCount = If(rows Is Nothing, 0, rows.Count)

            For Each h As String In headers
                Dim cp As New PptColumnProfile()
                cp.Name = h
                Dim numericVals As New List(Of Double)()
                Dim dateCount As Integer = 0
                Dim totalCount As Integer = 0
                Dim unique As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                For Each row As Dictionary(Of String, Object) In rows
                    If row Is Nothing OrElse Not row.ContainsKey(h) Then Continue For
                    Dim raw As Object = row(h)
                    If raw Is Nothing OrElse raw Is DBNull.Value Then Continue For
                    Dim s As String = Convert.ToString(raw)
                    If String.IsNullOrWhiteSpace(s) Then Continue For
                    totalCount += 1
                    unique.Add(s.Trim())
                    Dim d As Double
                    If Double.TryParse(s, d) Then
                        numericVals.Add(d)
                    End If
                    Dim dt As DateTime
                    If DateTime.TryParse(s, dt) Then
                        dateCount += 1
                    End If
                Next

                cp.UniqueCount = unique.Count
                If numericVals.Count > 0 AndAlso numericVals.Count >= totalCount * 0.7 Then
                    cp.Kind = "numeric"
                    cp.HasNumeric = True
                    cp.Sum = numericVals.Sum()
                    cp.Average = cp.Sum / numericVals.Count
                    cp.Min = numericVals.Min()
                    cp.Max = numericVals.Max()
                    cp.FirstValue = numericVals(0)
                    cp.LastValue = numericVals(numericVals.Count - 1)
                    If cp.FirstValue <> 0 Then
                        cp.TrendPct = (cp.LastValue - cp.FirstValue) / Math.Abs(cp.FirstValue) * 100
                    End If
                    If cp.TrendPct > 3 Then
                        cp.TrendDirection = "up"
                    ElseIf cp.TrendPct < -3 Then
                        cp.TrendDirection = "down"
                    Else
                        cp.TrendDirection = "flat"
                    End If
                ElseIf dateCount > 0 AndAlso dateCount >= totalCount * 0.7 Then
                    cp.Kind = "date"
                    cp.IsTimeSeries = True
                ElseIf totalCount > 0 AndAlso unique.Count <= Math.Max(2, totalCount * 0.3) Then
                    cp.Kind = "categorical"
                    Dim freq As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
                    For Each row As Dictionary(Of String, Object) In rows
                        If row Is Nothing OrElse Not row.ContainsKey(h) Then Continue For
                        Dim raw As Object = row(h)
                        If raw Is Nothing OrElse raw Is DBNull.Value Then Continue For
                        Dim s As String = Convert.ToString(raw).Trim()
                        If String.IsNullOrWhiteSpace(s) Then Continue For
                        If freq.ContainsKey(s) Then freq(s) += 1 Else freq(s) = 1
                    Next
                    cp.TopValues = freq.OrderByDescending(Function(kv) kv.Value).Take(5).ToList()
                Else
                    cp.Kind = "text"
                End If
                profile.Columns.Add(cp)
            Next

            ' ID 列识别：名称特征 + 高唯一性 → 不作为主指标/主分类
            For Each cp As PptColumnProfile In profile.Columns
                If IsIdLikeColumn(cp, profile.RowCount) Then
                    cp.IsIdLike = True
                End If
            Next

            ' 分类汇总
            For Each cp As PptColumnProfile In profile.Columns
                Select Case cp.Kind
                    Case "numeric"
                        profile.NumericColumns.Add(cp)
                    Case "categorical"
                        profile.CategoricalColumns.Add(cp)
                    Case "date"
                        profile.DateColumns.Add(cp)
                End Select
            Next

            ' 主数值列：跳过 ID 列，优先非首列（首列常为类别/时间）
            If profile.NumericColumns.Count > 0 Then
                Dim candidates As List(Of PptColumnProfile) = profile.NumericColumns.Where(Function(c) Not c.IsIdLike).ToList()
                If candidates.Count = 0 Then candidates = profile.NumericColumns
                profile.MainNumericColumn = candidates(0)
                For Each cp As PptColumnProfile In candidates
                    If Not String.Equals(cp.Name, profile.Headers(0), StringComparison.OrdinalIgnoreCase) Then
                        profile.MainNumericColumn = cp
                        Exit For
                    End If
                Next
            End If

            ' 时间序列列：优先日期列
            If profile.DateColumns.Count > 0 Then
                profile.TimeSeriesColumn = profile.DateColumns(0)
            End If

            ' 主分类列：唯一值适中（2-10）优先，其次任意分类列；均跳过 ID 列
            For Each cp As PptColumnProfile In profile.CategoricalColumns
                If Not cp.IsIdLike AndAlso cp.UniqueCount >= 2 AndAlso cp.UniqueCount <= 10 Then
                    profile.MainCategoryColumn = cp
                    Exit For
                End If
            Next
            If profile.MainCategoryColumn Is Nothing Then
                For Each cp As PptColumnProfile In profile.CategoricalColumns
                    If Not cp.IsIdLike Then
                        profile.MainCategoryColumn = cp
                        Exit For
                    End If
                Next
            End If

            Return profile
        End Function

        ' ------------------------------------------------------------
        ' ID 列识别：名称特征（ID/编号/序号/No/Code）+ 高唯一性（≥90% 每行唯一）
        ' 防止 ID 列被当作主指标/主分类（如"订单ID 合计"这类无意义分析）
        ' ------------------------------------------------------------
        Private Shared Function IsIdLikeColumn(cp As PptColumnProfile, rowCount As Integer) As Boolean
            If cp Is Nothing Then Return False
            ' 名称启发式
            Dim name As String = If(cp.Name, "").Trim()
            Dim lower As String = name.ToLowerInvariant()
            If lower.EndsWith("id") OrElse lower.Contains("_id") OrElse lower.Contains("-id") OrElse
               lower.Contains("编号") OrElse lower.Contains("序号") OrElse lower.Contains("编码") OrElse
               lower.Contains("单号") OrElse lower.Contains("流水号") Then
                Return True
            End If
            If lower = "no" OrElse lower.StartsWith("no.") OrElse lower.StartsWith("no ") Then Return True
            If lower = "code" OrElse lower.Contains("code_") OrElse lower.Contains("-code") Then Return True
            ' 数据启发式：唯一值占比 ≥95% 且唯一值绝对量 ≥30 才判为标识列。
            ' 加绝对下限可避免小样本（如十几行）中"每行不同"的真实数值指标列被误判为 ID，
            ' 否则「销售额/利润」这类高唯一数值列会被踢出 KPI/对比，削弱数据驱动的版式组合。
            If rowCount > 0 AndAlso cp.UniqueCount >= rowCount * 0.95 AndAlso cp.UniqueCount >= 30 Then Return True
            Return False
        End Function

        ' ------------------------------------------------------------
        ' 聚合建模：枚举所有"分类维度 × 数值列"组合，智能判别聚合函数（sum/avg/count），
        ' 生成候选聚合视图集（数据驱动，供"默认模型"与"用户驱动模型"共用）。
        ' 返回要素：xField=分类维度、series=数值列、agg=聚合函数、categoryCount=唯一值数、
        ' chartType=建议图表、score=信息量评分、isComposite=多系列。
        ' 每个分类维度产出：前 N 个数值列各成一组聚合视图 + 数值列总和视图 + 计数视图。
        ' ------------------------------------------------------------
        Public Shared Function BuildAggregationModels(headers As List(Of String),
                                                      rows As List(Of Dictionary(Of String, Object))) As List(Of JObject)
            If headers Is Nothing OrElse headers.Count = 0 OrElse rows Is Nothing OrElse rows.Count = 0 Then
                Return New List(Of JObject)()
            End If
            Return BuildAggregationModels(AnalyzeData(headers, rows))
        End Function

        ''' <summary>profile 版聚合建模：无需原始行，供默认模型 BuildDesign 直接复用评分最佳聚合视图。</summary>
        Public Shared Function BuildAggregationModels(profile As PptDataProfile) As List(Of JObject)
            Dim models As New List(Of JObject)()
            If profile Is Nothing OrElse profile.Headers Is Nothing OrElse profile.Headers.Count = 0 Then
                Return models
            End If
            Dim headers As List(Of String) = profile.Headers
            Dim cats As List(Of PptColumnProfile) = profile.CategoricalColumns.
                Where(Function(c) Not c.IsIdLike AndAlso c.UniqueCount >= 2).ToList()
            If cats.Count = 0 Then
                For Each c As PptColumnProfile In profile.CategoricalColumns
                    If Not c.IsIdLike Then cats.Add(c)
                Next
            End If
            Dim numCols As List(Of PptColumnProfile) = profile.NumericColumns.
                Where(Function(c) Not c.IsIdLike).ToList()
            If numCols.Count = 0 Then numCols = profile.NumericColumns

            For Each cat As PptColumnProfile In cats
                ' 每个分类维度优先为主数值列做聚合 + 全部数值列合计
                Dim topNum As List(Of PptColumnProfile) = numCols.Take(3).ToList()
                For Each n As PptColumnProfile In topNum
                    models.Add(BuildAggModel(cat, n, headers, Nothing, profile))
                Next
                If numCols.Count > 1 Then
                    models.Add(BuildAggModel(cat, numCols(0), headers, Nothing, profile, includeAllNumeric:=True))
                End If
            Next
            ' 无分类列时：以首个数值列的取值本身为类别（作为兜底聚合即使只有一列也给出）
            If models.Count = 0 AndAlso numCols.Count >= 1 AndAlso cats.Count = 0 Then
                models.Add(BuildAggModel(Nothing, numCols(0), headers, Nothing, profile))
            End If
            Return models
        End Function

        Private Shared Function BuildAggModel(cat As PptColumnProfile,
                                              n As PptColumnProfile,
                                              headers As List(Of String),
                                              rows As List(Of Dictionary(Of String, Object)),
                                              profile As PptDataProfile,
                                              Optional includeAllNumeric As Boolean = False) As JObject
            Dim agg As String = ChooseAggregateFunction(n)
            Dim m As New JObject()
            m("xField") = If(cat Is Nothing, headers(0), cat.Name)
            If includeAllNumeric Then
                m("series") = New JArray(profile.NumericColumns.Where(Function(c) Not c.IsIdLike).Take(5).
                    Select(Function(c) c.Name).ToArray())
                m("xDim") = If(cat Is Nothing, headers(0), cat.Name)
            Else
                m("series") = New JArray(n.Name)
                m("xDim") = If(cat Is Nothing, headers(0), cat.Name)
            End If
            m("agg") = agg
            m("aggLabel") = AggLabel(agg)
            m("categoryCount") = If(cat Is Nothing, 0, cat.UniqueCount)
            ' 建议图表：分类少→饼图，否则柱状；日期类型分类维度在聚合场景较少，统一柱状
            Dim chartType As String = "column"
            If cat IsNot Nothing AndAlso cat.UniqueCount >= 2 AndAlso cat.UniqueCount <= 6 Then chartType = "pie"
            m("chartType") = chartType
            m("isComposite") = includeAllNumeric
            ' 信息量评分：分类唯一值适中(2-10)+主数值列 → 高分；用于默认模型排序
            Dim score As Double = 0
            If cat IsNot Nothing Then
                If cat.UniqueCount >= 2 AndAlso cat.UniqueCount <= 10 Then score += 40
                score += Math.Min(20, cat.UniqueCount) ' 更多分类→略高分，但不足以盖过适中区间
            End If
            If Not includeAllNumeric AndAlso profile.MainNumericColumn IsNot Nothing AndAlso
               String.Equals(n.Name, profile.MainNumericColumn.Name, StringComparison.OrdinalIgnoreCase) Then
                score += 30
            End If
            If includeAllNumeric Then score += 10
            m("score") = score
            Return m
        End Function

        ' 智能聚合函数判别：按列名语义 + 值特征
        Private Shared Function ChooseAggregateFunction(n As PptColumnProfile) As String
            If n Is Nothing Then Return "sum"
            Dim name As String = If(n.Name, "").ToLowerInvariant()
            ' 比率/占比/平均 语义列 → avg（金额占比类不适合求和）
            If name.Contains("率") OrElse name.Contains("占比") OrElse name.Contains("比例") OrElse
               name.Contains("百分比") OrElse name.Contains("mean") OrElse name.Contains("avg") OrElse
               name.Contains("平均") OrElse name.StartsWith("ratio") OrElse name.Contains("_ratio") Then
                Return "avg"
            End If
            ' 次数/数量/金额类 → sum
            Return "sum"
        End Function

        Private Shared Function AggLabel(agg As String) As String
            Select Case agg
                Case "avg" : Return "均值"
                Case "count" : Return "计数"
                Case Else : Return "求和"
            End Select
        End Function

        ''' <summary>
        ''' 从聚合候选集中挑"信息量最佳"的单系列聚合视图：优先主分类维度、非合计系列、含主数值列，
        ''' 按 score 降序取首个；用于默认模型构成页。找不到返回 Nothing（调用方回退启发式）。
        ''' </summary>
        Private Shared Function PickBestSingleAggModel(aggModels As List(Of JObject), mainCat As PptColumnProfile) As JObject
            If aggModels Is Nothing OrElse aggModels.Count = 0 Then Return Nothing
            Dim best As JObject = Nothing
            Dim bestScore As Double = Double.MinValue
            For Each m As JObject In aggModels
                If m Is Nothing Then Continue For
                ' 仅单系列聚合视图（多系列合计不适合作为构成主角）
                Dim sArr As JArray = TryCast(m("series"), JArray)
                If sArr Is Nothing OrElse sArr.Count <> 1 Then Continue For
                Dim isComp As Boolean = m("isComposite") IsNot Nothing AndAlso CBool(m("isComposite"))
                If isComp Then Continue For
                ' 优先主分类维度；否则任何分类维度
                Dim xField As String = If(m("xField") Is Nothing, "", m("xField").ToString())
                Dim score As Double = If(m("score") Is Nothing, 0, CDbl(m("score")))
                If mainCat IsNot Nothing AndAlso String.Equals(xField, mainCat.Name, StringComparison.OrdinalIgnoreCase) Then
                    score += 50
                End If
                If score > bestScore Then
                    bestScore = score
                    best = m
                End If
            Next
            Return best
        End Function

        Public Shared Function RecommendOrientation(profile As PptDataProfile) As String
            If profile Is Nothing Then Return "landscape"
            Return RecommendOrientation(profile.Headers, profile.RowCount)
        End Function

        Public Shared Function RecommendOrientation(headers As List(Of String), rowCount As Integer) As String
            Dim cols As Integer = If(headers Is Nothing, 0, headers.Count)
            If cols >= 6 Then Return "landscape"                          ' 宽表：多列需要横向空间
            If rowCount >= 12 AndAlso rowCount > cols * 2 Then Return "portrait"  ' 长表：行数远多于列数
            Return "landscape"
        End Function

        ' ------------------------------------------------------------
        ' 图表类型推荐：时间→折线，分类少→饼图，分类多→柱状
        ' ------------------------------------------------------------
        Public Shared Function RecommendChartType(profile As PptDataProfile,
                                                  xCol As PptColumnProfile,
                                                  yCol As PptColumnProfile) As String
            If xCol Is Nothing Then Return "column"
            If xCol.Kind = "date" OrElse xCol.IsTimeSeries Then Return "line"
            If xCol.Kind = "categorical" Then
                If xCol.UniqueCount >= 2 AndAlso xCol.UniqueCount <= 6 Then Return "pie"
                Return "column"
            End If
            Return "column"
        End Function

        ' ------------------------------------------------------------
        ' 洞察文案：基于聚合/趋势/极值的规则化叙述
        ' ------------------------------------------------------------
        ''' <summary>
        ''' 结构化洞察（供前端"洞察预览"分区渲染）：
        ''' kind = kpi（指标卡）/ trend（趋势）/ composition（分布），text 为完整叙述兜底。
        ''' </summary>
        Public Shared Function GenerateInsightObjects(profile As PptDataProfile) As List(Of JObject)
            Dim items As New List(Of JObject)()
            If profile Is Nothing OrElse profile.NumericColumns.Count = 0 Then
                items.Add(New JObject() From {
                    { "kind", "note" },
                    { "text", "当前选区未检测到可量化的数值列，建议补充数值字段后重新生成。" }
                })
                Return items
            End If

            For Each cp As PptColumnProfile In profile.NumericColumns.Take(3)
                items.Add(New JObject() From {
                    { "kind", "kpi" },
                    { "name", cp.Name },
                    { "total", FormatNumber(cp.Sum, 0) },
                    { "avg", FormatNumber(cp.Average, 1) },
                    { "min", FormatNumber(cp.Min, 1) },
                    { "max", FormatNumber(cp.Max, 1) },
                    { "text", String.Format("「{0}」合计 {1}，平均 {2}，区间 {3} ~ {4}",
                        cp.Name, FormatNumber(cp.Sum, 0), FormatNumber(cp.Average, 1),
                        FormatNumber(cp.Min, 1), FormatNumber(cp.Max, 1)) }
                })
            Next

            If profile.TimeSeriesColumn IsNot Nothing AndAlso profile.MainNumericColumn IsNot Nothing Then
                Dim num As PptColumnProfile = profile.MainNumericColumn
                Dim dirText As String = ""
                Select Case num.TrendDirection
                    Case "up"
                        dirText = "整体呈上升趋势"
                    Case "down"
                        dirText = "整体呈下降趋势"
                    Case Else
                        dirText = "整体保持平稳"
                End Select
                items.Add(New JObject() From {
                    { "kind", "trend" },
                    { "name", num.Name },
                    { "axis", profile.TimeSeriesColumn.Name },
                    { "direction", If(String.IsNullOrEmpty(num.TrendDirection), "flat", num.TrendDirection) },
                    { "first", FormatNumber(num.FirstValue, 1) },
                    { "last", FormatNumber(num.LastValue, 1) },
                    { "pct", FormatNumber(num.TrendPct, 1) },
                    { "text", String.Format("按「{0}」观察，「{1}」{2}：从 {3} 变化到 {4}（{5}%）",
                        profile.TimeSeriesColumn.Name, num.Name, dirText,
                        FormatNumber(num.FirstValue, 1), FormatNumber(num.LastValue, 1),
                        FormatNumber(num.TrendPct, 1)) }
                })
            End If

            If profile.MainCategoryColumn IsNot Nothing AndAlso profile.MainCategoryColumn.TopValues.Count > 0 Then
                Dim cat As PptColumnProfile = profile.MainCategoryColumn
                Dim top As KeyValuePair(Of String, Integer) = cat.TopValues(0)
                Dim total As Integer = 0
                For Each kv As KeyValuePair(Of String, Integer) In cat.TopValues
                    total += kv.Value
                Next
                If total > 0 Then
                    Dim pct As Double = top.Value / total * 100
                    items.Add(New JObject() From {
                        { "kind", "composition" },
                        { "name", cat.Name },
                        { "top", top.Key },
                        { "count", top.Value },
                        { "pct", FormatNumber(pct, 0) },
                        { "text", String.Format("「{0}」中「{1}」出现最多（{2} 次，占比 {3}%）",
                            cat.Name, top.Key, top.Value, FormatNumber(pct, 0)) }
                    })
                End If
            End If

            Return items
        End Function

        Public Shared Function GenerateInsights(profile As PptDataProfile) As List(Of String)
            Return GenerateInsightObjects(profile).
                Select(Function(o) o.Value(Of String)("text")).ToList()
        End Function

        ' ------------------------------------------------------------
        ' 版式变体选择：按数据特征为每类页面挑选视觉变体（不再固定套用一种版式）
        ' 变体 id 与 PptGenerationService 各 builder 的 Select Case 分支一一对应
        ' ------------------------------------------------------------
        Private Shared Function PickTitleVariant(profile As PptDataProfile) As String
            If profile IsNot Nothing AndAlso profile.TimeSeriesColumn IsNot Nothing Then Return "modern"
            If profile IsNot Nothing AndAlso profile.CategoricalColumns.Count > 0 Then Return "split"
            Return "centered"
        End Function

        Private Shared Function PickKpiVariant(count As Integer) As String
            If count <= 3 Then Return "hero"
            If count = 4 Then Return "row"
            Return "grid"
        End Function

        Private Shared Function PickChartVariant(chartType As String) As String
            If String.Equals(chartType, "line", StringComparison.OrdinalIgnoreCase) Then Return "insight"
            Return "standard"
        End Function

        Private Shared Function PickAgendaVariant(count As Integer) As String
            If count <= 4 Then Return "cards"
            If count <= 6 Then Return "timeline"
            Return "numbered"
        End Function

        Private Shared Function PickTextVariant(count As Integer) As String
            If count <= 3 Then Return "cards"
            If count <= 6 Then Return "numbered"
            Return "bullets"
        End Function

        Private Shared Function PickSectionVariant(number As Integer) As String
            If number = 1 Then Return "centered"
            If number Mod 2 = 0 Then Return "bar"
            Return "number"
        End Function

        ' ------------------------------------------------------------
        ' 数据驱动强调：把趋势/占比等关键发现注入图表页洞察提示
        ' ------------------------------------------------------------
        Private Shared Function BuildTrendInsight(profile As PptDataProfile) As String
            If profile Is Nothing OrElse profile.MainNumericColumn Is Nothing Then Return ""
            Dim num As PptColumnProfile = profile.MainNumericColumn
            Select Case num.TrendDirection
                Case "up"
                    Return String.Format("「{0}」整体呈上升趋势，较期初增长 {1}%", num.Name, FormatNumber(num.TrendPct, 1))
                Case "down"
                    Return String.Format("「{0}」整体呈下降趋势，较期初下降 {1}%", num.Name, FormatNumber(Math.Abs(num.TrendPct), 1))
                Case Else
                    Return String.Format("「{0}」整体保持平稳，区间 {1} ~ {2}", num.Name, FormatNumber(num.Min, 1), FormatNumber(num.Max, 1))
            End Select
        End Function

        Private Shared Function BuildCompositionInsight(profile As PptDataProfile) As String
            If profile Is Nothing OrElse profile.MainCategoryColumn Is Nothing OrElse profile.MainCategoryColumn.TopValues.Count = 0 Then Return ""
            Dim cat As PptColumnProfile = profile.MainCategoryColumn
            Dim top As KeyValuePair(Of String, Integer) = cat.TopValues(0)
            Dim total As Integer = 0
            For Each kv As KeyValuePair(Of String, Integer) In cat.TopValues
                total += kv.Value
            Next
            If total = 0 Then Return ""
            Dim pct As Double = top.Value / total * 100
            Return String.Format("「{0}」中「{1}」占比最高（{2}%），是构成主体", cat.Name, top.Key, FormatNumber(pct, 0))
        End Function

        ' ------------------------------------------------------------
        ' 完整设计：数据画像 → 叙事结构 + 版式 + 图表 + 洞察
        ' 输出 slides 规格（与 PptGenerationService 布局引擎契约一致）
        ' 需求偏置：从用户需求原文提取结构倾向（趋势/对比/构成/明细/图表类型），
        ' 让规则引擎的兜底结构也随用户要求变化，而不是固定套路
        ' ------------------------------------------------------------
        Private Class RequirementBias
            Public Property WantTrend As Boolean = False
            Public Property WantComparison As Boolean = False
            Public Property WantComposition As Boolean = False
            Public Property WantDetail As Boolean = False   ' 数据明细默认不生成（选择性输出），需求明确要才加
            Public Property SkipTrend As Boolean = False
            Public Property SkipComposition As Boolean = False
            Public Property ChartType As String = ""
        End Class

        Private Shared Function ParseRequirement(req As String) As RequirementBias
            Dim b As New RequirementBias()
            If String.IsNullOrWhiteSpace(req) Then Return b
            If Regex.IsMatch(req, "趋势|走势|变化|增长|下降|trend") Then b.WantTrend = True
            If Regex.IsMatch(req, "对比|比较|对照|竞品|comparison") Then b.WantComparison = True
            If Regex.IsMatch(req, "占比|构成|份额|结构|pie|饼") Then b.WantComposition = True
            If Regex.IsMatch(req, "明细|详情|完整数据|全部数据|原始数据|表格") Then b.WantDetail = True
            If Regex.IsMatch(req, "不要明细|不要表格|去掉明细|去掉表格|无需明细") Then b.WantDetail = False
            If Regex.IsMatch(req, "只要趋势|只看趋势|仅趋势") Then b.SkipComposition = True
            If Regex.IsMatch(req, "只要构成|只看构成|只要占比|只看占比") Then b.SkipTrend = True
            If Regex.IsMatch(req, "折线|趋势图|走势图|line") Then b.ChartType = "line"
            If Regex.IsMatch(req, "柱状|柱形|条形|对比图|bar|column") Then b.ChartType = "column"
            If Regex.IsMatch(req, "饼图|占比|构成|pie") Then b.ChartType = "pie"
            Return b
        End Function

        ''' <summary>构建完整叙事结构（封面/目录/章节/结尾），按数据画像与用户需求偏置生成</summary>
        Public Shared Function BuildDesign(profile As PptDataProfile,
                                           reportTitle As String,
                                           sourceDesc As String,
                                           Optional requirement As String = "") As JArray
            Dim slides As New JArray()
            If profile Is Nothing Then Return slides
            If String.IsNullOrWhiteSpace(reportTitle) Then reportTitle = "数据分析报告"
            Dim bias As RequirementBias = ParseRequirement(requirement)

            ' 默认模型：先按数据生成全部分类×数值聚合候选视图，再按信息量 score 选取最佳视图供图表页使用。
            ' 这是"先出聚合关系、再排图表"的落点：趋势页/构成页的 xField/series/agg 数据驱动自评分最佳聚合视图。
            Dim aggModels As List(Of JObject) = BuildAggregationModels(profile)

            ' 1. 封面
            slides.Add(New JObject() From {
                { "type", "title" },
                { "title", reportTitle },
                { "subtitle", sourceDesc },
                { "variant", PickTitleVariant(profile) }
            })

            ' 2. 目录（依据数据能力生成章节，章节名融入实际数据列名，避免千篇一律；
            '    选择性输出：章节条目 < 3 时不生成目录页，避免空目录）
            Dim sections As New List(Of String)()
            Dim kpiSectionName As String = ""
            Dim trendSectionName As String = ""
            Dim compSectionName As String = ""
            If profile.NumericColumns.Count > 0 Then
                kpiSectionName = BuildKpiSectionName(profile)
                sections.Add(kpiSectionName)
            End If
            If profile.TimeSeriesColumn IsNot Nothing AndAlso profile.MainNumericColumn IsNot Nothing AndAlso Not bias.SkipTrend Then
                trendSectionName = BuildTrendSectionName(profile)
                sections.Add(trendSectionName)
            End If
            If profile.MainCategoryColumn IsNot Nothing AndAlso profile.MainNumericColumn IsNot Nothing AndAlso Not bias.SkipComposition Then
                compSectionName = BuildCompositionSectionName(profile)
                sections.Add(compSectionName)
            End If
            If bias.WantComparison Then sections.Add("对比分析")
            If profile.RowCount > 8 AndAlso bias.WantDetail Then sections.Add("数据明细")
            sections.Add("总结与建议")
            ' 选择性输出：目录页仅在章节条目足够（≥3）时生成，否则直接省略
            If sections.Count >= 3 Then
                Dim agendaItems As New JArray()
                For Each s As String In sections
                    agendaItems.Add(s)
                Next
                slides.Add(New JObject() From {
                    { "type", "agenda" },
                    { "title", "目录" },
                    { "items", agendaItems },
                    { "variant", PickAgendaVariant(sections.Count) }
                })
            End If

            ' 3. 关键指标（过滤 ID 列后的有效数值列）
            Dim sectionNo As Integer = 1
            Dim kpiColumns As List(Of PptColumnProfile) = profile.NumericColumns.Where(Function(c) Not c.IsIdLike).ToList()
            If kpiColumns.Count = 0 Then kpiColumns = profile.NumericColumns
            If kpiColumns.Count > 0 Then
                slides.Add(New JObject() From {
                    { "type", "section" },
                    { "number", sectionNo.ToString("00") },
                    { "title", kpiSectionName },
                    { "subtitle", "核心数据一览" },
                    { "variant", PickSectionVariant(sectionNo) }
                })
                sectionNo += 1
                ' 数据驱动强调：数值最大的指标标记 highlight（视觉上突出）
                Dim kpiCols As List(Of PptColumnProfile) = kpiColumns.Take(4).ToList()
                Dim maxSum As Double = Double.MinValue
                Dim maxIdx As Integer = -1
                For i As Integer = 0 To kpiCols.Count - 1
                    If kpiCols(i).Sum > maxSum Then
                        maxSum = kpiCols(i).Sum
                        maxIdx = i
                    End If
                Next
                Dim kpiItems As New JArray()
                For i As Integer = 0 To kpiCols.Count - 1
                    Dim cp As PptColumnProfile = kpiCols(i)
                    kpiItems.Add(New JObject() From {
                        { "label", cp.Name },
                        { "value", FormatNumber(cp.Sum, 0) },
                        { "highlight", (i = maxIdx) }
                    })
                Next
                slides.Add(New JObject() From {
                    { "type", "kpi" },
                    { "title", kpiSectionName },
                    { "items", kpiItems },
                    { "variant", PickKpiVariant(kpiItems.Count) }
                })
            End If

            ' 4. 趋势分析（时间序列 → 折线图；需求偏置可裁剪）
            If profile.TimeSeriesColumn IsNot Nothing AndAlso profile.MainNumericColumn IsNot Nothing AndAlso Not bias.SkipTrend Then
                slides.Add(New JObject() From {
                    { "type", "section" },
                    { "number", sectionNo.ToString("00") },
                    { "title", trendSectionName },
                    { "subtitle", "时间维度变化" },
                    { "variant", PickSectionVariant(sectionNo) }
                })
                sectionNo += 1
                Dim series As New JArray()
                series.Add(profile.MainNumericColumn.Name)
                Dim trendChartType As String = "line"
                If Not String.IsNullOrWhiteSpace(bias.ChartType) Then trendChartType = bias.ChartType
                slides.Add(New JObject() From {
                    { "type", "chart" },
                    { "title", profile.MainNumericColumn.Name & "趋势变化" },
                    { "chartType", trendChartType },
                    { "xField", profile.TimeSeriesColumn.Name },
                    { "series", series },
                    { "agg", ChooseAggregateFunction(profile.MainNumericColumn) },
                    { "variant", PickChartVariant(trendChartType) },
                    { "insight", BuildTrendInsight(profile) }
                })
            End If

            ' 5. 构成分析（分类 + 数值 → 饼图/柱状；需求偏置可裁剪/指定图表类型）
            If profile.MainCategoryColumn IsNot Nothing AndAlso profile.MainNumericColumn IsNot Nothing AndAlso Not bias.SkipComposition Then
                slides.Add(New JObject() From {
                    { "type", "section" },
                    { "number", sectionNo.ToString("00") },
                    { "title", compSectionName },
                    { "subtitle", "结构占比" },
                    { "variant", PickSectionVariant(sectionNo) }
                })
                sectionNo += 1
                ' 构成页优先取自按信息量评分最高的单系列聚合视图（数据驱动），找不到再回退主分类×主数值启发式
                Dim compModel As JObject = PickBestSingleAggModel(aggModels, profile.MainCategoryColumn)
                Dim compXField As String = profile.MainCategoryColumn.Name
                Dim compSeries As New JArray()
                compSeries.Add(profile.MainNumericColumn.Name)
                Dim compAgg As String = ChooseAggregateFunction(profile.MainNumericColumn)
                Dim chartType As String = RecommendChartType(profile, profile.MainCategoryColumn, profile.MainNumericColumn)
                If compModel IsNot Nothing Then
                    compXField = compModel("xField").ToString()
                    compSeries = TryCast(compModel("series"), JArray)
                    If compSeries Is Nothing OrElse compSeries.Count = 0 Then
                        compSeries = New JArray()
                        compSeries.Add(profile.MainNumericColumn.Name)
                    End If
                    compAgg = If(compModel("agg") Is Nothing, compAgg, compModel("agg").ToString())
                    chartType = If(compModel("chartType") Is Nothing, chartType, compModel("chartType").ToString())
                End If
                ' 需求偏置显式指定图表类型时覆盖聚合模型建议
                If Not String.IsNullOrWhiteSpace(bias.ChartType) Then chartType = bias.ChartType
                slides.Add(New JObject() From {
                    { "type", "chart" },
                    { "title", "按" & profile.MainCategoryColumn.Name & profile.MainNumericColumn.Name & "占比" },
                    { "chartType", chartType },
                    { "xField", compXField },
                    { "series", compSeries },
                    { "agg", compAgg },
                    { "variant", PickChartVariant(chartType) },
                    { "insight", BuildCompositionInsight(profile) }
                })
            End If

            ' 5b. 对比分析（选择性输出：仅当用户需求明确要求对比时才加入，不再按列数自动触发）
            Dim cmpCols As List(Of PptColumnProfile) = profile.NumericColumns.Where(Function(c) Not c.IsIdLike).Take(2).ToList()
            If bias.WantComparison AndAlso cmpCols.Count = 2 Then
                    slides.Add(New JObject() From {
                        { "type", "section" },
                        { "number", sectionNo.ToString("00") },
                        { "title", "对比分析" },
                        { "subtitle", "核心指标对比" },
                        { "variant", PickSectionVariant(sectionNo) }
                    })
                    sectionNo += 1
                    slides.Add(New JObject() From {
                        { "type", "comparison" },
                        { "title", "对比分析" },
                        { "leftTitle", cmpCols(0).Name },
                        { "rightTitle", cmpCols(1).Name },
                        { "leftItems", New JArray() From { "合计 " & FormatNumber(cmpCols(0).Sum, 0), "平均 " & FormatNumber(cmpCols(0).Average, 1), "区间 " & FormatNumber(cmpCols(0).Min, 1) & " ~ " & FormatNumber(cmpCols(0).Max, 1) } },
                        { "rightItems", New JArray() From { "合计 " & FormatNumber(cmpCols(1).Sum, 0), "平均 " & FormatNumber(cmpCols(1).Average, 1), "区间 " & FormatNumber(cmpCols(1).Min, 1) & " ~ " & FormatNumber(cmpCols(1).Max, 1) } }
                    })
                End If

                ' 6. 数据明细（行数较多时；需求偏置可裁剪）
            If profile.RowCount > 8 AndAlso bias.WantDetail Then
                slides.Add(New JObject() From {
                    { "type", "section" },
                    { "number", sectionNo.ToString("00") },
                    { "title", "数据明细" },
                    { "subtitle", "完整数据" },
                    { "variant", PickSectionVariant(sectionNo) }
                })
                sectionNo += 1
                slides.Add(New JObject() From {
                    { "type", "table" },
                    { "title", "数据明细" },
                    { "columns", JArray.FromObject(profile.Headers) },
                    { "rowCount", profile.RowCount },
                    { "startRow", 2 }
                })
            End If

            ' 7. 总结与建议（洞察文案）
            slides.Add(New JObject() From {
                { "type", "section" },
                { "number", sectionNo.ToString("00") },
                { "title", "总结与建议" },
                { "subtitle", "关键发现" },
                { "variant", PickSectionVariant(sectionNo) }
            })
            Dim insightLines As New JArray()
            For Each line As String In GenerateInsights(profile)
                insightLines.Add(line)
            Next
            slides.Add(New JObject() From {
                { "type", "text" },
                { "title", "关键发现" },
                { "bullets", insightLines },
                { "variant", PickTextVariant(insightLines.Count) }
            })

            ' 8. 结尾
            slides.Add(New JObject() From {
                { "type", "closing" },
                { "title", "谢谢" },
                { "subtitle", "数据驱动决策" },
                { "variant", "thanks" }
            })

            Return slides
        End Function

        ' ------------------------------------------------------------
        ' 顶层入口：headers/rows → 设计规格（供 Generate 与 Recommend 复用）
        ' ------------------------------------------------------------
        Public Shared Function DesignFromData(headers As List(Of String),
                                              rows As List(Of Dictionary(Of String, Object)),
                                              reportTitle As String,
                                              sourceDesc As String,
                                              Optional requirement As String = "") As JArray
            Dim profile As PptDataProfile = AnalyzeData(headers, rows)
            Return BuildDesign(profile, reportTitle, sourceDesc, requirement)
        End Function

        ' ------------------------------------------------------------
        ' LLM 列理解提示应用：前端让大模型思考"该拿什么列"后下发 columnHints，
        ' 覆盖规则引擎的启发式选列。hints 缺失/字段无效时静默忽略（规则引擎兜底）。
        ' ------------------------------------------------------------
        Public Shared Sub ApplyColumnHints(profile As PptDataProfile, hints As JObject)
            If profile Is Nothing OrElse hints Is Nothing Then Return

            ' 1. ID 列覆盖：LLM 识别的标识列强制排除出指标/分类
            Dim idCols As JArray = TryCast(hints("idColumns"), JArray)
            If idCols IsNot Nothing Then
                For Each c As PptColumnProfile In profile.Columns
                    For Each tok As JToken In idCols
                        If String.Equals(c.Name, tok.ToString(), StringComparison.OrdinalIgnoreCase) Then
                            c.IsIdLike = True
                        End If
                    Next
                Next
            End If

            ' 2. 主数值指标覆盖（主角）
            Dim mainMetric As String = hints.Value(Of String)("mainMetric")
            If Not String.IsNullOrWhiteSpace(mainMetric) Then
                For Each c As PptColumnProfile In profile.NumericColumns
                    If String.Equals(c.Name, mainMetric, StringComparison.OrdinalIgnoreCase) Then
                        profile.MainNumericColumn = c
                        Exit For
                    End If
                Next
            End If

            ' 3. 主分类维度覆盖
            Dim mainCategory As String = hints.Value(Of String)("mainCategory")
            If Not String.IsNullOrWhiteSpace(mainCategory) Then
                For Each c As PptColumnProfile In profile.CategoricalColumns
                    If String.Equals(c.Name, mainCategory, StringComparison.OrdinalIgnoreCase) Then
                        profile.MainCategoryColumn = c
                        Exit For
                    End If
                Next
            End If

            ' 4. 时间/日期列覆盖
            Dim timeColumn As String = hints.Value(Of String)("timeColumn")
            If Not String.IsNullOrWhiteSpace(timeColumn) Then
                For Each c As PptColumnProfile In profile.DateColumns
                    If String.Equals(c.Name, timeColumn, StringComparison.OrdinalIgnoreCase) Then
                        profile.TimeSeriesColumn = c
                        Exit For
                    End If
                Next
            End If
        End Sub

        ' ------------------------------------------------------------
        ' 数据感知章节命名：把实际数据列名融入章节标题，避免每次千篇一律
        ' ------------------------------------------------------------
        Private Shared Function BuildKpiSectionName(profile As PptDataProfile) As String
            Dim names As List(Of String) = profile.NumericColumns.Where(Function(c) Not c.IsIdLike).Take(2).Select(Function(c) c.Name).ToList()
            If names.Count = 0 Then
                names = profile.NumericColumns.Take(2).Select(Function(c) c.Name).ToList()
            End If
            If names.Count = 0 Then Return "关键指标"
            If names.Count = 1 Then Return names(0) & "概览"
            Return names(0) & "与" & names(1) & "概览"
        End Function

        Private Shared Function BuildTrendSectionName(profile As PptDataProfile) As String
            If profile.MainNumericColumn Is Nothing Then Return "趋势分析"
            Return profile.MainNumericColumn.Name & "趋势分析"
        End Function

        Private Shared Function BuildCompositionSectionName(profile As PptDataProfile) As String
            If profile.MainCategoryColumn Is Nothing OrElse profile.MainNumericColumn Is Nothing Then Return "构成分析"
            Return "按" & profile.MainCategoryColumn.Name & profile.MainNumericColumn.Name & "构成"
        End Function

        Private Shared Function FormatNumber(value As Double, digits As Integer) As String
            Try
                Return value.ToString("N" & digits.ToString())
            Catch ex As Exception  ' 
                Return value.ToString()
            End Try
        End Function

    End Class
End Namespace
