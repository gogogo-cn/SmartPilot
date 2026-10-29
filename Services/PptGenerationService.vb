' ============================================================
' PptGenerationService - 后端原生生成 PPT（不生成/不编译 VBA）
'
' 设计目标：把"生成 PPT"收敛为固定、可验证、少报错的后端执行内核。
'  - 选区数据复用 ExcelDataBridge.GetRangeData（已在 Excel UI 线程安全读取）。
'  - 使用 CreateObject 晚绑定驱动 PowerPoint，无需新增 PowerShell Interop 引用。
'  - 参数先校验；选区/表头缺失时返回明确错误，绝不生成坏文件。
'  - 主题由 Config/ppt_templates.json + Config/custom_ppt_themes.json 合并驱动（缺省回退内置商务蓝），
'    支持 styleDirection 风格方向（tech/academic/business/minimal/vivid）与 WCAG 校验，不依赖"当场拼 VBA 再编译"。
' ============================================================
Option Explicit On
Option Infer On
Option Strict Off

Imports System
Imports System.Collections.Generic
Imports System.Threading
Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Linq
Imports System.Text
Imports System.Xml.Linq
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json.Serialization
Imports ExcelAIVSTOAddIn.Utils
Imports ExcelAIVSTOAddIn.Bridges

Namespace Services

    ''' <summary>
    ''' 后端原生生成 PPT 的服务（静态入口，供 ExcelBridge 门面委托调用）
    ''' </summary>
    Public Partial Class PptGenerationService

        Private Sub New()
        End Sub

        ' ------------------------------------------------------------
        ' 主题模型（JSON 驱动，缺省回退内置商务蓝）
        ' 单一数据源：Config/ppt_templates.json
        ' ------------------------------------------------------------
        Private Class PptTheme
            Public Property Id As String
            Public Property Name As String
            Public Property Font As String
            Public Property Primary As Long
            Public Property Secondary As Long
            Public Property Accent As Long
            Public Property Bg As Long
            Public Property Light As Long
            Public Property Text As Long
            Public Property Gray As Long
            Public Property LayoutProfile As PptLayoutProfile
            Public Property DesignTokens As PptDesignTokens
        End Class

        Private Class PptTemplatesConfig
            Public Property Version As Integer
            Public Property DefaultTheme As String
            Public Property Themes As List(Of PptThemeJson)
        End Class

        Private Class PptThemeJson
            Public Property Id As String
            Public Property Name As String
            Public Property Font As String
            Public Property Colors As Dictionary(Of String, List(Of Integer))
            Public Property Layout As PptLayoutProfile
            Public Property DesignTokens As JToken
        End Class

        ' ------------------------------------------------------------
        ' 版式画像模型：从公司模板 PPT 提取的布局结构
        ' 标题/正文占位符位置、表格位置与列宽、装饰形状位置（设计空间坐标）
        ' ------------------------------------------------------------
        Private Class PptRect
            Public Property Left As Integer
            Public Property Top As Integer
            Public Property Width As Integer
            Public Property Height As Integer
        End Class

        Private Class PptTableProfile
            Public Property Left As Integer
            Public Property Top As Integer
            Public Property Width As Integer
            Public Property Height As Integer
            Public Property Columns As List(Of Integer)
            Public Property HeaderFill As Long
        End Class

        Private Class PptShapeProfile
            Public Property Kind As String
            Public Property Left As Integer
            Public Property Top As Integer
            Public Property Width As Integer
            Public Property Height As Integer
            Public Property Fill As Long
            Public Property LineColor As Long = -1       ' 线型边框色（-1=无线条）
            Public Property ImagePath As String          ' 图片形状（logo 等），相对 Config 目录
        End Class

        ''' <summary>
        ''' 母版背景画像：从 slideMaster 提取的背景（选择性：仅当模板显式定义了背景才生效）
        ''' </summary>
        Private Class PptBackgroundProfile
            Public Property Kind As String               ' "color" / "gradient" / "picture"
            Public Property Color As Long                ' Kind=color：背景色
            Public Property GradientColor1 As Long       ' Kind=gradient：渐变起色
            Public Property GradientColor2 As Long       ' Kind=gradient：渐变止色
            Public Property GradientAngle As Integer     ' Kind=gradient：线性渐变角度（0~360）
            Public Property ImagePath As String          ' Kind=picture：背景图片，相对 Config 目录
        End Class

        ' 版式容量限制：由模板内容区尺寸推导（正文区高度→要点数，表格区高度→行数）
        ' 供 LLM 规划约束每页内容量，并在构建器渲染时强制兜底，避免内容溢出
        Private Class PptCapacityProfile
            Public Property MaxBullets As Integer = 8          ' 每页要点上限
            Public Property MaxTableRows As Integer = 25       ' 表格数据行上限
            Public Property MaxTableCols As Integer = 12       ' 表格列上限
            Public Property MaxChartSeries As Integer = 5      ' 图表系列上限
            Public Property MaxChartCategories As Integer = 8  ' 图表分类上限
        End Class

        Private Class PptLayoutProfile
            Public Property Title As PptRect
            Public Property Body As PptRect
            Public Property Table As PptTableProfile
            Public Property Shapes As List(Of PptShapeProfile)
            Public Property Background As PptBackgroundProfile    ' 母版背景（每页套用）
            Public Property MasterShapes As List(Of PptShapeProfile)  ' 母版背景图形（logo/装饰，每页复现）
            Public Property ContentTopInset As Integer = 0   ' 顶部装饰区高度（横幅等），内容区下移避让
            Public Property Capacity As PptCapacityProfile
        End Class

        Private Shared _themeConfig As PptTemplatesConfig
        Private Shared _themeLoadAttempted As Boolean
        Private Shared _themeLoadLock As New Object()

        ' 内置兜底主题（商务蓝，与旧版固定主题一致）
        Private Shared ReadOnly FALLBACK_THEME As New PptTheme() With {
            .Id = "business_blue",
            .Name = "商务蓝",
            .Font = "微软雅黑",
            .Primary = RGB(31, 78, 121),
            .Secondary = RGB(190, 167, 91),
            .Accent = RGB(190, 167, 91),
            .Bg = RGB(245, 247, 250),
            .Light = RGB(239, 243, 248),
            .Text = RGB(33, 37, 41),
            .Gray = RGB(150, 150, 150)
        }

        ' 布局引擎支持的幻灯片类型白名单（与 Config/ppt_templates.json slideLayouts 对齐）
        Private Shared ReadOnly SUPPORTED_SLIDE_TYPES As New HashSet(Of String)() From {
            "title", "cover", "table", "text", "bullets", "chart",
            "kpi", "comparison", "timeline", "agenda", "section", "closing", "end",
            "shapes", "picture", "smartart"
        }

        ''' <summary>页型白名单查询（供 MCP 工具层分步构建时校验）</summary>
        Public Shared Function IsSupportedSlideType(slideType As String) As Boolean
            If String.IsNullOrWhiteSpace(slideType) Then Return False
            Return SUPPORTED_SLIDE_TYPES.Contains(slideType.Trim().ToLowerInvariant())
        End Function

        ' ------------------------------------------------------------
        ' 设计令牌（designTokens）：会话级可变，由主题 JSON 的 designTokens 驱动。
        ' 默认值与内置常量一致（舒适密度），无令牌主题行为与改造前完全一致（向后兼容）。
        ' 会话开始（GenerateImpl）时由 ApplyDesignTokens(theme.DesignTokens) 设置；
        ' 并发已被 Generate 入口的 Monitor.TryEnter 拒绝，无跨会话污染。
        ' ------------------------------------------------------------
        Private Const DEFAULT_MARGIN As Integer = 40
        Private Const DEFAULT_GRID As Integer = 20
        Private Const DEFAULT_TITLE_SIZE As Integer = 26       ' 页标题
        Private Const DEFAULT_HEADING_SIZE As Integer = 22     ' 栏目标题
        Private Const DEFAULT_BODY_SIZE As Integer = 16        ' 正文
        Private Const DEFAULT_SMALL_SIZE As Integer = 12       ' 辅助说明
        Private Const DEFAULT_FOOTER_SIZE As Integer = 10
        Private Const DEFAULT_TITLE_TOP As Integer = 20
        Private Const DEFAULT_CONTENT_TOP As Integer = 100
        Private Shared MARGIN As Integer = DEFAULT_MARGIN
        Private Shared GRID As Integer = DEFAULT_GRID
        Private Shared TITLE_SIZE As Integer = DEFAULT_TITLE_SIZE
        Private Shared HEADING_SIZE As Integer = DEFAULT_HEADING_SIZE
        Private Shared BODY_SIZE As Integer = DEFAULT_BODY_SIZE
        Private Shared SMALL_SIZE As Integer = DEFAULT_SMALL_SIZE
        Private Shared FOOTER_SIZE As Integer = DEFAULT_FOOTER_SIZE
        Private Shared TITLE_TOP As Integer = DEFAULT_TITLE_TOP
        Private Shared CONTENT_TOP As Integer = DEFAULT_CONTENT_TOP
        Private Shared RIBBON_STYLE As String = "line"   ' 标题装饰线：line / bar / none
        Private Shared RIBBON_WIDTH As Integer = 160     ' 标题装饰线宽（line=细线）
        ' 当前会话的设计令牌对象（供页型变体注入等读取；ApplyDesignTokens 维护）
        Private Shared _activeTokens As PptDesignTokens = Nothing

        ' 设计空间 960x720（4:3），页面方向数据感知：横版 960x720 / 竖版 720x960，
        ' 布局坐标经 SX/SY/SW/SH 变换适配（CanvasW/CanvasH 会话内设置）。
        Private Const DESIGN_W As Integer = 960
        Private Const DESIGN_H As Integer = 720
        Private Shared CanvasW As Integer = DESIGN_W
        Private Shared CanvasH As Integer = DESIGN_H

        ''' <summary>
        ''' 设计令牌：主题 JSON 的 designTokens 规范化后的运行时表示。
        ''' 所有字段可缺省（Nothing / 默认构造 = 舒适密度，与内置常量一致）。
        ''' </summary>
        Private Class PptDesignTokens
            Public Property TitleSize As Integer = DEFAULT_TITLE_SIZE
            Public Property HeadingSize As Integer = DEFAULT_HEADING_SIZE
            Public Property BodySize As Integer = DEFAULT_BODY_SIZE
            Public Property SmallSize As Integer = DEFAULT_SMALL_SIZE
            Public Property FooterSize As Integer = DEFAULT_FOOTER_SIZE
            Public Property Margin As Integer = DEFAULT_MARGIN
            Public Property Grid As Integer = DEFAULT_GRID
            Public Property TitleTop As Integer = DEFAULT_TITLE_TOP
            Public Property ContentTop As Integer = DEFAULT_CONTENT_TOP
            Public Property RibbonStyle As String = "line"
            Public Property RibbonWidth As Integer = 160
            Public Property Dense As Boolean = False
            ' 页型变体（主题级默认，spec 页级 variant 可覆盖）：
            ' key=页型（title/text/kpi/chart/section），value=合法变体
            Public Property PageVariants As Dictionary(Of String, String)
        End Class

        ''' <summary>
        ''' 会话开始应用设计令牌；tokens 为 Nothing 时恢复内置默认（舒适密度）。
        ''' 由 GenerateImpl 在解析主题后调用；并发已由 Generate 入口锁拒绝。
        ''' </summary>
        Private Shared Sub ApplyDesignTokens(tokens As PptDesignTokens)
            _activeTokens = tokens
            If tokens Is Nothing Then
                MARGIN = DEFAULT_MARGIN
                GRID = DEFAULT_GRID
                TITLE_SIZE = DEFAULT_TITLE_SIZE
                HEADING_SIZE = DEFAULT_HEADING_SIZE
                BODY_SIZE = DEFAULT_BODY_SIZE
                SMALL_SIZE = DEFAULT_SMALL_SIZE
                FOOTER_SIZE = DEFAULT_FOOTER_SIZE
                TITLE_TOP = DEFAULT_TITLE_TOP
                CONTENT_TOP = DEFAULT_CONTENT_TOP
                RIBBON_STYLE = "line"
                RIBBON_WIDTH = 160
                Return
            End If
            MARGIN = tokens.Margin
            GRID = tokens.Grid
            TITLE_SIZE = tokens.TitleSize
            HEADING_SIZE = tokens.HeadingSize
            BODY_SIZE = tokens.BodySize
            SMALL_SIZE = tokens.SmallSize
            FOOTER_SIZE = tokens.FooterSize
            TITLE_TOP = tokens.TitleTop
            CONTENT_TOP = tokens.ContentTop
            RIBBON_STYLE = tokens.RibbonStyle
            RIBBON_WIDTH = tokens.RibbonWidth
        End Sub

        ' 设计空间坐标 → 实际画布坐标（竖版时按比例变换，横版为恒等）
        Private Shared Function SX(x As Integer) As Integer
            Return CInt(Math.Round(x * CanvasW / DESIGN_W))
        End Function
        Private Shared Function SY(y As Integer) As Integer
            Return CInt(Math.Round(y * CanvasH / DESIGN_H))
        End Function
        Private Shared Function SW(w As Integer) As Integer
            Return CInt(Math.Round(w * CanvasW / DESIGN_W))
        End Function
        Private Shared Function SH(h As Integer) As Integer
            Return CInt(Math.Round(h * CanvasH / DESIGN_H))
        End Function

        ' ------------------------------------------------------------
        ' 版式画像套用：当前生成会话激活的布局画像（来自公司模板导入）
        ' 有画像时覆盖默认排版网格，无画像时回退内置常量（行为不变）
        ' ------------------------------------------------------------
        Private Shared _activeLayout As PptLayoutProfile

        Private Shared Function LayoutInset() As Integer
            If _activeLayout IsNot Nothing Then Return _activeLayout.ContentTopInset
            Return 0
        End Function

        Private Shared Function TitleLeft() As Integer
            If _activeLayout IsNot Nothing AndAlso _activeLayout.Title IsNot Nothing Then Return _activeLayout.Title.Left
            Return 40
        End Function
        Private Shared Function TitleTop() As Integer
            ' 内容避让：标题下移到母版顶部装饰区之下
            If _activeLayout IsNot Nothing AndAlso _activeLayout.Title IsNot Nothing Then
                Return Math.Max(_activeLayout.Title.Top, _activeLayout.ContentTopInset)
            End If
            Return TITLE_TOP + LayoutInset()
        End Function
        Private Shared Function TitleWidth() As Integer
            If _activeLayout IsNot Nothing AndAlso _activeLayout.Title IsNot Nothing Then Return _activeLayout.Title.Width
            Return 600
        End Function
        Private Shared Function TitleHeight() As Integer
            If _activeLayout IsNot Nothing AndAlso _activeLayout.Title IsNot Nothing Then Return _activeLayout.Title.Height
            Return 40
        End Function
        Private Shared Function ContentTop() As Integer
            ' 内容避让：正文区下移到母版顶部装饰区之下
            If _activeLayout IsNot Nothing AndAlso _activeLayout.Body IsNot Nothing Then
                Return Math.Max(_activeLayout.Body.Top, _activeLayout.ContentTopInset)
            End If
            Return CONTENT_TOP + LayoutInset()
        End Function
        Private Shared Function TitleRibbonTop() As Integer
            ' 标题装饰线：紧贴标题框底部，随标题下移避让母版顶部装饰（默认 20+40+6=66，与原常量一致）
            Return TitleTop() + TitleHeight() + 6
        End Function
        Private Shared Function ContentAreaTop() As Integer
            ' 正文区起点：先避让母版顶部装饰，再确保位于标题与装饰线之下（默认 max(100, 72)=100，与原常量一致）
            Return Math.Max(ContentTop(), TitleRibbonTop() + 6)
        End Function
        Private Shared Function TableLeft() As Integer
            If _activeLayout IsNot Nothing AndAlso _activeLayout.Table IsNot Nothing Then Return _activeLayout.Table.Left
            Return 40
        End Function
        Private Shared Function TableTop() As Integer
            ' 内容避让：表格区下移到母版顶部装饰区之下，且不越过标题底部（默认 max(80, 76)=80，与原常量一致）
            Dim base As Integer = 80 + LayoutInset()
            If _activeLayout IsNot Nothing AndAlso _activeLayout.Table IsNot Nothing Then
                base = Math.Max(_activeLayout.Table.Top, _activeLayout.ContentTopInset)
            End If
            Return Math.Max(base, TitleTop() + TitleHeight() + 16)
        End Function
        Private Shared Function TableWidth() As Integer
            If _activeLayout IsNot Nothing AndAlso _activeLayout.Table IsNot Nothing Then Return _activeLayout.Table.Width
            Return 880
        End Function
        Private Shared Function TableHeight() As Integer
            If _activeLayout IsNot Nothing AndAlso _activeLayout.Table IsNot Nothing Then Return _activeLayout.Table.Height
            Return 360
        End Function

        ''' <summary>
        ''' 套用版式画像的表格列宽（按比例缩放到当前表格宽度）
        ''' </summary>
        Private Shared Sub ApplyTableColumnWidths(tbl As Object, colCount As Integer)
            If _activeLayout Is Nothing OrElse _activeLayout.Table Is Nothing OrElse
               _activeLayout.Table.Columns Is Nothing OrElse _activeLayout.Table.Columns.Count = 0 Then
                Return
            End If
            Dim cols As List(Of Integer) = _activeLayout.Table.Columns
            Dim take As Integer = Math.Min(cols.Count, colCount)
            Dim total As Integer = 0
            For i As Integer = 0 To take - 1
                total += cols(i)
            Next
            If total <= 0 Then Return
            Dim tableW As Integer = TableWidth()
            For c As Integer = 1 To take
                Try
                    tbl.Columns(c).Width = CInt(Math.Round(cols(c - 1) * tableW / total))
                Catch
                End Try
            Next
        End Sub

        ''' <summary>
        ''' 页型统计：按 slides 规格统计各类页数量（含别名归一），返回"类型×数量"列表（保持出现顺序）
        ''' </summary>
        Private Shared Function BuildSlideTypeStats(slidesSpec As JArray) As List(Of String)
            Dim result As New List(Of String)()
            If slidesSpec Is Nothing OrElse slidesSpec.Count = 0 Then Return result
            Dim order As New List(Of String)()
            Dim counts As New Dictionary(Of String, Integer)()
            For Each item As JToken In slidesSpec
                Dim spec As JObject = TryCast(item, JObject)
                If spec Is Nothing Then Continue For
                Dim st As String = If(spec("type") Is Nothing, "", Convert.ToString(spec("type"))).Trim().ToLowerInvariant()
                Dim label As String = ""
                Select Case st
                    Case "title", "cover" : label = "封面"
                    Case "agenda" : label = "目录"
                    Case "section" : label = "章节"
                    Case "chart" : label = "图表"
                    Case "kpi" : label = "KPI"
                    Case "table" : label = "表格"
                    Case "comparison" : label = "对比"
                    Case "timeline" : label = "时间线"
                    Case "text", "bullets" : label = "要点"
                    Case "shapes" : label = "形状"
                    Case "picture" : label = "图文"
                    Case "smartart" : label = "图示"
                    Case "closing", "end" : label = "结尾"
                    Case Else : label = "其他"
                End Select
                If Not counts.ContainsKey(label) Then
                    counts(label) = 0
                    order.Add(label)
                End If
                counts(label) += 1
            Next
            For Each label As String In order
                result.Add(label & "×" & counts(label).ToString())
            Next
            Return result
        End Function

        ''' <summary>
        ''' 统计指定页型数量（用于图表形式描述的图表页计数）
        ''' </summary>
        Private Shared Function CountSlideType(slidesSpec As JArray, targetType As String) As Integer
            If slidesSpec Is Nothing Then Return 0
            Dim n As Integer = 0
            For Each item As JToken In slidesSpec
                Dim spec As JObject = TryCast(item, JObject)
                If spec Is Nothing Then Continue For
                Dim st As String = If(spec("type") Is Nothing, "", Convert.ToString(spec("type"))).Trim().ToLowerInvariant()
                If st = targetType Then n += 1
            Next
            Return n
        End Function

        ''' <summary>
        ''' 图表形式描述：按 chartMode 归一为可读文案，并附图表页数
        ''' </summary>
        Private Shared Function DescribeChartMode(chartMode As String, chartPages As Integer) As String
            Dim m As String = If(chartMode Is Nothing, "", chartMode.Trim().ToLowerInvariant())
            Dim modeDesc As String = ""
            Select Case m
                Case "native" : modeDesc = "带数据可编辑图（OLE 嵌入，双击进入 Excel 编辑）"
                Case "picture" : modeDesc = "静态图片"
                Case "auto", "" : modeDesc = "自动（原生图表优先，失败自动降级）"
                Case Else : modeDesc = "自动（" & m & "）"
            End Select
            If chartPages > 0 Then
                Return "图表形式：" & modeDesc & "（图表页 ×" & chartPages.ToString() & "）"
            End If
            Return "图表形式：无图表页"
        End Function

        ' ------------------------------------------------------------
        ' 入口：生成 PPT
        ' argsJson: { reportTitle?, rangeAddress, sourceLabel? }
        ' ------------------------------------------------------------
        ' 并发拒绝：同一时间只允许一个 Generate 在执行中。
        ' 典型场景：orchestrator 路由 race —— code-pipeline（等配置卡按钮）与 agent-loop
        ' （命中 ppt_generate 工具）可能同时调用 Generate，导致三个 PPT 窗口依次弹出。
        ' SyncLock 会排队依次生成（还是三个），这里用 flag + Monitor.TryEnter 做"拒绝式"防重入。
        Private Shared _genLock As New Object()
        Private Shared _isGenerating As Boolean = False
        Public Shared Function Generate(argsJson As String) As String
            If Not Monitor.TryEnter(_genLock, 0) Then
                ' 已有一个 Generate 在执行中 —— 直接拒绝，不排队
                Debug.WriteLine("[PptGenerationService] ⚠️ 拒绝并发 Generate（已有一个在执行）")
                Return "{""success"":false,""error"":""PPT 正在生成中，请稍候再试。""}"
            End If
            _isGenerating = True
            Try
                Return GenerateImpl(argsJson)
            Finally
                _isGenerating = False
                Monitor.Exit(_genLock)
            End Try
        End Function

        ''' <summary>
        ''' 队列入口（P1-1）：阻塞等待生成锁，供 PptTaskQueue 后台线程调用，
        ''' 与同步 Generate 的 TryEnter 拒绝式互斥，保证任意时刻只有一个 PPT 生成执行。
        ''' </summary>
        Friend Shared Function GenerateQueued(argsJson As String, progress As Action(Of Integer, String)) As String
            Monitor.Enter(_genLock)
            Try
                _isGenerating = True
                Try
                    Return GenerateImpl(argsJson, progress)
                Finally
                    _isGenerating = False
                End Try
            Finally
                Monitor.Exit(_genLock)
            End Try
        End Function

        Private Shared Function GenerateImpl(argsJson As String, Optional progress As Action(Of Integer, String) = Nothing) As String
            Dim reportTitle As String = ""
            Try
                Debug.WriteLine("[PptGenerationService] ========== Generate 开始 ==========")
                Debug.WriteLine("[PptGenerationService] argsJson = " & If(String.IsNullOrWhiteSpace(argsJson), "(空)", argsJson))

                Dim args As JObject = Nothing
                If Not String.IsNullOrWhiteSpace(argsJson) Then
                    Try
                        args = JObject.Parse(argsJson)
                    Catch
                        args = Nothing
                        Debug.WriteLine("[PptGenerationService] ⚠️ argsJson 解析失败")
                    End Try
                End If

                Dim rangeAddress As String = If(args Is Nothing, "", args.Value(Of String)("rangeAddress"))
                reportTitle = If(args Is Nothing, "", args.Value(Of String)("reportTitle"))
                Dim sourceLabel As String = If(args Is Nothing, "", args.Value(Of String)("sourceLabel"))
                Dim themeId As String = If(args Is Nothing, "", args.Value(Of String)("theme"))
                Dim theme As PptTheme = ResolveTheme(themeId)
                ' applyLayout=false 时跳过版式画像套用（标题/表格位置 + 装饰形状），仅用配色字体
                Dim applyLayout As Boolean = If(args Is Nothing OrElse args("applyLayout") Is Nothing, True, args.Value(Of Boolean)("applyLayout"))
                _activeLayout = If(applyLayout, theme.LayoutProfile, Nothing)
                ' 设计令牌（designTokens）独立于 applyLayout 生效：它是主题的设计语言（字号/密度/装饰），
                ' 不随"是否套用公司模板画像"开关而变化。
                ApplyDesignTokens(theme.DesignTokens)
                Dim slidesSpec As JArray = If(args Is Nothing, Nothing, args.Value(Of JArray)("slides"))
                Dim autoDesign As Boolean = If(args Is Nothing, False, args.Value(Of Boolean)("autoDesign"))
                Dim orientation As String = If(args Is Nothing, "", args.Value(Of String)("orientation"))
                ' 图表形式：auto=原生优先失败自动降级（默认）；picture=一律图片；native=仅带数据原生图
                Dim chartMode As String = If(args Is Nothing, "", args.Value(Of String)("chartMode"))
                ' 版式比例：wide16x9=1280x720（默认）；4x3=960x720；portrait=竖版 720x960（经 orientation 处理）
                Dim ratio As String = If(args Is Nothing, "", args.Value(Of String)("ratio"))
                ' 聚合函数（全局默认，spec 页级 agg 可覆盖）：sum=求和（默认）/avg=均值/count=计数
                Dim chartAgg As String = If(args Is Nothing, "", args.Value(Of String)("agg"))

                Debug.WriteLine("[PptGenerationService] 参数: rangeAddress='" & rangeAddress & "' slidesSpec=" & If(slidesSpec Is Nothing, "Nothing", slidesSpec.Count & "项") & " autoDesign=" & autoDesign)

                progress?.Invoke(5, "解析参数")

                ' ---- 布局引擎：幻灯片类型白名单校验 ----
                If slidesSpec IsNot Nothing Then
                    Dim invalidTypes As List(Of String) = ValidateSlideSpecs(slidesSpec)
                    If invalidTypes.Count > 0 Then
                        Return JsonError("不支持的幻灯片类型：" & String.Join("、", invalidTypes.ToArray()) &
                                         "。支持的 type：title/table/text/chart/kpi/comparison/timeline/agenda/section/closing/shapes/picture/smartart")
                    End If
                End If

                ' ---- 参数守卫：纯结构 PPT（无数据页面）允许省略选区；数据型页面必须有数据来源 ----
                ' AI 自由搓 PPT：slides 全为非数据页（title/text/agenda/section/timeline/comparison/closing）时，
                ' 允许不传 rangeAddress/headers/rows，生成纯文字/结构型演示文稿。
                Dim slidesNeedData As Boolean = False
                If slidesSpec IsNot Nothing Then
                    For Each s In slidesSpec
                        Dim st = If(s Is Nothing OrElse s("type") Is Nothing, "", Convert.ToString(s("type")))
                        If st.Equals("table", StringComparison.OrdinalIgnoreCase) OrElse
                           st.Equals("chart", StringComparison.OrdinalIgnoreCase) OrElse
                           st.Equals("kpi", StringComparison.OrdinalIgnoreCase) Then
                            slidesNeedData = True
                            Exit For
                        End If
                    Next
                End If
                Dim pureStructureMode As Boolean = slidesSpec IsNot Nothing AndAlso slidesSpec.Count > 0 AndAlso Not slidesNeedData

                If String.IsNullOrWhiteSpace(reportTitle) Then
                    reportTitle = "数据分析报告"
                End If

                ' ---- 数据来源：优先使用前端 context 直接传入的表头+数据 ----
                ' （前端参考区域/选区已带 headers/rows，直接传数据可完全绕开"按地址读 Excel"，
                '   避免工作簿名/工作表名匹配失败导致的 0x800A03EC）
                Dim headers As List(Of String) = Nothing
                Dim rows As List(Of Dictionary(Of String, Object)) = Nothing
                If args IsNot Nothing Then
                    Try
                        headers = args("headers")?.ToObject(Of List(Of String))()
                    Catch
                        headers = Nothing
                    End Try
                    Try
                        rows = args("rows")?.ToObject(Of List(Of Dictionary(Of String, Object)))()
                    Catch
                        rows = Nothing
                    End Try
                End If

                If Not pureStructureMode Then
                    ' 数据型流程：必须有数据来源（rangeAddress 或 直接传入的 headers+rows）
                    If String.IsNullOrWhiteSpace(rangeAddress) AndAlso
                       (headers Is Nothing OrElse headers.Count = 0 OrElse rows Is Nothing OrElse rows.Count = 0) Then
                        Debug.WriteLine("[PptGenerationService] rangeAddress 空 + headers/rows 都空 → 进入自动兜底")
                        ' ✅ 自动兜底：从 Excel 当前活动选区直接读 headers+rows
                        Try
                            Dim reader = BridgeRegistry.Instance.Resolve(Of IExcelDataReader)()
                            Debug.WriteLine("[PptGenerationService] reader = " & If(reader Is Nothing, "Nothing ✗", "OK ✓"))
                            If reader IsNot Nothing Then
                                Dim selJson = reader.GetSelectedRangeData()
                                Debug.WriteLine("[PptGenerationService] GetSelectedRangeData 返回: " & If(String.IsNullOrWhiteSpace(selJson), "(空)", selJson.Substring(0, Math.Min(200, selJson.Length))))
                                If Not String.IsNullOrWhiteSpace(selJson) Then
                                    Dim selObj As JObject = JObject.Parse(selJson)
                                    Dim selSuccess = If(selObj("success") IsNot Nothing, selObj.Value(Of Boolean)("success"), True)
                                    Debug.WriteLine("[PptGenerationService] 选区 success=" & selSuccess)
                                    If selSuccess Then
                                        Try
                                            headers = selObj("headers")?.ToObject(Of List(Of String))()
                                        Catch
                                        End Try
                                        Try
                                            rows = selObj("rows")?.ToObject(Of List(Of Dictionary(Of String, Object)))()
                                        Catch
                                        End Try
                                        Debug.WriteLine("[PptGenerationService] 选区读取结果: headers=" & If(headers Is Nothing, "Nothing", headers.Count & "列") & " rows=" & If(rows Is Nothing, "Nothing", rows.Count & "行"))
                                    End If
                                End If

                                ' ✅ 二级兜底：选区拿不到 → 从活动工作表 UsedRange 读（上一轮 dashboard/代码生成可能清空了选区）
                                If headers Is Nothing OrElse headers.Count = 0 OrElse rows Is Nothing OrElse rows.Count = 0 Then
                                    Debug.WriteLine("[PptGenerationService] 选区还是空 → 尝试 UsedRange 兜底")
                                    Try
                                        Dim usedJson = reader.GetUsedRangeData()
                                        Debug.WriteLine("[PptGenerationService] GetUsedRangeData 返回: " & If(String.IsNullOrWhiteSpace(usedJson), "(空)", usedJson.Substring(0, Math.Min(200, usedJson.Length))))
                                        If Not String.IsNullOrWhiteSpace(usedJson) Then
                                            Dim usedObj As JObject = JObject.Parse(usedJson)
                                            Dim usedSuccess = If(usedObj("success") IsNot Nothing, usedObj.Value(Of Boolean)("success"), True)
                                            Debug.WriteLine("[PptGenerationService] UsedRange success=" & usedSuccess)
                                            If usedSuccess Then
                                                Try
                                                    headers = usedObj("headers")?.ToObject(Of List(Of String))()
                                                Catch
                                                End Try
                                                Try
                                                    rows = usedObj("rows")?.ToObject(Of List(Of Dictionary(Of String, Object)))()
                                                Catch
                                                End Try
                                                Debug.WriteLine("[PptGenerationService] UsedRange 读取结果: headers=" & If(headers Is Nothing, "Nothing", headers.Count & "列") & " rows=" & If(rows Is Nothing, "Nothing", rows.Count & "行"))
                                                If headers IsNot Nothing AndAlso headers.Count > 0 Then
                                                    Debug.WriteLine("[PptGenerationService] ✅ 二级兜底成功: 从 UsedRange 读取 " & headers.Count & "列 × " & rows?.Count & "行")
                                                End If
                                            Else
                                                Debug.WriteLine("[PptGenerationService] ⚠️ UsedRange 返回失败: " & usedObj("error")?.ToString())
                                            End If
                                        End If
                                    Catch ex2 As Exception
                                        Debug.WriteLine("[PptGenerationService] ❌ UsedRange 兜底异常: " & ex2.Message)
                                    End Try
                                Else
                                    Debug.WriteLine("[PptGenerationService] ✅ 选区兜底成功，跳过 UsedRange")
                                End If
                            End If
                        Catch ex As Exception
                            Debug.WriteLine("[PptGenerationService] ❌ 自动兜底异常: " & ex.Message)
                        End Try
                    End If

                    ' ---- 数据感知设计 · 前置：有数据 + 无 slides 规格时自动开启设计引擎 ----
                    ' 【关键】必须在"缺少数据来源"检查（L521 附近）之前执行，
                    ' 否则 slidesSpec=Nothing 时会先 Return JsonError 而永远走不到设计引擎。
                    Dim hasSlidesParam = args IsNot Nothing AndAlso args("slides") IsNot Nothing AndAlso args("slides").ToString().Trim().Length > 0
                    Dim hasDataForDesign = headers IsNot Nothing AndAlso headers.Count > 0 AndAlso rows IsNot Nothing AndAlso rows.Count > 0
                    If Not hasSlidesParam AndAlso hasDataForDesign AndAlso slidesSpec Is Nothing Then
                        autoDesign = True
                        Debug.WriteLine("[PptGenerationService] autoDesign 自动开启：有数据(" & headers.Count & "列 × " & rows.Count & "行) 但无 slides 规格")
                    End If
                    If autoDesign AndAlso (slidesSpec Is Nothing OrElse slidesSpec.Count = 0) Then
                        Dim requirement As String = If(args Is Nothing, "", args.Value(Of String)("requirement"))
                        slidesSpec = PptDesignEngine.DesignFromData(headers, rows, reportTitle, "Excel 选区", requirement)
                        If slidesSpec IsNot Nothing AndAlso slidesSpec.Count > 0 Then
                            ChatLogger.Instance.LogInfo("PPT_GEN", "数据感知设计引擎生成 " & slidesSpec.Count.ToString() & " 页（前置执行）")
                        End If
                    End If

                    If String.IsNullOrWhiteSpace(rangeAddress) AndAlso
                       (headers Is Nothing OrElse headers.Count = 0 OrElse rows Is Nothing OrElse rows.Count = 0) Then
                        If slidesSpec Is Nothing OrElse slidesSpec.Count = 0 Then
                            Return JsonError("缺少数据来源。请先在 Excel 中选中数据区域，或传入 rangeAddress / headers+rows。")
                        End If
                        Return JsonError("页面包含数据表格/图表/KPI 页，需要数据来源：rangeAddress、headers+rows、或先选中 Excel 数据区域。")
                    End If

                    ' 前端没传数据（或数据不完整）时，回退：按地址从 Excel 读取（经接口解析，避免依赖具体数据桥）
                    If (headers Is Nothing OrElse headers.Count = 0) OrElse (rows Is Nothing OrElse rows.Count = 0) Then
                        Dim reader = BridgeRegistry.Instance.Resolve(Of IExcelDataReader)()
                        Dim dataJson As String = If(reader IsNot Nothing, reader.GetRangeData(rangeAddress), Nothing)
                        If String.IsNullOrWhiteSpace(dataJson) Then
                            Return JsonError("读取选区数据失败")
                        End If

                        Dim data As JObject = Nothing
                        Try
                            data = JObject.Parse(dataJson)
                        Catch
                            Return JsonError("读取选区数据解析失败")
                        End Try

                        If Not data.Value(Of Boolean)("success") Then
                            ' 传播后端既有错误（如 "未找到工作簿"/"地址错误"）
                            Dim errMsg As String = If(data.Value(Of String)("error"), "")
                            Return JsonError(If(String.IsNullOrWhiteSpace(errMsg), "读取选区数据失败", errMsg))
                        End If

                        Try
                            headers = data("headers")?.ToObject(Of List(Of String))()
                        Catch
                            headers = Nothing
                        End Try
                        Try
                            rows = data("rows")?.ToObject(Of List(Of Dictionary(Of String, Object)))()
                        Catch
                            rows = Nothing
                        End Try
                    End If

                    If headers Is Nothing OrElse headers.Count = 0 Then
                        Return JsonError("选区没有有效的表头（第一行），请确认选区域包含字段名")
                    End If
                End If

                progress?.Invoke(20, "数据就绪")

                ' ---- 数据感知页面方向：前端未显式传时按数据形状推断（竖版长表 → 竖版） ----
                If String.IsNullOrWhiteSpace(orientation) Then
                    orientation = PptDesignEngine.RecommendOrientation(headers, If(rows Is Nothing, 0, rows.Count))
                End If
                Dim isPortrait As Boolean = String.Equals(orientation, "portrait", StringComparison.OrdinalIgnoreCase)
                ' 版式比例（横版）：wide16x9=1280x720；4x3=960x720；竖版固定 720x960
                Dim ratioLower As String = If(ratio Is Nothing, "", ratio.ToLowerInvariant())
                Dim canvasW As Integer = If(ratioLower = "16x9" OrElse ratioLower = "wide" OrElse ratioLower = "169" OrElse ratioLower = "16:9", 1280, DESIGN_W)
                CanvasW = If(isPortrait, 720, canvasW)
                CanvasH = If(isPortrait, 960, DESIGN_H)

                ' 来源描述（展示用）
                Dim sheetName As String = ""
                Dim addr As String = ""
                If args IsNot Nothing Then
                    sheetName = If(args.Value(Of String)("sheetName"), "")
                    addr = If(args.Value(Of String)("address"), "")
                End If
                Dim src As String = If(String.IsNullOrWhiteSpace(sheetName), "", sheetName)
                If Not String.IsNullOrWhiteSpace(addr) Then src = If(src = "", addr, src & " · " & addr)
                If String.IsNullOrWhiteSpace(src) Then src = If(String.IsNullOrWhiteSpace(sourceLabel), "Excel 选区", sourceLabel)

                ' autoDesign 逻辑已在上方（数据检查之前）执行完毕，此处不再重复

                ' ---- 保存目录：优先工作簿所在目录，其次"我的文档" ----
                Dim saveDir As String = GetWorkbookDirectory()
                Dim safeName As String = SanitizeFileName(reportTitle)
                Dim filePath As String = Path.Combine(saveDir, safeName & "_" & Date.Now.ToString("yyyyMMdd_HHmmss") & ".pptx")

                ' ---- 生成 PPT（失败抛异常，由本方法统一转错误） ----
                ' 注意：BuildPresentation 内部已完成"保存 + 用同一个 PowerPoint 实例打开文件"，
                ' 不再需要 Process.Start 重开，避免 CreateObject→Quit→Shell 三段式闪烁。
                progress?.Invoke(30, "启动 PowerPoint 渲染")
                BuildPresentation(headers, rows, reportTitle, src, filePath, theme, slidesSpec, isPortrait, chartMode, CanvasW, chartAgg, progress)

                ' ---- 生成历史日志：成功记录 ----
                AppendGenerationLog(True, reportTitle, theme.Name, If(slidesSpec Is Nothing, 0, slidesSpec.Count), filePath, "")

                ' ---- 结果总结：文件路径、页数/页型统计、主题、比例、图表形式 ----
                Dim typeStats As List(Of String) = BuildSlideTypeStats(slidesSpec)
                Dim chartPages As Integer = CountSlideType(slidesSpec, "chart")
                Dim pageCount As Integer = If(slidesSpec Is Nothing, 0, slidesSpec.Count)
                Dim ratioDesc As String = If(isPortrait, "竖版 A4（720×960）", If(canvasW = 1280, "16:9 宽屏（1280×720）", "4:3 标准（960×720）"))
                Dim summary As New StringBuilder()
                summary.AppendLine("PPT 已生成并已打开")
                summary.AppendLine("文件：" & filePath)
                summary.AppendLine("主题：" & theme.Name)
                summary.AppendLine("比例：" & ratioDesc)
                If pageCount > 0 Then
                    summary.AppendLine("页数：" & pageCount.ToString() & " 页（" & String.Join("、", typeStats.ToArray()) & "）")
                Else
                    summary.AppendLine("页数：标题页 + 数据页")
                End If
                summary.AppendLine(DescribeChartMode(chartMode, chartPages))

                progress?.Invoke(100, "完成")
                Return JsonSuccess(summary.ToString().TrimEnd(), filePath)
            Catch ex As Exception
                ChatLogger.Instance.LogError("PPT_GEN", "生成 PPT 异常: " & ex.ToString())
                ' ---- 生成历史日志：失败记录 ----
                AppendGenerationLog(False, reportTitle, "", 0, "", ex.Message)
                Return JsonError("生成 PPT 失败：" & ex.Message)
            End Try
        End Function

        ' ------------------------------------------------------------
        ' 主题加载：ppt_templates.json（内置） + custom_ppt_themes.json（用户导入）合并 → 内置商务蓝兜底（缓存）
        ' 辅助方法已迁移至 PptGenerationService.Theme.vb（partial class），文档随方法保留
        ' ------------------------------------------------------------

        ' ------------------------------------------------------------
        ' 公司模板主题导入：从已有 .pptx 提取配色 + 字体，注册为自定义主题
        ' 数据源：ppt/theme/theme1.xml（clrScheme + fontScheme）
        ' ------------------------------------------------------------

        ' ------------------------------------------------------------
        ' 核心：用 PowerPoint COM 生成演示文稿
        ' slidesSpec 为空时走旧版固定结构（标题 + 表格 + 概览），否则按布局引擎逐页生成
        ' ------------------------------------------------------------
        Private Shared Sub BuildPresentation(headers As List(Of String),
                                             rows As List(Of Dictionary(Of String, Object)),
                                             title As String,
                                             sourceDesc As String,
                                             filePath As String,
                                             theme As PptTheme,
                                             slidesSpec As JArray,
                                             isPortrait As Boolean,
                                             Optional chartMode As String = "",
                                             Optional canvasWidth As Integer = 960,
                                             Optional chartAgg As String = "sum",
                                             Optional progress As Action(Of Integer, String) = Nothing)
            Dim pptApp As Object = Nothing
            Dim pres As Object = Nothing
            Dim toRelease As New List(Of Object)()
            Dim stepName As String = "初始化"
            Try
                stepName = "创建 PowerPoint 应用"
                ChatLogger.Instance.LogInfo("PPT_GEN", "步骤1/4：创建 PowerPoint 应用")
                pptApp = CreateObject("PowerPoint.Application")
                If pptApp Is Nothing Then
                    Throw New Exception("无法创建 PowerPoint 对象，请确认已安装 PowerPoint")
                End If
                ' 【静默生成 2026-10-07】全程隐藏 PowerPoint 窗口：
                ' 此前 Visible=True 会让 PPT 窗口在生成过程中弹出/抢焦点（多次生成=多个窗口闪烁），
                ' 而 Finally 里本来就会 pres.Close() + pptApp.Quit()，可见窗口最终也会消失，
                ' 可见性对结果没有任何价值，只会打扰用户。生成结束只回报文件路径。
                Try
                    pptApp.Visible = False ' msoFalse
                Catch
                    ' 部分环境不允许改可见性，忽略即可（不影响后台生成）
                End Try
                Try
                    ChatLogger.Instance.LogInfo("PPT_GEN", "PowerPoint 版本: " & pptApp.Version.ToString())
                Catch
                End Try
                pres = pptApp.Presentations.Add(False)
                If pres Is Nothing Then
                    Throw New Exception("无法创建演示文稿（PowerPoint 初始化失败）")
                End If
                progress?.Invoke(40, "创建 PowerPoint 应用")

                ' ---- 显式设置页面尺寸与方向（数据感知：竖版长表 → 竖版，否则横版） ----
                ' 设计空间 960x720：用自定义尺寸精确对齐。
                ' 注意：PpSlideSizeType 枚举中 14 是 HagakiCard 明信片(360x234)，不是 4:3 屏幕，
                ' 必须用 ppSlideSizeCustom(7) + SlideWidth/Height 显式指定。
                Try
                    pres.PageSetup.SlideSize = 7 ' ppSlideSizeCustom
                    pres.PageSetup.SlideWidth = canvasWidth
                    pres.PageSetup.SlideHeight = 720
                    pres.PageSetup.SlideOrientation = If(isPortrait, 2, 1) ' 2=portrait(交换宽高→720x960), 1=landscape(画布宽x720)
                Catch ex As Exception
                    ChatLogger.Instance.LogWarning("PPT_GEN", "设置页面方向失败（使用默认）: " & ex.Message)
                End Try

                If slidesSpec IsNot Nothing AndAlso slidesSpec.Count > 0 Then
                    ' ---- 布局引擎：按 slides 规格逐页生成 ----
                    stepName = "按规格生成幻灯片"
                    ChatLogger.Instance.LogInfo("PPT_GEN", "步骤2/4：布局引擎生成 " & slidesSpec.Count.ToString() & " 页")
                    Dim totalSlides As Integer = If(slidesSpec Is Nothing, 0, slidesSpec.Count)
                    Dim slideIdx As Integer = 0
                    For Each item As JToken In slidesSpec
                        Dim spec As JObject = TryCast(item, JObject)
                        If spec Is Nothing Then Continue For
                        slideIdx += 1
                        progress?.Invoke(40 + CInt(40 * slideIdx / Math.Max(1, totalSlides)), "生成第 " & slideIdx & "/" & totalSlides & " 页")
                        BuildSlideBySpec(pres, toRelease, spec, theme, headers, rows, sourceDesc, chartMode, chartAgg)
                    Next
                Else
                    ' ---- 旧版固定结构：标题 + 表格 + 概览 ----
                    stepName = "生成标题页"
                    ChatLogger.Instance.LogInfo("PPT_GEN", "步骤2/4：生成标题页")
                    progress?.Invoke(50, "生成标题页")
                    Dim legacyTitleSpec As New JObject()
                    legacyTitleSpec("title") = title
                    BuildTitleSlide(pres, toRelease, legacyTitleSpec, theme, sourceDesc)

                    Dim rowCount As Integer = If(rows Is Nothing, 0, rows.Count)
                    If rowCount > 0 AndAlso headers.Count > 0 Then
                        stepName = "生成数据表格页"
                        ChatLogger.Instance.LogInfo("PPT_GEN", "步骤3/4：生成数据表格页（" & rowCount.ToString() & " 行）")
                        progress?.Invoke(65, "生成数据表格页")
                        BuildTableSlide(pres, toRelease, headers, rows, "数据明细", sourceDesc, theme)
                    End If

                    stepName = "生成数据概览页"
                    ChatLogger.Instance.LogInfo("PPT_GEN", "步骤3b/4：生成数据概览页")
                    progress?.Invoke(75, "生成数据概览页")
                    BuildOverviewSlide(pres, toRelease, headers, rows, sourceDesc, theme)
                End If

                ' ---- 保存（显式指定 OpenXML 格式，避免按扩展名推断出错） ----
                stepName = "保存 PPT"
                ChatLogger.Instance.LogInfo("PPT_GEN", "步骤4/4：保存到 " & filePath)
                pres.SaveAs(filePath, 24) ' ppSaveAsOpenXMLPresentation = 24
                ChatLogger.Instance.LogInfo("PPT_GEN", "PPT 保存成功")
                progress?.Invoke(90, "保存文件")

                ' ---- 让当前 PowerPoint 实例直接打开刚保存的文件（不再 Quit 重开） ----
                ' 改前：SaveAs → Finally Quit → 外层 Process.Start（Shell 重开），
                '       形成 CreateObject→Quit→Shell 三段式，用户看到窗口"一闪而过再弹出"。
                ' 改后：同一个 pptApp 实例 SaveAs 完毕后直接打开刚保存的文件并 Visible=True，
                '       整个流程只启动一次 PowerPoint。
                ChatLogger.Instance.LogInfo("PPT_GEN", "步骤5/5：打开生成的 PPT")
                Try
                    pres.Close()
                Catch
                End Try
                Try
                    ' 先尝试获取已存在的文件（SaveAs 刚写完，PowerPoint 可能会缓存）
                    pptApp.Visible = True
                    pres = pptApp.Presentations.Open(filePath, ReadOnly:=False, WithWindow:=True)
                    ChatLogger.Instance.LogInfo("PPT_GEN", "已打开生成的 PPT（复用当前 PowerPoint 实例）")
                    progress?.Invoke(95, "打开生成的 PPT")
                Catch opEx As Exception
                    ChatLogger.Instance.LogWarning("PPT_GEN", "PowerPoint 实例内打开失败，回退 Shell 方式：" & opEx.Message)
                    Try
                        Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
                    Catch
                    End Try
                End Try
            Catch ex As Exception
                Throw New Exception(stepName & " 失败：" & ex.Message, ex)
            Finally
                ' 释放本地新建的 COM 对象；注意绝不释放宿主全局 Application/Workbook
                For i As Integer = toRelease.Count - 1 To 0 Step -1
                    Try
                        ComReleaseHelper.Release(toRelease(i))
                    Catch
                    End Try
                Next
                toRelease.Clear()
                ' 【不 Quit pptApp】让 PowerPoint 实例保持运行打开刚保存的文件。
                ' pres 对象可能已被上面的 Close()/Open() 替换；如果 Open 失败回退 Shell，
                ' 这里才需要关闭 pptApp 残留。
                If pres Is Nothing Then
                    Try
                        If pptApp IsNot Nothing Then pptApp.Quit()
                    Catch
                    End Try
                    pptApp = Nothing
                End If
                ' 不强制 GC —— Visible=True 后用户可能继续操作，让 RCW 自行管理生命周期
            End Try
        End Sub

        Private Shared Sub BuildTableSlide(pres As Object,
                                           toRelease As List(Of Object),
                                           headers As List(Of String),
                                           rows As List(Of Dictionary(Of String, Object)),
                                           reportTitle As String,
                                           sourceDesc As String,
                                           theme As PptTheme)
            Dim s As Object = pres.Slides.Add(pres.Slides.Count + 1, 12)
            SetSlideBackground(s, toRelease, theme.Bg)
            Dim slideTitle As String = If(String.IsNullOrWhiteSpace(reportTitle), "数据明细", reportTitle)
            AddTextBox(s, toRelease, slideTitle, TitleLeft(), TitleTop(), TitleWidth(), TitleHeight(), 26, True, theme.Primary, theme.Font)

            Dim cap As PptCapacityProfile = CapacityOf(_activeLayout)
            Dim colMax As Integer = If(cap Is Nothing, 12, cap.MaxTableCols)
            Dim rowMax As Integer = If(cap Is Nothing, 25, cap.MaxTableRows)
            Dim colCount As Integer = Math.Min(headers.Count, colMax)
            Dim dataRows As Integer = Math.Min(rows.Count, rowMax)
            Dim shape As Object = Nothing
            Try
                shape = s.Shapes.AddTable(dataRows + 1, colCount, SX(TableLeft()), SY(TableTop()), SW(TableWidth()), SH(Math.Max(TableHeight(), dataRows * 14 + 20)))
            Catch
                shape = Nothing
            End Try
            If shape Is Nothing Then
                AddTextBox(s, toRelease, "表格数据过多，请在 PowerPoint 中调整", 40, ContentAreaTop(), 880, 30, 12, False, theme.Gray, theme.Font)
                AddFooter(s, toRelease, pres.Slides.Count, theme.Gray, theme.Font)
                Exit Sub
            End If
            Dim tbl As Object = shape.Table
            toRelease.Add(shape)
            ApplyTableColumnWidths(tbl, colCount)

            ' 表头
            For c As Integer = 0 To colCount - 1
                Try
                    Dim cell As Object = tbl.Cell(1, c + 1)
                    cell.Shape.Fill.ForeColor.RGB = theme.Primary
                    cell.Shape.TextFrame.TextRange.Text = CStr(headers(c))
                    cell.Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
                    cell.Shape.TextFrame.TextRange.Font.Bold = True
                    cell.Shape.TextFrame.TextRange.Font.Name = theme.Font
                    cell.Shape.TextFrame.TextRange.Font.Size = 11
                Catch
                End Try
            Next

            ' 数据行（斑马纹）
            Dim key As String
            Dim val As Object
            Dim cellText As String = ""
            For r As Integer = 0 To dataRows - 1
                Dim row As Dictionary(Of String, Object) = rows(r)
                For c As Integer = 0 To colCount - 1
                    key = headers(c)
                    val = Nothing
                    If row IsNot Nothing AndAlso row.ContainsKey(key) Then val = row(key)
                    If val Is Nothing OrElse val Is DBNull.Value Then
                        cellText = ""
                    Else
                        Try
                            cellText = CStr(val)
                        Catch
                            cellText = ""
                        End Try
                    End If
                    Try
                        Dim cell As Object = tbl.Cell(r + 2, c + 1)
                        cell.Shape.TextFrame.TextRange.Text = cellText
                        cell.Shape.TextFrame.TextRange.Font.Name = theme.Font
                        cell.Shape.TextFrame.TextRange.Font.Size = 10
                        If r Mod 2 = 0 Then
                            cell.Shape.Fill.ForeColor.RGB = theme.Light
                        End If
                    Catch
                    End Try
                Next
            Next

            AddFooter(s, toRelease, pres.Slides.Count, theme.Gray, theme.Font)
        End Sub

        Private Shared Sub BuildOverviewSlide(pres As Object,
                                              toRelease As List(Of Object),
                                              headers As List(Of String),
                                              rows As List(Of Dictionary(Of String, Object)),
                                              sourceDesc As String,
                                              theme As PptTheme)
            Dim s As Object = pres.Slides.Add(pres.Slides.Count + 1, 12)
            SetSlideBackground(s, toRelease, theme.Bg)
            AddTextBox(s, toRelease, "数据概览", TitleLeft(), TitleTop(), TitleWidth(), TitleHeight(), 26, True, theme.Primary, theme.Font)
            AddRibbon(s, toRelease, theme.Secondary, 40, TitleRibbonTop(), 160, 4)

            Dim fieldPreview As String = ""
            For i As Integer = 0 To Math.Min(headers.Count - 1, 8)
                fieldPreview &= headers(i) & "、"
            Next
            If fieldPreview.EndsWith("、") Then fieldPreview = fieldPreview.Substring(0, fieldPreview.Length - 1)

            Dim bulletLines As New List(Of String)()
            bulletLines.Add("数据区域：" & sourceDesc)
            bulletLines.Add("字段（" & headers.Count.ToString() & " 个）：" & fieldPreview)
            bulletLines.Add("" & If(rows Is Nothing, 0, rows.Count).ToString() & " 行数据")

            Dim bullets As String = "• " & String.Join(vbCrLf & "• ", bulletLines.ToArray())

            AddTextBox(s, toRelease, bullets, 40, ContentAreaTop(), 880, 300, 16, False, theme.Text, theme.Font)

            AddFooter(s, toRelease, pres.Slides.Count, theme.Gray, theme.Font)
        End Sub

        ' ------------------------------------------------------------
        ' 布局引擎：按 slides 规格逐页调度
        ' ------------------------------------------------------------
        Private Shared Sub BuildSlideBySpec(pres As Object, toRelease As List(Of Object),
                                            spec As JObject, theme As PptTheme,
                                            headers As List(Of String),
                                            rows As List(Of Dictionary(Of String, Object)),
                                            sourceDesc As String,
                                            Optional chartMode As String = "",
                                            Optional chartAgg As String = "sum")
            Dim slideType As String = spec.Value(Of String)("type")
            If String.IsNullOrWhiteSpace(slideType) Then slideType = "text"
            ' 页型变体注入：spec 未声明 variant 时，套用主题 designTokens.pageVariants 的默认变体
            InjectPageVariant(spec, slideType)
            Select Case slideType.ToLowerInvariant()
                Case "title", "cover"
                    BuildTitleSlide(pres, toRelease, spec, theme, sourceDesc)
                Case "table"
                    Dim tableTitle As String = spec.Value(Of String)("title")
                    If String.IsNullOrWhiteSpace(tableTitle) Then tableTitle = "数据明细"
                    BuildTableSlide(pres, toRelease, headers, rows, tableTitle, sourceDesc, theme)
                Case "chart"
                    BuildChartSlide(pres, toRelease, spec, theme, headers, rows, chartMode, chartAgg)
                Case "kpi"
                    BuildKpiSlide(pres, toRelease, spec, theme, headers, rows)
                Case "text", "bullets"
                    BuildTextSlide(pres, toRelease, spec, theme)
                Case "comparison"
                    BuildComparisonSlide(pres, toRelease, spec, theme)
                Case "timeline"
                    BuildTimelineSlide(pres, toRelease, spec, theme)
                Case "agenda"
                    BuildAgendaSlide(pres, toRelease, spec, theme)
                Case "section"
                    BuildSectionSlide(pres, toRelease, spec, theme)
                Case "closing", "end"
                    BuildClosingSlide(pres, toRelease, spec, theme)
                Case "shapes"
                    BuildShapesSlide(pres, toRelease, spec, theme)
                Case "picture"
                    BuildPictureSlide(pres, toRelease, spec, theme)
                Case "smartart"
                    BuildSmartArtSlide(pres, toRelease, spec, theme)
                Case Else
                    BuildTextSlide(pres, toRelease, spec, theme)
            End Select

            ' 动画挂载：默认对全部页型挂载切换 + 进入动画，spec 可显式关闭（animate=false / animation=none）
            Dim animate As Boolean = True
            If spec("animate") IsNot Nothing Then animate = spec.Value(Of Boolean)("animate")
            Dim animSetting As String = spec.Value(Of String)("animation")
            If String.Equals(animSetting, "none", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(animSetting, "off", StringComparison.OrdinalIgnoreCase) Then
                animate = False
            End If
            If animate Then
                Try
                    Dim cur As Object = pres.Slides(pres.Slides.Count)
                    If cur IsNot Nothing Then
                        ApplySlideAnimation(cur, spec, theme)
                    End If
                Catch
                End Try
            End If
        End Sub

        ''' <summary>
        ''' 页型变体注入：spec 未声明 variant 时，注入主题 designTokens.pageVariants 的默认变体。
        ''' 别名归一：cover→title、bullets→text（渲染分发处的别名）。spec 显式 variant 优先。
        ''' </summary>
        Private Shared Sub InjectPageVariant(spec As JObject, slideType As String)
            If spec Is Nothing OrElse spec("variant") IsNot Nothing Then Return
            Dim tokens = _activeTokens
            If tokens Is Nothing OrElse tokens.PageVariants Is Nothing Then Return
            Dim key As String = slideType.ToLowerInvariant()
            If key = "cover" Then key = "title"
            If key = "bullets" Then key = "text"
            Dim v As String = ""
            If tokens.PageVariants.TryGetValue(key, v) Then
                spec("variant") = v
            End If
        End Sub

        Private Shared Function ValidateSlideSpecs(slidesSpec As JArray) As List(Of String)
            Dim invalid As New List(Of String)()
            If slidesSpec Is Nothing Then Return invalid
            For Each item As JToken In slidesSpec
                Dim spec As JObject = TryCast(item, JObject)
                If spec Is Nothing Then
                    invalid.Add("(非对象)")
                    Continue For
                End If
                Dim slideType As String = spec.Value(Of String)("type")
                If String.IsNullOrWhiteSpace(slideType) Then
                    invalid.Add("(缺少 type)")
                    Continue For
                End If
                If Not SUPPORTED_SLIDE_TYPES.Contains(slideType.ToLowerInvariant()) Then
                    invalid.Add(slideType)
                End If
            Next
            Return invalid
        End Function

        ' ------------------------------------------------------------
        ' 图表类型映射：名称 → PowerPoint AddChart2 枚举
        ' ------------------------------------------------------------

        ' ------------------------------------------------------------
        ' 各版式 builder
        ' ------------------------------------------------------------
        Private Shared Sub BuildTitleSlide(pres As Object, toRelease As List(Of Object),
                                           spec As JObject, theme As PptTheme, sourceDesc As String)
            Dim s As Object = pres.Slides.Add(pres.Slides.Count + 1, 12)
            SetSlideBackground(s, toRelease, theme.Bg)
            Dim title As String = spec.Value(Of String)("title")
            If String.IsNullOrWhiteSpace(title) Then title = "数据报告"
            Dim subtitle As String = spec.Value(Of String)("subtitle")
            If String.IsNullOrWhiteSpace(subtitle) Then subtitle = "数据来源：" & sourceDesc
            Dim slideVariant As String = spec.Value(Of String)("variant")
            If String.IsNullOrWhiteSpace(slideVariant) Then slideVariant = "split"

            Select Case slideVariant.ToLowerInvariant()
                Case "centered"
                    ' 居中构图：顶部渐变带 + 两侧装饰圆 + 居中标题
                    AddGradientRect(s, toRelease, 0, 0, DESIGN_W, 8, theme.Primary, theme.Secondary, 0)
                    AddDecorativeCircle(s, toRelease, 60, 60, 220, theme.Light, 0.9)
                    AddDecorativeCircle(s, toRelease, 700, 420, 260, theme.Light, 0.9)
                    AddDecorativeCircle(s, toRelease, 120, 480, 120, theme.Secondary, 0.92)
                    AddCenteredTextBox(s, toRelease, title, 80, 240, 800, 90, 44, True, theme.Primary, theme.Font)
                    AddCenteredTextBox(s, toRelease, subtitle, 120, 350, 720, 40, 18, False, theme.Secondary, theme.Font)
                    AddCenteredTextBox(s, toRelease, Date.Now.ToString("yyyy-MM-dd"), 360, 420, 240, 30, SMALL_SIZE, False, theme.Gray, theme.Font)
                    AddRibbon(s, toRelease, theme.Secondary, 0, DESIGN_H - 6, DESIGN_W, 6)
                Case "modern"
                    ' 现代风：顶部渐变带 + 左侧主标题 + 右侧装饰圆 + 日期徽标
                    AddGradientRect(s, toRelease, 0, 0, DESIGN_W, 10, theme.Primary, theme.Secondary, 0)
                    AddRibbon(s, toRelease, theme.Primary, 0, 0, 10, DESIGN_H)
                    AddDecorativeCircle(s, toRelease, 640, 60, 300, theme.Light, 0.9)
                    AddDecorativeCircle(s, toRelease, 760, 180, 160, theme.Secondary, 0.93)
                    AddTextBox(s, toRelease, title, 80, 250, 560, 90, 42, True, theme.Primary, theme.Font)
                    AddRibbon(s, toRelease, theme.Accent, 80, 360, 90, 6)
                    AddTextBox(s, toRelease, subtitle, 80, 390, 560, 40, 18, False, theme.Secondary, theme.Font)
                    AddBadge(s, toRelease, Date.Now.ToString("yyyy-MM-dd"), 80, 460, 150, 34, theme.Light, theme.Primary, SMALL_SIZE, theme.Font)
                    AddRibbon(s, toRelease, theme.Secondary, 0, DESIGN_H - 6, DESIGN_W, 6)
                Case Else ' split
                    ' 左右分栏：左侧浅色圆角面板 + 装饰圆 + 右侧标题
                    AddRibbon(s, toRelease, theme.Primary, 0, 0, DESIGN_W, 8)
                    AddRibbon(s, toRelease, theme.Primary, 0, 0, 12, DESIGN_H)
                    AddRoundedRect(s, toRelease, 40, 60, 300, 600, theme.Light, 16)
                    AddDecorativeCircle(s, toRelease, 60, 100, 200, theme.Primary, 0.94)
                    AddDecorativeCircle(s, toRelease, 130, 420, 140, theme.Secondary, 0.93)
                    AddTextBox(s, toRelease, title, 400, 230, 500, 80, 40, True, theme.Primary, theme.Font)
                    AddTextBox(s, toRelease, subtitle, 400, 330, 500, 40, 18, False, theme.Secondary, theme.Font)
                    AddTextBox(s, toRelease, Date.Now.ToString("yyyy-MM-dd"), 400, 395, 300, 30, SMALL_SIZE, False, theme.Gray, theme.Font)
                    AddRibbon(s, toRelease, theme.Secondary, 0, DESIGN_H - 6, DESIGN_W, 6)
            End Select
            AddFooter(s, toRelease, 1, theme.Gray, theme.Font)
        End Sub




        ' 图表辅助方法已迁移至 PptGenerationService.Charts.vb（partial class），文档随方法保留

        ' KPI 网格变体：2 行 x 3 列圆角卡片，顶部渐变强调条，高亮项加 TOP 徽标

        ' KPI 主角变体：首个指标为左侧大卡（渐变），其余右侧堆叠

        ' KPI 横排变体：一行卡片，高亮项用强调色


        ' 文本要点变体：圆角方块标记 + 行距优化

        ' 文本卡片变体：每条要点一张圆角卡片（适合少量要点）

        ' 文本编号变体：编号圆点 + 步骤式排版（适合流程/步骤）

        Private Shared Sub BuildComparisonSlide(pres As Object, toRelease As List(Of Object),
                                                spec As JObject, theme As PptTheme)
            Dim s As Object = pres.Slides.Add(pres.Slides.Count + 1, 12)
            SetSlideBackground(s, toRelease, theme.Bg)
            Dim title As String = spec.Value(Of String)("title")
            If String.IsNullOrWhiteSpace(title) Then title = "对比"
            AddTextBox(s, toRelease, title, TitleLeft(), TitleTop(), TitleWidth(), TitleHeight(), 26, True, theme.Primary, theme.Font)
            AddRibbon(s, toRelease, theme.Secondary, 40, TitleRibbonTop(), 160, 4)

            Dim leftTitle As String = spec.Value(Of String)("leftTitle")
            Dim rightTitle As String = spec.Value(Of String)("rightTitle")
            If String.IsNullOrWhiteSpace(leftTitle) Then leftTitle = "方案 A"
            If String.IsNullOrWhiteSpace(rightTitle) Then rightTitle = "方案 B"

            Dim colTop As Integer = ContentAreaTop()
            AddTextBox(s, toRelease, leftTitle, 40, colTop, 420, 40, 22, True, theme.Primary, theme.Font)
            AddTextBox(s, toRelease, JoinBullets(spec, "leftItems"), 40, colTop + 50, 420, 400, 16, False, theme.Text, theme.Font)
            AddTextBox(s, toRelease, rightTitle, 500, colTop, 420, 40, 22, True, theme.Secondary, theme.Font)
            AddTextBox(s, toRelease, JoinBullets(spec, "rightItems"), 500, colTop + 50, 420, 400, 16, False, theme.Text, theme.Font)

            AddFooter(s, toRelease, pres.Slides.Count, theme.Gray, theme.Font)
        End Sub

        Private Shared Function JoinBullets(spec As JObject, field As String) As String
            Dim arr As JArray = spec.Value(Of JArray)(field)
            Dim lines As New List(Of String)()
            If arr IsNot Nothing Then
                For Each b As JToken In arr
                    Dim txt As String = FormatBulletItem(b)
                    If Not String.IsNullOrWhiteSpace(txt) Then lines.Add(txt)
                Next
            End If
            If lines.Count = 0 Then Return "（无内容）"
            Return "• " & String.Join(vbCrLf & "• ", lines.ToArray())
        End Function

        ''' <summary>
        ''' 把 bullet item 规范成展示文本：字符串原样返回；对象（{label, value}/{text, ...}）
        ''' 取 label/text 与 value 格式化为 "label：value"，避免把整段 JSON 字符串直接打上幻灯片。
        ''' </summary>
        Friend Shared Function FormatBulletItem(b As JToken) As String
            If b Is Nothing Then Return ""
            If b.Type = JTokenType.String Then
                Return b.ToString()
            End If
            Dim obj As JObject = TryCast(b, JObject)
            If obj Is Nothing Then Return b.ToString()

            Dim label As String = obj.Value(Of String)("label")
            If String.IsNullOrEmpty(label) Then label = obj.Value(Of String)("text")
            If String.IsNullOrEmpty(label) Then label = obj.Value(Of String)("name")

            Dim valueTok As JToken = obj("value")
            Dim hasValue As Boolean = valueTok IsNot Nothing AndAlso valueTok.Type <> JTokenType.Null
            If Not hasValue Then
                ' 对象只有 label 时也至少显示 label（不输出杂 JSON）
                Return If(String.IsNullOrWhiteSpace(label), "", label)
            End If

            If Not String.IsNullOrEmpty(label) Then
                Return label & "：" & valueTok.ToString()
            End If
            ' 无 label 的裸值对象 → 直接显示值
            Return valueTok.ToString()
        End Function

        Private Shared Sub BuildTimelineSlide(pres As Object, toRelease As List(Of Object),
                                              spec As JObject, theme As PptTheme)
            Dim s As Object = pres.Slides.Add(pres.Slides.Count + 1, 12)
            SetSlideBackground(s, toRelease, theme.Bg)
            Dim title As String = spec.Value(Of String)("title")
            If String.IsNullOrWhiteSpace(title) Then title = "时间线"
            AddTextBox(s, toRelease, title, TitleLeft(), TitleTop(), TitleWidth(), TitleHeight(), 26, True, theme.Primary, theme.Font)
            AddRibbon(s, toRelease, theme.Secondary, 40, TitleRibbonTop(), 160, 4)

            Dim items As JArray = spec.Value(Of JArray)("items")
            Dim count As Integer = If(items Is Nothing, 0, Math.Min(items.Count, 6))
            If count = 0 Then
                AddTextBox(s, toRelease, "（无时间线内容）", 40, ContentAreaTop(), 880, 40, 16, False, theme.Gray, theme.Font)
                AddFooter(s, toRelease, pres.Slides.Count, theme.Gray, theme.Font)
                Return
            End If

            ' 水平时间线：轴线 + 圆点 + 时间 + 描述
            Dim lineY As Integer = 300
            Dim startX As Integer = 80
            Dim endX As Integer = 880
            Dim stepX As Integer = (endX - startX) / count

            Dim axis As Object = Nothing
            Try
                axis = s.Shapes.AddShape(1, SX(startX), SY(lineY), SW(endX - startX), SH(4))
                axis.Fill.ForeColor.RGB = theme.Secondary
                Try
                    axis.Line.Visible = False
                Catch
                End Try
                toRelease.Add(axis)
            Catch
            End Try

            For i As Integer = 0 To count - 1
                Dim obj As JObject = TryCast(items(i), JObject)
                Dim timeText As String = ""
                Dim desc As String = ""
                If obj IsNot Nothing Then
                    timeText = If(obj.Value(Of String)("time"), "")
                    desc = If(obj.Value(Of String)("text"), "")
                End If
                Dim cx As Integer = startX + i * stepX
                Dim dot As Object = Nothing
                Try
                    dot = s.Shapes.AddShape(9, SX(cx - 8), SY(lineY - 8), SW(16), SH(16)) ' msoShapeOval = 9
                    dot.Fill.ForeColor.RGB = theme.Primary
                    Try
                        dot.Line.Visible = False
                    Catch
                    End Try
                    toRelease.Add(dot)
                Catch
                End Try
                AddTextBox(s, toRelease, timeText, cx - 60, lineY - 90, 120, 30, 14, True, theme.Primary, theme.Font)
                AddTextBox(s, toRelease, desc, cx - 60, lineY + 30, 120, 80, 12, False, theme.Text, theme.Font)
            Next
            AddFooter(s, toRelease, pres.Slides.Count, theme.Gray, theme.Font)
        End Sub

        Private Shared Sub BuildAgendaSlide(pres As Object, toRelease As List(Of Object),
                                            spec As JObject, theme As PptTheme)
            Dim s As Object = pres.Slides.Add(pres.Slides.Count + 1, 12)
            SetSlideBackground(s, toRelease, theme.Bg)
            Dim title As String = spec.Value(Of String)("title")
            If String.IsNullOrWhiteSpace(title) Then title = "目录"
            AddTextBox(s, toRelease, title, TitleLeft(), TitleTop(), TitleWidth(), TitleHeight(), TITLE_SIZE, True, theme.Primary, theme.Font)
            If RIBBON_STYLE <> "none" Then
                If RIBBON_STYLE = "bar" Then
                    AddRibbon(s, toRelease, theme.Secondary, MARGIN, TitleRibbonTop(), Math.Min(RIBBON_WIDTH * 2, DESIGN_W - 2 * MARGIN), 8)
                Else
                    AddRibbon(s, toRelease, theme.Secondary, MARGIN, TitleRibbonTop(), RIBBON_WIDTH, 4)
                End If
            End If

            Dim lines As New List(Of String)()
            Dim items As JArray = spec.Value(Of JArray)("items")
            If items IsNot Nothing Then
                For Each b As JToken In items
                    Dim txt As String = b.ToString()
                    If Not String.IsNullOrWhiteSpace(txt) Then lines.Add(txt)
                Next
            End If
            If lines.Count = 0 Then lines.Add("（无目录内容）")

            Dim slideVariant As String = spec.Value(Of String)("variant")
            If String.IsNullOrWhiteSpace(slideVariant) Then slideVariant = "numbered"
            Select Case slideVariant.ToLowerInvariant()
                Case "cards"
                    BuildAgendaCards(s, toRelease, lines, theme)
                Case "timeline"
                    BuildAgendaTimeline(s, toRelease, lines, theme)
                Case Else
                    BuildAgendaNumbered(s, toRelease, lines, theme)
            End Select
            AddFooter(s, toRelease, pres.Slides.Count, theme.Gray, theme.Font)
        End Sub

        ' 目录编号变体：圆角编号块 + 标题（两列，每列最多 5 项）
        Private Shared Sub BuildAgendaNumbered(s As Object, toRelease As List(Of Object),
                                               lines As List(Of String), theme As PptTheme)
            Dim itemH As Integer = 64
            Dim startY As Integer = Math.Max(140, ContentAreaTop() + 20)
            For i As Integer = 0 To lines.Count - 1
                Dim col As Integer = i \ 5
                Dim rowIdx As Integer = i Mod 5
                Dim x As Integer = MARGIN + col * 440
                Dim y As Integer = startY + rowIdx * itemH
                AddRoundedRect(s, toRelease, x, y, 36, 36, theme.Primary, 8)
                AddCenteredTextBox(s, toRelease, (i + 1).ToString(), x, y, 36, 36, 16, True, RGB(255, 255, 255), theme.Font)
                AddTextBox(s, toRelease, lines(i), x + 52, y + 5, 360, 28, 18, False, theme.Text, theme.Font)
            Next
        End Sub

        ' 目录卡片变体：章节卡片（适合少量章节）
        Private Shared Sub BuildAgendaCards(s As Object, toRelease As List(Of Object),
                                            lines As List(Of String), theme As PptTheme)
            Dim count As Integer = Math.Min(lines.Count, 6)
            Dim cardW As Integer = 420
            Dim cardH As Integer = 96
            Dim gapX As Integer = 40
            Dim gapY As Integer = 24
            Dim startY As Integer = Math.Max(140, ContentAreaTop() + 20)
            For i As Integer = 0 To count - 1
                Dim col As Integer = i Mod 2
                Dim rowIdx As Integer = i \ 2
                Dim x As Integer = MARGIN + col * (cardW + gapX)
                Dim y As Integer = startY + rowIdx * (cardH + gapY)
                AddRoundedRect(s, toRelease, x, y, cardW, cardH, theme.Light, 12)
                AddGradientRect(s, toRelease, x, y, 8, cardH, theme.Primary, theme.Secondary, 0)
                AddCenteredTextBox(s, toRelease, (i + 1).ToString("00"), x + 24, y + 24, 48, 48, 20, True, theme.Secondary, theme.Font)
                AddTextBox(s, toRelease, lines(i), x + 92, y + 30, cardW - 120, 40, 18, False, theme.Text, theme.Font)
            Next
        End Sub

        ' 目录时间线变体：纵向时间线（适合 4-6 个章节）
        Private Shared Sub BuildAgendaTimeline(s As Object, toRelease As List(Of Object),
                                               lines As List(Of String), theme As PptTheme)
            Dim count As Integer = Math.Min(lines.Count, 6)
            Dim startY As Integer = Math.Max(150, ContentAreaTop() + 30)
            Dim lineH As Integer = 84
            Dim axisX As Integer = 90
            Dim axis As Object = Nothing
            Try
                axis = s.Shapes.AddShape(1, SX(axisX - 2), SY(startY), SW(4), SH((count - 1) * lineH + 40)) ' msoShapeRectangle
                axis.Fill.ForeColor.RGB = theme.Light
                Try
                    axis.Line.Visible = False
                Catch
                End Try
                toRelease.Add(axis)
            Catch
            End Try
            For i As Integer = 0 To count - 1
                Dim y As Integer = startY + i * lineH
                Dim dot As Object = Nothing
                Try
                    dot = s.Shapes.AddShape(9, SX(axisX - 12), SY(y), SW(24), SH(24)) ' msoShapeOval = 9
                    dot.Fill.ForeColor.RGB = theme.Primary
                    Try
                        dot.Line.Visible = False
                    Catch
                    End Try
                    toRelease.Add(dot)
                Catch
                End Try
                AddCenteredTextBox(s, toRelease, (i + 1).ToString(), axisX - 12, y, 24, 24, 11, True, RGB(255, 255, 255), theme.Font)
                AddTextBox(s, toRelease, lines(i), axisX + 40, y + 2, 700, 40, 18, False, theme.Text, theme.Font)
            Next
        End Sub

        Private Shared Sub BuildSectionSlide(pres As Object, toRelease As List(Of Object),
                                             spec As JObject, theme As PptTheme)
            Dim s As Object = pres.Slides.Add(pres.Slides.Count + 1, 12)
            SetSlideBackground(s, toRelease, theme.Bg)
            Dim title As String = spec.Value(Of String)("title")
            If String.IsNullOrWhiteSpace(title) Then title = "章节"
            Dim subtitle As String = spec.Value(Of String)("subtitle")
            Dim number As String = spec.Value(Of String)("number")
            If String.IsNullOrWhiteSpace(number) Then
                number = (pres.Slides.Count).ToString("00")
            End If
            Dim slideVariant As String = spec.Value(Of String)("variant")
            If String.IsNullOrWhiteSpace(slideVariant) Then slideVariant = "number"

            Select Case slideVariant.ToLowerInvariant()
                Case "centered"
                    ' 居中章节页：顶部渐变带 + 居中大标题
                    AddGradientRect(s, toRelease, 0, 0, DESIGN_W, 8, theme.Primary, theme.Secondary, 0)
                    AddDecorativeCircle(s, toRelease, 60, 60, 200, theme.Light, 0.9)
                    AddDecorativeCircle(s, toRelease, 720, 460, 240, theme.Light, 0.9)
                    AddCenteredTextBox(s, toRelease, number, 80, 200, 800, 90, 64, True, theme.Light, theme.Font)
                    AddCenteredTextBox(s, toRelease, title, 80, 300, 800, 80, 40, True, theme.Primary, theme.Font)
                    If Not String.IsNullOrWhiteSpace(subtitle) Then
                        AddCenteredTextBox(s, toRelease, subtitle, 80, 400, 800, 40, 18, False, theme.Secondary, theme.Font)
                    End If
                    AddRibbon(s, toRelease, theme.Accent, 430, 470, 100, 5)
                Case "bar"
                    ' 左侧强调条章节页：大号章节字 + 左竖条
                    AddRibbon(s, toRelease, theme.Primary, 0, 0, 14, DESIGN_H)
                    AddGradientRect(s, toRelease, 0, 0, DESIGN_W, 8, theme.Primary, theme.Secondary, 0)
                    AddDecorativeCircle(s, toRelease, 700, 60, 260, theme.Light, 0.92)
                    AddTextBox(s, toRelease, number, MARGIN, 200, 200, 120, 72, True, theme.Light, theme.Font)
                    AddTextBox(s, toRelease, title, MARGIN, 320, 800, 80, 40, True, theme.Primary, theme.Font)
                    If Not String.IsNullOrWhiteSpace(subtitle) Then
                        AddTextBox(s, toRelease, subtitle, MARGIN, 400, 800, 40, 18, False, theme.Secondary, theme.Font)
                    End If
                    AddRibbon(s, toRelease, theme.Accent, MARGIN, 460, 100, 5)
                Case Else ' number
                    ' 大号章节序号 + 标题（默认）
                    AddRibbon(s, toRelease, theme.Primary, 0, 0, DESIGN_W, 8)
                    AddDecorativeCircle(s, toRelease, 60, 500, 200, theme.Light, 0.9)
                    AddTextBox(s, toRelease, number, MARGIN, 210, 200, 110, 72, True, theme.Light, theme.Font)
                    AddTextBox(s, toRelease, title, MARGIN, 320, 880, 80, 40, True, theme.Primary, theme.Font)
                    If Not String.IsNullOrWhiteSpace(subtitle) Then
                        AddTextBox(s, toRelease, subtitle, MARGIN, 400, 880, 40, 18, False, theme.Secondary, theme.Font)
                    End If
                    AddRibbon(s, toRelease, theme.Secondary, MARGIN, 460, 180, 5)
            End Select
            AddFooter(s, toRelease, pres.Slides.Count, theme.Gray, theme.Font)
        End Sub

        Private Shared Sub BuildClosingSlide(pres As Object, toRelease As List(Of Object),
                                             spec As JObject, theme As PptTheme)
            Dim s As Object = pres.Slides.Add(pres.Slides.Count + 1, 12)
            SetSlideBackground(s, toRelease, theme.Bg)
            Dim title As String = spec.Value(Of String)("title")
            If String.IsNullOrWhiteSpace(title) Then title = "谢谢"
            Dim subtitle As String = spec.Value(Of String)("subtitle")
            Dim slideVariant As String = spec.Value(Of String)("variant")
            If String.IsNullOrWhiteSpace(slideVariant) Then slideVariant = "thanks"

            Select Case slideVariant.ToLowerInvariant()
                Case "centered"
                    ' 居中结束页：渐变带 + 装饰圆 + 居中致谢
                    AddGradientRect(s, toRelease, 0, 0, DESIGN_W, 8, theme.Primary, theme.Secondary, 0)
                    AddDecorativeCircle(s, toRelease, 60, 60, 220, theme.Light, 0.9)
                    AddDecorativeCircle(s, toRelease, 700, 420, 260, theme.Light, 0.9)
                    AddDecorativeCircle(s, toRelease, 120, 480, 120, theme.Secondary, 0.92)
                    AddCenteredTextBox(s, toRelease, title, 80, 260, 800, 90, 48, True, theme.Primary, theme.Font)
                    If Not String.IsNullOrWhiteSpace(subtitle) Then
                        AddCenteredTextBox(s, toRelease, subtitle, 120, 380, 720, 40, 18, False, theme.Secondary, theme.Font)
                    End If
                    AddRibbon(s, toRelease, theme.Accent, 430, 460, 100, 5)
                Case Else ' thanks
                    ' 左右分栏致谢页（与封面呼应）
                    AddRibbon(s, toRelease, theme.Primary, 0, 0, DESIGN_W, 8)
                    AddRibbon(s, toRelease, theme.Primary, 0, 0, 12, DESIGN_H)
                    AddRoundedRect(s, toRelease, 40, 60, 300, 600, theme.Light, 16)
                    AddDecorativeCircle(s, toRelease, 60, 100, 200, theme.Primary, 0.94)
                    AddDecorativeCircle(s, toRelease, 130, 420, 140, theme.Secondary, 0.93)
                    AddTextBox(s, toRelease, title, 400, 300, 500, 80, 44, True, theme.Primary, theme.Font)
                    If Not String.IsNullOrWhiteSpace(subtitle) Then
                        AddTextBox(s, toRelease, subtitle, 400, 400, 500, 40, 18, False, theme.Secondary, theme.Font)
                    End If
                    AddRibbon(s, toRelease, theme.Secondary, 0, DESIGN_H - 6, DESIGN_W, 6)
            End Select
            AddFooter(s, toRelease, pres.Slides.Count, theme.Gray, theme.Font)
        End Sub

        ' ------------------------------------------------------------
        ' 布局辅助
        ' ------------------------------------------------------------



        ' ------------------------------------------------------------
        ' 视觉增强辅助：圆角卡片 / 渐变 / 装饰圆 / 徽标 / 垂直居中文本
        ' 所有坐标均为设计空间（960x720）坐标，经 SX/SY/SW/SH 变换到实际画布。
        ' ------------------------------------------------------------







    End Class
End Namespace
