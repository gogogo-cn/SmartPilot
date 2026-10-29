Imports System.Collections.Generic

Namespace Services

    ''' <summary>
    ''' OCR 结构化元素类型
    ''' </summary>
    Public Enum OcrElementType
        Paragraph       ' 普通段落
        Heading         ' 标题
        ListItem        ' 列表项
        Table           ' 表格
        Image           ' 图片
        Formula         ' 公式（LaTeX）
        LayoutRegion    ' 版面区域（来自版面分析）
        Seal            ' 印章
        PageBreak       ' 分页符
        HorizontalRule  ' 分隔线

        ' === 机械图纸专项（GDT 扩展）===
        GdtSymbol       ' 形位公差符号（⊕⊥⌓⌒∥⌖∠▱◎↗ 等）
        Balloon         ' 球标圆圈（内部有数字的编号圆圈）
        DatumCircle     ' 基准圆圈（内部有 A/B/C/D 等基准字母的圆圈）
        Roughness       ' 表面粗糙度符号（√ 三角 + 横线 + Ra/Rz/Ra 值）
        Dimension       ' 尺寸标注（Ø32H7、170±0.3 等带数值的标注）
    End Enum

    ''' <summary>
    ''' OCR 元素（一个文本块/表格等结构化单元）
    ''' </summary>
    Public Class OcrElement
        ''' <summary>元素类型</summary>
        Public Property Type As OcrElementType = OcrElementType.Paragraph

        ''' <summary>纯文本内容</summary>
        Public Property Text As String = ""

        ''' <summary>标题级别（当 Type=Heading 时）1-6</summary>
        Public Property HeadingLevel As Integer = 0

        ''' <summary>列表标记（当 Type=ListItem 时）</summary>
        Public Property ListMarker As String = ""

        ''' <summary>列表层级（0=顶层）</summary>
        Public Property ListLevel As Integer = 0

        ''' <summary>边界框（左上右下坐标）</summary>
        Public Property BoundingBox As OcrBoundingBox = Nothing

        ''' <summary>
        ''' 引擎返回的原始四边形（4 点，按阅读方向排序）。空表示没有——
        ''' 这时 UI 用 BoundingBox 的四个角画。竖排文字只有靠它才画得出带朝向的盒子，
        ''' 否则永远是个正上下左右对齐的矩形（左上角到右下角），看不出文字是横是竖。
        ''' 注意：专项改写/合并盒子后此字段会失效，改写处负责清掉它。
        ''' </summary>
        Public Property Quad As List(Of Point) = Nothing

        ''' <summary>
        ''' 竖排文字阅读方向：up=从下往上读，down=从上往下读，空=水平文字。
        ''' 由引擎对竖排框旋转重识别时定出。合并被拆分的竖排标注按它接文本顺序。
        ''' </summary>
        Public Property VerticalDir As String = ""

        ''' <summary>识别置信度 0-1</summary>
        Public Property Confidence As Double = 0.0

        ''' <summary>机械图标注种类：直径、半径、线性尺寸、形位公差等。空表示不是一条尺寸。</summary>
        Public Property Kind As String = ""

        ''' <summary>标称值，如 Ø25、M8×1.25、120。</summary>
        Public Property Nominal As String = ""

        ''' <summary>下公差，如 -0.01、0。</summary>
        Public Property Lower As String = ""

        ''' <summary>上公差。形位公差的公差带也写在这里。</summary>
        Public Property Upper As String = ""

        ''' <summary>基准字母，多个用空格分开。</summary>
        Public Property Datum As String = ""

        ''' <summary>子表格数据（当 Type=Table 时）</summary>
        Public Property TableData As OcrTableData = Nothing

        ''' <summary>公式数据（当 Type=Formula 时）LaTeX 格式</summary>
        Public Property FormulaData As OcrFormulaData = Nothing

        ''' <summary>版面区域标签（当 Type=LayoutRegion 时），如 "text", "title", "figure"</summary>
        Public Property LayoutLabel As String = ""

        ''' <summary>版面区域标签 ID</summary>
        Public Property LayoutLabelId As Integer = 0
    End Class

    ''' <summary>
    ''' OCR 边界框
    ''' </summary>
    Public Class OcrBoundingBox
        Public Property Left As Integer
        Public Property Top As Integer
        Public Property Right As Integer
        Public Property Bottom As Integer

        Public Sub New()
        End Sub

        Public Sub New(l As Integer, t As Integer, r As Integer, b As Integer)
            Left = l
            Top = t
            Right = r
            Bottom = b
        End Sub
    End Class

    ''' <summary>
    ''' OCR 表格数据
    ''' </summary>
    Public Class OcrTableData
        ''' <summary>表头行</summary>
        Public Property Header As New List(Of String)()

        ''' <summary>数据行</summary>
        Public Property Rows As New List(Of List(Of String))()

        ''' <summary>跨行合并映射 (row, col) -> rowspan。元组键无法稳定穿过进程边界，独立服务往返时忽略。</summary>
        <Newtonsoft.Json.JsonIgnore>
        Public Property RowSpanMap As New Dictionary(Of Tuple(Of Integer, Integer), Integer)()

        ''' <summary>表格 HTML 表示（来自 PP-StructureV3 的 pred_html）</summary>
        Public Property HtmlContent As String = ""

        ''' <summary>单元格坐标列表（行列索引）</summary>
        Public Property CellCoordinates As New List(Of OcrCellCoordinate)()

        ''' <summary>列数</summary>
        Public Property ColumnCount As Integer

        ''' <summary>行数</summary>
        Public Property RowCount As Integer

        Public Function ToPlainText() As String
            Dim lines As New List(Of String)()
            If Header.Count > 0 Then
                lines.Add(String.Join(vbTab, Header))
            End If
            For Each row In Rows
                lines.Add(String.Join(vbTab, row))
            Next
            Return String.Join(vbCrLf, lines)
        End Function
    End Class

    ''' <summary>
    ''' OCR 单元格坐标
    ''' </summary>
    Public Class OcrCellCoordinate
        Public Property RowIndex As Integer
        Public Property ColIndex As Integer
        Public Property RowSpan As Integer = 1
        Public Property ColSpan As Integer = 1
    End Class

    ''' <summary>
    ''' OCR 公式数据（LaTeX 格式）
    ''' </summary>
    Public Class OcrFormulaData
        ''' <summary>LaTeX 格式的公式</summary>
        Public Property Latex As String = ""

        ''' <summary>置信度</summary>
        Public Property Confidence As Double = 0.0
    End Class

    ''' <summary>
    ''' OCR 识别文档（结构化结果）
    ''' 包含页信息、元素列表、元数据
    ''' </summary>
    Public Class OcrDocument
        ''' <summary>源文件名</summary>
        Public Property SourceFileName As String = ""

        ''' <summary>源文件路径</summary>
        Public Property SourceFilePath As String = ""

        ''' <summary>总页数</summary>
        Public Property TotalPages As Integer = 1

        ''' <summary>当前页码（从 1 开始）</summary>
        Public Property CurrentPage As Integer = 1

        ''' <summary>结构化元素列表</summary>
        Public Property Elements As New List(Of OcrElement)()

        ''' <summary>
        ''' 原始单元格级元素列表（每个 OCR 文本框一个元素，未经按行合并）
        ''' 供表格重建和 raw_ocr_data 使用，避免 BuildDocumentFromPositionedResults 的按行合并丢失坐标精度
        ''' </summary>
        Public Property RawCellElements As New List(Of OcrElement)()

        ''' <summary>原始识别文本（未经结构化处理的完整文本）</summary>
        Public Property RawText As String = ""

        ''' <summary>Markdown 格式的结果（来自 PP-StructureV3 的 markdown 输出）</summary>
        Public Property MarkdownText As String = ""

        ''' <summary>版面区域列表（来自 PP-StructureV3 的 layout_det_res）</summary>
        Public Property LayoutRegions As New List(Of OcrLayoutRegion)()

        ''' <summary>使用的 OCR 引擎名称</summary>
        Public Property EngineName As String = ""

        ''' <summary>识别所用时间（秒）</summary>
        Public Property ElapsedSeconds As Double = 0.0

        ''' <summary>
        ''' 引擎失败原因。空字符串表示这次调用没有引擎错误，结果文本可以为空。
        ''' </summary>
        Public Property ErrorMessage As String = ""

        ''' <summary>
        ''' 引擎原始 JSON（坐标、文字、置信度）。识别在独立进程里完成时，随文档一起带回插件。
        ''' </summary>
        Public Property RawEngineResponse As String = Nothing

        ''' <summary>
        ''' 判断文本是不是引擎失败信息。空白识别结果不是失败。
        ''' </summary>
        Public Shared Function IsEngineErrorText(text As String) As Boolean
            If String.IsNullOrEmpty(text) Then Return False
            Dim trimmed = text.TrimStart()
            Return trimmed.StartsWith("【错误") OrElse
                   trimmed.StartsWith("【OCR识别失败】") OrElse
                   trimmed.StartsWith("【PaddleOCR识别失败】") OrElse
                   trimmed.StartsWith("【识别异常】") OrElse
                   trimmed.StartsWith("【处理失败】") OrElse
                   trimmed.StartsWith("【解析错误】")
        End Function

        ''' <summary>总字符数</summary>
        Public ReadOnly Property TotalChars As Integer
            Get
                Return RawText.Length
            End Get
        End Property

        ''' <summary>元素数量</summary>
        Public ReadOnly Property ElementCount As Integer
            Get
                Return Elements.Count
            End Get
        End Property
    End Class

    ''' <summary>
    ''' OCR 版面区域（来自 PP-StructureV3 的布局分析）
    ''' </summary>
    Public Class OcrLayoutRegion
        ''' <summary>区域标签，如 "text", "title", "figure", "table", "formula"</summary>
        Public Property Label As String = ""

        ''' <summary>置信度</summary>
        Public Property Score As Double = 0.0

        ''' <summary>区域坐标 [x1,y1,x2,y2,...]</summary>
        Public Property Coordinate As New List(Of Integer)()

        ''' <summary>区域内的识别文本</summary>
        Public Property Text As String = ""
    End Class

    ''' <summary>
    ''' OCR 批量处理任务状态
    ''' </summary>
    Public Enum BatchTaskStatus
        Pending
        Processing
        Completed
        Failed
        Cancelled
    End Enum

    ''' <summary>
    ''' OCR 批量处理任务项
    ''' </summary>
    Public Class OcrBatchTask
        ''' <summary>任务唯一 ID</summary>
        Public Property TaskId As String = Guid.NewGuid().ToString("N")

        ''' <summary>文件名</summary>
        Public Property FileName As String = ""

        ''' <summary>文件完整路径</summary>
        Public Property FilePath As String = ""

        ''' <summary>任务状态</summary>
        Public Property Status As BatchTaskStatus = BatchTaskStatus.Pending

        ''' <summary>识别结果文档</summary>
        Public Property Document As OcrDocument = Nothing

        ''' <summary>错误信息</summary>
        Public Property ErrorMessage As String = ""

        ''' <summary>开始时间</summary>
        Public Property StartedAt As DateTime? = Nothing

        ''' <summary>完成时间</summary>
        Public Property CompletedAt As DateTime? = Nothing

        ''' <summary>重试次数</summary>
        Public Property RetryCount As Integer = 0
    End Class

    ''' <summary>
    ''' OCR 批量处理进度报告
    ''' </summary>
    Public Class OcrBatchProgress
        ''' <summary>任务总数</summary>
        Public Property TotalTasks As Integer = 0

        ''' <summary>已完成任务数</summary>
        Public Property CompletedTasks As Integer = 0

        ''' <summary>失败任务数</summary>
        Public Property FailedTasks As Integer = 0

        ''' <summary>执行中的任务数</summary>
        Public Property ProcessingTasks As Integer = 0

        ''' <summary>已完成百分比 0-100</summary>
        Public ReadOnly Property Percent As Integer
            Get
                If TotalTasks = 0 Then Return 100
                Return CInt(Math.Round(CDbl(CompletedTasks + FailedTasks) / CDbl(TotalTasks) * 100.0))
            End Get
        End Property

        ''' <summary>当前正在处理的文件名</summary>
        Public Property CurrentFileName As String = ""

        ''' <summary>状态文本描述</summary>
        Public Property StatusMessage As String = ""
    End Class

    ''' <summary>
    ''' OCR 输出格式
    ''' </summary>
    Public Enum OcrOutputFormat
        PlainText   ' 纯文本
        Markdown    ' Markdown 格式
        Html        ' HTML 格式
        Excel       ' Excel 表格格式
    End Enum

End Namespace
