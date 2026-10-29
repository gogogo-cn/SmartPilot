Imports System.IO
Imports System.Threading.Tasks

Namespace Services.Parsers

    ''' <summary>
    ''' 统一文档解析器接口 - 所有格式 Parser 实现此接口
    ''' 下游的 SemanticChunker / RAG 只依赖此接口，不关心具体文件格式
    ''' </summary>
    Public Interface IDocumentParser

        ''' <summary>此 Parser 支持的文件扩展名（不含点号）</summary>
        ReadOnly Property SupportedExtensions As List(Of String)

        ''' <summary>解析器显示名称</summary>
        ReadOnly Property DisplayName As String

        ''' <summary>提取全文文本 + 基础结构信息</summary>
        Function ExtractFullText(base64Data As String) As Task(Of UnifiedDocumentResult)

        ''' <summary>提取文档结构（目录/标题树/表格位置等）</summary>
        Function ExtractStructure(base64Data As String) As Task(Of DocumentStructure)

        ''' <summary>提取文档元数据</summary>
        Function ExtractMetadata(base64Data As String) As DocumentMetadata

        ''' <summary>提取指定范围内容（用于大文件分页加载）</summary>
        Function ExtractRange(base64Data As String, startPosition As Long, maxLength As Integer) As Task(Of String)

    End Interface

    ' ==================== 数据模型 ====================

    ''' <summary>统一文档输出格式 - 所有 Parser 输出相同结构</summary>
    Public Class UnifiedDocumentResult
        Public Property Success As Boolean = False
        Public Property FileType As String = ""                  ' "pdf"|"docx"|"xlsx"|"doc"|"xls"
        Public Property FullText As String = ""                 ' 全文拼接（纯文本）
        Public Property TotalChars As Long = 0                  ' 总字符数
        Public Property TotalPages As Integer = 0               ' 页数/Sheet数/幻灯片数
        Public Property [Structure] As DocumentStructure           ' 结构化信息
        Public Property Metadata As DocumentMetadata             ' 元数据
        Public Property Pages As List(Of PageContent)            ' 按页/段的内容列表
        Public Property ErrorMessage As String = ""
        Public Property IsScanned As Boolean = False             ' 是否扫描版（仅PDF）
        Public Property ParseTimeMs As Long = 0                 ' 解析耗时
    End Class

    ''' <summary>单页/单段内容</summary>
    Public Class PageContent
        Public Property Index As Integer                         ' 页码/Sheet索引(从1开始)
        Public Property Title As String = ""                     ' 标题（如Sheet名、章节名）
        Public Property Text As String = ""                      ' 本页文本
        Public Property CharCount As Integer = 0
        Public Property StartPosition As Long = 0               ' 在全文中的起始偏移
    End Class

    ''' <summary>文档结构信息</summary>
    Public Class DocumentStructure
        Public Property HasTableOfContents As Boolean = False
        Public Property Headings As List(Of HeadingNode)         ' 标题层级树
        Public Property Tables As List(Of TableInfo)             ' 表格位置与摘要
        Public Property Images As List(Of ImageInfo)             ' 图片位置
        Public Property SectionCount As Integer = 0              ' 章节数量
    End Class

    ''' <summary>标题节点</summary>
    Public Class HeadingNode
        Public Property Level As Integer = 1                     ' H1=1, H2=2, H3=3...
        Public Property Title As String = ""
        Public Property StartPosition As Long = 0                ' 在全文中的字符偏移
        Public Property EndPosition As Long = 0
        Public Property PageIndex As Integer = 0                 ' 所在页码
        Public Property Children As List(Of HeadingNode)
    End Class

    ''' <summary>表格信息</summary>
    Public Class TableInfo
        Public Property StartPosition As Long = 0
        Public Property EndPosition As Long = 0
        Public Property RowCount As Integer = 0
        Public Property ColCount As Integer = 0
        Public Property HeaderRow As List(Of String)
        Public Property Summary As String = ""
    End Class

    ''' <summary>图片信息</summary>
    Public Class ImageInfo
        Public Property StartPosition As Long = 0
        Public Property AltText As String = ""
        Public Property Width As Integer = 0
        Public Property Height As Integer = 0
    End Class

    ''' <summary>文档元数据</summary>
    Public Class DocumentMetadata
        Public Property Title As String = ""
        Public Property Author As String = ""
        Public Property Subject As String = ""
        Public Property CreatedDate As DateTime? = Nothing
        Public Property ModifiedDate As DateTime? = Nothing
        Public Property FileSizeBytes As Long = 0
        Public Property PageCount As Integer = 0
        Public Property WordCount As Integer = 0
        Public Property Properties As New Dictionary(Of String, String)()
    End Class

End Namespace
