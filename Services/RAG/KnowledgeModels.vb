Imports System
Imports System.Collections.Generic

Namespace Services.RAG

    ''' <summary>
    ''' 知识库定义
    ''' </summary>
    Public Class KnowledgeBase
        Public Property Id As String                        ' GUID
        Public Property Name As String                      ' 知识库名称
        Public Property Description As String               ' 描述
        Public Property QdrantCollection As String           ' Qdrant 集合名
        Public Property EmbeddingModel As String             ' 使用的嵌入模型
        Public Property ChunkSize As Integer                 ' 分块大小
        Public Property Overlap As Integer                   ' 重叠字符数
        Public Property DocumentCount As Integer             ' 文档数量
        Public Property ChunkCount As Integer                ' 总块数
        Public Property CreatedAt As DateTime                ' 创建时间
        Public Property UpdatedAt As DateTime                ' 更新时间
        Public Property Status As String                     ' active/archived
    End Class

    ''' <summary>
    ''' 文档记录
    ''' </summary>
    Public Class KnowledgeDocument
        Public Property Id As String                         ' GUID
        Public Property KbId As String                       ' 所属知识库 ID
        Public Property Title As String                      ' 文档标题
        Public Property OssUrl As String                     ' OSS 地址（原始文件）
        Public Property LocalPath As String                  ' 本地文件路径（本地导入时使用）
        Public Property FileSize As Long                     ' 文件大小
        Public Property FileType As String                   ' 文件扩展名
        Public Property ChunkCount As Integer                ' 分块数量
        Public Property ParsedTextLength As Integer          ' 解析后文本长度
        Public Property Status As String                     ' pending/processing/completed/failed
        Public Property ErrorMessage As String               ' 错误信息
        Public Property CreatedAt As DateTime
        Public Property ProcessedAt As DateTime?
    End Class

    ''' <summary>
    ''' 分块记录（带向量 ID）
    ''' </summary>
    Public Class KnowledgeChunk
        Public Property Id As String                         ' 本地 ID
        Public Property QdrantPointId As Guid?               ' Qdrant 点 ID（写入后返回）
        Public Property DocId As String                      ' 所属文档 ID
        Public Property KbId As String                       ' 所属知识库 ID
        Public Property Text As String                       ' 块文本内容
        Public Property TextHash As String                   ' 内容哈希（去重用）
        Public Property ChunkIndex As Integer                ' 块序号
        Public Property SourceSection As String              ' 来源章节
        Public Property StartPosition As Long
        Public Property EndPosition As Long
        Public Property EmbeddingVector As Single()          ' 向量数据（1024 维）
        Public Property Metadata As Dictionary(Of String, String) ' 元数据
    End Class

    ''' <summary>
    ''' RAG 配置
    ''' </summary>
    Public Class RagConfig
        Public Property DefaultEmbeddingProvider As String   ' local_bge / openai / azure / qianwen
        Public Property LocalBgeEndpoint As String           ' Ollama endpoint (默认 http://localhost:11434)
        Public Property LocalBgeModel As String              ' 默认 bge-m3
        Public Property CloudEndpoint As String              ' 云端 API 地址
        Public Property CloudApiKey As String                ' 云端 API Key
        Public Property CloudModel As String                 ' 云端模型名
        Public Property QdrantEndpoint As String             ' Qdrant 地址 (默认 http://localhost:6333)
        Public Property QdrantApiKey As String              ' Qdrant API Key（可选）
        Public Property DefaultTopK As Integer = 5           ' 默认检索数量
        Public Property ScoreThreshold As Single = 0.7F      ' 相似度阈值
        Public Property EnableRagForChat As Boolean = True   ' 是否在对话中启用 RAG
        Public Property MaxContextChunks As Integer = 3      ' 注入上下文的最大块数
        Public Property DefaultKnowledgeBaseId As String = "" ' 对话增强使用的默认知识库 ID（空=不增强）
    End Class

    ''' <summary>
    ''' 知识库持久化文件结构（rag_store.json）
    ''' 仅持久化元数据；向量数据存于 Qdrant，本地文件内容不落盘
    ''' </summary>
    Public Class RagStoreFile
        Public Property KnowledgeBases As List(Of KnowledgeBase)
        Public Property Documents As Dictionary(Of String, List(Of KnowledgeDocument))
    End Class

    ''' <summary>
    ''' 检索结果
    ''' </summary>
    Public Class SearchResult
        Public Property ChunkId As String
        Public Property Text As String
        Public Property Score As Single                       ' 相似度分数 0-1
        Public Property DocTitle As String                    ' 来源文档标题
        Public Property SourceSection As String
        Public Property Metadata As Dictionary(Of String, String)
    End Class

    ''' <summary>
    ''' RAG 结果（用于对话增强）
    ''' </summary>
    Public Class RagResult
        Public Property HasContext As Boolean                  ' 是否找到相关知识
        Public Property ContextText As String                  ' 组装后的上下文文本
        Public Property SourceReferences As List(Of SourceRef) ' 来源引用列表
        Public Property RetrievedChunks As Integer             ' 检索到的块数
        Public Property ShouldFallback As Boolean              ' 是否应该降级到普通 chat
    End Class

    ''' <summary>
    ''' 来源引用
    ''' </summary>
    Public Class SourceRef
        Public Property DocTitle As String
        Public Property Section As String
        Public Property Score As Single
    End Class

    ''' <summary>
    ''' 连接测试结果
    ''' </summary>
    Public Class ConnectionTestResult
        Public Property Provider As String
        Public Property Success As Boolean
        Public Property Message As String
        Public Property LatencyMs As Long
        Public Property Endpoint As String
    End Class

    ''' <summary>
    ''' RAG 测试结果（用于检索测试台）
    ''' </summary>
    Public Class RagTestResult
        Public Property Success As Boolean
        Public Property Question As String
        Public Property Results As List(Of SearchResult)
        Public Property EnhancedAnswer As String
        Public Property TotalRetrieved As Integer
        Public Property ProcessingTimeMs As Long
    End Class

    ''' <summary>
    ''' 搜索选项
    ''' </summary>
    Public Class SearchOptions
        Public Property TopK As Integer = 5
        Public Property ScoreThreshold As Single = 0.0F
        Public Property HybridSearch As Boolean = False       ' 是否启用混合检索
        Public Property FilterByDocIds As List(Of String)     ' 按文档 ID 过滤
    End Class

End Namespace
