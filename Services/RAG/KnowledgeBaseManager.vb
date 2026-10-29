Imports System
Imports System.Collections.Concurrent
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ExcelAIVSTOAddIn.Utils

Namespace Services.RAG

    ''' <summary>
    ''' 知识库统一管理器 - 编排所有 RAG 子模块
    ''' 单例模式，线程安全
    '''
    ''' 【历史坑 · 已修复】仓库自带的 Config/rag_config.json 模板是嵌套 schema
    ''' （localBge.endpoint / qdrant.endpoint / retrieval.* / chatEnhancement.*），
    ''' 而 RagConfig 类是扁平 schema（LocalBgeEndpoint / QdrantEndpoint ...）。
    ''' 直接反序列化后扁平字段全为 Nothing → QdrantClient 构造函数 NRE →
    ''' InitializeClients 把异常吞掉、字段留 Nothing → 之后每次导入文档都报
    ''' "未将对象引用设置到对象的实例"（块数 0）。现在 LoadConfig 会做 schema
    ''' 归一化（嵌套→扁平迁移 + 默认值兜底），EnsureClients 初始化失败时明确报错。
    ''' </summary>
    Public Class KnowledgeBaseManager
        Implements IDisposable

        Private Shared ReadOnly _instance As New Lazy(Of KnowledgeBaseManager)(Function() New KnowledgeBaseManager())

        Private _qdrant As QdrantClient
        Private _embedding As EmbeddingService
        Private _processor As DocumentProcessor
        Private _config As RagConfig
        Private _configPath As String
        Private _storePath As String
        Private _lockObj As New Object()
        Private _storeLock As New Object()
        Private _isDisposed As Boolean = False

        ' 内存缓存（运行时，启动时从 rag_store.json 恢复）
        Private _knowledgeBases As New ConcurrentDictionary(Of String, KnowledgeBase)()
        Private _documents As New ConcurrentDictionary(Of String, ConcurrentDictionary(Of String, KnowledgeDocument))()

        ' 支持的导入格式（与 RagBridge.BrowseLocalFile 的过滤器保持一致）
        Private Shared ReadOnly SupportedExtensions As String() = {".pdf", ".doc", ".docx", ".xls", ".xlsx", ".csv", ".txt", ".md"}

        Public Shared ReadOnly Property Instance As KnowledgeBaseManager
            Get
                Return _instance.Value
            End Get
        End Property

        Private Sub New()
            LoadConfig()
            LoadStore()
            ' 初始化失败不在此处抛出（单例构造炸掉会让 Instance 永远不可用），
            ' 由各公开方法调用 EnsureClients 时明确报错。
            Try
                EnsureClients()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[KBManager] 构造期客户端初始化失败（将在首次调用时重试）: {ex.Message}")
            End Try
        End Sub

#Region "初始化"

        ''' <summary>
        ''' 加载配置文件（兼容扁平/嵌套两种 schema，迁移后写回标准扁平结构）
        ''' </summary>
        Private Sub LoadConfig()
            Try
                _configPath = Path.Combine(Utils.PathHelper.ConfigDirectory, "rag_config.json")
                _storePath = Path.Combine(Utils.PathHelper.ConfigDirectory, "rag_store.json")

                If File.Exists(_configPath) Then
                    Dim json = File.ReadAllText(_configPath)
                    _config = NormalizeConfig(json)
                    ' 迁移/补默认值后立即扁平化写回，后续启动无需再迁移
                    SaveConfigToFile()
                    System.Diagnostics.Debug.WriteLine($"[KBManager] 配置已加载: {_configPath}")
                Else
                    _config = CreateDefaultConfig()
                    SaveConfigToFile()
                    System.Diagnostics.Debug.WriteLine("[KBManager] 使用默认配置（配置文件不存在）")
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[KBManager] 加载配置失败: {ex.Message}，使用默认配置")
                _config = CreateDefaultConfig()
                If String.IsNullOrEmpty(_configPath) Then _configPath = Path.Combine(Utils.PathHelper.ConfigDirectory, "rag_config.json")
                If String.IsNullOrEmpty(_storePath) Then _storePath = Path.Combine(Utils.PathHelper.ConfigDirectory, "rag_store.json")
            End Try
        End Sub

        ''' <summary>
        ''' 创建默认配置
        ''' </summary>
        Private Function CreateDefaultConfig() As RagConfig
            Return New RagConfig With {
                .DefaultEmbeddingProvider = "local_bge",
                .LocalBgeEndpoint = Utils.ConfigConstants.RAG.DEFAULT_LOCAL_BGE_ENDPOINT,
                .LocalBgeModel = "bge-m3",
                .CloudEndpoint = "",
                .CloudApiKey = "",
                .CloudModel = "text-embedding-v3",
                .QdrantEndpoint = Utils.ConfigConstants.RAG.DEFAULT_QDRANT_ENDPOINT,
                .QdrantApiKey = Nothing,
                .DefaultTopK = 5,
                .ScoreThreshold = 0.35F,
                .EnableRagForChat = True,
                .MaxContextChunks = 3,
                .DefaultKnowledgeBaseId = ""
            }
        End Function

        ''' <summary>
        ''' 配置归一化：反序列化 → 嵌套 schema 迁移 → 默认值兜底。
        ''' 关键字段（Qdrant/LocalBge Endpoint 等）为空时一律从嵌套结构读取或回落默认值，
        ''' 绝不让 Nothing 流入客户端构造函数。
        ''' </summary>
        Private Function NormalizeConfig(json As String) As RagConfig
            Dim config As RagConfig = Nothing
            Try
                config = JsonConvert.DeserializeObject(Of RagConfig)(json)
            Catch jex As Exception
                System.Diagnostics.Debug.WriteLine($"[KBManager] 配置反序列化失败: {jex.Message}")
            End Try
            If config Is Nothing Then config = New RagConfig()

            Dim jo As JObject = Nothing
            Try
                jo = JObject.Parse(json)
            Catch
                jo = Nothing
            End Try

            If jo IsNot Nothing Then
                ' --- 嵌套 schema 迁移（仅当扁平字段缺失时补齐，不覆盖已有值）---
                If String.IsNullOrEmpty(config.DefaultEmbeddingProvider) Then
                    config.DefaultEmbeddingProvider = ReadStr(jo, "defaultEmbeddingProvider", "DefaultEmbeddingProvider")
                End If

                Dim localBge = jo("localBge")
                If localBge IsNot Nothing Then
                    If String.IsNullOrEmpty(config.LocalBgeEndpoint) Then config.LocalBgeEndpoint = ReadStr(localBge, "endpoint", "Endpoint")
                    If String.IsNullOrEmpty(config.LocalBgeModel) Then config.LocalBgeModel = ReadStr(localBge, "model", "Model")
                End If

                Dim cloud = jo("cloudEmbedding")
                If cloud IsNot Nothing Then
                    If String.IsNullOrEmpty(config.CloudEndpoint) Then config.CloudEndpoint = ReadStr(cloud, "endpoint", "Endpoint")
                    If String.IsNullOrEmpty(config.CloudApiKey) Then config.CloudApiKey = ReadStr(cloud, "apiKey", "ApiKey")
                    If String.IsNullOrEmpty(config.CloudModel) Then config.CloudModel = ReadStr(cloud, "model", "Model")
                End If

                Dim qdrant = jo("qdrant")
                If qdrant IsNot Nothing Then
                    If String.IsNullOrEmpty(config.QdrantEndpoint) Then config.QdrantEndpoint = ReadStr(qdrant, "endpoint", "Endpoint")
                    If String.IsNullOrEmpty(config.QdrantApiKey) Then config.QdrantApiKey = ReadStr(qdrant, "apiKey", "ApiKey")
                End If

                Dim retrieval = jo("retrieval")
                If retrieval IsNot Nothing Then
                    If config.DefaultTopK <= 0 Then config.DefaultTopK = ReadInt(retrieval, "defaultTopK", "DefaultTopK")
                    If config.ScoreThreshold <= 0F Then config.ScoreThreshold = ReadFloat(retrieval, "scoreThreshold", "ScoreThreshold")
                    If config.MaxContextChunks <= 0 Then config.MaxContextChunks = ReadInt(retrieval, "maxContextChunks", "MaxContextChunks")
                End If

                Dim chatEnh = jo("chatEnhancement")
                If chatEnh IsNot Nothing Then
                    config.EnableRagForChat = ReadBool(chatEnh, "enabled", "Enabled")
                    If String.IsNullOrEmpty(config.DefaultKnowledgeBaseId) Then
                        config.DefaultKnowledgeBaseId = ReadStr(chatEnh, "defaultKbId", "DefaultKbId")
                    End If
                End If

                System.Diagnostics.Debug.WriteLine("[KBManager] 已按嵌套 schema 兼容模式归一化配置")
            End If

            SanitizeConfig(config)
            Return config
        End Function

        ''' <summary>
        ''' 默认值兜底：任何关键字段为空/非正都回落默认值，杜绝 Nothing 流入下游
        ''' </summary>
        Private Sub SanitizeConfig(config As RagConfig)
            If config Is Nothing Then Return
            If String.IsNullOrEmpty(config.DefaultEmbeddingProvider) Then config.DefaultEmbeddingProvider = "local_bge"
            If String.IsNullOrEmpty(config.LocalBgeEndpoint) Then config.LocalBgeEndpoint = Utils.ConfigConstants.RAG.DEFAULT_LOCAL_BGE_ENDPOINT
            If String.IsNullOrEmpty(config.LocalBgeModel) Then config.LocalBgeModel = "bge-m3"
            If String.IsNullOrEmpty(config.QdrantEndpoint) Then config.QdrantEndpoint = Utils.ConfigConstants.RAG.DEFAULT_QDRANT_ENDPOINT
            If String.IsNullOrEmpty(config.CloudModel) Then config.CloudModel = "text-embedding-v3"
            If config.DefaultTopK <= 0 Then config.DefaultTopK = 5
            If config.ScoreThreshold <= 0F Then config.ScoreThreshold = 0.35F
            If config.MaxContextChunks <= 0 Then config.MaxContextChunks = 3
            If config.DefaultKnowledgeBaseId Is Nothing Then config.DefaultKnowledgeBaseId = ""
        End Sub

        ' ---- 嵌套 schema 读取辅助（大小写键都试） ----
        Private Function ReadStr(parent As JToken, ParamArray keys As String()) As String
            If parent Is Nothing Then Return ""
            For Each key In keys
                Dim t = parent(key)
                If t IsNot Nothing AndAlso t.Type <> JTokenType.Null Then
                    Dim s = t.ToString()
                    If Not String.IsNullOrEmpty(s) Then Return s
                End If
            Next
            Return ""
        End Function

        Private Function ReadInt(parent As JToken, ParamArray keys As String()) As Integer
            If parent Is Nothing Then Return 0
            For Each key In keys
                Dim t = parent(key)
                If t IsNot Nothing AndAlso t.Type <> JTokenType.Null Then
                    Dim v As Integer = 0
                    If Integer.TryParse(t.ToString(), v) AndAlso v > 0 Then Return v
                End If
            Next
            Return 0
        End Function

        Private Function ReadFloat(parent As JToken, ParamArray keys As String()) As Single
            If parent Is Nothing Then Return 0F
            For Each key In keys
                Dim t = parent(key)
                If t IsNot Nothing AndAlso t.Type <> JTokenType.Null Then
                    Dim v As Single = 0F
                    If Single.TryParse(t.ToString(), v) AndAlso v > 0F Then Return v
                End If
            Next
            Return 0F
        End Function

        Private Function ReadBool(parent As JToken, ParamArray keys As String()) As Boolean
            If parent Is Nothing Then Return True
            For Each key In keys
                Dim t = parent(key)
                If t IsNot Nothing AndAlso t.Type <> JTokenType.Null Then
                    Dim v As Boolean = True
                    If Boolean.TryParse(t.ToString(), v) Then Return v
                End If
            Next
            Return True
        End Function

        ''' <summary>
        ''' 确保三个客户端已初始化（幂等，失败时重试）。
        ''' 【历史坑】原 InitializeClients 吞异常后字段留 Nothing，后续所有调用变成
        ''' "未将对象引用设置到对象的实例"。现在初始化失败会记录原因并抛出带中文说明的异常。
        ''' </summary>
        Private Sub EnsureClients()
            If _isDisposed Then Throw New ObjectDisposedException("KnowledgeBaseManager")
            If _qdrant IsNot Nothing AndAlso _embedding IsNot Nothing AndAlso _processor IsNot Nothing Then Return

            SyncLock _lockObj
                If _qdrant IsNot Nothing AndAlso _embedding IsNot Nothing AndAlso _processor IsNot Nothing Then Return
                Try
                    If _qdrant Is Nothing Then _qdrant = New QdrantClient(_config.QdrantEndpoint, _config.QdrantApiKey)
                    If _embedding Is Nothing Then _embedding = New EmbeddingService(_config)
                    If _processor Is Nothing Then _processor = New DocumentProcessor()
                    System.Diagnostics.Debug.WriteLine("[KBManager] 客户端初始化完成")
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"[KBManager] 客户端初始化失败: {ex.Message}")
                    Throw New InvalidOperationException("RAG 组件初始化失败（请检查 RAG 配置中的 Qdrant / Embedding 地址）: " & ex.Message, ex)
                End Try
            End SyncLock
        End Sub

#End Region

#Region "持久化（rag_store.json）"

        ''' <summary>
        ''' 从磁盘恢复知识库与文档元数据（向量在 Qdrant，此处只存元数据）
        ''' </summary>
        Private Sub LoadStore()
            Try
                If String.IsNullOrEmpty(_storePath) OrElse Not File.Exists(_storePath) Then Return
                Dim json = File.ReadAllText(_storePath)
                If String.IsNullOrWhiteSpace(json) Then Return

                Dim store = JsonConvert.DeserializeObject(Of RagStoreFile)(json)
                If store Is Nothing Then Return

                If store.KnowledgeBases IsNot Nothing Then
                    For Each kb In store.KnowledgeBases
                        If kb Is Nothing OrElse String.IsNullOrEmpty(kb.Id) Then Continue For
                        _knowledgeBases(kb.Id) = kb
                        _documents(kb.Id) = New ConcurrentDictionary(Of String, KnowledgeDocument)()
                    Next
                End If

                If store.Documents IsNot Nothing Then
                    For Each kvp In store.Documents
                        If String.IsNullOrEmpty(kvp.Key) OrElse kvp.Value Is Nothing Then Continue For
                        If Not _documents.ContainsKey(kvp.Key) Then
                            _documents(kvp.Key) = New ConcurrentDictionary(Of String, KnowledgeDocument)()
                        End If
                        For Each doc In kvp.Value
                            If doc Is Nothing OrElse String.IsNullOrEmpty(doc.Id) Then Continue For
                            ' 上次进程被杀死时可能停留在 processing，重置为 pending 以便重试
                            If doc.Status = "processing" Then doc.Status = "pending"
                            _documents(kvp.Key)(doc.Id) = doc
                        Next
                    Next
                End If

                ' 校正计数（防止历史数据不一致）
                For Each kb In _knowledgeBases.Values
                    Dim docs As ConcurrentDictionary(Of String, KnowledgeDocument) = Nothing
                    If _documents.TryGetValue(kb.Id, docs) Then
                        kb.DocumentCount = docs.Count
                    Else
                        kb.DocumentCount = 0
                    End If
                    If kb.DocumentCount = 0 Then kb.ChunkCount = 0
                Next

                System.Diagnostics.Debug.WriteLine($"[KBManager] 已从磁盘恢复 {_knowledgeBases.Count} 个知识库")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[KBManager] 加载知识库存储失败: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' 将知识库与文档元数据写入磁盘（原子写入：先 .tmp 再替换）
        ''' </summary>
        Private Sub SaveStore()
            Try
                If String.IsNullOrEmpty(_storePath) Then Return
                SyncLock _storeLock
                    Dim store As New RagStoreFile With {
                        .KnowledgeBases = _knowledgeBases.Values.ToList(),
                        .Documents = New Dictionary(Of String, List(Of KnowledgeDocument))()
                    }
                    For Each kvp In _documents
                        store.Documents(kvp.Key) = kvp.Value.Values.ToList()
                    Next

                    Dim dir = Path.GetDirectoryName(_storePath)
                    If Not String.IsNullOrEmpty(dir) AndAlso Not Directory.Exists(dir) Then
                        Directory.CreateDirectory(dir)
                    End If

                    Dim tmpPath = _storePath & ".tmp"
                    File.WriteAllText(tmpPath, JsonConvert.SerializeObject(store, Formatting.Indented))
                    File.Copy(tmpPath, _storePath, True)
                    File.Delete(tmpPath)
                End SyncLock
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[KBManager] 保存知识库存储失败: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' 线程安全地调整知识库计数（永不为负）
        ''' </summary>
        Private Sub AdjustKbCounts(kb As KnowledgeBase, docDelta As Integer, chunkDelta As Integer)
            If kb Is Nothing Then Return
            SyncLock kb
                kb.DocumentCount = Math.Max(0, kb.DocumentCount + docDelta)
                kb.ChunkCount = Math.Max(0, kb.ChunkCount + chunkDelta)
                kb.UpdatedAt = DateTime.Now
            End SyncLock
        End Sub

#End Region

#Region "配置管理"

        ''' <summary>
        ''' 获取当前配置（返回副本）
        ''' </summary>
        Public Function GetConfig() As RagConfig
            SyncLock _lockObj
                Return JsonConvert.DeserializeObject(Of RagConfig)(
                    JsonConvert.SerializeObject(_config))
            End SyncLock
        End Function

        ''' <summary>
        ''' 测试 Qdrant 连接性
        ''' </summary>
        Public Async Function TestQdrantConnectionAsync() As Task(Of ConnectionTestResult)
            EnsureClients()
            Try
                Dim sw = Diagnostics.Stopwatch.StartNew()
                Dim isHealthy = Await _qdrant.HealthCheckAsync()
                sw.Stop()

                If isHealthy Then
                    Return New ConnectionTestResult With {
                        .Provider = "Qdrant",
                        .Success = True,
                        .Message = $"连接成功 ({_config.QdrantEndpoint})",
                        .LatencyMs = sw.ElapsedMilliseconds,
                        .Endpoint = _config.QdrantEndpoint
                    }
                Else
                    Return New ConnectionTestResult With {
                        .Provider = "Qdrant",
                        .Success = False,
                        .Message = "服务不可达或响应异常",
                        .LatencyMs = sw.ElapsedMilliseconds,
                        .Endpoint = _config.QdrantEndpoint
                    }
                End If
            Catch ex As Exception
                Return New ConnectionTestResult With {
                    .Provider = "Qdrant",
                    .Success = False,
                    .Message = ex.Message,
                    .LatencyMs = -1,
                    .Endpoint = _config.QdrantEndpoint
                }
            End Try
        End Function

        ''' <summary>
        ''' 保存配置到文件和内存，并重建客户端（Qdrant endpoint / Embedding provider 可能已变）
        ''' </summary>
        Public Async Function SaveConfigAsync(newConfig As RagConfig) As Task(Of Boolean)
            ' ConfigureAwait(False)：宿主在 UI 线程上用 .Result 等待本方法。
            ' 若延续回到 UI 同步上下文，会和 .Result 互相等待，按钮会停在按下状态。
            Return Await Task.Run(Function()
                Try
                    If newConfig Is Nothing Then Return False
                    SanitizeConfig(newConfig)

                    ' API Key 掩码保护：回传 "***" 或空值表示"未修改"，保留原值
                    SyncLock _lockObj
                        If String.IsNullOrEmpty(newConfig.CloudApiKey) OrElse newConfig.CloudApiKey = "***" Then
                            newConfig.CloudApiKey = _config.CloudApiKey
                        End If
                        If String.IsNullOrEmpty(newConfig.QdrantApiKey) OrElse newConfig.QdrantApiKey = "***" Then
                            newConfig.QdrantApiKey = _config.QdrantApiKey
                        End If
                        _config = newConfig
                    End SyncLock

                    ' 保存到文件
                    SaveConfigToFile()

                    ' 重建客户端，使 endpoint / provider 变更立即生效
                    _qdrant = Nothing
                    _embedding = Nothing
                    _processor = Nothing
                    EnsureClients()

                    System.Diagnostics.Debug.WriteLine("[KBManager] 配置已保存并重新初始化服务")
                    Return True
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"[KBManager] 保存配置失败: {ex.Message}")
                    Return False
                End Try
            End Function).ConfigureAwait(False)
        End Function

        ''' <summary>
        ''' 将配置写入文件
        ''' </summary>
        Private Sub SaveConfigToFile()
            Try
                If String.IsNullOrEmpty(_configPath) Then Return
                Dim dir = Path.GetDirectoryName(_configPath)
                If Not String.IsNullOrEmpty(dir) AndAlso Not Directory.Exists(dir) Then
                    Directory.CreateDirectory(dir)
                End If

                Dim json = JsonConvert.SerializeObject(_config, Formatting.Indented)
                File.WriteAllText(_configPath, json)

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[KBManager] 写入配置文件失败: {ex.Message}")
            End Try
        End Sub

#End Region

#Region "知识库 CRUD"

        ''' <summary>
        ''' 创建知识库（同时在 Qdrant 创建集合）
        ''' </summary>
        Public Async Function CreateKnowledgeBaseAsync(name As String,
                                                        description As String,
                                                        Optional options As Dictionary(Of String, Object) = Nothing) As Task(Of KnowledgeBase)
            EnsureClients()
            Try
                Dim kbId = Guid.NewGuid().ToString("N")
                Dim collectionName = $"kb_{kbId}"

                ' 尝试在 Qdrant 创建集合（不阻塞知识库创建）
                Dim vectorSize As Integer = 1024  ' BGE 默认维度
                If options IsNot Nothing AndAlso options.ContainsKey("vectorSize") Then
                    vectorSize = Convert.ToInt32(options("vectorSize"))
                End If

                Dim qdrantReady As Boolean = True
                ' 先探测 Qdrant 可用性（3 秒快速失败）：不可达时直接置 pending，
                ' 不再做「exists 检查 + 创建」两次网络尝试（系统代理环境下每次尝试都是秒级）
                Try
                    qdrantReady = Await _qdrant.HealthCheckAsync()
                Catch hex As Exception
                    qdrantReady = False
                    System.Diagnostics.Debug.WriteLine($"[KBManager] Qdrant 健康检查异常: {hex.Message}")
                End Try

                If qdrantReady Then
                    Try
                        Dim created = Await _qdrant.CreateCollectionIfNotExists(collectionName, vectorSize)
                        If Not created Then
                            qdrantReady = False
                            System.Diagnostics.Debug.WriteLine($"[KBManager] Qdrant 集合已存在或创建失败: {collectionName}")
                        End If
                    Catch qex As Exception
                        qdrantReady = False
                        System.Diagnostics.Debug.WriteLine($"[KBManager] Qdrant 不可用（非致命）: {qex.Message}")
                    End Try
                Else
                    System.Diagnostics.Debug.WriteLine($"[KBManager] Qdrant 当前不可达，知识库先置 pending: {collectionName}")
                End If

                ' 构建知识库对象
                Dim chunkSize As Integer = 1500
                Dim overlap As Integer = 200
                If options IsNot Nothing Then
                    If options.ContainsKey("chunkSize") Then chunkSize = Convert.ToInt32(options("chunkSize"))
                    If options.ContainsKey("overlap") Then overlap = Convert.ToInt32(options("overlap"))
                End If

                Dim kb As New KnowledgeBase With {
                    .Id = kbId,
                    .Name = name,
                    .Description = description,
                    .QdrantCollection = collectionName,
                    .EmbeddingModel = _config.LocalBgeModel,
                    .ChunkSize = chunkSize,
                    .Overlap = overlap,
                    .DocumentCount = 0,
                    .ChunkCount = 0,
                    .CreatedAt = DateTime.Now,
                    .UpdatedAt = DateTime.Now,
                    .Status = If(qdrantReady, "active", "pending")
                }

                ' 添加到内存缓存
                _knowledgeBases(kbId) = kb
                _documents(kbId) = New ConcurrentDictionary(Of String, KnowledgeDocument)()
                SaveStore()

                System.Diagnostics.Debug.WriteLine($"[KBManager] 知识库创建成功: {name} (ID: {kbId})")

                Return kb
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[KBManager] 创建知识库失败: {ex.Message}")
                Throw
            End Try
        End Function

        ''' <summary>
        ''' 获取所有知识库列表
        ''' </summary>
        Public Function GetAllKnowledgeBases() As List(Of KnowledgeBase)
            Return _knowledgeBases.Values.ToList()
        End Function

        ''' <summary>
        ''' 获取单个知识库详情
        ''' </summary>
        Public Function GetKnowledgeBase(kbId As String) As KnowledgeBase
            If String.IsNullOrEmpty(kbId) Then Return Nothing
            If _knowledgeBases.ContainsKey(kbId) Then
                Return _knowledgeBases(kbId)
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' 删除知识库（级联删除 Qdrant 集合和所有文档记录）
        ''' </summary>
        Public Async Function DeleteKnowledgeBaseAsync(kbId As String) As Task(Of Boolean)
            Try
                Dim kb = GetKnowledgeBase(kbId)
                If kb Is Nothing Then
                    Throw New InvalidOperationException($"知识库不存在 (ID: {kbId})")
                End If

                ' 从 Qdrant 删除集合（尽力而为，失败不阻塞本地删除）
                If _qdrant IsNot Nothing Then
                    ' 先探测（TCP 500ms）：Qdrant 未运行时跳过 HTTP 尝试，删除秒完成
                    Dim qdrantUp As Boolean
                    Try
                        qdrantUp = Await _qdrant.HealthCheckAsync()
                    Catch hex As Exception
                        qdrantUp = False
                        System.Diagnostics.Debug.WriteLine($"[KBManager] 删除前健康检查异常: {hex.Message}")
                    End Try

                    If qdrantUp Then
                        Try
                            Dim qdrantDeleted = Await _qdrant.DeleteCollectionAsync(kb.QdrantCollection)
                            If Not qdrantDeleted Then
                                System.Diagnostics.Debug.WriteLine($"[Qdrant] 集合删除未成功（可能不存在或不可达）: {kb.QdrantCollection}")
                            End If
                        Catch qex As Exception
                            System.Diagnostics.Debug.WriteLine($"[Qdrant] 集合删除异常（已跳过）: {kb.QdrantCollection} - {qex.Message}")
                        End Try
                    Else
                        System.Diagnostics.Debug.WriteLine($"[KBManager] Qdrant 不可达，跳过集合删除（仅删本地记录）: {kb.QdrantCollection}")
                    End If
                End If

                ' 从内存移除
                _knowledgeBases.TryRemove(kbId, Nothing)
                _documents.TryRemove(kbId, Nothing)
                SaveStore()

                System.Diagnostics.Debug.WriteLine($"[KBManager] 知识库已删除: {kb.Name}")

                Return True
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[KBManager] 删除知识库失败: {ex.Message}")
                Throw
            End Try
        End Function

        ''' <summary>
        ''' 将"待同步"（创建时 Qdrant 不可达）的知识库重新尝试创建集合并置为活跃。
        ''' 返回成功同步的数量。
        ''' </summary>
        Public Async Function SyncPendingKnowledgeBasesAsync() As Task(Of Integer)
            EnsureClients()
            Dim synced As Integer = 0

            ' 先探测 Qdrant 可用性（3 秒快速失败）：不可达时直接返回，
            ' 不再对每个 pending 知识库逐一尝试（N 个 pending = N 次秒级网络等待）
            Dim qdrantUp As Boolean
            Try
                qdrantUp = Await _qdrant.HealthCheckAsync()
            Catch hex As Exception
                qdrantUp = False
                System.Diagnostics.Debug.WriteLine($"[KBManager] 同步前健康检查异常: {hex.Message}")
            End Try
            If Not qdrantUp Then
                System.Diagnostics.Debug.WriteLine("[KBManager] Qdrant 不可达，跳过待同步知识库")
                Return 0
            End If

            For Each kb In _knowledgeBases.Values.ToList()
                If kb.Status <> "pending" Then Continue For
                Try
                    Dim created = Await _qdrant.CreateCollectionIfNotExists(kb.QdrantCollection, 1024)
                    If created Then
                        kb.Status = "active"
                        kb.UpdatedAt = DateTime.Now
                        synced += 1
                        System.Diagnostics.Debug.WriteLine($"[KBManager] 知识库已同步为活跃: {kb.Name}")
                    End If
                Catch qex As Exception
                    System.Diagnostics.Debug.WriteLine($"[KBManager] 同步知识库失败（非致命）: {kb.Name} - {qex.Message}")
                End Try
            Next
            If synced > 0 Then SaveStore()
            Return synced
        End Function

#End Region

#Region "文档管理"

        ''' <summary>
        ''' 导入文档（核心流程，OSS URL）
        ''' 1. 创建文档记录（status=pending）
        ''' 2. 异步处理: 下载→解析→分块→向量化→写入 Qdrant
        ''' 3. 更新文档状态为 completed/failed
        ''' </summary>
        Public Async Function ImportDocumentAsync(kbId As String,
                                                   ossUrl As String,
                                                   Optional title As String = "",
                                                   Optional metadata As Dictionary(Of String, String) = Nothing) As Task(Of KnowledgeDocument)
            EnsureClients()
            Return Await Task.Run(Function()
                Try
                    ' 验证知识库存在
                    Dim kb = GetKnowledgeBase(kbId)
                    If kb Is Nothing Then
                        Throw New Exception($"知识库不存在: {kbId}")
                    End If

                    ' 校验 URL 中的扩展名
                    Dim urlExt = GetExtensionFromUrl(ossUrl)
                    If Not SupportedExtensions.Contains(urlExt) Then
                        Dim extDisplay = If(String.IsNullOrEmpty(urlExt), "未知", urlExt)
                        Throw New Exception($"不支持的文件格式: {extDisplay}（支持: PDF/Word/Excel/CSV/TXT/MD）")
                    End If

                    ' 如果未提供标题，从 URL 提取
                    If String.IsNullOrWhiteSpace(title) Then
                        Try
                            Dim uri = New Uri(ossUrl)
                            title = Path.GetFileNameWithoutExtension(uri.LocalPath)
                            If String.IsNullOrWhiteSpace(title) Then
                                title = $"文档_{DateTime.Now:yyyyMMdd_HHmmss}"
                            End If
                        Catch ex As Exception
                            title = $"文档_{DateTime.Now:yyyyMMdd_HHmmss}"
                        End Try
                    End If

                    ' 创建文档记录
                    Dim docId = Guid.NewGuid().ToString("N")
                    Dim doc As New KnowledgeDocument With {
                        .Id = docId,
                        .KbId = kbId,
                        .Title = title,
                        .OssUrl = ossUrl,
                        .FileSize = 0,
                        .FileType = urlExt,
                        .ChunkCount = 0,
                        .ParsedTextLength = 0,
                        .Status = "pending",
                        .ErrorMessage = "",
                        .CreatedAt = DateTime.Now,
                        .ProcessedAt = Nothing
                    }

                    ' 添加到内存
                    If Not _documents.ContainsKey(kbId) Then
                        _documents(kbId) = New ConcurrentDictionary(Of String, KnowledgeDocument)()
                    End If
                    _documents(kbId)(docId) = doc

                    ' 更新知识库统计
                    AdjustKbCounts(kb, 1, 0)
                    SaveStore()

                    System.Diagnostics.Debug.WriteLine($"[KBManager] 文档导入请求已接收: {title} (ID: {docId})")

                    ' 异步处理文档（不阻塞调用方，fire-and-forget模式）
                    Task.Run(Async Function()
                        Try
                            Await ProcessDocumentInBackgroundAsync(doc, kb)
                        Catch ex As Exception
                            System.Diagnostics.Debug.WriteLine($"[KBManager] 后台处理异常: {ex.Message}")
                            doc.Status = "failed"
                            doc.ErrorMessage = ex.Message
                            doc.ProcessedAt = DateTime.Now
                            SaveStore()
                        End Try
                    End Function).ConfigureAwait(False)

                    Return doc

                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"[KBManager] 导入文档失败: {ex.Message}")
                    Throw
                End Try
            End Function)
        End Function

        ''' <summary>
        ''' 导入本地文件到知识库（核心流程）
        ''' 1. 创建文档记录（status=pending，保存本地路径）
        ''' 2. 异步处理: 解析→分块→向量化→写入 Qdrant
        ''' 3. 更新文档状态为 completed/failed
        ''' </summary>
        Public Async Function ImportLocalDocumentAsync(kbId As String,
                                                       localPath As String,
                                                       Optional title As String = "",
                                                       Optional metadata As Dictionary(Of String, String) = Nothing) As Task(Of KnowledgeDocument)
            EnsureClients()
            Return Await Task.Run(Function()
                Try
                    ' 验证知识库存在
                    Dim kb = GetKnowledgeBase(kbId)
                    If kb Is Nothing Then
                        Throw New Exception($"知识库不存在: {kbId}")
                    End If

                    ' 验证本地文件存在
                    If Not File.Exists(localPath) Then
                        Throw New Exception($"本地文件不存在: {localPath}")
                    End If

                    ' 校验扩展名
                    Dim ext = Path.GetExtension(localPath).ToLower()
                    If Not SupportedExtensions.Contains(ext) Then
                        Dim extDisplay2 = If(String.IsNullOrEmpty(ext), "未知", ext)
                        Throw New Exception($"不支持的文件格式: {extDisplay2}（支持: PDF/Word/Excel/CSV/TXT/MD）")
                    End If

                    ' 如果未提供标题，从路径提取
                    If String.IsNullOrWhiteSpace(title) Then
                        title = Path.GetFileNameWithoutExtension(localPath)
                        If String.IsNullOrWhiteSpace(title) Then
                            title = $"文档_{DateTime.Now:yyyyMMdd_HHmmss}"
                        End If
                    End If

                    ' 创建文档记录
                    Dim docId = Guid.NewGuid().ToString("N")
                    Dim fileInfo As New FileInfo(localPath)
                    Dim doc As New KnowledgeDocument With {
                        .Id = docId,
                        .KbId = kbId,
                        .Title = title,
                        .OssUrl = "",
                        .LocalPath = localPath,
                        .FileSize = fileInfo.Length,
                        .FileType = ext,
                        .ChunkCount = 0,
                        .ParsedTextLength = 0,
                        .Status = "pending",
                        .ErrorMessage = "",
                        .CreatedAt = DateTime.Now,
                        .ProcessedAt = Nothing
                    }

                    ' 添加到内存
                    If Not _documents.ContainsKey(kbId) Then
                        _documents(kbId) = New ConcurrentDictionary(Of String, KnowledgeDocument)()
                    End If
                    _documents(kbId)(docId) = doc

                    ' 更新知识库统计
                    AdjustKbCounts(kb, 1, 0)
                    SaveStore()

                    System.Diagnostics.Debug.WriteLine($"[KBManager] 本地文档导入请求已接收: {title} (路径: {localPath}, ID: {docId})")

                    ' 异步处理文档
                    Task.Run(Async Function()
                        Try
                            Await ProcessDocumentInBackgroundAsync(doc, kb)
                        Catch ex As Exception
                            System.Diagnostics.Debug.WriteLine($"[KBManager] 本地文档后台处理异常: {ex.Message}")
                            doc.Status = "failed"
                            doc.ErrorMessage = ex.Message
                            doc.ProcessedAt = DateTime.Now
                            SaveStore()
                        End Try
                    End Function).ConfigureAwait(False)

                    Return doc

                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"[KBManager] 导入本地文档失败: {ex.Message}")
                    Throw
                End Try
            End Function)
        End Function

        ''' <summary>
        ''' 重试失败的文档（重置状态并重新走后台处理，不产生新文档记录）
        ''' </summary>
        Public Async Function RetryDocumentAsync(docId As String) As Task(Of KnowledgeDocument)
            EnsureClients()
            Return Await Task.Run(Function()
                Dim doc As KnowledgeDocument = Nothing
                Dim targetKb As KnowledgeBase = Nothing

                For Each kvp In _documents
                    If kvp.Value.ContainsKey(docId) Then
                        doc = kvp.Value(docId)
                        targetKb = GetKnowledgeBase(kvp.Key)
                        Exit For
                    End If
                Next

                If doc Is Nothing Then
                    Throw New Exception($"文档不存在: {docId}")
                End If
                If targetKb Is Nothing Then
                    Throw New Exception("文档所属的知识库不存在，无法重试")
                End If

                ' 本地文档：确认原文件还在；OSS 文档：确认有 URL
                If Not String.IsNullOrWhiteSpace(doc.LocalPath) AndAlso Not File.Exists(doc.LocalPath) Then
                    Throw New Exception($"原始文件已不存在，无法重试: {doc.LocalPath}")
                End If
                If String.IsNullOrWhiteSpace(doc.LocalPath) AndAlso String.IsNullOrWhiteSpace(doc.OssUrl) Then
                    Throw New Exception("文档没有可用的来源（本地路径与 OSS URL 均缺失）")
                End If

                doc.Status = "pending"
                doc.ErrorMessage = ""
                doc.ProcessedAt = Nothing
                SaveStore()

                System.Diagnostics.Debug.WriteLine($"[KBManager] 重试文档: {doc.Title} (ID: {docId})")

                Task.Run(Async Function()
                    Try
                        Await ProcessDocumentInBackgroundAsync(doc, targetKb)
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine($"[KBManager] 重试后台处理异常: {ex.Message}")
                        doc.Status = "failed"
                        doc.ErrorMessage = ex.Message
                        doc.ProcessedAt = DateTime.Now
                        SaveStore()
                    End Try
                End Function).ConfigureAwait(False)

                Return doc
            End Function)
        End Function

        ''' <summary>
        ''' 获取知识库下的文档列表
        ''' </summary>
        Public Function GetDocuments(kbId As String) As List(Of KnowledgeDocument)
            If String.IsNullOrEmpty(kbId) Then Return New List(Of KnowledgeDocument)()
            If _documents.ContainsKey(kbId) Then
                Return _documents(kbId).Values.ToList()
            End If
            Return New List(Of KnowledgeDocument)()
        End Function

        ''' <summary>
        ''' 删除文档（级联删除 Qdrant 中的向量点）
        ''' </summary>
        Public Async Function DeleteDocumentAsync(docId As String) As Task(Of Boolean)
            Try
                ' 找到文档所属的知识库
                Dim targetKbId As String = ""
                Dim targetDoc As KnowledgeDocument = Nothing

                For Each kvp In _documents
                    If kvp.Value.ContainsKey(docId) Then
                        targetKbId = kvp.Key
                        targetDoc = kvp.Value(docId)
                        Exit For
                    End If
                Next

                If targetDoc Is Nothing Then
                    Return False
                End If

                Dim kb = GetKnowledgeBase(targetKbId)
                If kb IsNot Nothing AndAlso _qdrant IsNot Nothing Then
                    ' 从 Qdrant 删除该文档的所有向量点（通过 doc_id 过滤）
                    ' 先探测（TCP 500ms）：Qdrant 未运行时跳过 HTTP 尝试，删除秒完成
                    Dim qdrantUp As Boolean
                    Try
                        qdrantUp = Await _qdrant.HealthCheckAsync()
                    Catch hex As Exception
                        qdrantUp = False
                        System.Diagnostics.Debug.WriteLine($"[KBManager] 删文档前健康检查异常: {hex.Message}")
                    End Try

                    If qdrantUp Then
                        Try
                            Await _qdrant.DeleteByFilterAsync(kb.QdrantCollection, targetDoc.Id)
                        Catch ex As Exception
                            System.Diagnostics.Debug.WriteLine($"[KBManager] Qdrant 删除文档点失败（非致命）: {ex.Message}")
                        End Try
                    Else
                        System.Diagnostics.Debug.WriteLine($"[KBManager] Qdrant 不可达，跳过文档向量点删除（仅删本地记录）: {targetDoc.Title}")
                    End If
                End If

                ' 从内存移除
                Dim docDict As ConcurrentDictionary(Of String, KnowledgeDocument) = Nothing
                If _documents.TryGetValue(targetKbId, docDict) Then
                    docDict.TryRemove(docId, Nothing)
                End If

                ' 更新知识库统计
                AdjustKbCounts(kb, -1, -targetDoc.ChunkCount)
                SaveStore()

                System.Diagnostics.Debug.WriteLine($"[KBManager] 文档已删除: {targetDoc.Title}")

                Return True
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[KBManager] 删除文档失败: {ex.Message}")
                Throw
            End Try
        End Function

        ''' <summary>
        ''' 从 URL 中提取扩展名（去掉 query/fragment）
        ''' </summary>
        Private Shared Function GetExtensionFromUrl(url As String) As String
            If String.IsNullOrWhiteSpace(url) Then Return ""
            Try
                Dim pathPart = url.Split({"?"c, "#"c}, StringSplitOptions.None)(0)
                Return Path.GetExtension(pathPart).ToLower()
            Catch
                Return ""
            End Try
        End Function

#End Region

#Region "检索"

        ''' <summary>
        ''' 智能检索 - 向量搜索
        ''' 未显式传入阈值时使用配置的 ScoreThreshold（历史缺陷：此前恒为 0，用户设置的阈值不生效）
        ''' </summary>
        Public Async Function SearchAsync(kbId As String,
                                           query As String,
                                           Optional topK As Integer = 5,
                                           Optional options As SearchOptions = Nothing) As Task(Of List(Of SearchResult))
            Dim results As New List(Of SearchResult)()

            Try
                Dim kb = GetKnowledgeBase(kbId)
                If kb Is Nothing Then
                    System.Diagnostics.Debug.WriteLine($"[KBManager] SearchAsync: 知识库不存在 {kbId}")
                    Return results
                End If

                ' 使用配置的 topK 或传入值
                Dim actualTopK = If(options IsNot Nothing, options.TopK, topK)
                If actualTopK <= 0 Then actualTopK = _config.DefaultTopK

                ' 向量化查询文本
                Dim queryVector = Await _embedding.EmbedAsync(query)
                If queryVector Is Nothing OrElse queryVector.Length = 0 Then
                    System.Diagnostics.Debug.WriteLine("[KBManager] SearchAsync: 查询向量化失败")
                    Return results
                End If

                ' 设置分数阈值：options 显式指定优先，否则用配置值
                Dim threshold As Single = _config.ScoreThreshold
                If options IsNot Nothing AndAlso options.ScoreThreshold > 0F Then
                    threshold = options.ScoreThreshold
                End If

                ' 在 Qdrant 中执行向量搜索
                results = Await _qdrant.SearchAsync(
                    kb.QdrantCollection,
                    queryVector,
                    actualTopK,
                    threshold
                )

                System.Diagnostics.Debug.WriteLine($"[KBManager] 搜索完成: 返回 {results.Count} 个结果")

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[KBManager] SearchAsync 异常: {ex.Message}")
            End Try

            Return results
        End Function

        ''' <summary>
        ''' 测试 RAG 查询（用于检索测试台）
        ''' 返回检索结果 + 组装的参考答案上下文
        ''' </summary>
        Public Async Function TestQueryAsync(kbId As String,
                                              question As String,
                                              Optional topK As Integer = 5) As Task(Of RagTestResult)
            Dim result As New RagTestResult With {
                .Success = False,
                .Question = question,
                .Results = New List(Of SearchResult)(),
                .EnhancedAnswer = "",
                .TotalRetrieved = 0,
                .ProcessingTimeMs = 0
            }

            Try
                Dim sw = Diagnostics.Stopwatch.StartNew()

                ' 执行检索
                result.Results = Await SearchAsync(kbId, question, topK)
                result.TotalRetrieved = result.Results.Count

                sw.Stop()
                result.ProcessingTimeMs = sw.ElapsedMilliseconds
                result.Success = True

                ' 组装参考答案上下文（不调用 LLM，只展示检索到的内容）
                If result.Results.Count > 0 Then
                    Dim contextBuilder As New Text.StringBuilder()
                    contextBuilder.AppendLine("【检索到的参考资料】")
                    contextBuilder.AppendLine()

                    ' 修复 off-by-one：原写法 Math.Min(Count-1, Max)-1 会少展示一个块
                    Dim takeCount As Integer = Math.Min(result.Results.Count, Math.Max(1, _config.MaxContextChunks))
                    For i As Integer = 0 To takeCount - 1
                        Dim chunk = result.Results(i)
                        contextBuilder.AppendLine($"--- 资料 {i + 1} (相关度: {chunk.Score:P0}, 来源: {chunk.DocTitle}) ---")
                        contextBuilder.AppendLine(chunk.Text)
                        contextBuilder.AppendLine()
                    Next

                    result.EnhancedAnswer = contextBuilder.ToString()
                Else
                    result.EnhancedAnswer = "未检索到相关内容"
                End If

            Catch ex As Exception
                result.EnhancedAnswer = $"测试异常: {ex.Message}"
                System.Diagnostics.Debug.WriteLine($"[KBManager] TestQueryAsync 异常: {ex.Message}")
            End Try

            Return result
        End Function

#End Region

#Region "内部方法"

        ''' <summary>
        ''' 异步处理文档（下载→解析→分块→向量化→入库）
        ''' 在后台线程执行，不阻塞调用方
        ''' </summary>
        Private Async Function ProcessDocumentInBackgroundAsync(doc As KnowledgeDocument,
                                                                 kb As KnowledgeBase) As Task
            Try
                If kb Is Nothing Then Throw New Exception("文档所属的知识库不存在")

                ' 更新状态为处理中
                doc.Status = "processing"
                SaveStore()

                System.Diagnostics.Debug.WriteLine($"[KBManager] 开始处理文档: {doc.Title}")

                ' Step 1: 处理文档（下载+解析+分块，或本地文件直接解析）
                Dim chunks As List(Of KnowledgeChunk)
                If Not String.IsNullOrWhiteSpace(doc.LocalPath) Then
                    ' 本地文件路径：跳过下载直接解析
                    chunks = Await _processor.ProcessLocalDocumentAsync(
                        doc.LocalPath,
                        kb.Id,
                        doc.Id,
                        kb.ChunkSize,
                        kb.Overlap
                    )
                Else
                    ' OSS URL：下载+解析
                    chunks = Await _processor.ProcessDocumentAsync(
                        doc.OssUrl,
                        kb.Id,
                        doc.Id,
                        kb.ChunkSize,
                        kb.Overlap
                    )
                End If

                If chunks Is Nothing OrElse chunks.Count = 0 Then
                    Throw New Exception("文档处理后无有效分块")
                End If

                ' 更新文档信息
                doc.ParsedTextLength = chunks.Sum(Function(c) c.Text.Length)
                doc.ChunkCount = chunks.Count
                ' 尝试从OSS获取文件大小，失败时保持0（应在DocumentProcessor下载阶段记录）
                Try
                    If Not String.IsNullOrWhiteSpace(doc.OssUrl) Then
                        Dim req As Net.WebRequest = Net.WebRequest.Create(doc.OssUrl)
                        req.Method = "HEAD"
                        req.Timeout = 5000   ' 探文件大小不阻塞主流程（默认 100 秒会明显拖慢处理）
                        Using resp = req.GetResponse()
                            doc.FileSize = resp.ContentLength
                        End Using
                    End If
                Catch ex As Exception  '  获取失败不影响主流程
                End Try

                System.Diagnostics.Debug.WriteLine($"[KBManager] 文档分块完成: {chunks.Count} 个块")

                ' Step 2: 批量向量化
                Await VectorizeChunksAsync(chunks)

                ' Step 3: 写入 Qdrant 并更新 Point ID
                Await StoreChunksToQdrantAsync(chunks, kb.QdrantCollection, doc.Title)

                ' Step 4: 更新状态为完成
                doc.Status = "completed"
                doc.ProcessedAt = DateTime.Now

                ' 更新知识库统计
                AdjustKbCounts(kb, 0, chunks.Count)
                SaveStore()

                System.Diagnostics.Debug.WriteLine($"[KBManager] 文档处理完成: {doc.Title} ({chunks.Count} 块)")

            Catch ex As Exception
                doc.Status = "failed"
                doc.ErrorMessage = ex.Message
                doc.ProcessedAt = DateTime.Now
                SaveStore()
                System.Diagnostics.Debug.WriteLine($"[KBManager] 文档处理失败: {doc.Title} - {ex.Message}")
            End Try
        End Function

        ''' <summary>
        ''' 批量向量化（自动批次控制，避免超出 API 限制）
        ''' </summary>
        Private Async Function VectorizeChunksAsync(chunks As List(Of KnowledgeChunk)) As Task
            Try
                Dim batchSize As Integer = 10  ' 每批处理的文本数量
                Dim texts As New List(Of String)()

                For Each chunk In chunks
                    texts.Add(chunk.Text)
                Next

                ' 分批向量化
                Dim allVectors As New List(Of Single())()
                For i As Integer = 0 To texts.Count - 1 Step batchSize
                    Dim batchEnd = Math.Min(i + batchSize, texts.Count)
                    Dim batch = texts.GetRange(i, batchEnd - i)

                    Dim vectors = Await _embedding.EmbedBatchAsync(batch)
                    allVectors.AddRange(vectors)

                    System.Diagnostics.Debug.WriteLine($"[KBManager] 向量化进度: {Math.Min(batchEnd, texts.Count)}/{texts.Count}")
                Next

                ' 将向量赋回给 chunk 对象
                ' 防御：向量数与块数不一致、或存在空向量，必须让文档失败并暴露根因，
                ' 绝不允许「部分块静默丢失」或「全空向量却假成功」（历史缺陷：
                ' Ollama 未装 embedding 模型时 404 返回空列表 → 块向量全 Nothing →
                ' upsert 被跳过但文档显示 completed，用户看到「3 块成功」却检索 0 命中）
                If allVectors.Count <> chunks.Count Then
                    Throw New Exception($"向量化数量不匹配: {chunks.Count} 个分块仅获得 {allVectors.Count} 个向量（Ollama 返回异常）")
                End If
                For Each vec In allVectors
                    If vec Is Nothing OrElse vec.Length = 0 Then
                        Throw New Exception("向量化结果存在空向量（请检查 Ollama 是否已安装向量模型，如 bge-m3；可执行 ollama pull bge-m3）")
                    End If
                Next

                For i As Integer = 0 To chunks.Count - 1
                    chunks(i).EmbeddingVector = allVectors(i)
                Next

                System.Diagnostics.Debug.WriteLine($"[KBManager] 向量化完成: {allVectors.Count} 个向量")

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[KBManager] VectorizeChunksAsync 异常: {ex.Message}")
                Throw
            End Try
        End Function

        ''' <summary>
        ''' 写入 Qdrant 并更新 Point ID
        ''' </summary>
        Private Async Function StoreChunksToQdrantAsync(chunks As List(Of KnowledgeChunk),
                                                           collectionName As String,
                                                           Optional docTitle As String = "") As Task
            Try
                ' 构建 Qdrant 点列表
                Dim points As New List(Of QdrantPoint)()

                For Each chunk In chunks
                    If chunk.EmbeddingVector IsNot Nothing AndAlso chunk.EmbeddingVector.Length > 0 Then
                        Dim pointId = Guid.NewGuid()

                        ' 构建 Payload
                        Dim payload As New Dictionary(Of String, Object) From {
                            {"text", chunk.Text},
                            {"chunk_id", chunk.Id},
                            {"doc_id", chunk.DocId},
                            {"kb_id", chunk.KbId},
                            {"doc_title", docTitle}
                        }

                        If Not String.IsNullOrWhiteSpace(chunk.SourceSection) Then
                            payload("source_section") = chunk.SourceSection
                        End If

                        ' 合并自定义 Metadata
                        If chunk.Metadata IsNot Nothing Then
                            For Each kvp In chunk.Metadata
                                payload(kvp.Key) = kvp.Value
                            Next
                        End If

                        points.Add(New QdrantPoint With {
                            .Id = pointId,
                            .Vector = chunk.EmbeddingVector,
                            .Payload = payload
                        })

                        ' 记录 Qdrant Point ID
                        chunk.QdrantPointId = pointId
                    End If
                Next

                ' 防御：有分块但零个有效向量点时，必须抛异常让文档状态如实变为 failed。
                ' （历史缺陷：此场景静默跳过 upsert，文档显示 completed 但 Qdrant 无数据，
                '  用户看到「导入成功、N 块」却在检索时 0 命中，且没有任何错误提示）
                If chunks.Count > 0 AndAlso points.Count = 0 Then
                    Throw New Exception("所有分块均未获得有效向量，已取消写入 Qdrant（请检查 Ollama 向量模型是否可用，如 ollama pull bge-m3）")
                End If

                If points.Count > 0 Then
                    ' 分批 Upsert（每批 100 个点）
                    Dim upsertBatchSize As Integer = 100
                    For i As Integer = 0 To points.Count - 1 Step upsertBatchSize
                        Dim batchEnd = Math.Min(i + upsertBatchSize, points.Count)
                        Dim batch = points.GetRange(i, batchEnd - i)

                        Dim upsertResult = Await _qdrant.UpsertPointsAsync(collectionName, batch)

                        If Not upsertResult.Success Then
                            System.Diagnostics.Debug.WriteLine($"[KBManager] Qdrant Upsert 部分失败: {upsertResult.ErrorMessage}")
                        End If
                    Next

                    System.Diagnostics.Debug.WriteLine($"[KBManager] Qdrant 入库完成: {points.Count} 个点")
                End If

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[KBManager] StoreChunksToQdrantAsync 异常: {ex.Message}")
                Throw
            End Try
        End Function

#End Region

#Region "IDisposable Support"

        Protected Overloads Sub Dispose(disposing As Boolean)
            If Not _isDisposed Then
                If disposing Then
                    ' 关闭前落盘，确保元数据不丢
                    Try
                        SaveStore()
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine($"[KBManager] Dispose 前保存存储失败: {ex.Message}")
                    End Try

                    If _qdrant IsNot Nothing Then
                        _qdrant.Dispose()
                        _qdrant = Nothing
                    End If
                    If _embedding IsNot Nothing Then
                        _embedding.Dispose()
                        _embedding = Nothing
                    End If
                End If
                _isDisposed = True
            End If
        End Sub

        Public Overloads Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

#End Region

    End Class

End Namespace
