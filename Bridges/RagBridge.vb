Option Explicit On
Option Infer On

Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Diagnostics
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Serialization
Imports ExcelAIVSTOAddIn.Services

Namespace Bridges

    ''' <summary>
    ''' RagBridge - RAG 知识库管理 API 桥梁类
    ''' 提供 COM 可见的方法供前端 JavaScript 调用，专门处理 RAG（检索增强生成）相关操作
    ''' </summary>
    <ComVisible(True)>
    <ClassInterface(ClassInterfaceType.AutoDispatch)>
    <ProgId("ExcelAIVSTOAddIn.RagBridge")>
    Public Class RagBridge
        Implements IBridge

        ' 单例实例 - 使用 Lazy(Of T) 确保线程安全
        Private Shared ReadOnly _lazyInstance As New Lazy(Of RagBridge)(Function() New RagBridge(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        ' RAG 组件（延迟初始化）
        Private _kbManager As Services.RAG.KnowledgeBaseManager
        Private _ragOrchestrator As Services.RAG.RagOrchestrator

        ' CamelCase JSON 序列化设置（前端 JavaScript 使用 camelCase 属性名）
        Private Shared ReadOnly _camelCaseSettings As New JsonSerializerSettings With {
            .ContractResolver = New CamelCasePropertyNamesContractResolver(),
            .NullValueHandling = NullValueHandling.Ignore,
            .Formatting = Formatting.None
        }

        ''' <summary>
        ''' Win32 P/Invoke 封装：把窗口置前台（弹 OpenFileDialog 时保证原生框出现在最上面，不会被任务栏闷死）
        ''' </summary>
        Private NotInheritable Class NativeMethods
            Private Sub New()
            End Sub
            <DllImport("user32.dll", SetLastError:=True)>
            Friend Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
            End Function
        End Class

#Region "单例模式"

        ''' <summary>
        ''' 获取 RagBridge 单例实例（线程安全 Lazy 初始化）
        ''' </summary>
        Public Shared ReadOnly Property Instance As RagBridge
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        ''' <summary>
        ''' 私有构造函数（单例模式）
        ''' </summary>
        Private Sub New()
            ' 统一生命周期管理：实例化即登记，由 BridgeRegistry.ShutdownAll 统一释放
            BridgeRegistry.Instance.Register(Me)
        End Sub

#End Region

#Region "私有辅助方法"

        ''' <summary>
        ''' 确保 KnowledgeBaseManager 已初始化
        ''' </summary>
        Private Function EnsureKbManager() As Services.RAG.KnowledgeBaseManager
            If _kbManager Is Nothing Then
                _kbManager = Services.RAG.KnowledgeBaseManager.Instance
            End If
            Return _kbManager
        End Function

#End Region

#Region "RAG 知识库管理 API"

        ''' <summary>
        ''' 打开本地文件选择对话框，返回选中文件的完整路径
        ''' 必须在 UI 线程调用（OpenFileDialog 需要 STA 线程）
        ''' 前端调用: var path = await ragBridge.BrowseLocalFile()
        ''' </summary>
        <ComVisible(True)>
        Public Function BrowseLocalFile() As String
            Using dlg As New OpenFileDialog()
                dlg.Title = "选择要导入的文档"
                dlg.Filter = "支持的文档类型|*.pdf;*.docx;*.doc;*.xlsx;*.xls;*.csv;*.txt;*.md|所有文件|*.*"
                dlg.FilterIndex = 1
                If dlg.ShowDialog() = DialogResult.OK Then
                    Return dlg.FileName
                End If
            End Using
            Return ""
        End Function

        ''' <summary>
        ''' 导入文档到知识库（接收 OSS URL）
        ''' 前端调用: await ragBridge.ImportDocumentAsync(ossUrl, kbId, title)
        ''' </summary>
        <ComVisible(True)>
        Public Function ImportDocumentAsync(ossUrl As String,
                                             kbId As String,
                                             Optional title As String = "",
                                             Optional configJson As String = "") As String
            Try
                Dim kbMgr = EnsureKbManager()
                Dim doc = kbMgr.ImportDocumentAsync(kbId, ossUrl, title).Result

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .document = doc,
                    .message = $"文档 '{If(title, doc.Id)}' 已导入，正在异步处理..."
                }, _camelCaseSettings)
            Catch ex As AggregateException
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = If(ex.InnerException?.Message, ex.Message)
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' 导入本地文件到知识库（接收本地文件路径）
        ''' 前端调用: await ragBridge.ImportLocalDocumentAsync(localPath, kbId, title)
        ''' </summary>
        <ComVisible(True)>
        Public Function ImportLocalDocumentAsync(localPath As String,
                                                  kbId As String,
                                                  Optional title As String = "") As String
            Try
                ' 验证本地文件存在
                If Not File.Exists(localPath) Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = $"本地文件不存在: {localPath}"
                    }, _camelCaseSettings)
                End If

                Dim kbMgr = EnsureKbManager()
                Dim doc = kbMgr.ImportLocalDocumentAsync(kbId, localPath, title).Result

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .document = doc,
                    .message = $"本地文档 '{doc.Title}' 已导入，正在异步处理..."
                }, _camelCaseSettings)
            Catch ex As AggregateException
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = If(ex.InnerException?.Message, ex.Message)
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' 搜索知识库
        ''' 前端调用: await ragBridge.SearchKnowledgeAsync(query, kbId, topK)
        ''' </summary>
        <ComVisible(True)>
        Public Function SearchKnowledgeAsync(query As String,
                                              kbId As String,
                                              Optional topK As Integer = 5,
                                              Optional optionsJson As String = "") As String
            Try
                Dim kbMgr = EnsureKbManager()
                Dim results = kbMgr.SearchAsync(kbId, query, topK).Result

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .results = results,
                    .count = If(results IsNot Nothing, results.Count, 0)
                }, _camelCaseSettings)
            Catch ex As AggregateException
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = If(ex.InnerException?.Message, ex.Message)
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' 获取所有知识库列表
        ''' 前端调用: await ragBridge.GetKnowledgeBasesAsync()
        ''' </summary>
        <ComVisible(True)>
        Public Function GetKnowledgeBasesAsync() As String
            Try
                Dim kbMgr = EnsureKbManager()
                Dim kbs = kbMgr.GetAllKnowledgeBases()
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .knowledgeBases = kbs,
                    .count = kbs.Count
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' 创建知识库
        ''' 前端调用: var result = await ragBridge.CreateKnowledgeBaseAsync(name, description)
        ''' 注意：同步方法，桥接层统一返回 Promise
        ''' </summary>
        <ComVisible(True)>
        Public Function CreateKnowledgeBaseAsync(name As String,
                                                  description As String,
                                                  Optional optionsJson As String = "") As String
            Try
                Dim kbMgr = EnsureKbManager()

                Dim options As Dictionary(Of String, Object) = Nothing
                If Not String.IsNullOrEmpty(optionsJson) Then
                    options = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(optionsJson)
                End If

                ' 同步等待 manager 异步方法（HTTP调用不涉及UI线程，安全使用 .Result）
                Dim kb = kbMgr.CreateKnowledgeBaseAsync(name, description, options).Result

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .knowledgeBase = kb,
                    .message = $"知识库 '{name}' 创建成功"
                }, _camelCaseSettings)
            Catch ex As AggregateException
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = If(ex.InnerException?.Message, ex.Message)
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' 删除知识库
        ''' 前端调用: await ragBridge.DeleteKnowledgeBaseAsync(kbId)
        ''' </summary>
        <ComVisible(True)>
        Public Function DeleteKnowledgeBaseAsync(kbId As String) As String
            Try
                Dim kbMgr = EnsureKbManager()
                Dim success = kbMgr.DeleteKnowledgeBaseAsync(kbId).Result

                Return JsonConvert.SerializeObject(New With {
                    .success = success,
                    .message = If(success, "知识库已删除", "删除失败"),
                    .error = If(success, Nothing, "删除失败: 操作未成功完成")
                }, _camelCaseSettings)
            Catch ex As AggregateException
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = If(ex.InnerException?.Message, ex.Message)
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' 获取知识库下的文档列表
        ''' 前端调用: await ragBridge.GetDocumentsAsync(kbId)
        ''' </summary>
        <ComVisible(True)>
        Public Function GetDocumentsAsync(kbId As String) As String
            Try
                Dim kbMgr = EnsureKbManager()
                Dim docs = kbMgr.GetDocuments(kbId)
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .documents = docs,
                    .count = docs.Count
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' 删除文档
        ''' 前端调用: await ragBridge.DeleteDocumentAsync(docId)
        ''' </summary>
        <ComVisible(True)>
        Public Function DeleteDocumentAsync(docId As String) As String
            Try
                Dim kbMgr = EnsureKbManager()
                Dim success = kbMgr.DeleteDocumentAsync(docId).Result

                Return JsonConvert.SerializeObject(New With {
                    .success = success,
                    .message = If(success, "文档已删除", "删除失败")
                }, _camelCaseSettings)
            Catch ex As AggregateException
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = If(ex.InnerException?.Message, ex.Message)
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' 重试失败的文档（重置状态并重新后台处理，不产生新记录）
        ''' 前端调用: await ragBridge.RetryDocumentAsync(docId)
        ''' </summary>
        <ComVisible(True)>
        Public Function RetryDocumentAsync(docId As String) As String
            Try
                Dim kbMgr = EnsureKbManager()
                Dim doc = kbMgr.RetryDocumentAsync(docId).Result

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .document = doc,
                    .message = $"文档 '{doc.Title}' 已重新加入处理队列"
                }, _camelCaseSettings)
            Catch ex As AggregateException
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = If(ex.InnerException?.Message, ex.Message)
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' 同步"待同步"状态的知识库（创建时 Qdrant 不可达，之后重试建集合并置为活跃）
        ''' 前端调用: await ragBridge.SyncKnowledgeBaseStatusAsync()
        ''' </summary>
        <ComVisible(True)>
        Public Function SyncKnowledgeBaseStatusAsync() As String
            Try
                Dim kbMgr = EnsureKbManager()
                Dim synced = kbMgr.SyncPendingKnowledgeBasesAsync().Result

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .synced = synced,
                    .message = If(synced > 0, $"已同步 {synced} 个知识库", "没有待同步的知识库")
                }, _camelCaseSettings)
            Catch ex As AggregateException
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = If(ex.InnerException?.Message, ex.Message)
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' RAG 对话增强：检索默认（或指定）知识库，返回可注入 LLM 的上下文本。
        ''' 未命中时 shouldFallback=True，调用方应忽略 contextText 走普通对话。
        ''' 前端调用: await ragBridge.EnhanceWithRagAsync(question, kbId?)
        ''' </summary>
        <ComVisible(True)>
        Public Function EnhanceWithRagAsync(question As String, Optional kbId As String = "") As String
            Try
                If String.IsNullOrWhiteSpace(question) Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = "问题不能为空",
                        .hasContext = False,
                        .shouldFallback = True
                    }, _camelCaseSettings)
                End If

                Dim orchestrator = Services.RAG.RagOrchestrator.Instance
                ' Task.Run 包一层：COM 调用来自 UI 线程，.Result 直接阻塞会让异步续体无法回帖导致死锁
                Dim result = Task.Run(Function() orchestrator.RetrieveAndEnhanceAsync(question, kbId)).Result

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .hasContext = result.HasContext,
                    .contextText = result.ContextText,
                    .sourceReferences = result.SourceReferences,
                    .retrievedChunks = result.RetrievedChunks,
                    .shouldFallback = result.ShouldFallback
                }, _camelCaseSettings)
            Catch ex As AggregateException
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = If(ex.InnerException?.Message, ex.Message),
                    .hasContext = False,
                    .shouldFallback = True
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message,
                    .hasContext = False,
                    .shouldFallback = True
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' RAG 检索测试（测试台专用）
        ''' 前端调用: await ragBridge.TestRagQueryAsync(question, kbId, topK)
        ''' </summary>
        <ComVisible(True)>
        Public Function TestRagQueryAsync(question As String,
                                           kbId As String,
                                           Optional topK As Integer = 5) As String
            Try
                Dim kbMgr = EnsureKbManager()
                Dim result = kbMgr.TestQueryAsync(kbId, question, topK).Result

                Return JsonConvert.SerializeObject(New With {
                    .success = result.Success,
                    .question = result.Question,
                    .results = result.Results,
                    .enhancedAnswer = result.EnhancedAnswer,
                    .totalRetrieved = result.TotalRetrieved,
                    .processingTimeMs = result.ProcessingTimeMs
                }, _camelCaseSettings)
            Catch ex As AggregateException
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = If(ex.InnerException?.Message, ex.Message)
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' 获取 RAG 配置
        ''' 前端调用: await ragBridge.GetRagConfigAsync()
        ''' </summary>
        <ComVisible(True)>
        Public Function GetRagConfigAsync() As String
            Try
                Dim kbMgr = EnsureKbManager()
                Dim config = kbMgr.GetConfig()

                Dim safeConfig = New With {
                    .defaultEmbeddingProvider = config.DefaultEmbeddingProvider,
                    .localBgeEndpoint = config.LocalBgeEndpoint,
                    .localBgeModel = config.LocalBgeModel,
                    .cloudEndpoint = config.CloudEndpoint,
                    .cloudApiKey = If(String.IsNullOrEmpty(config.CloudApiKey), "", "***"),
                    .cloudModel = config.CloudModel,
                    .qdrantEndpoint = config.QdrantEndpoint,
                    .qdrantApiKey = If(Not String.IsNullOrEmpty(config.QdrantApiKey), "***", ""),
                    .defaultTopK = config.DefaultTopK,
                    .scoreThreshold = config.ScoreThreshold,
                    .enableRagForChat = config.EnableRagForChat,
                    .maxContextChunks = config.MaxContextChunks,
                    .defaultKnowledgeBaseId = config.DefaultKnowledgeBaseId
                }

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .config = safeConfig
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' 保存 RAG 配置
        ''' 前端调用: await ragBridge.SaveRagConfigAsync(configJson)
        ''' </summary>
        <ComVisible(True)>
        Public Function SaveRagConfigAsync(configJson As String) As String
            Try
                Dim kbMgr = EnsureKbManager()
                Dim newConfig = JsonConvert.DeserializeObject(Of Services.RAG.RagConfig)(configJson)

                ' 如果传入的是掩码值（***），保留原有值
                Dim currentConfig = kbMgr.GetConfig()
                If newConfig.CloudApiKey = "***" Then newConfig.CloudApiKey = currentConfig.CloudApiKey
                If Not String.IsNullOrEmpty(newConfig.QdrantApiKey) AndAlso newConfig.QdrantApiKey = "***" Then
                    newConfig.QdrantApiKey = currentConfig.QdrantApiKey
                End If

                Dim success = kbMgr.SaveConfigAsync(newConfig).Result

                Return JsonConvert.SerializeObject(New With {
                    .success = success,
                    .message = If(success, "配置已保存", "保存失败")
                }, _camelCaseSettings)
            Catch ex As AggregateException
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = If(ex.InnerException?.Message, ex.Message)
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' 测试 Qdrant 连接
        ''' 前端调用: await ragBridge.TestQdrantConnectionAsync()
        ''' </summary>
        <ComVisible(True)>
        Public Function TestQdrantConnectionAsync() As String
            Try
                Dim kbMgr = EnsureKbManager()
                Dim config = kbMgr.GetConfig()
                Using tempClient As New Services.RAG.QdrantClient(config.QdrantEndpoint, config.QdrantApiKey)
                    Dim isHealthy = tempClient.HealthCheckAsync().Result

                    Return JsonConvert.SerializeObject(New With {
                        .success = isHealthy,
                        .message = If(isHealthy, "Qdrant 连接正常", "无法连接到 Qdrant"),
                        .endpoint = config.QdrantEndpoint
                    }, _camelCaseSettings)
                End Using
            Catch ex As AggregateException
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = If(ex.InnerException?.Message, ex.Message)
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' 测试 Embedding 服务连接
        ''' 前端调用: await ragBridge.TestEmbeddingConnectionAsync(provider)
        ''' </summary>
        <ComVisible(True)>
        Public Function TestEmbeddingConnectionAsync(provider As String) As String
            Try
                Dim kbMgr = EnsureKbManager()
                Dim config = kbMgr.GetConfig()

                Using tempService As New Services.RAG.EmbeddingService(config)
                    Dim testResult = tempService.TestConnectionAsync(provider).Result

                    Return JsonConvert.SerializeObject(New With {
                        .success = testResult.Success,
                        .provider = testResult.Provider,
                        .message = testResult.Message,
                        .latencyMs = testResult.LatencyMs
                    }, _camelCaseSettings)
                End Using
            Catch ex As AggregateException
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = If(ex.InnerException?.Message, ex.Message)
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' 获取 Qdrant 安装状态
        ''' 前端调用：await ragBridge.GetQdrantInstallStatus()
        ''' </summary>
        <ComVisible(True)>
        Public Function GetQdrantInstallStatus() As String
            Try
                Dim installer = New Services.RAG.QdrantInstaller()
                Dim status = installer.GetStatus()
        
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .installed = status("installed"),
                    .running = status("running"),
                    .installDir = status("installDir"),
                    .exePath = status("exePath")
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function
        
        ''' <summary>
        ''' 打开文件选择对话框（Windows 原生）
        ''' 前端调用：await ragBridge.ShowOpenFileDialogAsync()
        ''' </summary>
        <ComVisible(True)>
        Public Function ShowOpenFileDialog() As String
            Try
                Using dlg As New Windows.Forms.OpenFileDialog()
                    dlg.Title = "选择 Qdrant 安装包"
                    dlg.Filter = "ZIP 文件 (*.zip)|*.zip|所有文件 (*.*)|*.*"
                    dlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) & "\Downloads"
                            
                    If dlg.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        Return JsonConvert.SerializeObject(New With {
                            .success = True,
                            .filePath = dlg.FileName
                        }, _camelCaseSettings)
                    Else
                        Return JsonConvert.SerializeObject(New With {
                            .success = False,
                            .message = "用户取消选择"
                        }, _camelCaseSettings)
                    End If
                End Using
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' 弹出 Windows 原生文件选择对话框，用于选目标 Excel 文件。
        ''' 与 ExcelBridge.PickExcelFile 一致，供 HTTP API 复用。
        ''' 注意：HTTP 请求来自线程池（MTA），ShowDialog 必须在 STA 线程 + 有消息循环的上下文执行，
        ''' 否则对话框不会弹出来。这里新建专用 STA Thread + ThreadPool 超时兜底 + SetForegroundWindow，
        ''' 避免"默默挂 2.5s，前端误判成浏览器沙箱阻止"。
        ''' </summary>
        <ComVisible(True)>
        Public Shared Function BrowseExcelFile(Optional filter As String = "", Optional title As String = "", Optional initialDir As String = "") As String
            Dim result As String = Nothing
            Dim captured As Exception = Nothing
            If String.IsNullOrEmpty(title) Then title = "选择目标 Excel 文件"
            If String.IsNullOrEmpty(filter) Then filter = "Excel 文件 (*.xlsx;*.xls;*.xlsm;*.csv)|*.xlsx;*.xls;*.xlsm;*.csv|Excel 工作簿 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*"
            If String.IsNullOrEmpty(initialDir) Then
                Try : initialDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) : Catch : initialDir = "" : End Try
            End If
            Dim fFilter = filter, fTitle = title, fInit = initialDir

            Dim worker = New Threading.Thread(
                Sub()
                    Try
                        Application.EnableVisualStyles()
                        Using dlg As New Windows.Forms.OpenFileDialog()
                            dlg.Title = fTitle
                            dlg.Filter = fFilter
                            dlg.CheckFileExists = False
                            dlg.CheckPathExists = True
                            Try
                                If Directory.Exists(fInit) Then dlg.InitialDirectory = fInit
                            Catch : End Try

                            ' 尝试把对话框顶到最前：借用一个临时 Form 作为父窗口
                            Dim owner As Windows.Forms.IWin32Window = Nothing
                            Dim frm As Windows.Forms.Form = Nothing
                            Try
                                frm = New Windows.Forms.Form() With {
                                    .TopMost = True,
                                    .StartPosition = Windows.Forms.FormStartPosition.CenterScreen,
                                    .ShowInTaskbar = False,
                                    .FormBorderStyle = Windows.Forms.FormBorderStyle.None,
                                    .Opacity = 0D,
                                    .Size = New Drawing.Size(1, 1)
                                }
                                frm.Show()
                                Try
                                    NativeMethods.SetForegroundWindow(frm.Handle)
                                Catch : End Try
                                frm.Activate()
                                owner = frm
                            Catch : End Try

                            Dim dr As Windows.Forms.DialogResult
                            Try
                                If owner IsNot Nothing Then dr = dlg.ShowDialog(owner) Else dr = dlg.ShowDialog()
                            Finally
                                Try
                                    If frm IsNot Nothing Then
                                        frm.Close()
                                        frm.Dispose()
                                    End If
                                Catch : End Try
                            End Try

                            If dr = Windows.Forms.DialogResult.OK Then
                                result = JsonConvert.SerializeObject(New With {
                                    .success = True,
                                    .filePath = dlg.FileName
                                }, _camelCaseSettings)
                            Else
                                result = JsonConvert.SerializeObject(New With {
                                    .success = False,
                                    .message = "用户取消选择"
                                }, _camelCaseSettings)
                            End If
                        End Using
                    Catch ex As Exception
                        captured = ex
                    End Try
                End Sub)
            worker.SetApartmentState(Threading.ApartmentState.STA)
            worker.IsBackground = True
            worker.Start()
            ' 30s 兜底：防止 STA 线程挂住整个 HTTP 请求
            If Not worker.Join(TimeSpan.FromSeconds(30)) Then
                Try : worker.Interrupt() : Catch : End Try
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .message = "文件选择对话框超时，请重试"
                }, _camelCaseSettings)
            End If
            If captured IsNot Nothing Then
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .[error] = captured.Message,
                    .message = "弹出文件选择框失败：" & captured.Message
                }, _camelCaseSettings)
            End If
            If String.IsNullOrEmpty(result) Then
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .message = "未能获取选中的文件路径"
                }, _camelCaseSettings)
            End If
            Return result
        End Function

        ''' <summary>
        ''' 获取宿主（当前 Excel 打开的）工作簿路径信息。
        ''' 返回 JSON: {success, hasHost, path, directory, name}
        ''' </summary>
        <ComVisible(True)>
        Public Shared Function GetHostWorkbook() As String
            Try
                Dim app As Object = Nothing
                Try
                    app = Marshal.GetActiveObject("Excel.Application")
                Catch : End Try

                Dim wb As Object = Nothing
                If app IsNot Nothing Then
                    Try : wb = Interaction.CallByName(app, "ActiveWorkbook", CallType.Get) : Catch : End Try
                End If

                If wb Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = True,
                        .hasHost = False,
                        .path = "",
                        .directory = "",
                        .name = ""
                    }, _camelCaseSettings)
                End If

                Dim wbName As String = ""
                Try : wbName = CStr(Interaction.CallByName(wb, "Name", CallType.Get)) : Catch : End Try

                Dim fullPath As String = ""
                Try : fullPath = CStr(Interaction.CallByName(wb, "FullName", CallType.Get)) : Catch : End Try

                Dim isSaved As Boolean = False
                Try : isSaved = Not String.IsNullOrEmpty(fullPath) AndAlso fullPath.Contains(Path.DirectorySeparatorChar.ToString()) : Catch : End Try

                Dim dir As String = ""
                If isSaved Then
                    Try : dir = Path.GetDirectoryName(fullPath) : Catch : End Try
                End If

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .hasHost = isSaved,
                    .path = If(isSaved, fullPath, ""),
                    .directory = If(dir, ""),
                    .name = wbName
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message,
                    .hasHost = False,
                    .path = "",
                    .directory = "",
                    .name = ""
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' 安装 Qdrant
        ''' 前端调用: await ragBridge.InstallQdrantAsync()
        ''' </summary>
        <ComVisible(True)>
        Public Async Function InstallQdrantAsync() As Task(Of String)
            Dim installer = New Services.RAG.QdrantInstaller()

            ' 简单进度跟踪（实际进度通过返回值传递）
            Dim lastMessage = ""
            Dim lastPercentage = 0

            Dim result = Await installer.InstallAsync(
                Sub(message, percentage)
                    lastMessage = message
                    lastPercentage = percentage
                    Debug.WriteLine($"[QdrantInstall] {message} ({percentage}%)")
                End Sub
            )

            If result Then
                ' 安装成功后自动启动
                Dim started = installer.StartService()

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .installed = True,
                    .running = started,
                    .message = If(started, "安装并启动成功", "安装成功，启动失败"),
                    .lastMessage = lastMessage,
                    .percentage = lastPercentage
                }, _camelCaseSettings)
            Else
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .installed = False,
                    .message = lastMessage,
                    .percentage = lastPercentage
                }, _camelCaseSettings)
            End If
        End Function

        ''' <summary>
        ''' 从本地文件安装 Qdrant
        ''' 前端调用: await ragBridge.InstallQdrantFromLocalFileAsync(zipFilePath)
        ''' </summary>
        <ComVisible(True)>
        Public Function InstallQdrantFromLocalFileAsync(zipFilePath As String) As String
            Dim installer = New Services.RAG.QdrantInstaller()

            Dim lastMessage = ""
            Dim lastPercentage = 0

            Dim result = installer.InstallFromLocalFile(zipFilePath,
                Sub(message, percentage)
                    lastMessage = message
                    lastPercentage = percentage
                    Debug.WriteLine($"[QdrantInstall] {message} ({percentage}%)")
                End Sub
            )

            If result Then
                ' 安装成功后自动启动
                Dim started = installer.StartService()

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .installed = True,
                    .running = started,
                    .message = If(started, "安装并启动成功", "安装成功，启动失败"),
                    .lastMessage = lastMessage,
                    .percentage = lastPercentage
                }, _camelCaseSettings)
            Else
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .installed = False,
                    .message = lastMessage,
                    .percentage = lastPercentage
                }, _camelCaseSettings)
            End If
        End Function

        ''' <summary>
        ''' 启动 Qdrant 服务
        ''' 前端调用: await ragBridge.StartQdrantAsync()
        ''' </summary>
        <ComVisible(True)>
        Public Function StartQdrantAsync() As String
            Try
                Dim installer = New Services.RAG.QdrantInstaller()
                Dim started = installer.StartService()

                Return JsonConvert.SerializeObject(New With {
                    .success = started,
                    .running = installer.IsRunning(),
                    .message = If(started, "Qdrant 启动成功", "Qdrant 启动失败")
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' 停止 Qdrant 服务
        ''' 前端调用: await ragBridge.StopQdrantAsync()
        ''' </summary>
        <ComVisible(True)>
        Public Function StopQdrantAsync() As String
            Try
                Dim installer = New Services.RAG.QdrantInstaller()
                Dim stopped = installer.StopService()

                Return JsonConvert.SerializeObject(New With {
                    .success = stopped,
                    .running = installer.IsRunning(),
                    .message = If(stopped, "Qdrant 已停止", "Qdrant 停止失败")
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' 检测 RAG 本地环境：Ollama 安装/运行状态 + embedding 模型是否就绪。
        ''' 前端调用: var env = await ragBridge.CheckRagEnvironmentAsync()
        ''' 返回: { ollamaInstalled, ollamaRunning, modelInstalled, ready, modelName, endpoint, nextStep }
        ''' nextStep: install_ollama / start_ollama / pull_model / ready
        ''' </summary>
        <ComVisible(True)>
        Public Async Function CheckRagEnvironmentAsync() As Task(Of String)
            Try
                Dim cfg = _kbManager.GetConfig()
                Dim endpoint = If(String.IsNullOrWhiteSpace(cfg.LocalBgeEndpoint), "http://127.0.0.1:11434", cfg.LocalBgeEndpoint)
                Dim model = If(String.IsNullOrWhiteSpace(cfg.LocalBgeModel), "bge-m3", cfg.LocalBgeModel)

                Dim status = Await Services.RAG.OllamaEnvironment.CheckEnvironmentAsync(endpoint, model)

                Return JsonConvert.SerializeObject(status, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' 一键拉起 embedding 模型（异步进程 ollama pull）。
        ''' 前端调用: var r = await ragBridge.StartModelPull()
        ''' 返回: { success, started, message }
        ''' </summary>
        <ComVisible(True)>
        Public Function StartModelPull() As String
            Try
                Dim cfg = _kbManager.GetConfig()
                Dim model = If(String.IsNullOrWhiteSpace(cfg.LocalBgeModel), "bge-m3", cfg.LocalBgeModel)

                Dim started = Services.RAG.OllamaEnvironment.StartModelPull(model)

                Return JsonConvert.SerializeObject(New With {
                    .success = started,
                    .started = started,
                    .model = model,
                    .message = If(started, "已开始拉取 " & model, Services.RAG.OllamaEnvironment.GetLastError())
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .started = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' 查询模型拉起进度（前端 1~2s 轮询）。
        ''' 前端调用: var st = await ragBridge.GetModelPullProgress()
        ''' 返回: { running, percent, lastLine, error, elapsedSeconds }
        ''' </summary>
        <ComVisible(True)>
        Public Function GetModelPullProgress() As String
            Try
                Dim state = Services.RAG.OllamaEnvironment.GetPullState()
                Return JsonConvert.SerializeObject(state, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' 取消正在进行的模型拉起。
        ''' 前端调用: var r = await ragBridge.StopModelPull()
        ''' </summary>
        <ComVisible(True)>
        Public Function StopModelPull() As String
            Try
                Services.RAG.OllamaEnvironment.StopModelPull()
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .message = "已取消"
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function

        ''' <summary>
        ''' 用系统默认浏览器打开外部链接（用于 Ollama 下载页等）。
        ''' 前端调用: await ragBridge.OpenExternalUrl("https://ollama.com/download")
        ''' </summary>
        <ComVisible(True)>
        Public Function OpenExternalUrl(url As String) As String
            Try
                If String.IsNullOrWhiteSpace(url) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "链接为空"}, _camelCaseSettings)
                End If
                ' 仅允许 http/https，防止 file: 等危险 scheme
                Dim uri As Uri = Nothing
                If Not Uri.TryCreate(url.Trim(), UriKind.Absolute, uri) OrElse
                   (uri.Scheme <> Uri.UriSchemeHttp AndAlso uri.Scheme <> Uri.UriSchemeHttps) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "仅支持 http/https 链接"}, _camelCaseSettings)
                End If
                Dim psi As New ProcessStartInfo()
                psi.FileName = uri.AbsoluteUri
                psi.UseShellExecute = True
                Process.Start(psi)
                Return JsonConvert.SerializeObject(New With {.success = True}, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                }, _camelCaseSettings)
            End Try
        End Function

#End Region

#Region "IBridge 生命周期"

        Public ReadOnly Property BridgeName As String Implements IBridge.BridgeName
            Get
                Return "RagBridge"
            End Get
        End Property

        Private _disposed As Boolean = False
        Public Sub Dispose() Implements IBridge.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not _disposed Then
                _disposed = True
                If disposing Then
                    ' 【关键修复】释放 KnowledgeBaseManager 以关闭 RAG 相关 HttpClient
                    ' 否则 QdrantClient/EmbeddingService 的 HttpClient 连接不会被释放，
                    ' 导致 socket 泄漏影响下次启动
                    Try
                        If _kbManager IsNot Nothing Then
                            _kbManager.Dispose()
                            _kbManager = Nothing
                        End If
                        ' 同时释放单例（可能被其他地方直接访问）
                        Services.RAG.KnowledgeBaseManager.Instance.Dispose()
                        System.Diagnostics.Debug.WriteLine("[RagBridge] KnowledgeBaseManager 已释放")
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine($"[RagBridge] 释放 KnowledgeBaseManager 失败: {ex.Message}")
                    End Try
                End If
            End If
        End Sub

#End Region

    End Class

End Namespace