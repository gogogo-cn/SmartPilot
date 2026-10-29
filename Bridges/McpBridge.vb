Option Strict On
Option Explicit On
Option Infer On

Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading.Tasks
Imports System.Diagnostics
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ExcelAIVSTOAddIn.Services.Mcp

Namespace Bridges

    ''' <summary>
    ''' McpBridge - MCP (Model Context Protocol) API 桥梁类
    ''' 从 ExcelBridge 中提取的 MCP 相关完整 API 方法
    ''' 提供 COM 可见的方法供前端 JavaScript 调用
    ''' </summary>
    <ComVisible(True)>
    <ClassInterface(ClassInterfaceType.AutoDispatch)>
    <ProgId("ExcelAIVSTOAddIn.McpBridge")>
    Public Class McpBridge
        Implements IBridge, IMcpBridge
        ' 注意: 此类实现 IMcpBridge 接口契约（见 IBridges.vb）。
        ' 契约方法以同名公共方法显式 Implements 绑定（见各方法签名行），供依赖注入与单元测试按接口解析。

        ' 单例实例 - 使用 Lazy(Of T) 确保线程安全
        Private Shared ReadOnly _lazyInstance As New Lazy(Of McpBridge)(Function() New McpBridge(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        Private _mcpManager As McpManager
        Private _uiSyncContext As System.Threading.SynchronizationContext

#Region "单例模式"

        ''' <summary>
        ''' 获取 McpBridge 单例实例（线程安全 Lazy 初始化）
        ''' </summary>
        Public Shared ReadOnly Property Instance As McpBridge
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        ''' <summary>
        ''' 私有构造函数（单例模式）
        ''' </summary>
        Private Sub New()
            ' 加入统一生命周期管理：此 Bridge 被实例化后即登记，
            ' 由 ExcelBridge.ShutdownAll → BridgeRegistry.ShutdownAll 统一释放
            BridgeRegistry.Instance.Register(Me)

            ' 捕获UI线程的SynchronizationContext，确保后续COM调用在UI线程执行
            _uiSyncContext = System.Threading.SynchronizationContext.Current
            ' 【P2-5 修复】若在非 UI 线程初始化，SynchronizationContext.Current 为 Nothing，
            ' 必须保持 Nothing 让 RunOnUiThread 走「直接执行 + COM 自动 marshaling」的降级路径。
            ' 切勿在此处 new 一个空 SynchronizationContext 冒充 UI 上下文——
            ' 空上下文的 Send/Post 只会就地同步执行、并不会真正调度回 UI 线程。
        End Sub

#End Region

#Region "私有辅助方法"

        ''' <summary>
        ''' 在UI线程上执行操作（确保Excel COM调用在主线程执行）
        ''' </summary>
        Private Function RunOnUiThread(Of T)(func As Func(Of T)) As T
            ' 如果已经在UI线程（SynchronizationContext匹配），直接执行
            If _uiSyncContext IsNot Nothing AndAlso System.Threading.SynchronizationContext.Current Is _uiSyncContext Then
                Return func()
            End If

            ' 如果没有捕获到UI SyncContext，尝试直接执行（COM互操作会自动marshaling）
            If _uiSyncContext Is Nothing Then
                Return func()
            End If

            ' 通过SynchronizationContext.Send同步调度到UI线程
            Dim result As T = Nothing
            Dim exception As Exception = Nothing
            _uiSyncContext.Send(Sub()
                                   Try
                                       result = func()
                                   Catch ex As Exception
                                       exception = ex
                                   End Try
                               End Sub, Nothing)
            If exception IsNot Nothing Then Throw exception
            Return result
        End Function

        ''' <summary>
        ''' 获取或初始化 McpManager（McpBridge 内部使用）
        ''' </summary>
        Private Function GetMcpManager() As McpManager
            If _mcpManager Is Nothing Then
                _mcpManager = New McpManager(ExcelAIVSTOAddIn.ExcelBridge.Instance)
            End If
            Return _mcpManager
        End Function

        ''' <summary>
        ''' 动态构建 MCP 能力描述文本（用于注入系统提示词）
        ''' 【修复】将所有可用工具（含本地工具）暴露给LLM，让其通过工具列表自我识别MCP
        ''' Friend：供 Skills 引擎在执行声明式 SkillCreator 时把实时工具清单注入提示词，
        ''' 替代该 Skill 提示词里写死的工具列表（单一来源，不复制粘贴）。
        ''' </summary>
        Friend Function BuildMcpCapabilitiesText() As String
            Try
                Dim toolsJson = GetMcpManager().GetAllAvailableTools()
                Dim toolsObj = JObject.Parse(toolsJson)
                Dim tools As JArray = Nothing
                Dim totalCount As Integer = 0

                Dim successToken = DirectCast(toolsObj("success"), JToken)
                If successToken IsNot Nothing AndAlso successToken.ToString() = "True" Then
                    tools = DirectCast(toolsObj("tools"), JArray)
                    If tools IsNot Nothing AndAlso tools.HasValues Then
                        Dim countToken = DirectCast(toolsObj("count"), JToken)
                        If countToken IsNot Nothing Then totalCount = countToken.ToObject(Of Integer)()
                    End If
                End If

                Dim sb As New StringBuilder()
                ' 收集全部真实工具名，用于末尾输出"合法工具名白名单"
                Dim allNames As New List(Of String)()

                ' 有工具就全部展示，让LLM通过工具列表理解MCP
                If tools IsNot Nothing AndAlso tools.HasValues Then
                    Dim sourceGroups As New Dictionary(Of String, List(Of JToken))()
                    For Each tool As JToken In tools
                        Dim srcToken = DirectCast(tool("source"), JToken)
                        Dim source As String = If(srcToken IsNot Nothing, srcToken.ToString(), Nothing)
                        If String.IsNullOrEmpty(source) Then source = "unknown"
                        If Not sourceGroups.ContainsKey(source) Then
                            sourceGroups.Add(source, New List(Of JToken)())
                        End If
                        sourceGroups(source).Add(tool)
                    Next

                    sb.AppendLine("- 当前共有 " & totalCount.ToString() & " 个可用工具（通过MCP协议提供）:")
                    For Each kvp In sourceGroups
                        Dim sourceName = TranslateSourceName(kvp.Key)
                        sb.AppendLine(String.Format("  - **{0}**: {1} 个工具", sourceName, kvp.Value.Count))
                        For Each tool As JToken In kvp.Value
                            Dim nameToken = DirectCast(tool("name"), JToken)
                            Dim descToken = DirectCast(tool("description"), JToken)
                            Dim toolName As String = If(nameToken IsNot Nothing, nameToken.ToString(), Nothing)
                            Dim toolDesc As String = If(descToken IsNot Nothing, descToken.ToString(), Nothing)
                            If Not String.IsNullOrEmpty(toolName) Then
                                sb.AppendLine(String.Format("    - `{0}`: {1}", toolName,
                                    If(Not String.IsNullOrEmpty(toolDesc), toolDesc, "(无描述)")))
                                allNames.Add(toolName)
                            End If
                        Next
                    Next
                    sb.AppendLine(vbCrLf & "- MCP即上方这些工具所使用的协议(Model Context Protocol)")

                    ' ── 合法工具名白名单（反臆造硬约束）────────────────────────────
                    ' 实测（2026-10-02）：只逐条列 "name: 描述" 时，LLM 仍会臆造不存在的工具名
                    ' 去"实现"清单外的能力（如 excel_send_email / excel_template_render 用于批量发邮件），
                    ' 生成的 Skill 保存时/执行时才暴露"工具不存在"，用户得到的是不可用技能。
                    ' 显式白名单 + 禁止声明，比散文式描述更能压住臆造。
                    If allNames.Count > 0 Then
                        sb.AppendLine(vbCrLf & "【合法工具名白名单（共 " & allNames.Count.ToString() & " 个，必须逐字使用）】")
                        sb.AppendLine(String.Join(", ", allNames))
                        sb.AppendLine("上表之外的任何工具名都是不存在的（例如 excel_send_email、excel_template_render、web_search 均不存在）。")
                        sb.AppendLine("若用户需求超出上表能力（发邮件/联网/读写本地文件/调用外部系统），不得编造工具名——" &
                                      "改为声明 LLM 提示词形态的技能，只产出内容草稿（正文/清单/步骤），不宣称能实际执行。")
                    End If
                Else
                    sb.AppendLine("- 当前未检测到任何可用的MCP工具")
                End If

                Debug.WriteLine(String.Format("[BuildMcpCapabilitiesText] {0}字符, {1}个工具", sb.Length, totalCount))
                Return sb.ToString()

            Catch ex As Exception
                Debug.WriteLine("[BuildMcpCapabilitiesText] 异常: " & ex.Message)
                Return "- MCP工具信息获取异常"
            End Try
        End Function

        ''' <summary>
        ''' 将工具来源标识翻译为可读名称
        ''' </summary>
        Private Function TranslateSourceName(source As String) As String
            ' 先从注册中心的 sourceNames 查（内置工具自动登记）
            Dim reg = Services.Mcp.McpCapabilityRegistry.Instance
            Dim translated = reg.TranslateSource(source)
            ' 外部客户端兜底
            If translated = source Then
                Select Case source.ToLower()
                    Case "filesystem" : Return "文件系统访问(filesystem)"
                    Case Else : Return source
                End Select
            End If
            Return translated
        End Function

#End Region

#Region "状态查询"

        ''' <summary>
        ''' 获取 MCP 整体状态
        ''' </summary>
        <ComVisible(True)>
        Public Function GetMcpStatus() As String Implements IMcpBridge.GetMcpStatus
            Try
                Return GetMcpManager().GetStatus()
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 获取 MCP Server 状态
        ''' </summary>
        <ComVisible(True)>
        Public Function GetMcpServerStatus() As String Implements IMcpBridge.GetMcpServerStatus
            Try
                Return GetMcpManager().GetServerStatus()
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

#End Region

#Region "服务器控制"

        ''' <summary>
        ''' 启动 MCP Server
        ''' </summary>
        <ComVisible(True)>
        Public Async Function StartMcpServerAsync(Optional port As Integer = -1) As Task(Of String) Implements IMcpBridge.StartMcpServerAsync
            Try
                Return Await GetMcpManager().StartServerAsync(port)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 停止 MCP Server
        ''' </summary>
        <ComVisible(True)>
        Public Async Function StopMcpServerAsync() As Task(Of String) Implements IMcpBridge.StopMcpServerAsync
            Try
                Return Await GetMcpManager().StopServerAsync()
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

#End Region

#Region "配置管理"

        ''' <summary>
        ''' 获取 MCP Server 配置
        ''' </summary>
        <ComVisible(True)>
        Public Function GetMcpServerConfig() As String Implements IMcpBridge.GetMcpServerConfig
            Try
                Return GetMcpManager().GetServerConfigs()
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 保存 MCP Server 配置
        ''' </summary>
        <ComVisible(True)>
        Public Async Function SaveMcpServerConfigAsync(configJson As String) As Task(Of String) Implements IMcpBridge.SaveMcpServerConfigAsync
            Try
                Return Await GetMcpManager().SaveServerConfigAsync(configJson)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

#End Region

#Region "客户端管理"

        ''' <summary>
        ''' 添加并连接 MCP Client
        ''' </summary>
        <ComVisible(True)>
        Public Async Function AddMcpClientAsync(clientConfigJson As String) As Task(Of String) Implements IMcpBridge.AddMcpClientAsync
            Try
                Return Await GetMcpManager().AddClientAsync(clientConfigJson)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 移除 MCP Client
        ''' </summary>
        <ComVisible(True)>
        Public Async Function RemoveMcpClientAsync(name As String) As Task(Of String) Implements IMcpBridge.RemoveMcpClientAsync
            Try
                Return Await GetMcpManager().RemoveClientAsync(name)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 断开 MCP Client（保留配置）
        ''' </summary>
        <ComVisible(True)>
        Public Async Function DisconnectMcpClientAsync(name As String) As Task(Of String)
            Try
                Return Await GetMcpManager().DisconnectClientAsync(name)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

#End Region

#Region "客户端工具"

        ''' <summary>
        ''' 获取 MCP Client 的工具列表
        ''' </summary>
        <ComVisible(True)>
        Public Function GetMcpClientTools(name As String) As String Implements IMcpBridge.GetMcpClientTools
            Try
                Return GetMcpManager().GetClientTools(name)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 测试 MCP Client 连接（支持传入完整配置 JSON 或仅名称）
        ''' </summary>
        <ComVisible(True)>
        Public Async Function TestMcpClientConnectionAsync(configOrName As String) As Task(Of String) Implements IMcpBridge.TestMcpClientConnectionAsync
            Try
                Return Await GetMcpManager().TestClientConnectionAsync(configOrName)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

#End Region

#Region "工具调用"

        ''' <summary>
        ''' 调用外部 MCP 工具
        ''' </summary>
        <ComVisible(True)>
        Public Async Function CallMcpToolAsync(clientName As String, toolName As String, argumentsJson As String) As Task(Of String) Implements IMcpBridge.CallMcpToolAsync
            Try
                Return Await GetMcpManager().CallToolAsync(clientName, toolName, argumentsJson)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 同步版本 - COM 互操作使用（避免 async/await 导致 Excel 崩溃）
        ''' </summary>
        <ComVisible(True)>
        Public Function CallMcpTool(clientName As String, toolName As String, argumentsJson As String) As String Implements IMcpBridge.CallMcpTool
            Try
                Return GetMcpManager().CallToolAsync(clientName, toolName, argumentsJson).GetAwaiter().GetResult()
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 调用本地 Excel 工具（直接通过 McpServer.ExcelTools 执行，不经过 Client Pool）
        ''' 所有Excel COM操作必须在UI线程执行，否则可能静默失败
        ''' </summary>
        <ComVisible(True)>
        Public Function CallLocalTool(toolName As String, argumentsJson As String) As String Implements IMcpBridge.CallLocalTool
            Try
                Return RunOnUiThread(Function() GetMcpManager().CallLocalTool(toolName, argumentsJson))
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 重置选区上下文（新会话开始时由前端调用）
        ''' </summary>
        <ComVisible(True)>
        Public Function ResetRangeContext() As String
            Try
                Return RunOnUiThread(Function() GetMcpManager().ResetRangeContext())
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 获取所有可用 MCP 工具（供 AI 使用）
        ''' 优先从 McpCapabilityRegistry 获取完整工具列表，确保 AI 看到全部 59 个工具
        ''' </summary>
        <ComVisible(True)>
        Public Function GetAvailableMcpTools() As String Implements IMcpBridge.GetAvailableMcpTools
            Try
                Dim registry = Services.Mcp.McpCapabilityRegistry.Instance
                If registry IsNot Nothing Then
                    Dim caps = registry.GetAllCapabilities()
                    ' 兜底：如果注册中心为空（初始化时序问题），从真源重新注册所有内置工具
                    If caps Is Nothing OrElse caps.Count = 0 Then
                        Services.Mcp.McpCapabilityRegistry.RegisterAllBuiltInToolSets()
                        caps = registry.GetAllCapabilities()
                    End If
                    If caps IsNot Nothing AndAlso caps.Count > 0 Then
                        Dim mcpTools = caps.Where(Function(c) c.Kind = Services.Mcp.CapabilityKind.McpTool).ToList()
                        Dim toolsList As New List(Of Object)()
                        For Each c In mcpTools
                            Dim paramsList As New List(Of Object)()
                            If c.InputSchema IsNot Nothing AndAlso c.InputSchema.Properties IsNot Nothing Then
                                Dim requiredSet As New HashSet(Of String)(If(c.InputSchema.Required, Enumerable.Empty(Of String)()))
                                For Each kvp In c.InputSchema.Properties
                                    paramsList.Add(New With {
                                        .name = kvp.Key,
                                        .type = kvp.Value.Type,
                                        .description = kvp.Value.Description,
                                        .required = requiredSet.Contains(kvp.Key)
                                    })
                                Next
                            End If
                            ' 【修复·元数据丢失】CapabilityDescriptor 里本就有 Keywords/IsWriteOperation/IsDangerous，
                            ' 但这里原先只发了 5 个字段 → 前端 buildMcpToolMeta 读不到 t.intentKeywords。
                            ' 后果：tool-registry 注册的工具没有关键词 → agent-loop 的"要求覆盖度自检"
                            ' (detectUncoveredRequirements) 对每个工具都 continue → 未完成项恒为空，
                            ' 用户一句话列 4 条要求只做 1 条也不会被纠正（诊断日志显示 带intentKeywords=0）。
                            toolsList.Add(New With {
                                .name = c.Id,
                                .displayName = c.DisplayName,
                                .source = c.Source,
                                .description = "【" + c.DisplayName + "】" + c.Description,
                                .parameters = paramsList,
                                .intentKeywords = If(c.Keywords, New List(Of String)()),
                                .isWrite = c.IsWriteOperation,
                                .isDangerous = c.IsDangerous,
                                .dangerSeverity = c.DangerSeverity,
                                .dangerMessage = c.DangerMessage
                            })
                        Next
                        Return JsonConvert.SerializeObject(New With {
                            .success = True,
                            .tools = toolsList,
                            .count = toolsList.Count
                        })
                    End If
                End If

                ' 降级：从 MCP Server 获取
                Return RunOnUiThread(Function() GetMcpManager().GetAllAvailableTools())
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        '''  根据用户消息匹配最佳能力（后端统一匹配，替代前端 KEYWORD_SCORE_BOOST）
        ''' </summary>
        <ComVisible(True)>
        Public Function MatchCapability(userMessage As String, Optional topN As Integer = 5) As String
            Try
                Return Services.Mcp.McpCapabilityRegistry.Instance.MatchJson(userMessage, topN)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        '''  获取所有能力的 JSON（供前端 toolRegistry）
        ''' </summary>
        <ComVisible(True)>
        Public Function GetAllCapabilities() As String
            Try
                Return Services.Mcp.McpCapabilityRegistry.Instance.GetAllCapabilitiesJson()
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 获取 MCP 配置（包括 Server 和 Clients 配置）
        ''' </summary>
        <ComVisible(True)>
        Public Function GetMcpConfig() As String Implements IMcpBridge.GetMcpConfig
            Try
                Return RunOnUiThread(Function() GetMcpManager().GetMcpConfig())
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

#End Region

#region "自动重连与清理"

        ''' <summary>
        ''' 自动重连所有已保存的 MCP Clients
        ''' </summary>
        <ComVisible(True)>
        Public Async Function AutoReconnectMcpClientsAsync() As Task(Of String) Implements IMcpBridge.AutoReconnectMcpClientsAsync
            Try
                Return Await GetMcpManager().AutoReconnectAsync()
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 关闭时释放所有 MCP 客户端连接
        ''' </summary>
        Public Sub ShutdownMcpClients() Implements IMcpBridge.ShutdownMcpClients
            Try
                If _mcpManager IsNot Nothing Then
                    _mcpManager.Dispose()
                    _mcpManager = Nothing
                End If
            Catch ex As Exception
                Debug.WriteLine($"[McpBridge] 关闭 MCP 客户端异常: {ex.Message}")
            End Try
        End Sub

#End Region

#Region "IBridge 生命周期"

        Public ReadOnly Property BridgeName As String Implements IBridge.BridgeName
            Get
                Return "McpBridge"
            End Get
        End Property

        Private _disposed As Boolean = False
        Public Sub Dispose() Implements IBridge.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not _disposed Then
                ' 释放 MCP 客户端连接
                Try
                    ShutdownMcpClients()
                Catch ex As Exception  ' 
                End Try
                _disposed = True
            End If
        End Sub

#End Region

    End Class

End Namespace