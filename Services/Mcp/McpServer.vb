Option Strict Off
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services.Mcp

    '#Region "MCP Server 实现"

    ''' <summary>
    ''' MCP Server 核心 - 处理 MCP 协议的请求并响应
    ''' 通过 ExcelMCPTools 暴露 Excel 操作能力
    ''' </summary>
    Public Class McpServer
        Implements IDisposable

        Private _serverInfo As ServerInfo
        Private _serverCapabilities As ServerCapabilities
        Private _excelTools As ExcelMcpTools
        Private _pptTools As PptMcpTools
    Private _computerTools As ComputerAgentTools
        Private _isRunning As Boolean = False
        Private _isInitialized As Boolean = False
        Private _isDisposed As Boolean = False

        ''' <summary>
        ''' 是否正在运行
        ''' </summary>
        Public ReadOnly Property IsRunning As Boolean
            Get
                Return _isRunning
            End Get
        End Property

        ''' <summary>
        ''' 是否已完成 MCP 初始化握手
        ''' </summary>
        Public ReadOnly Property IsInitialized As Boolean
            Get
                Return _isInitialized
            End Get
        End Property

        ''' <summary>
        ''' Server 信息
        ''' </summary>
        Public ReadOnly Property ServerInfo As ServerInfo
            Get
                Return _serverInfo
            End Get
        End Property

        ''' <summary>
        ''' Excel 工具注册表
        ''' </summary>
        Public ReadOnly Property ExcelTools As ExcelMcpTools
            Get
                Return _excelTools
            End Get
        End Property

        ''' <summary>
        ''' PPT 工具注册表（与 Excel 工具完全分离）
        ''' </summary>
        Public ReadOnly Property PptTools As PptMcpTools
            Get
                Return _pptTools
            End Get
        End Property

        Public ReadOnly Property ComputerTools As ComputerAgentTools
            Get
                Return _computerTools
            End Get
        End Property

        Public Sub New(bridge As Object)
            _serverInfo = New ServerInfo With {
                .Name = McpConstants.IMPLEMENTATION_NAME,
                .Version = McpConstants.IMPLEMENTATION_VERSION
            }

            _serverCapabilities = New ServerCapabilities() With {
                .Tools = New ToolsCapability() With {.ListChanged = False},
                .Resources = New ResourcesCapability() With {.Subscribe = False, .ListChanged = False},
                .Prompts = New PromptsCapability() With {.ListChanged = False}
            }

            _excelTools = New ExcelMcpTools(bridge)
            _pptTools = New PptMcpTools(bridge)
            ' 从真源一次性注册所有内置本地工具（加新工具只改 McpCapabilityRegistry.RegisterAllBuiltInToolSets）
            McpCapabilityRegistry.RegisterAllBuiltInToolSets()
            Debug.WriteLine("[McpServer] 工具注册完成: " & McpCapabilityRegistry.Instance.GetAllCapabilities().Count & " 个本地工具")
            _isRunning = True
        End Sub

        '#Region "请求处理"

        ''' <summary>
        ''' 处理 JSON-RPC 请求并返回响应 JSON 字符串
        ''' </summary>
        ''' <param name="requestJson">JSON-RPC 请求</param>
        ''' <returns>JSON-RPC 响应（或 Nothing 如果是通知）</returns>
        Public Function HandleRequest(requestJson As String) As String
            Try
                Dim request = JToken.Parse(requestJson)

                ' 检查是否为通知（无 id）
                Dim idToken = request("id")
                If idToken Is Nothing Then
                    HandleNotification(request)
                    Return Nothing
                End If

                Dim id = idToken.ToObject(Of Object)()
                Dim methodToken = request("method")
                Dim method = If(methodToken IsNot Nothing, methodToken.ToString(), Nothing)
                Dim params = request("params")

                If String.IsNullOrEmpty(method) Then
                    Return BuildErrorResponse(id, McpConstants.INVALID_REQUEST, "缺少 method 字段")
                End If

                ' 路由到对应处理方法
                Select Case method
                    Case McpConstants.METHOD_INITIALIZE
                        Return HandleInitialize(id, params)

                    Case McpConstants.METHOD_INITIALIZED
                        _isInitialized = True
                        Return BuildSuccessResponse(id, New JObject())

                    Case McpConstants.METHOD_PING
                        Return BuildSuccessResponse(id, New JObject())

                    Case McpConstants.METHOD_TOOLS_LIST
                        Return HandleToolsList(id, params)

                    Case McpConstants.METHOD_TOOLS_CALL
                        Return HandleToolsCall(id, params)

                    Case McpConstants.METHOD_RESOURCES_LIST
                        Return HandleResourcesList(id, params)

                    Case McpConstants.METHOD_RESOURCES_READ
                        Return HandleResourcesRead(id, params)

                    Case McpConstants.METHOD_RESOURCES_TEMPLATES_LIST
                        Return HandleResourcesTemplatesList(id)

                    Case McpConstants.METHOD_PROMPTS_LIST
                        Return HandlePromptsList(id)

                    Case Else
                        Return BuildErrorResponse(id, McpConstants.METHOD_NOT_FOUND, $"未知方法: {method}")
                End Select

            Catch ex As JsonException
                Return BuildErrorResponse(Nothing, McpConstants.PARSE_ERROR, $"JSON 解析错误: {ex.Message}")
            Catch ex As Exception
                Debug.WriteLine($"MCP Server 处理请求异常: {ex.Message}")
                Return BuildErrorResponse(Nothing, McpConstants.INTERNAL_ERROR, $"内部错误: {ex.Message}")
            End Try
        End Function

        ''' <summary>
        ''' 处理通知（无 id 的请求）
        ''' </summary>
        Private Sub HandleNotification(request As JToken)
            Dim methodToken = request("method")
            Dim method = If(methodToken IsNot Nothing, methodToken.ToString(), Nothing)
            Select Case method
                Case McpConstants.METHOD_INITIALIZED
                    _isInitialized = True
                    Debug.WriteLine("MCP Server: 收到 Initialized 通知")
                Case McpConstants.METHOD_CANCELLED
                    Debug.WriteLine("MCP Server: 收到取消通知")
                Case Else
                    Debug.WriteLine($"MCP Server: 收到未知通知: {method}")
            End Select
        End Sub

        '#End Region

        '#Region "具体方法处理"

        Private Function HandleInitialize(id As Object, params As JToken) As String
            Try
                Dim initParams As InitializeParams = Nothing
                If params IsNot Nothing Then
                    initParams = params.ToObject(Of InitializeParams)()
                End If

                ' 协商协议版本
                Dim agreedVersion = McpConstants.PROTOCOL_VERSION
                If initParams IsNot Nothing AndAlso Not String.IsNullOrEmpty(initParams.ProtocolVersion) Then
                    ' 使用客户端支持的版本
                    agreedVersion = initParams.ProtocolVersion
                End If

                Dim result = New InitializeResult With {
                    .ProtocolVersion = agreedVersion,
                    .Capabilities = _serverCapabilities,
                    .ServerInfo = _serverInfo,
                    .Instructions = "Excel AI 助手 MCP Server - 提供 Excel 工作表读/写、VBA 执行、搜索、图表创建等工具。"
                }

                _isRunning = True

                Return BuildSuccessResponse(id, JToken.FromObject(result))

            Catch ex As Exception
                Return BuildErrorResponse(id, McpConstants.INTERNAL_ERROR, $"初始化失败: {ex.Message}")
            End Try
        End Function

        Private Function HandleToolsList(id As Object, params As JToken) As String
            Try
                Dim cursor As String = Nothing
                If params IsNot Nothing Then
                    Dim cursorToken = params.SelectToken("cursor")
                    If cursorToken IsNot Nothing Then
                        cursor = cursorToken.ToString()
                    End If
                End If

                ' Excel + PPT 两套独立工具集合并暴露
                Dim tools As New List(Of McpTool)()
                tools.AddRange(ExcelMcpTools.GetToolDefinitions())
                tools.AddRange(PptMcpTools.GetToolDefinitions())
                tools.AddRange(ComputerAgentTools.GetToolDefinitions())
                tools.AddRange(WebSearchTools.GetToolDefinitions())
                tools.AddRange(TimeTools.GetToolDefinitions())

                ' 简单的分页支持（每页最多 50 个工具）
                Dim nextCursor As String = Nothing
                If Not String.IsNullOrEmpty(cursor) AndAlso tools.Count > 50 Then
                    tools = tools.GetRange(50, Math.Min(50, tools.Count - 50))
                End If

                Dim result = New ToolsListResult With {
                    .Tools = tools,
                    .NextCursor = nextCursor
                }

                Return BuildSuccessResponse(id, JToken.FromObject(result))

            Catch ex As Exception
                Return BuildErrorResponse(id, McpConstants.INTERNAL_ERROR, $"获取工具列表失败: {ex.Message}")
            End Try
        End Function

        Private Function HandleToolsCall(id As Object, params As JToken) As String
            Try
                Dim callParams As ToolsCallParams = Nothing
                If params IsNot Nothing Then
                    callParams = params.ToObject(Of ToolsCallParams)()
                End If
                If callParams Is Nothing OrElse String.IsNullOrEmpty(callParams.Name) Then
                    Return BuildErrorResponse(id, McpConstants.INVALID_PARAMS, "缺少工具名称")
                End If

                ' 按工具名前缀分发：ppt_ → PPT 工具集，computer_ → ComputerAgent，
                ' web_ → 联网搜索工具集，time_ → 系统时间工具集，其余 → Excel
                Dim toolResult As ToolsCallResult
                If callParams.Name.StartsWith("ppt_", StringComparison.OrdinalIgnoreCase) Then
                    toolResult = _pptTools.ExecuteTool(callParams.Name, callParams.Arguments)
                ElseIf callParams.Name.StartsWith("computer_", StringComparison.OrdinalIgnoreCase) Then
                    toolResult = ComputerAgentTools.ExecuteTool(callParams.Name, callParams.Arguments)
                ElseIf callParams.Name.StartsWith("web_", StringComparison.OrdinalIgnoreCase) Then
                    toolResult = WebSearchTools.ExecuteTool(callParams.Name, callParams.Arguments)
                ElseIf callParams.Name.StartsWith("time_", StringComparison.OrdinalIgnoreCase) Then
                    toolResult = TimeTools.ExecuteTool(callParams.Name, callParams.Arguments)
                Else
                    toolResult = _excelTools.ExecuteTool(callParams.Name, callParams.Arguments)
                End If

                Return BuildSuccessResponse(id, JToken.FromObject(toolResult))

            Catch ex As Exception
                ' 【P2 修复】接入 ErrorRecovery 统一错误分类：为异常附加恢复动作建议
                ' （re_extract=参数重提取 / rollback=已回滚 / partial_result=超时部分结果），
                ' 让前端与 LLM 能据此引导重试，而不是只拿到一句裸异常文本。
                ' 一处接线覆盖 Excel + PPT 全部工具的执行异常。
                Dim recoveryAction As String = "unknown_error"
                Dim recoveryMsg As String = ex.Message
                Try
                    ' 注意：callParams 声明在 Try 块内，Catch 中不可见（VB 块级作用域），
                    ' 从原始 params(JToken) 提取工具名与参数
                    Dim recTool As String = ""
                    If params IsNot Nothing AndAlso params("name") IsNot Nothing Then
                        recTool = params("name").ToString()
                    End If
                    Dim recCtx As New Dictionary(Of String, Object) From {
                        {"tool", recTool},
                        {"arguments", If(params IsNot Nothing, params("arguments"), Nothing)}
                    }
                    Dim recovery = Services.ErrorRecovery.Instance.Handle(ex, recCtx).GetAwaiter().GetResult()
                    If recovery IsNot Nothing Then
                        If recovery.ContainsKey("action") Then recoveryAction = CStr(recovery("action"))
                        If recovery.ContainsKey("message") Then recoveryMsg = CStr(recovery("message"))
                    End If
                Catch recEx As Exception
                    Debug.WriteLine($"[McpServer] ErrorRecovery 分类失败: {recEx.Message}")
                End Try

                Dim errorResult = New ToolsCallResult() With {
                    .IsError = True
                }
                errorResult.Content.Add(New McpTextContent With {.Text = $"工具调用异常: {ex.Message}（恢复建议: {recoveryAction} — {recoveryMsg}）"})
                Return BuildSuccessResponse(id, JToken.FromObject(errorResult))
            End Try
        End Function

        Private Function HandleResourcesList(id As Object, params As JToken) As String
            Try
                Dim resources = _excelTools.GetResourceDefinitions()
                Dim result = New ResourcesListResult With {
                    .Resources = resources
                }
                Return BuildSuccessResponse(id, JToken.FromObject(result))
            Catch ex As Exception
                Return BuildErrorResponse(id, McpConstants.INTERNAL_ERROR, $"获取资源列表失败: {ex.Message}")
            End Try
        End Function

        Private Function HandleResourcesRead(id As Object, params As JToken) As String
            Try
                Dim readParams As ResourcesReadParams = Nothing
                If params IsNot Nothing Then
                    readParams = params.ToObject(Of ResourcesReadParams)()
                End If
                If readParams Is Nothing OrElse String.IsNullOrEmpty(readParams.Uri) Then
                    Return BuildErrorResponse(id, McpConstants.INVALID_PARAMS, "缺少资源 URI")
                End If

                Dim result = New ResourcesReadResult()

                Select Case readParams.Uri
                    Case "workbook://metadata"
                        Dim excelApp = Globals.ThisAddIn.Application
                        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
                        Try
                            If excelApp IsNot Nothing AndAlso excelApp.ActiveWorkbook IsNot Nothing Then
                                wb = excelApp.ActiveWorkbook
                                Dim meta = New Dictionary(Of String, Object) From {
                                    {"name", If(wb.Name, "")},
                                    {"fullName", If(wb.FullName, "")},
                                    {"path", If(wb.Path, "")},
                                    {"sheetCount", wb.Worksheets.Count}
                                }
                                result.Contents.Add(New McpTextContent With {
                                    .Text = JsonConvert.SerializeObject(meta, Formatting.Indented)
                                })
                            End If
                        Finally
                            Utils.ComReleaseHelper.Release(wb)
                        End Try

                    Case Else
                        ' 尝试解析 URI 模式
                        If readParams.Uri.StartsWith("workbook://sheets/") Then
                            Dim remaining = readParams.Uri.Substring("workbook://sheets/".Length)
                            Dim parts = remaining.Split("/"c)
                            Dim sheetName = Uri.UnescapeDataString(parts(0))

                            ' 读取工作表数据
                            Dim excelApp2 = Globals.ThisAddIn.Application
                            Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                            Dim rng As Object = Nothing
                            Try
                                ws = TryCast(excelApp2.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                                If ws IsNot Nothing Then
                                    rng = ws.UsedRange
                                    Dim data = DirectCast(rng.Value2, Object(,))
                                    Dim dataStr = JsonConvert.SerializeObject(data, Formatting.Indented)
                                    result.Contents.Add(New McpTextContent With {.Text = dataStr})
                                End If
                            Finally
                                Utils.ComReleaseHelper.Release(rng)
                                Utils.ComReleaseHelper.Release(ws)
                            End Try
                        End If
                End Select

                Return BuildSuccessResponse(id, JToken.FromObject(result))

            Catch ex As Exception
                Return BuildErrorResponse(id, McpConstants.INTERNAL_ERROR, $"读取资源失败: {ex.Message}")
            End Try
        End Function

        Private Function HandleResourcesTemplatesList(id As Object) As String
            Try
                Dim templates = _excelTools.GetResourceTemplateDefinitions()
                Dim result = New With {.resourceTemplates = templates}
                Return BuildSuccessResponse(id, JToken.FromObject(result))
            Catch ex As Exception
                Return BuildErrorResponse(id, McpConstants.INTERNAL_ERROR, $"获取资源模板失败: {ex.Message}")
            End Try
        End Function

        Private Function HandlePromptsList(id As Object) As String
            ' 目前不提供 Prompts 功能
            Dim result = New PromptsListResult()
            Return BuildSuccessResponse(id, JToken.FromObject(result))
        End Function

        '#End Region

        '#Region "响应构建"

        ''' <summary>
        ''' 构建成功响应
        ''' </summary>
        Private Function BuildSuccessResponse(id As Object, result As JToken) As String
            Dim response = New JsonRpcResponse With {
                .Id = id,
                .Result = result
            }
            Return JsonConvert.SerializeObject(response, New JsonSerializerSettings With {
                .NullValueHandling = NullValueHandling.Ignore
            })
        End Function

        ''' <summary>
        ''' 构建错误响应
        ''' </summary>
        Private Function BuildErrorResponse(id As Object, code As Integer, message As String) As String
            Dim errorResponse = New JsonRpcError With {
                .Id = id,
                .[Error] = New JsonRpcErrorDetail With {
                    .Code = code,
                    .Message = message
                }
            }
            Return JsonConvert.SerializeObject(errorResponse, New JsonSerializerSettings With {
                .NullValueHandling = NullValueHandling.Ignore
            })
        End Function

        '#End Region

        '#Region "获取工具数量"

        Public Function GetToolCount() As Integer
            Dim excelCount = If(_excelTools IsNot Nothing, ExcelMcpTools.GetToolDefinitions().Count, 0)
            Dim pptCount = If(_pptTools IsNot Nothing, PptMcpTools.GetToolDefinitions().Count, 0)
            Dim computerCount = ComputerAgentTools.GetToolDefinitions().Count
            Dim webCount = WebSearchTools.GetToolDefinitions().Count
            Dim timeCount = TimeTools.GetToolDefinitions().Count
            Return excelCount + pptCount + computerCount + webCount + timeCount
        End Function

        '#End Region

        Public Sub Dispose() Implements IDisposable.Dispose
            ' 【修复】完善Dispose方法，确保正确释放所有资源
            If _isDisposed Then Return
            _isDisposed = True

            If _isRunning Then
                _isRunning = False

                ' 停止TCP监听器（如果存在）
                ' 注意：McpServer本身不直接持有TCP监听器，但需要确保状态正确重置
                Try
                    Debug.WriteLine("[McpServer] 正在释放资源...")
                    
                    ' 释放Excel工具资源
                    If _excelTools IsNot Nothing Then
                        Try
                            ' 如果ExcelTools实现了IDisposable，则释放它
                            Dim disposableTool = TryCast(_excelTools, IDisposable)
                            If disposableTool IsNot Nothing Then
                                disposableTool.Dispose()
                                Debug.WriteLine("[McpServer] ExcelTools 已释放")
                            End If
                        Catch ex As Exception
                            Debug.WriteLine($"[McpServer] 释放ExcelTools异常: {ex.Message}")
                        End Try
                        _excelTools = Nothing
                    End If

                    ' 释放PPT工具资源
                    If _pptTools IsNot Nothing Then
                        Try
                            Dim disposablePpt = TryCast(_pptTools, IDisposable)
                            If disposablePpt IsNot Nothing Then
                                disposablePpt.Dispose()
                                Debug.WriteLine("[McpServer] PptTools 已释放")
                            End If
                        Catch ex As Exception
                            Debug.WriteLine($"[McpServer] 释放PptTools异常: {ex.Message}")
                        End Try
                        _pptTools = Nothing
                    End If
                    
                    Debug.WriteLine("[McpServer] 资源释放完成")
                Catch ex As Exception
                    Debug.WriteLine($"[McpServer] Dispose过程异常: {ex.Message}")
                End Try
            End If

            _isInitialized = False
            _serverInfo = Nothing
            _serverCapabilities = Nothing
        End Sub

    End Class

    '#End Region

End Namespace

