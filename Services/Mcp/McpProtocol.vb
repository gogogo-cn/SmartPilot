Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services.Mcp

    '#Region "MCP 协议处理器"

    ''' <summary>
    ''' MCP 协议处理器 - 封装传输层并实现 MCP 协议生命周期
    ''' 负责 initialize、ping、tools/list、tools/call 等协议方法
    ''' </summary>
    Public Class McpProtocolHandler
        Implements IDisposable

        Private _transport As IMcpTransport
        Private _serverInfo As ServerInfo
        Private _serverCapabilities As ServerCapabilities
        Private _clientInfo As ClientInfo
        Private _clientCapabilities As ClientCapabilities
        Private _protocolVersion As String
        Private _isInitialized As Boolean = False
        Private _nextRequestId As Integer = 1
        Private _lockObj As New Object()

        ''' <summary>
        ''' 服务端信息（此客户端连接到的 Server）
        ''' </summary>
        Public ReadOnly Property ServerInfo As ServerInfo
            Get
                Return _serverInfo
            End Get
        End Property

        ''' <summary>
        ''' 服务端能力
        ''' </summary>
        Public ReadOnly Property ServerCapabilities As ServerCapabilities
            Get
                Return _serverCapabilities
            End Get
        End Property

        ''' <summary>
        ''' 是否已完成初始化
        ''' </summary>
        Public ReadOnly Property IsInitialized As Boolean
            Get
                Return _isInitialized
            End Get
        End Property

        ''' <summary>
        ''' 协议版本
        ''' </summary>
        Public ReadOnly Property ProtocolVersion As String
            Get
                Return _protocolVersion
            End Get
        End Property

        Public Sub New(transport As IMcpTransport, Optional clientInfo As ClientInfo = Nothing)
            _transport = transport
            _clientInfo = If(clientInfo, New ClientInfo With {
                .Name = McpConstants.IMPLEMENTATION_NAME,
                .Version = McpConstants.IMPLEMENTATION_VERSION
            })
            _clientCapabilities = New ClientCapabilities()
        End Sub

        '#Region "协议生命周期"

        ''' <summary>
        ''' 执行 MCP 协议握手：连接 + Initialize + Initialized 通知
        ''' </summary>
        Public Async Function InitializeAsync() As Task(Of InitializeResult)
            ' 1. 建立传输连接
            Await _transport.ConnectAsync()

            ' 2. 发送 Initialize 请求
            Dim initParams = New InitializeParams With {
                .ProtocolVersion = McpConstants.PROTOCOL_VERSION,
                .Capabilities = _clientCapabilities,
                .ClientInfo = _clientInfo
            }

            Dim requestMethod = McpConstants.METHOD_INITIALIZE
            Dim requestParams = JToken.FromObject(initParams)
            Dim result = Await SendRequestAsync(requestMethod, requestParams)

            ' 解析 Initialize 结果
            Dim capsToken = result("capabilities")
            If capsToken IsNot Nothing Then
                _serverCapabilities = capsToken.ToObject(Of ServerCapabilities)()
            End If
            Dim serverInfoToken = result("serverInfo")
            If serverInfoToken IsNot Nothing Then
                _serverInfo = serverInfoToken.ToObject(Of ServerInfo)()
            End If
            Dim protoVerToken = result("protocolVersion")
            If protoVerToken IsNot Nothing Then
                _protocolVersion = protoVerToken.ToString()
            End If

            Dim instructionsToken = result("instructions")
            Dim instructions As String = Nothing
            If instructionsToken IsNot Nothing Then
                instructions = instructionsToken.ToString()
            End If

            Dim initResult = New InitializeResult With {
                .ProtocolVersion = _protocolVersion,
                .Capabilities = _serverCapabilities,
                .ServerInfo = _serverInfo,
                .Instructions = instructions
            }

            ' 3. 发送 Initialized 通知
            Await SendNotificationAsync(McpConstants.METHOD_INITIALIZED, Nothing)
            _isInitialized = True

            Return initResult
        End Function

        ''' <summary>
        ''' 断开协议连接
        ''' </summary>
        Public Async Function ShutdownAsync() As Task
            _isInitialized = False
            Await _transport.DisconnectAsync()
        End Function

        '#End Region

        '#Region "MCP 标准方法"

        ''' <summary>
        ''' Ping - 检查连接是否存活
        ''' </summary>
        Public Async Function PingAsync() As Task(Of Boolean)
            Try
                Await SendRequestAsync(McpConstants.METHOD_PING, Nothing)
                Return True
            Catch ex As Exception  ' 
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Tools/List - 获取可用工具列表
        ''' </summary>
        Public Async Function ListToolsAsync(Optional cursor As String = Nothing) As Task(Of ToolsListResult)
            Dim params As JToken = Nothing
            If Not String.IsNullOrEmpty(cursor) Then
                params = JToken.FromObject(New ToolsListParams With {.Cursor = cursor})
            End If
            Dim result = Await SendRequestAsync(McpConstants.METHOD_TOOLS_LIST, params)
            Return result.ToObject(Of ToolsListResult)()
        End Function

        ''' <summary>
        ''' Tools/Call - 调用指定工具
        ''' </summary>
        Public Async Function CallToolAsync(toolName As String, arguments As Dictionary(Of String, Object)) As Task(Of ToolsCallResult)
            Try
                ' 手动构建 JSON 参数，避免 JToken.FromObject 对 Object Dictionary 的序列化问题
                Dim argsObj = New JObject()
                If arguments IsNot Nothing Then
                    For Each kvp In arguments
                        argsObj(kvp.Key) = If(kvp.Value IsNot Nothing, JToken.FromObject(kvp.Value), JValue.CreateNull())
                    Next
                End If
                Dim callParams = New JObject From {
                    {"name", toolName},
                    {"arguments", argsObj}
                }
                Dim result = Await SendRequestAsync(McpConstants.METHOD_TOOLS_CALL, callParams)
                Return result.ToObject(Of ToolsCallResult)()
            Catch ex As Exception
                Dim errorResult = New ToolsCallResult() With {.IsError = True}
                errorResult.Content.Add(New McpTextContent With {.Text = $"工具调用失败: {ex.Message}"})
                Return errorResult
            End Try
        End Function

        ''' <summary>
        ''' Resources/List - 获取资源列表
        ''' </summary>
        Public Async Function ListResourcesAsync(Optional cursor As String = Nothing) As Task(Of ResourcesListResult)
            Dim params As JToken = Nothing
            If Not String.IsNullOrEmpty(cursor) Then
                params = JToken.FromObject(New With {.cursor = cursor})
            End If
            Dim result = Await SendRequestAsync(McpConstants.METHOD_RESOURCES_LIST, params)
            Return result.ToObject(Of ResourcesListResult)()
        End Function

        ''' <summary>
        ''' Resources/Read - 读取资源内容
        ''' </summary>
        Public Async Function ReadResourceAsync(uri As String) As Task(Of ResourcesReadResult)
            Dim readParams = New ResourcesReadParams With {.Uri = uri}
            Dim result = Await SendRequestAsync(McpConstants.METHOD_RESOURCES_READ, JToken.FromObject(readParams))
            Return result.ToObject(Of ResourcesReadResult)()
        End Function

        ''' <summary>
        ''' Resources/Templates/List - 获取资源模板列表
        ''' </summary>
        Public Async Function ListResourceTemplatesAsync() As Task(Of List(Of McpResourceTemplate))
            Dim result = Await SendRequestAsync(McpConstants.METHOD_RESOURCES_TEMPLATES_LIST, Nothing)
            If result IsNot Nothing AndAlso result("resourceTemplates") IsNot Nothing Then
                Return result("resourceTemplates").ToObject(Of List(Of McpResourceTemplate))()
            End If
            Return New List(Of McpResourceTemplate)()
        End Function

        ''' <summary>
        ''' Prompts/List - 获取提示模板列表
        ''' </summary>
        Public Async Function ListPromptsAsync(Optional cursor As String = Nothing) As Task(Of PromptsListResult)
            Dim params As JToken = Nothing
            If Not String.IsNullOrEmpty(cursor) Then
                params = JToken.FromObject(New With {.cursor = cursor})
            End If
            Dim result = Await SendRequestAsync(McpConstants.METHOD_PROMPTS_LIST, params)
            Return result.ToObject(Of PromptsListResult)()
        End Function

        ''' <summary>
        ''' Prompts/Get - 获取指定提示模板
        ''' </summary>
        Public Async Function GetPromptAsync(promptName As String, Optional arguments As Dictionary(Of String, String) = Nothing) As Task(Of Dictionary(Of String, Object))
            Dim promptParams = New Dictionary(Of String, Object) From {
                {"name", promptName}
            }
            If arguments IsNot Nothing Then
                promptParams("arguments") = arguments
            End If
            Dim result = Await SendRequestAsync(McpConstants.METHOD_PROMPTS_GET, JToken.FromObject(promptParams))
            Return result.ToObject(Of Dictionary(Of String, Object))()
        End Function

        '#End Region

        '#Region "请求/响应核心"

        ''' <summary>
        ''' 发送 JSON-RPC 请求并等待响应
        ''' </summary>
        Private Async Function SendRequestAsync(method As String, params As JToken) As Task(Of JToken)
            Dim requestId = GetNextId()
            Dim request = New JsonRpcRequest With {
                .Id = requestId,
                .Method = method,
                .Params = params
            }
            Dim requestJson = JsonConvert.SerializeObject(request, New JsonSerializerSettings With {
                .NullValueHandling = NullValueHandling.Ignore
            })

            Debug.WriteLine($"MCP SendRequest: method={method}, request={requestJson}")
            Try
                Dim logDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ExcelMCP")
                If Not System.IO.Directory.Exists(logDir) Then System.IO.Directory.CreateDirectory(logDir)
                System.IO.File.AppendAllText(System.IO.Path.Combine(logDir, "debug.log"), $"[{DateTime.Now:HH:mm:ss.fff}] SEND method={method} body={requestJson}{vbCrLf}")
            Catch : End Try
            Dim responseJson = Await _transport.SendAsync(requestJson)
            Debug.WriteLine($"MCP SendRequest: response raw={(If(responseJson, "NULL"))}")
            Try
                Dim logDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ExcelMCP")
                System.IO.File.AppendAllText(System.IO.Path.Combine(logDir, "debug.log"), $"[{DateTime.Now:HH:mm:ss.fff}] RECV method={method} body={responseJson}{vbCrLf}")
            Catch : End Try
            Dim response = JToken.Parse(responseJson)

            ' 检查是否为错误响应
            If response("error") IsNot Nothing Then
                Dim errorDetail = response("error").ToObject(Of JsonRpcErrorDetail)()
                Throw New Exception($"MCP 协议错误 [{errorDetail.Code}]: {errorDetail.Message}")
            End If

            Dim resultToken = response("result")
            If resultToken Is Nothing Then
                Debug.WriteLine($"MCP SendRequest: WARNING - response has no 'result' field! Response: {responseJson}")
            End If
            Return resultToken
        End Function

        ''' <summary>
        ''' 发送 JSON-RPC 通知（无响应）
        ''' </summary>
        Private Async Function SendNotificationAsync(method As String, params As JToken) As Task
            Dim notification = New JsonRpcNotification With {
                .Method = method,
                .Params = params
            }
            Dim notificationJson = JsonConvert.SerializeObject(notification, New JsonSerializerSettings With {
                .NullValueHandling = NullValueHandling.Ignore
            })
            Await _transport.SendNotificationAsync(notificationJson)
        End Function

        ''' <summary>
        ''' 生成递增请求 ID
        ''' </summary>
        Private Function GetNextId() As Integer
            SyncLock _lockObj
                Dim id = _nextRequestId
                _nextRequestId += 1
                Return id
            End SyncLock
        End Function

        '#End Region

        '#Region "便捷方法"

        ''' <summary>
        ''' 创建传输并初始化协议处理器
        ''' </summary>
        Public Shared Async Function CreateAndInitializeAsync(
            transport As IMcpTransport,
            Optional clientInfo As ClientInfo = Nothing) As Task(Of McpProtocolHandler)

            Dim handler = New McpProtocolHandler(transport, clientInfo)
            Await handler.InitializeAsync()
            Return handler
        End Function

        ''' <summary>
        ''' 获取所有工具的简化列表（供前端使用）
        ''' </summary>
        Public Async Function GetSimplifiedToolsAsync() As Task(Of List(Of McpToolInfo))
            Dim result = New List(Of McpToolInfo)()
            Try
                Debug.WriteLine("MCP Protocol: 调用 tools/list...")
                Dim toolsResult = Await ListToolsAsync()
                Dim toolCount = If(toolsResult IsNot Nothing AndAlso toolsResult.Tools IsNot Nothing, toolsResult.Tools.Count, 0)
                Debug.WriteLine($"MCP Protocol: tools/list 返回 {toolCount} 个工具")
                Try
                    Dim logDir = IO.Path.Combine(IO.Path.GetTempPath(), "ExcelMCP")
                    IO.File.AppendAllText(IO.Path.Combine(logDir, "debug.log"), $"[{DateTime.Now:HH:mm:ss.fff}] tools/list returned {toolCount} tools{IIf(toolCount = 0, " (empty!)", "")}{vbCrLf}")
                Catch : End Try
                If toolsResult IsNot Nothing AndAlso toolsResult.Tools IsNot Nothing Then
                    For Each tool In toolsResult.Tools
                        Dim toolInfo = New McpToolInfo With {
                            .Name = tool.Name,
                            .Description = tool.Description
                        }

                        If tool.InputSchema IsNot Nothing AndAlso tool.InputSchema.Properties IsNot Nothing Then
                            Dim requiredSet As New HashSet(Of String)()
                            If tool.InputSchema.Required IsNot Nothing Then
                                For Each req In tool.InputSchema.Required
                                    requiredSet.Add(req)
                                Next
                            End If

                            For Each kvp In tool.InputSchema.Properties
                                toolInfo.Parameters.Add(New McpParameterInfo With {
                                    .Name = kvp.Key,
                                    .Type = kvp.Value.Type,
                                    .Description = kvp.Value.Description,
                                    .Required = requiredSet.Contains(kvp.Key)
                                })
                            Next
                        End If

                        result.Add(toolInfo)
                    Next
                End If
            Catch ex As Exception
                Debug.WriteLine($"获取简化工具列表失败: {ex.GetType().Name}: {ex.Message}")
                Try
                    Dim logDir = IO.Path.Combine(IO.Path.GetTempPath(), "ExcelMCP")
                    IO.File.AppendAllText(IO.Path.Combine(logDir, "debug.log"), $"[{DateTime.Now:HH:mm:ss.fff}] tools/list PARSE ERROR: {ex.GetType().Name}: {ex.Message}{vbCrLf}{ex.StackTrace}{vbCrLf}")
                Catch : End Try
            End Try
            Debug.WriteLine($"MCP Protocol: GetSimplifiedToolsAsync 返回 {result.Count} 个工具")
            Return result
        End Function

        '#End Region

        '#Region "IDisposable"

        Public Sub Dispose() Implements IDisposable.Dispose
            ' 不调 ShutdownAsync —— 会触发 WaitForExit + KillProcessTree 卡死
            ' McpClient.Dispose 已经直接 preserveProcess=True Dispose transport 零等待
            Try : _isInitialized = False : Catch : End Try
            _transport = Nothing
        End Sub

        '#End Region

    End Class

    '#End Region

End Namespace
