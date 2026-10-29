Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Threading.Tasks

Namespace Services.Mcp

    '#Region "MCP Client"

    ''' <summary>
    ''' MCP Client - 连接到外部 MCP Server 并调用其工具
    ''' </summary>
    Public Class McpClient
        Implements IDisposable

        Private _config As McpClientConfig
        Private _protocol As McpProtocolHandler
        Private _transport As IMcpTransport
        Private _isConnected As Boolean = False
        Private _cachedTools As List(Of McpToolInfo)
        Private _errorMessage As String

        ''' <summary>
        ''' Client 配置
        ''' </summary>
        Public ReadOnly Property Config As McpClientConfig
            Get
                Return _config
            End Get
        End Property

        ''' <summary>
        ''' 是否已连接
        ''' </summary>
        Public ReadOnly Property IsConnected As Boolean
            Get
                Return _isConnected
            End Get
        End Property

        ''' <summary>
        ''' 错误消息
        ''' </summary>
        Public ReadOnly Property ErrorMessage As String
            Get
                Return _errorMessage
            End Get
        End Property

        ''' <summary>
        ''' 协议处理器
        ''' </summary>
        Public ReadOnly Property Protocol As McpProtocolHandler
            Get
                Return _protocol
            End Get
        End Property

        Public Sub New(config As McpClientConfig)
            _config = config
            _cachedTools = New List(Of McpToolInfo)()
        End Sub

        '#Region "连接管理"

        ''' <summary>
        ''' 连接到外部 MCP Server
        ''' </summary>
        Public Async Function ConnectAsync() As Task(Of Boolean)
            If _isConnected Then Return True

            Try
                ' 根据传输类型创建传输层
                If _config.Transport = "stdio" Then
                    If String.IsNullOrEmpty(_config.Command) Then
                        Throw New Exception("stdio 传输需要 command 配置")
                    End If
                    _transport = New StdioTransport(_config.Command, _config.Args, Nothing, _config.Env)
                Else
                    If String.IsNullOrEmpty(_config.Url) Then
                        Throw New Exception("HTTP 传输需要 url 配置")
                    End If
                    _transport = New HttpSseTransport(_config.Url, _config.Headers)
                End If

                ' 创建协议处理器并初始化
                Dim clientInfo = New ClientInfo With {
                    .Name = $"{McpConstants.IMPLEMENTATION_NAME}-Client",
                    .Version = McpConstants.IMPLEMENTATION_VERSION
                }

                _protocol = New McpProtocolHandler(_transport, clientInfo)
                Dim initResult = Await _protocol.InitializeAsync()
                _isConnected = True
                _errorMessage = Nothing

                Dim srvName As String = "Unknown"
                Dim srvVer As String = "Unknown"
                If initResult.ServerInfo IsNot Nothing Then
                    srvName = If(initResult.ServerInfo.Name, "Unknown")
                    srvVer = If(initResult.ServerInfo.Version, "Unknown")
                End If
                Debug.WriteLine($"MCP Client [{_config.Name}]: 连接成功, Server: {srvName} v{srvVer}")

                ' 缓存工具列表
                Await RefreshToolsAsync()

                Return True

            Catch ex As Exception
                _isConnected = False
                _errorMessage = ex.Message
                Debug.WriteLine($"MCP Client [{_config.Name}]: 连接失败 - {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 断开连接
        ''' </summary>
        Public Async Function DisconnectAsync() As Task
            If Not _isConnected Then Return

            Try
                If _protocol IsNot Nothing Then
                    Await _protocol.ShutdownAsync()
                End If
            Catch ex As Exception
                Debug.WriteLine($"MCP Client [{_config.Name}]: 断开连接异常 - {ex.Message}")
            End Try

            _isConnected = False
            _cachedTools.Clear()
        End Function

        ''' <summary>
        ''' 重新连接
        ''' </summary>
        Public Async Function ReconnectAsync() As Task(Of Boolean)
            Await DisconnectAsync()
            Await Task.Delay(1000) ' 等待 1 秒
            Return Await ConnectAsync()
        End Function

        '#End Region

        '#Region "工具操作"

        ''' <summary>
        ''' 刷新缓存的工具列表
        ''' </summary>
        Public Async Function RefreshToolsAsync() As Task
            If Not _isConnected Then Return

            Try
                _cachedTools = Await _protocol.GetSimplifiedToolsAsync()
                Debug.WriteLine($"MCP Client [{_config.Name}]: 获取到 {_cachedTools.Count} 个工具")
                Try
                    Dim logDir = IO.Path.Combine(IO.Path.GetTempPath(), "ExcelMCP")
                    IO.File.AppendAllText(IO.Path.Combine(logDir, "debug.log"), $"[{DateTime.Now:HH:mm:ss.fff}] [{_config.Name}] RefreshToolsAsync: 缓存了 {_cachedTools.Count} 个工具{vbCrLf}")
                Catch : End Try
            Catch ex As Exception
                Debug.WriteLine($"MCP Client [{_config.Name}]: 刷新工具列表失败 - {ex.Message}")
                Try
                    Dim logDir = IO.Path.Combine(IO.Path.GetTempPath(), "ExcelMCP")
                    IO.File.AppendAllText(IO.Path.Combine(logDir, "debug.log"), $"[{DateTime.Now:HH:mm:ss.fff}] [{_config.Name}] RefreshToolsAsync ERROR: {ex.Message}{vbCrLf}")
                Catch : End Try
            End Try
        End Function

        ''' <summary>
        ''' 获取缓存的工具列表
        ''' </summary>
        Public Function GetTools() As List(Of McpToolInfo)
            Return New List(Of McpToolInfo)(_cachedTools)
        End Function

        ''' <summary>
        ''' 调用外部 MCP 工具
        ''' </summary>
        ''' <param name="toolName">工具名称</param>
        ''' <param name="arguments">工具参数（JSON 格式）</param>
        ''' <returns>工具执行结果</returns>
        Public Async Function CallToolAsync(toolName As String, arguments As Dictionary(Of String, Object)) As Task(Of ToolsCallResult)
            ' 惰性连接：没连就先连上（首次调用自动触发）
            If Not _isConnected Then
                Try
                    System.Diagnostics.Debug.WriteLine($"[McpClient] [{_config.Name}]: 惰性连接中...")
                    Await ConnectAsync()
                    System.Diagnostics.Debug.WriteLine($"[McpClient] [{_config.Name}]: 惰性连接成功")
                Catch ex As Exception
                    Dim errorResult = New ToolsCallResult() With {.IsError = True}
                    errorResult.Content.Add(New McpTextContent With {.Text = $"MCP Client 连接失败: {ex.Message}"})
                    Return errorResult
                End Try
                If Not _isConnected Then
                    Dim errorResult = New ToolsCallResult() With {.IsError = True}
                    errorResult.Content.Add(New McpTextContent With {.Text = "MCP Client 无法连接"})
                    Return errorResult
                End If
            End If

            Try
                Return Await _protocol.CallToolAsync(toolName, arguments)
            Catch ex As Exception
                Dim errorResult = New ToolsCallResult() With {.IsError = True}
                errorResult.Content.Add(New McpTextContent With {.Text = $"工具调用异常: {ex.Message}"})
                Return errorResult
            End Try
        End Function

        ''' <summary>
        ''' Ping 测试连接
        ''' </summary>
        Public Async Function PingAsync() As Task(Of Boolean)
            If Not _isConnected OrElse _protocol Is Nothing Then Return False
            Return Await _protocol.PingAsync()
        End Function

        '#End Region

        '#Region "资源操作"

        ''' <summary>
        ''' 获取资源列表
        ''' </summary>
        Public Async Function ListResourcesAsync() As Task(Of ResourcesListResult)
            If Not _isConnected Then Return New ResourcesListResult()
            Return Await _protocol.ListResourcesAsync()
        End Function

        ''' <summary>
        ''' 读取资源
        ''' </summary>
        Public Async Function ReadResourceAsync(uri As String) As Task(Of ResourcesReadResult)
            If Not _isConnected Then Return New ResourcesReadResult()
            Return Await _protocol.ReadResourceAsync(uri)
        End Function

        '#End Region

        '#Region "状态信息"

        Public Function GetStatus() As McpClientStatusInfo
            Return New McpClientStatusInfo With {
                .Name = _config.Name,
                .Description = If(_config.Description, ""),
                .Connected = _isConnected,
                .Transport = _config.Transport,
                .ToolCount = _cachedTools.Count,
                .ErrorMessage = _errorMessage,
                .Command = If(_config.Command, ""),
                .Args = If(_config.Args, New List(Of String)()),
                .Url = If(_config.Url, ""),
                .Headers = If(_config.Headers, New Dictionary(Of String, String)()),
                .Env = If(_config.Env, New Dictionary(Of String, String)()),
                .Enabled = _config.Enabled
            }
        End Function

        '#End Region

        Private _isDisposed As Boolean = False

        Public Sub Dispose() Implements IDisposable.Dispose
            If _isDisposed Then Return
            _isDisposed = True

            ' 直接关管道（preserveProcess=True 保留 python 进程），不走 Shutdown RPC 不 WaitForExit
            ' 之前的 disconnectTask.Wait(500ms) 在 UI 线程卡死导致 Excel 关闭慢
            Try
                If _transport IsNot Nothing Then
                    Dim stdio = TryCast(_transport, Services.Mcp.StdioTransport)
                    If stdio IsNot Nothing Then
                        stdio.Dispose(preserveProcess:=True)
                    Else
                        _transport.Dispose()
                    End If
                    _transport = Nothing
                End If
            Catch ex As Exception
                Debug.WriteLine($"[McpClient] [{_config.Name}]: Dispose transport 异常 - {ex.Message}")
            End Try

            ' 释放协议处理器
            If _protocol IsNot Nothing Then
                Try
                    _protocol.Dispose()
                    _protocol = Nothing
                Catch ex As Exception
                    Debug.WriteLine($"[McpClient] [{_config.Name}]: 释放 Protocol 异常 - {ex.Message}")
                End Try
            End If

            ' 清空缓存工具列表
            If _cachedTools IsNot Nothing Then
                _cachedTools.Clear()
                _cachedTools = Nothing
            End If

            ' 清理配置引用
            _config = Nothing
            _isConnected = False
            
            Debug.WriteLine($"[McpClient] [{_config?.Name}]: Dispose 完成")
        End Sub

    End Class

    '#End Region

End Namespace


