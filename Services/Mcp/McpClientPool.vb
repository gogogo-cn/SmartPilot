Imports System
Imports System.Collections.Concurrent
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks

Namespace Services.Mcp

    '#Region "MCP Client 连接池"

    ''' <summary>
    ''' MCP Client 连接池 - 管理多个外部 MCP Server 连接
    ''' </summary>
    Public Class McpClientPool
        Implements IDisposable

        Private _clients As ConcurrentDictionary(Of String, McpClient)
        Private _lockObj As New Object()
        Private _configManager As McpConfigManager

        ''' <summary>
        ''' 获取所有已连接的 Client
        ''' </summary>
        Public ReadOnly Property Clients As IEnumerable(Of McpClient)
            Get
                Return _clients.Values
            End Get
        End Property

        ''' <summary>
        ''' 已连接的 Client 数量
        ''' </summary>
        Public ReadOnly Property Count As Integer
            Get
                Return _clients.Count
            End Get
        End Property

        Public Sub New(configManager As McpConfigManager)
            _configManager = configManager
            _clients = New ConcurrentDictionary(Of String, McpClient)(StringComparer.OrdinalIgnoreCase)
        End Sub

        '#Region "Client 管理"

        ''' <summary>
        ''' 添加并连接 MCP Client
        ''' </summary>
        Public Async Function AddClientAsync(config As McpClientConfig) As Task(Of Boolean)
            If _clients.ContainsKey(config.Name) Then
                ' 如果已存在，先断开并移除
                Await RemoveClientAsync(config.Name)
            End If

            Dim client = New McpClient(config)
            Dim connected = Await client.ConnectAsync()

            ' 无论是否连接成功，都加入连接池以在 UI 中显示状态
            _clients.TryAdd(config.Name, client)
            If connected Then
                Debug.WriteLine($"MCP ClientPool: {config.Name} 已添加并连接")
            Else
                Debug.WriteLine($"MCP ClientPool: {config.Name} 连接失败 - {client.ErrorMessage}")
            End If

            Return connected
        End Function

        ''' <summary>
        ''' 仅将 Client 加入连接池（不尝试连接，用于自动加载场景）
        ''' </summary>
        Public Sub AddClientToPoolOnly(config As McpClientConfig)
            If _clients.ContainsKey(config.Name) Then Return
            Dim client = New McpClient(config)
            _clients.TryAdd(config.Name, client)
        End Sub

        ''' <summary>
        ''' 移除 MCP Client
        ''' </summary>
        Public Async Function RemoveClientAsync(name As String) As Task(Of Boolean)
            Dim client As McpClient = Nothing
            If _clients.TryRemove(name, client) Then
                Await client.DisconnectAsync()
                client.Dispose()
                Debug.WriteLine($"MCP ClientPool: {name} 已移除")
                Return True
            End If
            Return False
        End Function

        ''' <summary>
        ''' 断开 MCP Client（保留在池中，仅断开连接，不移除）
        ''' </summary>
        Public Async Function DisconnectClientAsync(name As String) As Task(Of Boolean)
            Dim client = GetClient(name)
            If client IsNot Nothing Then
                Await client.DisconnectAsync()
                Debug.WriteLine($"MCP ClientPool: {name} 已断开（保留在池中）")
                Return True
            End If
            Return False
        End Function

        ''' <summary>
        ''' 根据名称获取 Client
        ''' </summary>
        Public Function GetClient(name As String) As McpClient
            Dim client As McpClient = Nothing
            _clients.TryGetValue(name, client)
            Return client
        End Function

        ''' <summary>
        ''' 检查 Client 是否已连接
        ''' </summary>
        Public Function IsConnected(name As String) As Boolean
            Dim client = GetClient(name)
            Return client IsNot Nothing AndAlso client.IsConnected
        End Function

        '#End Region

        '#Region "工具操作"

        ''' <summary>
        ''' 调用指定 Client 的工具
        ''' </summary>
        Public Async Function CallToolAsync(clientName As String, toolName As String, arguments As Dictionary(Of String, Object)) As Task(Of ToolsCallResult)
            Dim client = GetClient(clientName)
            If client Is Nothing Then
                Dim errorResult = New ToolsCallResult() With {.IsError = True}
                errorResult.Content.Add(New McpTextContent With {.Text = $"MCP Client '{clientName}' 不存在"})
                Return errorResult
            End If
            Return Await client.CallToolAsync(toolName, arguments)
        End Function

        ''' <summary>
        ''' 获取指定 Client 的工具列表
        ''' </summary>
        Public Function GetClientTools(clientName As String) As List(Of McpToolInfo)
            Dim client = GetClient(clientName)
            If client IsNot Nothing Then
                Return client.GetTools()
            End If
            Return New List(Of McpToolInfo)()
        End Function

        ''' <summary>
        ''' 获取所有可用工具（从所有已连接的 Client）
        ''' </summary>
        Public Function GetAllAvailableTools() As List(Of McpToolInfo)
            Dim allTools = New List(Of McpToolInfo)()
            For Each client In _clients.Values
                If client.IsConnected Then
                    Dim tools = client.GetTools()
                    For Each tool In tools
                        tool.Source = client.Config.Name
                    Next
                    allTools.AddRange(tools)
                End If
            Next
            Return allTools
        End Function

        '#End Region

        '#Region "启动/停止"

        ''' <summary>
        ''' 自动连接所有配置中启用的 Server
        ''' </summary>
        Public Async Function AutoConnectAllAsync() As Task
            Debug.WriteLine($"MCP ClientPool: AutoConnectAllAsync 开始, 已有 {_clients.Count} 个client")
            For Each kvp In _configManager.GetConfig().McpServers
                Debug.WriteLine($"MCP ClientPool: 检查 {kvp.Key}, 在池中={_clients.ContainsKey(kvp.Key)}, 连接状态={If(_clients.ContainsKey(kvp.Key), _clients(kvp.Key).IsConnected.ToString(), "N/A")}")
                ' 如果已存在但未连接，先移除再重新添加（触发连接）
                If _clients.ContainsKey(kvp.Key) Then
                    Dim existingClient = _clients(kvp.Key)
                    If Not existingClient.IsConnected Then
                        Debug.WriteLine($"MCP ClientPool: {kvp.Key} 未连接，移除并重新连接")
                        _clients.TryRemove(kvp.Key, Nothing)
                    Else
                        Debug.WriteLine($"MCP ClientPool: {kvp.Key} 已连接，跳过")
                    End If
                End If
                
                If Not _clients.ContainsKey(kvp.Key) Then
                    Debug.WriteLine($"MCP ClientPool: 正在连接 {kvp.Key}...")
                    Dim connected = Await AddClientAsync(kvp.Value)
                    If connected Then
                        Debug.WriteLine($"MCP ClientPool: {kvp.Key} 连接成功")
                    Else
                        Debug.WriteLine($"MCP ClientPool: {kvp.Key} 连接失败")
                    End If
                End If
            Next
        End Function

        ''' <summary>
        ''' 断开所有连接
        ''' </summary>
        Public Async Function DisconnectAllAsync() As Task
            Dim names = _clients.Keys.ToArray()
            For Each name In names
                Await RemoveClientAsync(name)
            Next
        End Function

        ''' <summary>
        ''' 对所有 Client 执行 Ping 检查
        ''' </summary>
        Public Async Function PingAllAsync() As Task(Of Dictionary(Of String, Boolean))
            Dim results = New Dictionary(Of String, Boolean)()
            For Each kvp In _clients
                results(kvp.Key) = Await kvp.Value.PingAsync()
            Next
            Return results
        End Function

        '#End Region

        '#Region "状态"

        ''' <summary>
        ''' 获取所有 Client 的状态信息
        ''' </summary>
        Public Function GetAllStatus() As List(Of McpClientStatusInfo)
            Dim statusList = New List(Of McpClientStatusInfo)()
            For Each client In _clients.Values
                statusList.Add(client.GetStatus())
            Next
            Return statusList
        End Function

        '#End Region

        Public Sub Dispose() Implements IDisposable.Dispose
            Dim sw = Diagnostics.Stopwatch.StartNew()
            Debug.WriteLine($"MCP ClientPool: Dispose 开始, 共有 {_clients.Count} 个客户端")

            ' 【修复】并行释放所有客户端，避免串行等待导致 Excel 关闭卡死
            ' 每个 McpClient.Dispose 内部有 500ms 超时，串行时 N 个客户端需 N*500ms
            ' 并行后总耗时约等于单个客户端的耗时（~500ms）
            Try
                Dim clients = _clients.Values.ToArray()
                Parallel.ForEach(clients, Sub(client)
                    Try
                        client.Dispose()
                    Catch ex As Exception
                        Debug.WriteLine($"MCP ClientPool: 释放 {client.Config.Name} 异常 - {ex.Message}")
                    End Try
                End Sub)
            Catch ex As Exception
                Debug.WriteLine($"MCP ClientPool: 并行释放异常 - {ex.Message}")
            End Try

            _clients.Clear()
            Debug.WriteLine($"MCP ClientPool: Dispose 完成, 耗时 {sw.ElapsedMilliseconds}ms")
        End Sub

    End Class

    '#End Region

End Namespace
