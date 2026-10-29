Imports System
Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Diagnostics
Imports System.Linq

Namespace Services.Mcp

    '#Region "MCP Server Host"

    ''' <summary>
    ''' MCP Server 主机 - 管理 Server 生命周期
    ''' 支持 TCP 监听（用于本地或远程 MCP 客户端连接）
    ''' </summary>
    Public Class McpServerHost
        Implements IDisposable

        Private _server As McpServer
        Private _listener As TcpListener
        Private _port As Integer
        Private _isRunning As Boolean = False
        Private _cancellationTokenSource As CancellationTokenSource
        Private _acceptTask As Task
        Private _clientTasks As List(Of Task)

        ''' <summary>
        ''' Server 是否正在运行
        ''' </summary>
        Public ReadOnly Property IsRunning As Boolean
            Get
                Return _isRunning
            End Get
        End Property

        ''' <summary>
        ''' 监听端口
        ''' </summary>
        Public ReadOnly Property Port As Integer
            Get
                Return _port
            End Get
        End Property

        ''' <summary>
        ''' MCP Server 实例
        ''' </summary>
        Public ReadOnly Property Server As McpServer
            Get
                Return _server
            End Get
        End Property

        ''' <summary>
        ''' 已连接的客户端数量
        ''' </summary>
        Public ReadOnly Property ConnectedClientCount As Integer
            Get
                If _clientTasks Is Nothing Then Return 0
                Return _clientTasks.Where(Function(t) Not t.IsCompleted).Count()
            End Get
        End Property

        Public Sub New(bridge As Object, Optional port As Integer = Utils.ConfigConstants.MCP.DEFAULT_SERVER_PORT)
            _server = New McpServer(bridge)
            _port = port
            _clientTasks = New List(Of Task)()
        End Sub

        '#Region "启动/停止"

        ''' <summary>
        ''' 启动 MCP Server（在本地端口上监听）
        ''' </summary>
        Public Sub StartAsync()
            If _isRunning Then Return

            Try
                _cancellationTokenSource = New CancellationTokenSource()

                ' 启动 TCP 监听器
                _listener = New TcpListener(IPAddress.Loopback, _port)
                _listener.Start()
                _isRunning = True

                Debug.WriteLine($"MCP Server 已在端口 {_port} 启动")

                ' 后台接受客户端连接
                _acceptTask = Task.Run(Sub() AcceptClientsLoop(_cancellationTokenSource.Token))

            Catch ex As Exception
                _isRunning = False
                Throw New Exception($"MCP Server 启动失败: {ex.Message}", ex)
            End Try
        End Sub

        ''' <summary>
        ''' 停止 MCP Server
        ''' </summary>
        Public Async Function StopAsync() As Task
            If Not _isRunning Then Return

            Try
                If _cancellationTokenSource IsNot Nothing Then
                    _cancellationTokenSource.Cancel()
                End If

                If _listener IsNot Nothing Then
                    _listener.Stop()
                    _listener = Nothing
                End If

                _isRunning = False
                Debug.WriteLine("MCP Server 已停止")

            Catch ex As Exception
                Debug.WriteLine($"停止 MCP Server 异常: {ex.Message}")
            End Try

            Await Task.CompletedTask
        End Function

        '#End Region

        '#Region "客户端连接处理"

        ''' <summary>
        ''' 接受客户端连接的循环
        ''' </summary>
        Private Async Sub AcceptClientsLoop(token As CancellationToken)
            Try
                While Not token.IsCancellationRequested AndAlso _isRunning
                    Try
                        Dim client = Await _listener.AcceptTcpClientAsync()
                        client.NoDelay = True
                        client.ReceiveTimeout = 60000 ' 60秒超时
                        client.SendTimeout = 60000

                        ' 启动客户端处理任务
                        Dim processTask = Task.Run(Async Function()
                            Await ProcessClientAsync(client, token)
                        End Function, token)

                        SyncLock _clientTasks
                            _clientTasks.Add(processTask)
                            ' 清理已完成的任务
                            _clientTasks.RemoveAll(Function(t) t.IsCompleted)
                        End SyncLock

                    Catch ex As ObjectDisposedException
                        Exit While
                    Catch ex As OperationCanceledException
                        Exit While
                    Catch ex As Exception
                        If Not token.IsCancellationRequested Then
                            Debug.WriteLine($"接受客户端连接异常: {ex.Message}")
                        End If
                    End Try
                End While

            Catch ex As Exception
                Debug.WriteLine($"接受连接循环异常: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' 处理单个客户端连接
        ''' </summary>
        Private Async Function ProcessClientAsync(client As TcpClient, token As CancellationToken) As Task
            Try
                Using client
                    Using stream = client.GetStream()
                        Using reader As New System.IO.StreamReader(stream, Encoding.UTF8)
                            Using writer As New System.IO.StreamWriter(stream, Encoding.UTF8) With {.AutoFlush = True, .NewLine = vbLf}
                                Dim line As String
                                While Not token.IsCancellationRequested
                                    line = Await reader.ReadLineAsync()
                                    If line Is Nothing Then Exit While
                                    If String.IsNullOrWhiteSpace(line) Then Continue While

                                    ' 处理收到的行
                                    Dim response = ProcessLine(line)
                                    If Not String.IsNullOrEmpty(response) Then
                                        Await writer.WriteLineAsync(response)
                                    End If
                                End While
                            End Using
                        End Using
                    End Using
                End Using

            Catch ex As Exception
                If Not token.IsCancellationRequested Then
                    Debug.WriteLine($"处理客户端异常: {ex.Message}")
                End If
            End Try
        End Function

        ''' <summary>
        ''' 处理单行 JSON-RPC 消息
        ''' </summary>
        Private Function ProcessLine(line As String) As String
            Try
                ' 移除 BOM 和多余空白
                line = line.Trim()
                If String.IsNullOrEmpty(line) Then Return Nothing

                ' 交给 McpServer 处理
                Return _server.HandleRequest(line)

            Catch ex As Exception
                Debug.WriteLine($"处理行异常: {ex.Message}, 行: {line.Substring(0, Math.Min(100, line.Length))}...")
                Return Nothing
            End Try
        End Function

        '#End Region

        '#Region "查询状态"

        Public Function GetStatus() As McpStatusInfo
            Return New McpStatusInfo With {
                .ServerRunning = _isRunning,
                .ServerPort = _port,
                .ServerToolCount = _server.GetToolCount()
            }
        End Function

        '#End Region

        Public Sub Dispose() Implements IDisposable.Dispose
            Try
                ' 【修复】使用 Task.Run 避免 .Wait() 死锁，缩短超时时间
                Dim sw = Diagnostics.Stopwatch.StartNew()
                Dim stopTask = Task.Run(Async Function()
                    Await StopAsync()
                End Function)

                If Not stopTask.Wait(TimeSpan.FromMilliseconds(300)) Then
                    Debug.WriteLine($"McpServerHost: StopAsync 超时({sw.ElapsedMilliseconds}ms)，强制跳过")
                Else
                    Debug.WriteLine($"McpServerHost: StopAsync 完成({sw.ElapsedMilliseconds}ms)")
                End If

            Catch ex As Exception
                Debug.WriteLine($"McpServerHost: Dispose 异常 - {ex.Message}")
            End Try

            ' 强制释放资源
            Try
                If _listener IsNot Nothing Then
                    _listener.Stop()
                    _listener = Nothing
                End If
            Catch ex As Exception  ' 
            End Try

            If _cancellationTokenSource IsNot Nothing Then
                Try
                    _cancellationTokenSource.Dispose()
                Catch ex As Exception  ' 
                End Try
                _cancellationTokenSource = Nothing
            End If

            If _server IsNot Nothing Then
                Try
                    _server.Dispose()
                Catch ex As Exception  ' 
                End Try
                _server = Nothing
            End If

            Debug.WriteLine("McpServerHost: Dispose 完成")
        End Sub

    End Class

    '#End Region

End Namespace
