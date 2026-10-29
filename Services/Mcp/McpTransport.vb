Imports System
Imports System.Diagnostics
Imports System.IO
Imports System.Net.Http
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services.Mcp

    '#Region "传输层接口"

    ''' <summary>
    ''' MCP 传输层接口
    ''' </summary>
    Public Interface IMcpTransport
        Inherits IDisposable

        ''' <summary>
        ''' 连接到 MCP Server
        ''' </summary>
        Function ConnectAsync() As Task

        ''' <summary>
        ''' 断开连接
        ''' </summary>
        Function DisconnectAsync() As Task

        ''' <summary>
        ''' 发送消息并等待响应
        ''' </summary>
        Function SendAsync(message As String) As Task(Of String)

        ''' <summary>
        ''' 发送通知（无响应）
        ''' </summary>
        Function SendNotificationAsync(message As String) As Task

        ''' <summary>
        ''' 是否已连接
        ''' </summary>
        ReadOnly Property IsConnected As Boolean

        ''' <summary>
        ''' 传输类型
        ''' </summary>
        ReadOnly Property TransportType As String

        ''' <summary>
        ''' 消息接收事件
        ''' </summary>
        Event MessageReceived As EventHandler(Of String)

        ''' <summary>
        ''' 连接断开事件
        ''' </summary>
        Event Disconnected As EventHandler(Of Exception)

    End Interface

    '#End Region

    '#Region "Stdio 传输实现"

    ''' <summary>
    ''' 基于 stdio 的 MCP 传输（启动子进程通信）
    ''' 用于连接本地 MCP Server（如 filesystem, sqlite 等）
    ''' </summary>
    Public Class StdioTransport
        Implements IMcpTransport

        Private _process As Process
        Private _command As String
        Private _args As List(Of String)
        Private _envVars As Dictionary(Of String, String)
        Private _workingDirectory As String
        Private _reader As StreamReader
        Private _writer As StreamWriter
        Private _isConnected As Boolean = False
        Private _cancellationTokenSource As CancellationTokenSource
        Private _readTask As Task
        Private _pendingRequests As Dictionary(Of Object, TaskCompletionSource(Of String))
        Private _lockObj As New Object()
        ' 进程对象专用锁：DisconnectAsync / Dispose / Exited 事件可能并发访问 Process，
        ' 无锁时 Dispose 与后台断开任务竞态会抛 InvalidOperationException(EnsureState) 导致 Excel 崩溃
        Private _processLock As New Object()
        Private _disposed As Boolean = False
        Private _nextId As Integer = 1

        Public Event MessageReceived As EventHandler(Of String) Implements IMcpTransport.MessageReceived
        Public Event Disconnected As EventHandler(Of Exception) Implements IMcpTransport.Disconnected

        Public ReadOnly Property IsConnected As Boolean Implements IMcpTransport.IsConnected
            Get
                Return _isConnected
            End Get
        End Property

        Public ReadOnly Property TransportType As String Implements IMcpTransport.TransportType
            Get
                Return "stdio"
            End Get
        End Property

        ''' <summary>
        ''' 创建 Stdio 传输
        ''' </summary>
        ''' <param name="command">要执行的命令（如 "npx", "uvx", "python"）</param>
        ''' <param name="args">命令参数</param>
        ''' <param name="workingDirectory">工作目录</param>
        ''' <param name="envVars">环境变量（用于传递认证凭证等）</param>
        Public Sub New(command As String, args As List(Of String), Optional workingDirectory As String = Nothing, Optional envVars As Dictionary(Of String, String) = Nothing)
            _command = command
            _args = args
            _workingDirectory = workingDirectory
            _envVars = envVars
            _pendingRequests = New Dictionary(Of Object, TaskCompletionSource(Of String))()
        End Sub

        ''' <summary>
        ''' 展开字符串中的环境变量引用（支持 ${VAR} 和 %VAR% 两种格式）
        ''' </summary>
        Private Shared Function ExpandEnvVars(input As String) As String
            If String.IsNullOrEmpty(input) Then Return input

            Dim result = input
            ' 展开 ${VAR} 格式（如 ${USERPROFILE}）
            result = System.Text.RegularExpressions.Regex.Replace(result, "\$\{(\w+)\}", Function(m)
                Dim varName = m.Groups(1).Value
                Dim val = Environment.GetEnvironmentVariable(varName)
                Return If(val IsNot Nothing, val, m.Value)
            End Function)

            ' 展开 %VAR% 格式（如 %USERPROFILE%）
            result = System.Text.RegularExpressions.Regex.Replace(result, "%(\w+)%", Function(m)
                Dim varName = m.Groups(1).Value
                Dim val = Environment.GetEnvironmentVariable(varName)
                Return If(val IsNot Nothing, val, m.Value)
            End Function)

            Return result
        End Function

        Public Async Function ConnectAsync() As Task Implements IMcpTransport.ConnectAsync
            If _isConnected Then Return

            Try
                ' 使用 cmd /c 包装命令确保 PATH 解析正确（VSTO 环境下直接调用 npx 可能找不到）
                ' 自动 trim 每个参数，避免尾部空格导致 npm/npx 解析失败
                ' 展开环境变量引用（如 ${USERPROFILE}、%USERPROFILE%）
                Dim trimmedArgs = _args.Select(Function(a) ExpandEnvVars(a.Trim())).ToList()
                Dim fullArgs = String.Join(" ", trimmedArgs.Select(Function(a) If(a.Contains(" "), """" & a & """", a)))
                Dim wrappedCommand = ExpandEnvVars(_command)
                Dim wrappedArgs = fullArgs
                If _command = "npx" OrElse _command = "uvx" OrElse _command = "node" OrElse _command = "python" Then
                    wrappedCommand = "cmd"
                    wrappedArgs = "/c " & _command & " " & fullArgs
                End If

                ' stderr 使用系统默认编码避免乱码（cmd.exe 中文输出是 GBK）
                Dim stderrEncoding As Encoding = Nothing
                Try
                    stderrEncoding = Console.OutputEncoding
                Catch
                    Try : stderrEncoding = Encoding.GetEncoding("GBK") : Catch : stderrEncoding = Encoding.UTF8 : End Try
                End Try

                Dim startInfo = New ProcessStartInfo With {
                    .FileName = wrappedCommand,
                    .Arguments = wrappedArgs,
                    .UseShellExecute = False,
                    .RedirectStandardInput = True,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True,
                    .CreateNoWindow = True,
                    .StandardOutputEncoding = Encoding.UTF8,
                    .StandardErrorEncoding = stderrEncoding
                }

                If Not String.IsNullOrEmpty(_workingDirectory) Then
                    startInfo.WorkingDirectory = _workingDirectory
                End If

                ' 注入环境变量（如 ms365 的 AZURE_CLIENT_ID 等）
                If _envVars IsNot Nothing AndAlso _envVars.Count > 0 Then
                    For Each kvp In _envVars
                        startInfo.EnvironmentVariables(kvp.Key) = kvp.Value
                        Debug.WriteLine($"[StdioTransport] 注入环境变量: {kvp.Key}=***")
                    Next
                End If

                _process = New Process() With {.StartInfo = startInfo, .EnableRaisingEvents = True}

                AddHandler _process.Exited, Sub(s, e)
                    _isConnected = False
                    ' 【修复】Exited 事件在线程池线程触发，可能与 Dispose 并发：
                    ' Dispose 释放 Process 后访问 ExitCode 会抛 InvalidOperationException(EnsureState)，
                    ' 未捕获会直接终止整个 Excel 进程。加锁 + 异常保护。
                    Dim exitCode As String = "N/A"
                    Try
                        SyncLock _processLock
                            If _process IsNot Nothing Then
                                exitCode = _process.ExitCode.ToString()
                            End If
                        End SyncLock
                    Catch ex As Exception
                        Debug.WriteLine($"StdioTransport: 读取退出码失败(可能已被释放) - {ex.Message}")
                    End Try
                    Try
                        RaiseEvent Disconnected(Me, New Exception($"MCP进程退出，退出代码: {exitCode}"))
                    Catch ex As Exception
                        Debug.WriteLine($"StdioTransport: 触发 Disconnected 事件异常 - {ex.Message}")
                    End Try
                End Sub

                _process.Start()
                _reader = _process.StandardOutput
                _writer = _process.StandardInput
                _writer.AutoFlush = True

                ' 短暂延迟，检查进程是否立即退出（启动失败）
                Await Task.Delay(500)
                If _process.HasExited Then
                    Dim stderr = _process.StandardError.ReadToEnd()
                    Throw New Exception($"进程启动后立即退出（代码{_process.ExitCode}）。{If(Not String.IsNullOrEmpty(stderr), vbCrLf & "错误详情: " & stderr, "请检查命令和参数是否正确")}")
                End If

                _isConnected = True
                _cancellationTokenSource = New CancellationTokenSource()

                ' 启动后台读取循环
                _readTask = Task.Run(Sub() ReadLoop(_cancellationTokenSource.Token))

            Catch ex As Exception
                _isConnected = False
                Throw New Exception($"Stdio 传输连接失败: {ex.Message}", ex)
            End Try
        End Function

        ''' <summary>
        ''' 后台读取循环：从 stdout 逐行读取 JSON-RPC 消息
        ''' </summary>
        Private Async Sub ReadLoop(token As CancellationToken)
            Try
                While Not token.IsCancellationRequested AndAlso _process IsNot Nothing AndAlso Not _process.HasExited
                    Dim line As String = Await _reader.ReadLineAsync()
                    If line Is Nothing Then
                        Exit While
                    End If

                    line = line.Trim()
                    If String.IsNullOrEmpty(line) Then Continue While

                    Try
                        Dim json = JToken.Parse(line)

                        ' 检查是否有 id 字段（响应 vs 通知）
                        Dim idToken = json("id")
                        If idToken IsNot Nothing Then
                            ' 这是对某个请求的响应
                            Dim reqId = idToken.ToObject(Of Object)()
                            Dim tcs As TaskCompletionSource(Of String) = Nothing
                            SyncLock _lockObj
                                If _pendingRequests.TryGetValue(reqId, tcs) Then
                                    _pendingRequests.Remove(reqId)
                                End If
                            End SyncLock

                            If tcs IsNot Nothing Then
                                tcs.TrySetResult(line)
                            End If
                        Else
                            ' 这是通知或服务器主动推送的消息
                            RaiseEvent MessageReceived(Me, line)
                        End If

                    Catch ex As JsonException
                        ' 跳过非 JSON 行
                        Debug.WriteLine($"MCP Stdio: 跳过非 JSON 行: {line.Substring(0, Math.Min(100, line.Length))}")
                    End Try
                End While

            Catch ex As OperationCanceledException
                ' 正常取消
            Catch ex As Exception
                Debug.WriteLine($"MCP Stdio 读取异常: {ex.Message}")
                _isConnected = False
                RaiseEvent Disconnected(Me, ex)
            End Try
        End Sub

        ''' <summary>
        ''' 递归终止整个进程树（taskkill /T /F）
        ''' npx/npm 启动的 MCP 是 cmd -> npx -> node 多层嵌套，Process.Kill 只杀顶层，
        ''' 残留的 node 子进程会持有管道句柄导致 Excel 关闭被拖住。
        ''' </summary>
        Private Shared Sub KillProcessTree(process As Process)
            Try
                If process Is Nothing Then Return
                Dim pid = process.Id
                Dim psi = New ProcessStartInfo With {
                    .FileName = System.IO.Path.Combine(Environment.SystemDirectory, "taskkill.exe"),
                    .Arguments = $"/PID {pid} /T /F",
                    .UseShellExecute = False,
                    .CreateNoWindow = True
                }
                Dim killer = Process.Start(psi)
                If killer IsNot Nothing Then
                    killer.WaitForExit(2000)
                    killer.Dispose()
                End If
                Debug.WriteLine($"StdioTransport: 已终止进程树 PID={pid}")
            Catch ex As Exception
                Debug.WriteLine($"StdioTransport: 终止进程树失败 - {ex.Message}")
                Try
                    If process IsNot Nothing AndAlso Not process.HasExited Then process.Kill()
                Catch
                End Try
            End Try
        End Sub

        Public Async Function SendAsync(message As String) As Task(Of String) Implements IMcpTransport.SendAsync
            If Not _isConnected Then Throw New InvalidOperationException("传输未连接")

            Dim jsonObj = JToken.Parse(message)
            Dim idToken = jsonObj("id")
            Dim id As Object = Nothing
            If idToken IsNot Nothing Then
                id = idToken.ToObject(Of Object)()
            End If
            If id Is Nothing Then
                ' 如果消息没有 id，生成一个
                id = Interlocked.Increment(_nextId)
                jsonObj("id") = JToken.FromObject(id)
                message = jsonObj.ToString(Formatting.None)
            End If

            Dim tcs = New TaskCompletionSource(Of String)()
            SyncLock _lockObj
                _pendingRequests(id) = tcs
            End SyncLock

            ' 发送消息（json-rpc 使用换行符分隔）
            Await _writer.WriteLineAsync(message)

            ' 等待响应（超时 60 秒，Playwright 等重型 MCP 首次启动较慢）
            Dim timeoutTask = Task.Delay(60000)
            Dim completedTask = Await Task.WhenAny(tcs.Task, timeoutTask)

            If completedTask Is timeoutTask Then
                SyncLock _lockObj
                    _pendingRequests.Remove(id)
                End SyncLock
                Throw New TimeoutException($"MCP 请求超时: {message.Substring(0, Math.Min(100, message.Length))}...")
            End If

            Return Await tcs.Task
        End Function

        Public Async Function SendNotificationAsync(message As String) As Task Implements IMcpTransport.SendNotificationAsync
            If Not _isConnected Then Throw New InvalidOperationException("传输未连接")
            Await _writer.WriteLineAsync(message)
        End Function

        Public Async Function DisconnectAsync() As Task Implements IMcpTransport.DisconnectAsync
            Dim totalSw = Stopwatch.StartNew()
            If Not _isConnected Then
                Debug.WriteLine("[StdioTransport.Disconnect] 已断开，跳过")
                Return
            End If

            ' 1. 取消 ReadLoop
            If _cancellationTokenSource IsNot Nothing Then
                _cancellationTokenSource.Cancel()
            End If
            Debug.WriteLine("[StdioTransport.Disconnect] CTS.Cancel() done, {0}ms", totalSw.ElapsedMilliseconds)

            ' 2. 进程管理（加锁）
            SyncLock _processLock
                Debug.WriteLine("[StdioTransport.Disconnect] 拿到 _processLock, {0}ms", totalSw.ElapsedMilliseconds)

                ' 关 stdin
                Try
                    If _writer IsNot Nothing Then _writer.Close()
                Catch
                End Try
                Debug.WriteLine("[StdioTransport.Disconnect] _writer.Close() done, {0}ms", totalSw.ElapsedMilliseconds)

                ' 等进程退出
                Try
                    If _process IsNot Nothing AndAlso Not _process.HasExited Then
                        If _process.WaitForExit(500) Then
                            Debug.WriteLine("[StdioTransport.Disconnect] 进程优雅退出, {0}ms", totalSw.ElapsedMilliseconds)
                        Else
                            Debug.WriteLine("[StdioTransport.Disconnect] WaitForExit 500ms 超时，KillProcessTree, {0}ms", totalSw.ElapsedMilliseconds)
                            KillProcessTree(_process)
                            _process.WaitForExit(300)
                            Debug.WriteLine("[StdioTransport.Disconnect] Kill + Wait done, {0}ms", totalSw.ElapsedMilliseconds)
                        End If
                    Else
                        Debug.WriteLine("[StdioTransport.Disconnect] 进程已退出, {0}ms", totalSw.ElapsedMilliseconds)
                    End If
                Catch ex As Exception
                    Debug.WriteLine("[StdioTransport.Disconnect] 异常: {0}, {1}ms", ex.Message, totalSw.ElapsedMilliseconds)
                    Try : KillProcessTree(_process) : Catch : End Try
                End Try
            End SyncLock

            _isConnected = False

            ' 清理 pending requests
            SyncLock _lockObj
                For Each kvp In _pendingRequests
                    kvp.Value.TrySetException(New Exception("传输连接已断开"))
                Next
                _pendingRequests.Clear()
            End SyncLock

            Debug.WriteLine("[StdioTransport.Disconnect] 完成，总耗时 {0}ms", totalSw.ElapsedMilliseconds)
            Await Task.CompletedTask
        End Function

        ''' <summary>保存 PID 到文件（Dispose 时 preserveProcess=True 调用）</summary>
        Private Sub SavePidFile()
            Try
                If _process Is Nothing OrElse _process.HasExited Then Return
                Dim pidFile = GetPidFilePath()
                Directory.CreateDirectory(Path.GetDirectoryName(pidFile))
                File.WriteAllText(pidFile, _process.Id.ToString())
                Debug.WriteLine("[StdioTransport] 保存 PID={0} 到 {1}", _process.Id, pidFile)
            Catch ex As Exception
                Debug.WriteLine("[StdioTransport] SavePidFile 异常: " & ex.Message)
            End Try
        End Sub

        ''' <summary>删除 PID 文件（进程真的退出时调用）</summary>
        Private Sub DeletePidFile()
            Try
                Dim pidFile = GetPidFilePath()
                If File.Exists(pidFile) Then File.Delete(pidFile)
            Catch
            End Try
        End Sub

        Public Shared Function GetPidFilePath() As String
            Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "ExcelAIVSTOAddIn", "computer-agent", "mcp_server.pid")
        End Function

        ''' <summary>
        ''' 清理残留 PID 文件：进程死了或管道关了都删 PID
        ''' Dispose(preserveProcess=True) 后 PID 留着但管道关了——下次 ConnectAsync 必须重启
        ''' </summary>
        Private Shared Sub TryGetExistingProcessAndCleanup()
            Try
                Dim pidFile = GetPidFilePath()
                If Not File.Exists(pidFile) Then Return
                Dim pid As Integer
                If Not Integer.TryParse(File.ReadAllText(pidFile).Trim(), pid) Then
                    File.Delete(pidFile)
                    Return
                End If
                Dim p = Process.GetProcessById(pid)
                If p Is Nothing OrElse p.HasExited Then
                    File.Delete(pidFile)
                    Debug.WriteLine("[StdioTransport] 旧进程 PID={0} 已死，清理 PID 文件", pid)
                Else
                    File.Delete(pidFile)
                    Debug.WriteLine("[StdioTransport] 旧进程 PID={0} 还活着但管道已关，强制重启", pid)
                    Try
                        p.Kill()
                    Catch
                    End Try
                End If
            Catch
            End Try
        End Sub

        ''' <summary>检查上次的 python MCP server 是不是还活着，活着就返回进程，让启动时跳过冷启动</summary>
        Public Shared Function TryGetExistingProcess() As Process
            Try
                Dim pidFile = GetPidFilePath()
                If Not File.Exists(pidFile) Then Return Nothing
                Dim pid As Integer
                If Not Integer.TryParse(File.ReadAllText(pidFile).Trim(), pid) Then Return Nothing
                Dim p = Process.GetProcessById(pid)
                If p IsNot Nothing AndAlso Not p.HasExited Then
                    Debug.WriteLine("[StdioTransport] 发现存活的 MCP server PID={0}，尝试复用", pid)
                    Return p
                End If
            Catch ex As Exception
                Debug.WriteLine("[StdioTransport] TryGetExistingProcess: " & ex.Message)
            End Try
            Return Nothing
        End Function

        ''' <summary>
        ''' 释放资源（接口实现）
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(False)
        End Sub

        ''' <summary>
        ''' 释放资源
        ''' preserveProcess=True 时：关闭管道但保留 python.exe（下次 Excel 启动复用）
        ''' preserveProcess=False 时：Kill 进程（彻底清理）
        ''' </summary>
        Public Sub Dispose(preserveProcess As Boolean)
            SyncLock _processLock
                If _disposed Then Return
                _disposed = True

                If _cancellationTokenSource IsNot Nothing Then
                    Try
                        _cancellationTokenSource.Cancel()
                    Catch
                    End Try
                    Try
                        _cancellationTokenSource.Dispose()
                    Catch
                    End Try
                    _cancellationTokenSource = Nothing
                End If

                Try
                    If _reader IsNot Nothing Then _reader.Dispose()
                Catch
                End Try
                Try
                    If _writer IsNot Nothing Then _writer.Dispose()
                Catch
                End Try

                ' 进程处理
                Try
                    If _process IsNot Nothing Then
                        If Not _process.HasExited Then
                            If preserveProcess Then
                                SavePidFile()
                                Debug.WriteLine("[StdioTransport] 保留进程 PID={0}，关闭管道", _process.Id)
                            Else
                                KillProcessTree(_process)
                                DeletePidFile()
                            End If
                        Else
                            DeletePidFile()
                        End If
                        _process.Dispose()
                        _process = Nothing
                    End If
                Catch ex As Exception
                    Debug.WriteLine("[StdioTransport] Dispose 进程异常: " & ex.Message)
                    Try
                        KillProcessTree(_process)
                    Catch
                    End Try
                End Try
            End SyncLock
        End Sub

    End Class

    '#End Region

    '#Region "HTTP SSE 传输实现"

    ''' <summary>
    ''' 基于 HTTP SSE 的 MCP 传输（用于连接远程 MCP Server）
    ''' </summary>
    Public Class HttpSseTransport
        Implements IMcpTransport

        Private _httpClient As HttpClient
        Private _baseUrl As String
        Private _headers As Dictionary(Of String, String)
        Private _isConnected As Boolean = False
        Private _cancellationTokenSource As CancellationTokenSource
        Private _readTask As Task
        Private _pendingRequests As Dictionary(Of Object, TaskCompletionSource(Of String))
        Private _lockObj As New Object()
        Private _nextId As Integer = 1
        Private _sessionId As String
        Private _messageEndpoint As String

        Public Event MessageReceived As EventHandler(Of String) Implements IMcpTransport.MessageReceived
        Public Event Disconnected As EventHandler(Of Exception) Implements IMcpTransport.Disconnected

        Public ReadOnly Property IsConnected As Boolean Implements IMcpTransport.IsConnected
            Get
                Return _isConnected
            End Get
        End Property

        Public ReadOnly Property TransportType As String Implements IMcpTransport.TransportType
            Get
                Return "http-sse"
            End Get
        End Property

        ''' <summary>
        ''' 创建 HTTP SSE 传输
        ''' </summary>
        ''' <param name="baseUrl">MCP Server 的基础 URL</param>
        ''' <param name="headers">额外的 HTTP 头</param>
        Public Sub New(baseUrl As String, Optional headers As Dictionary(Of String, String) = Nothing)
            _baseUrl = baseUrl.TrimEnd("/"c)
            _headers = headers
            _httpClient = New HttpClient()
            _httpClient.Timeout = TimeSpan.FromSeconds(60)
            _pendingRequests = New Dictionary(Of Object, TaskCompletionSource(Of String))()

            ' 设置自定义 headers
            If _headers IsNot Nothing Then
                For Each kvp In _headers
                    _httpClient.DefaultRequestHeaders.TryAddWithoutValidation(kvp.Key, kvp.Value)
                Next
            End If
        End Sub

        Public Async Function ConnectAsync() As Task Implements IMcpTransport.ConnectAsync
            If _isConnected Then Return

            Try
                _cancellationTokenSource = New CancellationTokenSource()

                ' 建立 SSE 连接
                Dim request = New HttpRequestMessage(HttpMethod.Get, $"{_baseUrl}/sse")
                request.Headers.Add("Accept", "text/event-stream")

                Dim response = Await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, _cancellationTokenSource.Token)
                response.EnsureSuccessStatusCode()

                _isConnected = True

                ' 后台读取 SSE 事件流
                _readTask = Task.Run(Async Function()
                    Try
                        Using stream = Await response.Content.ReadAsStreamAsync()
                            Using reader = New StreamReader(stream)
                                Dim eventType As String = ""
                                Dim eventData As String = ""

                                While Not _cancellationTokenSource.Token.IsCancellationRequested
                                    Dim line = Await reader.ReadLineAsync()
                                    If line Is Nothing Then Exit While

                                    If line.StartsWith("event:") Then
                                        eventType = line.Substring(6).Trim()
                                    ElseIf line.StartsWith("data:") Then
                                        eventData = line.Substring(5).Trim()
                                    ElseIf String.IsNullOrEmpty(line) AndAlso Not String.IsNullOrEmpty(eventData) Then
                                        ' 空行表示事件结束
                                        HandleSseEvent(eventType, eventData)
                                        eventType = ""
                                        eventData = ""
                                    End If
                                End While
                            End Using
                        End Using
                    Catch ex As OperationCanceledException
                    Catch ex As Exception
                        Debug.WriteLine($"MCP SSE 读取异常: {ex.Message}")
                        _isConnected = False
                        RaiseEvent Disconnected(Me, ex)
                    End Try
                End Function)

            Catch ex As Exception
                _isConnected = False
                Throw New Exception($"HTTP SSE 传输连接失败: {ex.Message}", ex)
            End Try
        End Function

        ''' <summary>
        ''' 处理 SSE 事件
        ''' </summary>
        Private Sub HandleSseEvent(eventType As String, data As String)
            Try
                Select Case eventType
                    Case "endpoint"
                        ' 获取消息端点
                        _messageEndpoint = data.Trim()
                        Debug.WriteLine($"MCP SSE: 消息端点 = {_messageEndpoint}")

                    Case "message"
                        ' 处理 JSON-RPC 消息
                        Dim json = JToken.Parse(data)
                        Dim idToken = json("id")
                        If idToken IsNot Nothing Then
                            Dim reqId = idToken.ToObject(Of Object)()
                            Dim tcs As TaskCompletionSource(Of String) = Nothing
                            SyncLock _lockObj
                                If _pendingRequests.TryGetValue(reqId, tcs) Then
                                    _pendingRequests.Remove(reqId)
                                End If
                            End SyncLock
                            If tcs IsNot Nothing Then
                                tcs.TrySetResult(data)
                            End If
                        Else
                            RaiseEvent MessageReceived(Me, data)
                        End If

                    Case Else
                        Debug.WriteLine($"MCP SSE: 未处理的事件类型: {eventType}, 数据: {data.Substring(0, Math.Min(100, data.Length))}")
                End Select

            Catch ex As Exception
                Debug.WriteLine($"MCP SSE: 事件处理异常: {ex.Message}")
            End Try
        End Sub

        Public Async Function SendAsync(message As String) As Task(Of String) Implements IMcpTransport.SendAsync
            If Not _isConnected Then Throw New InvalidOperationException("传输未连接")
            If String.IsNullOrEmpty(_messageEndpoint) Then Throw New InvalidOperationException("尚未收到消息端点")

            Dim jsonObj = JToken.Parse(message)
            Dim idToken = jsonObj("id")
            Dim id As Object = Nothing
            If idToken IsNot Nothing Then
                id = idToken.ToObject(Of Object)()
            End If
            If id Is Nothing Then
                id = Interlocked.Increment(_nextId)
                jsonObj("id") = JToken.FromObject(id)
                message = jsonObj.ToString(Formatting.None)
            End If

            Dim tcs = New TaskCompletionSource(Of String)()
            SyncLock _lockObj
                _pendingRequests(id) = tcs
            End SyncLock

            ' 通过 HTTP POST 发送消息到消息端点
            Dim fullUrl = If(_messageEndpoint.StartsWith("http"), _messageEndpoint, $"{_baseUrl}{_messageEndpoint}")
            Dim content = New StringContent(message, Encoding.UTF8, "application/json")
            Dim cancelToken = If(_cancellationTokenSource IsNot Nothing, _cancellationTokenSource.Token, CancellationToken.None)
            Dim response = Await _httpClient.PostAsync(fullUrl, content, cancelToken)

            If response.IsSuccessStatusCode Then
                ' MCP SSE 规范：响应可能直接返回或通过 SSE 返回
                ' 如果直接返回 JSON，且包含所需 id，直接处理
                Dim responseBody = Await response.Content.ReadAsStringAsync()
                Try
                    Dim respJson = JToken.Parse(responseBody)
                    If respJson("id") IsNot Nothing Then
                        SyncLock _lockObj
                            _pendingRequests.Remove(id)
                        End SyncLock
                        tcs.TrySetResult(responseBody)
                        Return responseBody
                    End If
                Catch
                End Try

                ' 否则等待 SSE 推送
                Return Await tcs.Task

            Else
                SyncLock _lockObj
                    _pendingRequests.Remove(id)
                End SyncLock
                Dim errorBody = Await response.Content.ReadAsStringAsync()
                Throw New Exception($"MCP HTTP 请求失败 ({CInt(response.StatusCode)}): {errorBody}")
            End If
        End Function

        Public Async Function SendNotificationAsync(message As String) As Task Implements IMcpTransport.SendNotificationAsync
            If Not _isConnected Then Throw New InvalidOperationException("传输未连接")
            If String.IsNullOrEmpty(_messageEndpoint) Then Throw New InvalidOperationException("尚未收到消息端点")

            Dim fullUrl = If(_messageEndpoint.StartsWith("http"), _messageEndpoint, $"{_baseUrl}{_messageEndpoint}")
            Dim content = New StringContent(message, Encoding.UTF8, "application/json")
            Dim cancelToken = If(_cancellationTokenSource IsNot Nothing, _cancellationTokenSource.Token, CancellationToken.None)
            Await _httpClient.PostAsync(fullUrl, content, cancelToken)
        End Function

        Public Async Function DisconnectAsync() As Task Implements IMcpTransport.DisconnectAsync
            If Not _isConnected Then Return

            If _cancellationTokenSource IsNot Nothing Then
                _cancellationTokenSource.Cancel()
            End If
            _isConnected = False

            SyncLock _lockObj
                For Each kvp In _pendingRequests
                    kvp.Value.TrySetException(New Exception("传输连接已断开"))
                Next
                _pendingRequests.Clear()
            End SyncLock

            Await Task.CompletedTask
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            If _cancellationTokenSource IsNot Nothing Then
                _cancellationTokenSource.Cancel()
                _cancellationTokenSource.Dispose()
            End If
            If _httpClient IsNot Nothing Then _httpClient.Dispose()
        End Sub

    End Class

    '#End Region

End Namespace
