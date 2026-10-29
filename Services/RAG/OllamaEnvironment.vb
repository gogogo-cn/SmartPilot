Option Explicit On
Option Strict On

Imports System.Diagnostics
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Sockets
Imports System.Threading.Tasks

Namespace Services.RAG

    ''' <summary>
    ''' Ollama 环境检测与 embedding 模型一键拉起。
    ''' 面向「终端用户未装 Ollama / 未拉模型」场景，供 RAG 面板首次启动引导使用：
    '''   1. 检测 ollama.exe 是否安装（PATH + 常见安装目录）
    '''   2. 检测 Ollama 服务是否运行（TCP 500ms 预探测 + /api/tags 复核）
    '''   3. 检测 embedding 模型是否已拉取（/api/tags 模型清单前缀匹配）
    '''   4. 一键拉起模型：异步进程 + 进度轮询，防重复启动，输出解析百分比
    ''' 设计约束：
    '''   - 全程不抛异常给调用方（返回状态对象，由 UI 决定展示）
    '''   - TCP 探测超时分支绝不再 Await connectTask（其内部网络超时约 2s 会把 500ms 探测拖垮），
    '''     仅用 ContinueWith(OnlyOnFaulted) 标记异常已观察
    '''   - localhost 一律换 127.0.0.1：.NET 对 localhost 优先尝试 IPv6 ::1，未监听 v6 时每次固定多花约 2 秒
    ''' </summary>
    Public Class OllamaEnvironment

        Private Const DEFAULT_OLLAMA_PORT As Integer = 11434
        Private Const DEFAULT_OLLAMA_ENDPOINT As String = "http://127.0.0.1:11434"
        Private Const DEFAULT_MODEL As String = "bge-m3"

        ' ---- 模型拉起状态（Shared：跨 COM 桥调用保持）----
        Private Shared _pullLock As New Object()
        Private Shared _pullProcess As Process = Nothing
        Private Shared _pullRunning As Boolean = False
        Private Shared _pullPercent As Integer = 0
        Private Shared _pullLastLine As String = ""
        Private Shared _pullError As String = ""
        Private Shared _pullStartTime As DateTime = DateTime.MinValue

        ' =====================================================================
        ' 公共状态类型
        ' =====================================================================

        ''' <summary>RAG 本地环境检测结果。</summary>
        Public Class EnvironmentStatus
            Public Property OllamaInstalled As Boolean = False
            Public Property OllamaRunning As Boolean = False
            Public Property ModelInstalled As Boolean = False
            Public Property Ready As Boolean = False
            Public Property ModelName As String = ""
            Public Property Endpoint As String = ""
            ''' <summary>引导步骤：install_ollama / pull_model / ready。</summary>
            Public Property NextStep As String = "ready"
        End Class

        ''' <summary>模型拉起进度状态。</summary>
        Public Class PullState
            Public Property Running As Boolean = False
            Public Property Percent As Integer = 0
            Public Property LastLine As String = ""
            Public Property [Error] As String = ""
            Public Property ElapsedSeconds As Integer = 0
        End Class

        ' =====================================================================
        ' 环境检测
        ' =====================================================================

        ''' <summary>
        ''' 查找 ollama.exe：先 PATH，再常见安装位置（%LOCALAPPDATA%\Programs\Ollama、%ProgramFiles%\Ollama）。
        ''' 找不到返回 Nothing。
        ''' </summary>
        Public Shared Function FindOllamaExe() As String
            Try
                Dim pathVar = Environment.GetEnvironmentVariable("PATH")
                If Not String.IsNullOrWhiteSpace(pathVar) Then
                    For Each dirPath In pathVar.Split(";"c)
                        If String.IsNullOrWhiteSpace(dirPath) Then Continue For
                        Try
                            Dim exe = Path.Combine(dirPath.Trim(), "ollama.exe")
                            If File.Exists(exe) Then Return exe
                        Catch
                            ' 非法路径字符等，跳过该目录
                        End Try
                    Next
                End If
                Dim localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
                Dim programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
                Dim candidates As String() = {
                    Path.Combine(localAppData, "Programs", "Ollama", "ollama.exe"),
                    Path.Combine(programFiles, "Ollama", "ollama.exe")
                }
                For Each exe In candidates
                    If File.Exists(exe) Then Return exe
                Next
            Catch
                ' 检测失败按未安装处理
            End Try
            Return Nothing
        End Function

        ''' <summary>
        ''' 汇总检测：安装状态 + 运行状态 + 模型就绪状态 + 下一步引导。
        ''' endpoint 为配置里的 LocalBgeEndpoint；为空时用默认地址。
        ''' </summary>
        Public Shared Async Function CheckEnvironmentAsync(endpoint As String, modelName As String) As Task(Of EnvironmentStatus)
            Dim status As New EnvironmentStatus()
            If String.IsNullOrWhiteSpace(modelName) Then modelName = DEFAULT_MODEL
            status.ModelName = modelName
            status.Endpoint = If(String.IsNullOrWhiteSpace(endpoint), DEFAULT_OLLAMA_ENDPOINT, endpoint)

            status.OllamaInstalled = Not String.IsNullOrWhiteSpace(FindOllamaExe())

            Dim normalized = NormalizeLocalEndpoint(status.Endpoint)
            Dim port As Integer = DEFAULT_OLLAMA_PORT
            Try
                Dim uri As New Uri(normalized)
                If uri.Port > 0 Then port = uri.Port
            Catch
                ' 保持默认端口
            End Try

            status.OllamaRunning = Await TcpProbeAsync(port, 500)
            If status.OllamaRunning Then
                status.ModelInstalled = Await IsModelInstalledAsync(normalized, modelName)
            End If

            status.Ready = status.OllamaRunning AndAlso status.ModelInstalled
            If Not status.OllamaInstalled Then
                status.NextStep = "install_ollama"
            ElseIf Not status.OllamaRunning Then
                status.NextStep = "start_ollama"
            ElseIf Not status.ModelInstalled Then
                status.NextStep = "pull_model"
            Else
                status.NextStep = "ready"
            End If
            Return status
        End Function

        ''' <summary>localhost → 127.0.0.1（避开 .NET IPv6 ::1 优先探测的 2 秒级挂起）。</summary>
        Private Shared Function NormalizeLocalEndpoint(endpoint As String) As String
            Try
                Dim uri As New Uri(endpoint.Trim())
                If uri.Host = "localhost" Then
                    Dim builder As New UriBuilder(uri)
                    builder.Host = "127.0.0.1"
                    Return builder.Uri.ToString().TrimEnd("/"c)
                End If
                Return endpoint.Trim().TrimEnd("/"c)
            Catch
                Return endpoint.Trim().TrimEnd("/"c)
            End Try
        End Function

        ''' <summary>
        ''' TCP 500ms 预探测：端口通不通。
        ''' 超时分支绝不再 Await connectTask（其内部网络超时约 2 秒会把探测拖垮）。
        ''' </summary>
        Private Shared Async Function TcpProbeAsync(port As Integer, timeoutMs As Integer) As Task(Of Boolean)
            Try
                Using client As New TcpClient()
                    Dim connectTask = client.ConnectAsync("127.0.0.1", port)
                    Dim delayTask = Task.Delay(timeoutMs)
                    Dim completed = Await Task.WhenAny(connectTask, delayTask)
                    If completed Is delayTask Then
                        ' 仅挂续任务标记异常已观察（不 Await 返回值，防 BC42358）
                        Dim ignoreContinuation = connectTask.ContinueWith(
                            Sub(t)
                                Dim ignore = t.Exception
                            End Sub,
                            TaskContinuationOptions.OnlyOnFaulted)
                        Return False
                    End If
                    Await connectTask
                    Return client.Connected
                End Using
            Catch
                Return False
            End Try
        End Function

        ''' <summary>GET /api/tags，模型清单中前缀匹配（bge-m3 命中 bge-m3:latest）。</summary>
        Private Shared Async Function IsModelInstalledAsync(endpoint As String, modelName As String) As Task(Of Boolean)
            Try
                Dim handler As New HttpClientHandler()
                handler.UseProxy = False
                Using client As New HttpClient(handler)
                    client.Timeout = TimeSpan.FromSeconds(5)
                    Dim response = Await client.GetAsync(endpoint & "/api/tags")
                    If Not response.IsSuccessStatusCode Then Return False
                    Dim jsonStr = Await response.Content.ReadAsStringAsync()
                    Dim obj = Newtonsoft.Json.Linq.JObject.Parse(jsonStr)
                    Dim models = obj("models")
                    If models Is Nothing Then Return False
                    Dim prefix = modelName.Trim().ToLower() & ":"
                    For Each m In models
                        Dim nameToken = m("name")
                        If nameToken Is Nothing Then Continue For
                        Dim name = nameToken.ToString().Trim().ToLower()
                        If name = modelName.Trim().ToLower() OrElse name.StartsWith(prefix) Then
                            Return True
                        End If
                    Next
                    Return False
                End Using
            Catch
                Return False
            End Try
        End Function

        ' =====================================================================
        ' 模型一键拉起
        ' =====================================================================

        ''' <summary>启动 ollama pull（异步进程）。返回是否成功启动；已在进行返回 False。</summary>
        Public Shared Function StartModelPull(modelName As String) As Boolean
            If String.IsNullOrWhiteSpace(modelName) Then modelName = DEFAULT_MODEL

            ' 防参数注入：模型名不允许空白与 shell 元字符
            If modelName.IndexOfAny(New Char() {" "c, ControlChars.Tab, """"c, "&"c, "|"c, ">"c, "<"c}) >= 0 Then
                SyncLock _pullLock
                    _pullError = "模型名包含非法字符: " & modelName
                End SyncLock
                Return False
            End If

            SyncLock _pullLock
                If _pullRunning AndAlso _pullProcess IsNot Nothing AndAlso Not _pullProcess.HasExited Then
                    Return False
                End If
                Dim ollamaExe = FindOllamaExe()
                If String.IsNullOrWhiteSpace(ollamaExe) Then
                    _pullError = "未找到 ollama.exe，请先安装 Ollama"
                    Return False
                End If
                Try
                    Dim psi As New ProcessStartInfo()
                    psi.FileName = ollamaExe
                    psi.Arguments = "pull " & modelName
                    psi.UseShellExecute = False
                    psi.CreateNoWindow = True
                    psi.RedirectStandardOutput = True
                    psi.RedirectStandardError = True

                    _pullProcess = New Process()
                    _pullProcess.StartInfo = psi
                    _pullProcess.EnableRaisingEvents = True
                    AddHandler _pullProcess.OutputDataReceived, AddressOf OnPullOutput
                    AddHandler _pullProcess.ErrorDataReceived, AddressOf OnPullOutput
                    AddHandler _pullProcess.Exited, AddressOf OnPullExited

                    _pullPercent = 0
                    _pullLastLine = "正在启动 ollama pull " & modelName & " ..."
                    _pullError = ""
                    _pullStartTime = DateTime.Now
                    _pullRunning = True

                    _pullProcess.Start()
                    _pullProcess.BeginOutputReadLine()
                    _pullProcess.BeginErrorReadLine()
                    Return True
                Catch ex As Exception
                    _pullRunning = False
                    _pullError = "启动 ollama pull 失败: " & ex.Message
                    Return False
                End Try
            End SyncLock
        End Function

        ''' <summary>查询模型拉起进度（供前端轮询）。</summary>
        Public Shared Function GetPullState() As PullState
            SyncLock _pullLock
                Dim state As New PullState()
                state.Running = _pullRunning
                state.Percent = _pullPercent
                state.LastLine = _pullLastLine
                state.Error = _pullError
                If _pullProcess IsNot Nothing AndAlso _pullProcess.HasExited Then
                    state.Running = False
                    If _pullProcess.ExitCode <> 0 AndAlso String.IsNullOrEmpty(_pullError) Then
                        _pullError = "ollama pull 失败（退出码 " & _pullProcess.ExitCode & "）"
                        state.Error = _pullError
                    End If
                End If
                If _pullRunning AndAlso _pullStartTime <> DateTime.MinValue Then
                    state.ElapsedSeconds = CInt((DateTime.Now - _pullStartTime).TotalSeconds)
                End If
                Return state
            End SyncLock
        End Function

        ''' <summary>最近一次拉起失败原因（无则空串）。</summary>
        Public Shared Function GetLastError() As String
            SyncLock _pullLock
                Return _pullError
            End SyncLock
        End Function

        ''' <summary>停止正在进行的拉起（杀进程）。</summary>
        Public Shared Sub StopModelPull()
            SyncLock _pullLock
                Try
                    If _pullProcess IsNot Nothing AndAlso Not _pullProcess.HasExited Then
                        _pullProcess.Kill()
                    End If
                Catch
                    ' 进程可能已退出，忽略
                End Try
                _pullRunning = False
                If String.IsNullOrEmpty(_pullError) Then _pullError = "已被用户取消"
            End SyncLock
        End Sub

        ' =====================================================================
        ' 进程输出处理
        ' =====================================================================

        Private Shared Sub OnPullOutput(sender As Object, e As DataReceivedEventArgs)
            If e Is Nothing OrElse String.IsNullOrEmpty(e.Data) Then Return
            Dim line = e.Data.Trim()
            If line.Length = 0 Then Return
            SyncLock _pullLock
                _pullLastLine = line
                ' ollama 输出形如：pulling daec91ffb5dd:  45% ▕████▏ 540 MB/1.2 GB
                ' 多行进度混杂时取最大值（pulling manifest 与 pulling blob 交替出现）
                Dim m = System.Text.RegularExpressions.Regex.Match(line, "(\d{1,3})%")
                If m.Success AndAlso m.Groups.Count >= 2 Then
                    Dim pct As Integer = 0
                    If Integer.TryParse(m.Groups(1).Value, pct) Then
                        If pct >= 0 AndAlso pct <= 100 AndAlso pct > _pullPercent Then
                            _pullPercent = pct
                        End If
                    End If
                End If
                If line.ToLower().Contains("error") OrElse line.ToLower().Contains("failed") Then
                    _pullError = line
                End If
            End SyncLock
        End Sub

        Private Shared Sub OnPullExited(sender As Object, e As EventArgs)
            SyncLock _pullLock
                _pullRunning = False
                If _pullProcess IsNot Nothing Then
                    If _pullProcess.ExitCode = 0 Then
                        _pullPercent = 100
                        _pullLastLine = "模型拉取完成"
                    Else
                        If String.IsNullOrEmpty(_pullError) Then
                            _pullError = "ollama pull 失败（退出码 " & _pullProcess.ExitCode & "）"
                        End If
                        _pullLastLine = _pullError
                    End If
                End If
            End SyncLock
        End Sub

    End Class

End Namespace
