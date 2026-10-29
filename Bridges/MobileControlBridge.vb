Option Explicit On
Option Infer On

Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.NetworkInformation
Imports System.Net.Sockets
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ExcelAIVSTOAddIn.Utils
Imports ExcelAIVSTOAddIn.Services

Namespace Bridges

    ''' <summary>
    ''' MobileControlBridge - 移动端遥控桥梁
    ''' 在同一 WiFi 网络下，手机浏览器访问桌面端 HTTP 服务，
    ''' 通过摄像头拍照上传图片，触发桌面端 OCR 识别并返回结果。
    ''' 使用 TcpListener 实现，无需管理员权限。
    ''' </summary>
    <ComVisible(True)>
    <ClassInterface(ClassInterfaceType.AutoDispatch)>
    <ProgId("ExcelAIVSTOAddIn.MobileControlBridge")>
    Public Class MobileControlBridge
        Implements IBridge

        ' OCR 完成事件（通知桌面 OCR 面板刷新结果）
        Public Event OcrCompleted(text As String, fileName As String, imageBase64 As String)

        ' 移动端上传事件（通知桌面打开 OCR 面板）
        Public Event UploadReceived(fileName As String)

        ' 单例
        Private Shared ReadOnly _lazyInstance As New Lazy(Of MobileControlBridge)(
            Function() New MobileControlBridge(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        Public Shared ReadOnly Property Instance As MobileControlBridge
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        ' 防止监听 socket 句柄被子进程继承（否则 Excel 启动的子进程会持有句柄，
        ' Excel 退出后端口永不释放，重启后报"端口只允许使用一次"）
        Private Const HANDLE_FLAG_INHERIT As Integer = &H1

        <DllImport("kernel32.dll", SetLastError:=True)>
        Private Shared Function SetHandleInformation(hObject As IntPtr, dwMask As Integer, dwFlags As Integer) As Boolean
        End Function

        Private Shared Sub MakeSocketNonInheritable(s As Socket)
            Try
                If s IsNot Nothing AndAlso s.Handle <> IntPtr.Zero Then
                    SetHandleInformation(s.Handle, HANDLE_FLAG_INHERIT, 0)
                End If
            Catch
            End Try
        End Sub

        ' TCP 服务器（替代 HttpListener，无需管理员权限）
        Private _listener As TcpListener
        Private _listenThread As Thread
        Private _isRunning As Boolean = False
        Private _port As Integer = Utils.ConfigConstants.LocalServices.VSTO_HTTP_PORT
        Private Const STANDALONE_AGENT_PORT As Integer = Utils.ConfigConstants.LocalServices.STANDALONE_AGENT_PORT
        Private _localIp As String = ""
        Private _cancellationToken As CancellationTokenSource
        Private _sessionToken As String = ""
        ' 活跃客户端连接集合：StopServer 时全部断开，避免请求线程在 Excel 关闭时
        ' 持有 COM 引用，导致 GC.WaitForPendingFinalizers 与主线程死锁（"关闭 Excel 卡"）。
        Private ReadOnly _activeClients As New List(Of TcpClient)()
        Private ReadOnly _clientLock As New Object()

        ' === 番茄钟共享同步状态（独立版 ↔ Excel 内嵌版 的权威存储） ===
        Private Shared _pomodoroLock As New Object()
        Private Shared _pomodoroState As JObject = New JObject()
        Private Shared _pomodoroVersion As Long = 0
        Private Shared _pomodoroLastWriterId As String = ""
        Private Shared _pomodoroEvents As New List(Of String)()

        ' 配置
        Private _config As ConfigurationManager

        ' 临时图片存储目录
        Private _tempImageDir As String

#Region "构造 & 初始化"

        Private Sub New()
            ' 统一生命周期管理：实例化即登记，由 BridgeRegistry.ShutdownAll 统一释放
            BridgeRegistry.Instance.Register(Me)
            Try
                _config = ConfigurationManager.Instance
                _tempImageDir = Path.Combine(Path.GetTempPath(), "ExcelAIMobile")
                If Not Directory.Exists(_tempImageDir) Then
                    Directory.CreateDirectory(_tempImageDir)
                End If
                _localIp = GetLocalIPAddress()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[MobileControlBridge] 构造函数异常: " & ex.Message)
            End Try
        End Sub

#End Region

#Region "IBridge"

        Public ReadOnly Property BridgeName As String Implements IBridge.BridgeName
            Get
                Return "MobileControlBridge"
            End Get
        End Property

        Private _disposed As Boolean = False
        Public Sub Dispose() Implements IBridge.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not _disposed Then
                If disposing Then
                    StopServer()
                End If
                _disposed = True
            End If
        End Sub

#End Region

#Region "公开 API（供前端 JS 调用）"

        ''' <summary>
        ''' 启动 TCP HTTP 服务器（无需管理员权限）
        ''' </summary>
        <ComVisible(True)>
        Public Function StartServer(Optional port As Integer = -1) As String
            Try
                If _isRunning Then
                    ' 服务已在运行，重新探测 IP（网络环境可能已变化）
                    _localIp = GetLocalIPAddress()
                    Return JsonConvert.SerializeObject(New With {
                        .success = True,
                        .message = "服务已在运行中",
                        .ip = _localIp,
                        .port = _port,
                        .url = $"http://{_localIp}:{_port}/",
                        .sessionToken = _sessionToken
                    })
                End If

                If port > 0 Then _port = port

                ' 获取局域网 IP
                _localIp = GetLocalIPAddress()
                If String.IsNullOrEmpty(_localIp) Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = "无法获取本机局域网 IP 地址，请检查网络连接"
                    })
                End If

                ' 生成一次性 session token 用于 API 鉴权（进程内存，重启自动重置）
                _sessionToken = Guid.NewGuid().ToString("N") & DateTimeOffset.UtcNow.ToUnixTimeSeconds().ToString("X8")

                ' 绑定端口；若被占用（常见于上次 Excel 非正常退出后 socket 句柄被
                ' 子进程继承形成僵尸监听），从 _port 起依次向后回退，最多试 11 个
                Dim bindEx As Exception = Nothing
                Dim boundListener As TcpListener = Nothing
                Dim usedPort As Integer = _port
                For attempt As Integer = 0 To 10
                    Dim candidatePort As Integer = _port + attempt
                    Dim l As New TcpListener(IPAddress.Any, candidatePort)
                    Try
                        ' 必须在 Start 前禁止句柄继承
                        MakeSocketNonInheritable(l.Server)
                        l.Start()
                        boundListener = l
                        usedPort = candidatePort
                        Exit For
                    Catch ex As SocketException When ex.SocketErrorCode = SocketError.AddressAlreadyInUse
                        Try : l.Stop() : Catch : End Try
                        bindEx = ex
                    Catch
                        Try : l.Stop() : Catch : End Try
                        Throw
                    End Try
                Next

                If boundListener Is Nothing Then
                    System.Diagnostics.Debug.WriteLine($"[MobileControlBridge] 端口 {_port}-{_port + 10} 均被占用")
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = $"端口 {_port} 起连续 11 个端口均被占用，通常是残留的 Excel/子进程仍在监听（僵尸端口）。请关闭所有 Excel 进程后重试，必要时重启电脑。({bindEx.Message})"
                    })
                End If

                _listener = boundListener
                _port = usedPort
                _cancellationToken = New CancellationTokenSource()
                _isRunning = True

                ' 在后台线程处理请求
                _listenThread = New Thread(Sub() ListenLoop())
                _listenThread.IsBackground = True
                _listenThread.Name = "MobileControlBridge-Listener"
                _listenThread.Start()

                System.Diagnostics.Debug.WriteLine($"[MobileControlBridge] 服务已启动: http://{_localIp}:{_port}/")

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .message = "移动端遥控服务已启动",
                    .ip = _localIp,
                    .port = _port,
                    .url = $"http://{_localIp}:{_port}/",
                    .sessionToken = _sessionToken
                })
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[MobileControlBridge] StartServer 异常: " & ex.Message)
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 停止 HTTP 服务器
        ''' </summary>
        <ComVisible(True)>
        Public Function StopServer() As String
            Try
                If Not _isRunning Then
                    Return JsonConvert.SerializeObject(New With {.success = True, .message = "服务未运行"})
                End If

                _isRunning = False
                If _cancellationToken IsNot Nothing Then
                    _cancellationToken.Cancel()
                End If
                If _listener IsNot Nothing Then
                    Try
                        _listener.Stop()
                    Catch : End Try
                    _listener = Nothing
                End If

                ' 立即断开所有活跃连接：让 HandleClient 的读取立刻抛异常退出，
                ' 避免请求线程在 Excel 关闭时持有 COM 引用导致 GC 死锁（"关闭很卡"）
                SyncLock _clientLock
                    For Each c As TcpClient In _activeClients
                        Try : c.Close() : Catch : End Try
                    Next
                    _activeClients.Clear()
                End SyncLock

                System.Diagnostics.Debug.WriteLine("[MobileControlBridge] 服务已停止")
                Return JsonConvert.SerializeObject(New With {.success = True, .message = "服务已停止"})
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[MobileControlBridge] StopServer 异常: " & ex.Message)
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 获取服务状态（运行中时会重新探测本机 IP，便于网络切换后刷新二维码）
        ''' </summary>
        <ComVisible(True)>
        Public Function GetStatus() As String
            If _isRunning Then
                _localIp = GetLocalIPAddress()
            End If
            Return JsonConvert.SerializeObject(New With {
                .success = True,
                .isRunning = _isRunning,
                .ip = _localIp,
                .port = _port,
                .url = If(_isRunning, $"http://{_localIp}:{_port}/", ""),
                .ocrEngine = ReadOcrEngineNameSafe(),
                .sessionToken = _sessionToken
            })
        End Function

#End Region

#Region "TCP 请求处理"

        Private Sub ListenLoop()
            Try
                While _isRunning AndAlso _listener IsNot Nothing
                    Dim client As TcpClient = Nothing
                    Try
                        ' 异步等待客户端连接
                        Dim acceptTask = _listener.AcceptTcpClientAsync()
                        acceptTask.Wait(_cancellationToken.Token)
                        client = acceptTask.Result
                    Catch ex As OperationCanceledException
                        Exit While
                    Catch ex As ObjectDisposedException
                        Exit While
                    Catch ex As Exception
                        If _isRunning Then
                            System.Diagnostics.Debug.WriteLine("[MobileControlBridge] AcceptTcpClient 异常: " & ex.Message)
                        End If
                        Exit While
                    End Try

                    If client IsNot Nothing Then
                        ' 在后台线程池处理请求
                        Task.Run(Sub() HandleClient(client))
                    End If
                End While
            Catch ex As Exception
                If _isRunning Then
                    System.Diagnostics.Debug.WriteLine("[MobileControlBridge] ListenLoop 致命异常: " & ex.Message)
                End If
            End Try
            _isRunning = False
        End Sub

        Private Sub HandleClient(client As TcpClient)
            SyncLock _clientLock
                _activeClients.Add(client)
            End SyncLock
            Try
                client.ReceiveTimeout = 30000
                client.SendTimeout = 30000

                Using client
                    Using stream = client.GetStream()
                        ' 读取 HTTP 请求
                        Dim request = ReadHttpRequest(stream)
                        If request Is Nothing Then Return

                        Dim path = request.Path.TrimEnd("/"c)
                        Dim method = request.Method.ToUpper()
                        ' 本次响应允许的 CORS 来源（第三方站点 → 空串 = 不输出 ACAO）
                        Dim corsOrigin = ResolveCorsOrigin(request)

                        System.Diagnostics.Debug.WriteLine($"[MobileControlBridge] {method} {path}")

                        ' ============ 鉴权检查 ============
                        If Not IsPublicPath(path, method) Then
                            Dim tokenValid = ValidateRequestToken(request)
                            If Not tokenValid Then
                                ' 写入 401 JSON 响应后直接返回，不继续路由分发
                                Try
                                    Dim denyBody = JsonConvert.SerializeObject(New With {
                                        .success = False,
                                        .error = "未授权：缺少或无效的 session token",
                                        .hint = "sessionToken 仅下发给可信来源（本机宿主页面 / localhost / 服务自身局域网 IP 的同源页面），请从移动端页面同源获取后，在 X-Session-Token header 或 ?token= 查询参数中传递"
                                    })
                                    SendHttpResponse(stream, 401, "application/json; charset=utf-8", denyBody, corsOrigin)
                                Catch ex As Exception  ' 
                                End Try
                                Return
                            End If
                        End If

                        ' 路由分发
                        Dim responseBody As String = ""
                        Dim contentType As String = "application/json; charset=utf-8"
                        Dim statusCode As Integer = 200
                        Dim binaryBody As Byte() = Nothing
                        Dim useBinary As Boolean = False

                        Try
                            Select Case True
                                Case path = "" OrElse path = "/"
                                    responseBody = GetMobilePageHtml()
                                    contentType = "text/html; charset=utf-8"
                                Case method = "OPTIONS"
                                    ' CORS 预检请求：204 + 完整 CORS 头，避免 fetch 走到 ERR_EMPTY_RESPONSE
                                    statusCode = 204
                                    responseBody = ""
                                Case path = "/api/status" AndAlso method = "GET"
                                    responseBody = GetApiStatus(request)
                                Case path = "/api/upload" AndAlso method = "POST"
                                    responseBody = ProcessApiUpload(request)
                                Case path = "/api/ocr-result" AndAlso method = "GET"
                                    responseBody = GetApiOcrResult()
                                ' === 任务规划路由 ===
                                Case path = "/task-planner" AndAlso method = "GET"
                                    responseBody = GetTaskPlannerHtml()
                                    contentType = "text/html; charset=utf-8"
                                Case path = "/api/tasks" AndAlso method = "GET"
                                    responseBody = TaskSchedulerBridge.Instance.GetAllTasks()
                                Case path = "/api/tasks" AndAlso method = "POST"
                                    responseBody = TaskSchedulerBridge.Instance.CreateTask(request.Body)
                                Case path = "/api/tasks/execute" AndAlso method = "POST"
                                    ' 执行指定任务
                                    Dim body = ParseJsonBody(request.Body)
                                    Dim taskId = If(body IsNot Nothing, body("taskId")?.ToString(), "")
                                    If String.IsNullOrEmpty(taskId) Then
                                        responseBody = JsonConvert.SerializeObject(New With {.success = False, .message = "缺少 taskId 参数"})
                                    Else
                                        responseBody = TaskSchedulerBridge.Instance.ExecuteTask(taskId)
                                    End If
                                Case path.StartsWith("/api/tasks/") AndAlso path.EndsWith("/execute") AndAlso method = "POST"
                                    Dim tId = path.Replace("/api/tasks/", "").Replace("/execute", "")
                                    responseBody = TaskSchedulerBridge.Instance.ExecuteTask(tId)
                                Case path.StartsWith("/api/tasks/") AndAlso method = "PUT"
                                    Dim tId2 = path.Replace("/api/tasks/", "")
                                    responseBody = TaskSchedulerBridge.Instance.UpdateTask(tId2, request.Body)
                                Case path.StartsWith("/api/tasks/") AndAlso method = "DELETE"
                                    Dim tId3 = path.Replace("/api/tasks/", "")
                                    responseBody = TaskSchedulerBridge.Instance.DeleteTask(tId3)
                                ' === 番茄钟路由（独立版 ↔ Excel 内嵌版 双向同步） ===
                                Case path = "/pomodoro" AndAlso method = "GET"
                                    responseBody = GetPomodoroHtml()
                                    contentType = "text/html; charset=utf-8"
                                Case path = "/api/pomodoro/get" AndAlso method = "GET"
                                    responseBody = PomodoroGet()
                                Case path = "/api/pomodoro/set" AndAlso method = "POST"
                                    responseBody = PomodoroSet(request.Body)
                                Case path = "/api/pomodoro/poll" AndAlso method = "GET"
                                    responseBody = PomodoroPoll(request)
                                ' === 独立服务（9529）进程管理 ===
                                Case path = "/api/standalone/status" AndAlso method = "GET"
                                    responseBody = GetStandaloneAgentStatus()
                                Case path = "/api/standalone/start" AndAlso method = "POST"
                                    responseBody = StartStandaloneAgent()
                                Case path = "/api/standalone/stop" AndAlso method = "POST"
                                    responseBody = StopStandaloneAgent()
                                Case path = "/api/browse/excel-file" AndAlso method = "POST"
                                    Dim bfilter As String = "", btitle As String = "", binitDir As String = ""
                                    Try
                                        If Not String.IsNullOrEmpty(request.Body) Then
                                            Dim jbody = JObject.Parse(request.Body)
                                            bfilter = If(jbody("filter") IsNot Nothing, jbody("filter").ToString(), "")
                                            btitle = If(jbody("title") IsNot Nothing, jbody("title").ToString(), "")
                                            binitDir = If(jbody("initialDir") IsNot Nothing, jbody("initialDir").ToString(), "")
                                        End If
                                    Catch : End Try
                                    Try
                                        responseBody = Bridges.RagBridge.BrowseExcelFile(bfilter, btitle, binitDir)
                                    Catch browseEx As Exception
                                        responseBody = JsonConvert.SerializeObject(New With {
                                            .success = False,
                                            .message = "弹出文件选择框失败：" & browseEx.Message,
                                            .[error] = browseEx.Message
                                        })
                                    End Try
                                Case path = "/api/host/workbook" AndAlso method = "GET"
                                    responseBody = RagBridge.GetHostWorkbook()
                                Case path = "/api/llm/export-config" AndAlso method = "POST"
                                    responseBody = ConfigBridge.Instance.ExportLLMConfigForStandalone()
                                Case path = "/api/llm/save-config" AndAlso method = "POST"
                                    ' 1) 保存 LLM 配置到 VSTO 存储（ConfigurationManager）
                                    Dim saveBody1 As String = If(request.Body, "")
                                    Dim step1 As String = ConfigBridge.Instance.SaveLLMConfig(saveBody1)
                                    ' 2) 紧接着导出明文到独立服务可读的位置
                                    Dim step2 As String = ConfigBridge.Instance.ExportLLMConfigForStandalone()
                                    ' 合并返回
                                    Try
                                        Dim r1 = JObject.Parse(step1)
                                        Dim r2 = JObject.Parse(step2)
                                        If CBool(r1("success")) = False Then
                                            responseBody = JsonConvert.SerializeObject(New With {.success = False, .error = If(CStr(r1("error")), "保存失败")})
                                        Else
                                            r1("exported") = r2
                                            responseBody = r1.ToString(Formatting.None)
                                        End If
                                    Catch ex As Exception  ' 
                                        responseBody = JsonConvert.SerializeObject(New With {.success = False, .error = "保存配置并导出失败"})
                                    End Try
                                Case path = "/api/llm/config" AndAlso method = "GET"
                                    ' 读回当前 LLM 配置（不回显 apiKey 明文，只显示是否已填），供面板显示当前模型
                                    Try
                                        Dim cfg = ConfigurationManager.Instance
                                        If cfg Is Nothing Then Throw New Exception("未初始化")
                                        Dim cbi = ConfigBridge.Instance
                                        Dim provider = cbi.ReadLLMConfig(cfg, Utils.ConfigConstants.LLM.PROVIDER, "LLM_PROVIDER", "")
                                        Dim model = cbi.ReadLLMConfig(cfg, Utils.ConfigConstants.LLM.MODEL, "LLM_MODEL", "")
                                        Dim apiUrl = cbi.ReadLLMConfig(cfg, Utils.ConfigConstants.LLM.API_URL, "LLM_API_URL", "")
                                        Dim apiKey = cbi.ReadLLMConfig(cfg, Utils.ConfigConstants.LLM.API_KEY, "LLM_API_KEY", "")
                                        Dim masked = If(String.IsNullOrEmpty(apiKey), "", If(apiKey.Length <= 6, "******", apiKey.Substring(0, 3) & New String("*"c, Math.Max(0, apiKey.Length - 3))))
                                        responseBody = JsonConvert.SerializeObject(New With {
                                            .success = True,
                                            .provider = provider,
                                            .model = model,
                                            .apiUrl = apiUrl,
                                            .apiKeyMasked = masked,
                                            .hasApiKey = Not String.IsNullOrEmpty(apiKey)
                                        })
                                    Catch ex As Exception
                                        responseBody = JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
                                    End Try
                                Case method = "GET" AndAlso (path.StartsWith("/css/") OrElse path.StartsWith("/js/") OrElse path.StartsWith("/lib/") OrElse path.StartsWith("/images/"))
                                    Dim servedBytes As Byte() = Nothing
                                    Dim servedCT As String = Nothing
                                    Dim servedOk As Boolean = TryServeStaticFile(path, servedBytes, servedCT)
                                    If servedOk Then
                                        binaryBody = servedBytes
                                        contentType = servedCT
                                        useBinary = True
                                        responseBody = ""
                                    Else
                                        statusCode = 404
                                        responseBody = "{""error"":""Static Not Found"",""path"":" & JsonConvert.ToString(path) & "}"
                                    End If
                                Case Else
                                    statusCode = 404
                                    responseBody = JsonConvert.SerializeObject(New With {.error = "Not Found"})
                            End Select
                        Catch ex As Exception
                            ' 路由分发异常兜底：写成 500 JSON，走正常响应流程，别裸断 socket。
                            statusCode = 500
                            responseBody = JsonConvert.SerializeObject(New With {
                                .success = False,
                                .message = "路由处理异常: " & ex.Message,
                                .[error] = ex.Message
                            })
                            System.Diagnostics.Debug.WriteLine("[MobileControlBridge] 路由异常: " & ex.ToString())
                        End Try

                        ' 全局异常兜底：只要路由已经跑完还没响应，就保证一定写 HTTP 响应回浏览器，
                        ' 不能直接关 socket —— 否则前端会报 net::ERR_EMPTY_RESPONSE，CORS 预检也直接挂。
                        If useBinary AndAlso binaryBody IsNot Nothing Then
                            Try : SendHttpResponseBytes(stream, statusCode, contentType, binaryBody, corsOrigin)
                            Catch bEx As Exception
                                Try : SendHttpResponse(stream, 500, "application/json; charset=utf-8",
                                    JsonConvert.SerializeObject(New With {.success = False, .[error] = "响应写入失败: " & bEx.Message}), corsOrigin)
                                Catch : End Try
                            End Try
                        Else
                            Try : SendHttpResponse(stream, statusCode, contentType, responseBody, corsOrigin)
                            Catch bEx As Exception
                                System.Diagnostics.Debug.WriteLine("[MobileControlBridge] 发送响应失败: " & bEx.Message)
                            End Try
                        End If
                    End Using
                End Using
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[MobileControlBridge] HandleClient 异常: " & ex.Message)
                ' 路由层异常也要回一个合法 HTTP 响应，别让浏览器看到 ERR_EMPTY_RESPONSE
                Try
                    Dim fallback As String = JsonConvert.SerializeObject(New With {
                        .success = False,
                        .[error] = ex.Message,
                        .message = "服务器处理失败: " & ex.Message
                    })
                    ' 此时 stream / client 可能已释放，只能尽力
                    SendHttpResponse(Nothing, 500, "application/json; charset=utf-8", fallback)
                Catch : End Try
            Finally
                SyncLock _clientLock
                    _activeClients.Remove(client)
                End SyncLock
            End Try
        End Sub

        ''' <summary>
        ''' 简易 HTTP 请求解析（全部用 StreamReader 读取，避免缓冲区问题）
        ''' </summary>
        Private Function ReadHttpRequest(stream As NetworkStream) As HttpRequestInfo
            Try
                Dim reader As New StreamReader(stream, Encoding.UTF8, False, 1024, True)

                ' 读取请求行
                Dim requestLine = reader.ReadLine()
                If String.IsNullOrEmpty(requestLine) Then Return Nothing

                Dim parts = requestLine.Split(" "c)
                If parts.Length < 2 Then Return Nothing

                Dim info As New HttpRequestInfo()
                info.Method = parts(0)
                info.Path = parts(1)
                info.RawPath = info.Path  ' 保留含 query 的原始路径（给 poll 路由解析 sinceVersion）

                ' 去除查询字符串，避免缓存版本号影响路由
                Dim queryIdx = info.Path.IndexOf("?"c)
                If queryIdx >= 0 Then
                    info.Path = info.Path.Substring(0, queryIdx)
                End If

                System.Diagnostics.Debug.WriteLine($"[MobileControlBridge] 收到请求: {info.Method} {info.Path}")

                ' 读取请求头
                Dim headers As New Dictionary(Of String, String)()
                Dim contentLength As Integer = 0
                Dim line = reader.ReadLine()
                While Not String.IsNullOrEmpty(line)
                    Dim colonIdx = line.IndexOf(":"c)
                    If colonIdx > 0 Then
                        Dim key = line.Substring(0, colonIdx).Trim()
                        Dim val = line.Substring(colonIdx + 1).Trim()
                        headers(key.ToLower()) = val
                        If key.ToLower() = "content-length" Then
                            Integer.TryParse(val, contentLength)
                        End If
                    End If
                    line = reader.ReadLine()
                End While

                info.Headers = headers

                ' 读取请求体（继续用 reader 读取，避免 StreamReader 缓冲区导致数据丢失）
                If contentLength > 0 Then
                    Dim bodyChars(contentLength - 1) As Char
                    Dim totalRead = 0
                    While totalRead < contentLength
                        Dim read = reader.Read(bodyChars, totalRead, contentLength - totalRead)
                        If read = 0 Then Exit While
                        totalRead += read
                    End While
                    info.Body = New String(bodyChars, 0, totalRead)
                    System.Diagnostics.Debug.WriteLine($"[MobileControlBridge] 请求体大小: {totalRead} 字符")
                End If

                Return info
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[MobileControlBridge] ReadHttpRequest 异常: " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 发送 HTTP 响应
        ''' </summary>
        ''' <param name="corsOrigin">
        ''' 允许读取本响应的来源。传 "*" 表示不限制（仅用于无 Origin 的非浏览器客户端，
        ''' 如手机端同源请求 / 本机 curl）；传空字符串表示**不输出 ACAO 头**，
        ''' 浏览器会因此阻断响应体 —— 用于拦截不可信站点的跨站读取（见 ResolveCorsOrigin）。
        ''' </param>
        Private Sub SendHttpResponse(stream As NetworkStream, statusCode As Integer, contentType As String, body As String, Optional corsOrigin As String = "*")
            If stream Is Nothing Then Return
            Try
                Dim statusText = If(statusCode = 200, "OK", If(statusCode = 404, "Not Found", "Error"))
                Dim bodyBytes = Encoding.UTF8.GetBytes(body)

                Dim sb As New StringBuilder()
                sb.AppendLine($"HTTP/1.1 {statusCode} {statusText}")
                sb.AppendLine($"Content-Type: {contentType}")
                sb.AppendLine($"Content-Length: {bodyBytes.Length}")
                If Not String.IsNullOrEmpty(corsOrigin) Then sb.AppendLine($"Access-Control-Allow-Origin: {corsOrigin}")
                sb.AppendLine("Access-Control-Allow-Methods: GET, POST, PUT, DELETE, OPTIONS")
                sb.AppendLine("Access-Control-Allow-Headers: Content-Type, Authorization, Accept, X-Requested-With")
                sb.AppendLine("Access-Control-Expose-Headers: Content-Type, Content-Length")
                If statusCode = 204 Then sb.AppendLine("Access-Control-Max-Age: 86400")
                sb.AppendLine("Cache-Control: no-cache, no-store, must-revalidate")
                sb.AppendLine("Pragma: no-cache")
                sb.AppendLine("Expires: 0")
                sb.AppendLine("Connection: close")
                sb.AppendLine()

                Dim headerBytes = Encoding.UTF8.GetBytes(sb.ToString())
                stream.Write(headerBytes, 0, headerBytes.Length)
                stream.Write(bodyBytes, 0, bodyBytes.Length)
                stream.Flush()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[MobileControlBridge] SendHttpResponse 异常: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' 发送二进制 HTTP 响应（用于静态文件：CSS/JS/图片，避免 UTF-8 转码破坏字节）
        ''' </summary>
        Private Sub SendHttpResponseBytes(stream As NetworkStream, statusCode As Integer, contentType As String, body As Byte(), Optional corsOrigin As String = "*")
            If stream Is Nothing Then Return
            Try
                Dim statusText = If(statusCode = 200, "OK", If(statusCode = 404, "Not Found", "Error"))
                Dim length As Long = If(body IsNot Nothing, CLng(body.Length), 0L)
                Dim sb As New StringBuilder()
                sb.AppendLine($"HTTP/1.1 {statusCode} {statusText}")
                sb.AppendLine($"Content-Type: {contentType}")
                sb.AppendLine($"Content-Length: {length}")
                If Not String.IsNullOrEmpty(corsOrigin) Then sb.AppendLine($"Access-Control-Allow-Origin: {corsOrigin}")
                sb.AppendLine("Access-Control-Allow-Methods: GET, POST, PUT, DELETE, OPTIONS")
                sb.AppendLine("Access-Control-Allow-Headers: Content-Type")
                ' 静态文件允许短期缓存（10 秒），避免频繁刷新时重复拉取
                sb.AppendLine("Cache-Control: public, max-age=10")
                sb.AppendLine("Connection: close")
                sb.AppendLine()

                Dim headerBytes = Encoding.UTF8.GetBytes(sb.ToString())
                stream.Write(headerBytes, 0, headerBytes.Length)
                If length > 0 Then
                    stream.Write(body, 0, CInt(length))
                End If
                stream.Flush()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[MobileControlBridge] SendHttpResponseBytes 异常: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' 安全地提供 UI 目录下的静态文件（仅限 css/js/lib/images 白名单目录 + 白名单扩展名，防 ../ 穿越）
        ''' </summary>
        Private Function TryServeStaticFile(path As String, ByRef outBytes As Byte(), ByRef outContentType As String) As Boolean
            outBytes = Nothing
            outContentType = "application/octet-stream"
            Try
                If String.IsNullOrEmpty(path) Then Return False
                ' 路径规范化：阻止所有 ".." 片段
                If path.Contains("..") Then Return False
                If Not (path.StartsWith("/css/") OrElse path.StartsWith("/js/") OrElse path.StartsWith("/lib/") OrElse path.StartsWith("/images/")) Then
                    Return False
                End If
                Dim ext As String = IO.Path.GetExtension(path).ToLowerInvariant()
                Dim allowedExt As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                allowedExt.Add(".css") : allowedExt.Add(".js") : allowedExt.Add(".mjs") : allowedExt.Add(".map") : allowedExt.Add(".json")
                allowedExt.Add(".png") : allowedExt.Add(".jpg") : allowedExt.Add(".jpeg") : allowedExt.Add(".gif") : allowedExt.Add(".svg")
                allowedExt.Add(".ico") : allowedExt.Add(".webp") : allowedExt.Add(".bmp")
                allowedExt.Add(".woff") : allowedExt.Add(".woff2") : allowedExt.Add(".ttf") : allowedExt.Add(".otf") : allowedExt.Add(".eot")
                If Not allowedExt.Contains(ext) Then Return False

                ' 拼物理路径（相对于 BaseDirectory/AssemblyDir 的 UI 目录）
                Dim relativePart As String = path.TrimStart("/"c).Replace("/"c, IO.Path.DirectorySeparatorChar)
                Dim baseDir As String = AppDomain.CurrentDomain.BaseDirectory
                Dim asmDir As String = IO.Path.GetDirectoryName(GetType(MobileControlBridge).Assembly.Location)
                Dim candidates As String() = {
                    IO.Path.Combine(baseDir, "UI", relativePart),
                    IO.Path.Combine(baseDir, relativePart),
                    IO.Path.Combine(asmDir, "UI", relativePart)
                }
                Dim finalPath As String = Nothing
                For Each c As String In candidates
                    If IO.File.Exists(c) Then
                        finalPath = c
                        Exit For
                    End If
                Next
                If String.IsNullOrEmpty(finalPath) Then
                    System.Diagnostics.Debug.WriteLine("[MobileControlBridge] TryServeStaticFile 未找到: " & path)
                    Return False
                End If

                ' 二次确认：解析后路径必须仍位于 UI 根目录下（防御路径穿越的最后一道闸）
                Dim uiRoot1 As String = IO.Path.GetFullPath(IO.Path.Combine(baseDir, "UI")).TrimEnd(IO.Path.DirectorySeparatorChar)
                Dim uiRoot2 As String = Nothing
                If Not String.IsNullOrEmpty(asmDir) Then
                    uiRoot2 = IO.Path.GetFullPath(IO.Path.Combine(asmDir, "UI")).TrimEnd(IO.Path.DirectorySeparatorChar)
                End If
                Dim fullFinal As String = IO.Path.GetFullPath(finalPath).TrimEnd(IO.Path.DirectorySeparatorChar)
                Dim underUi As Boolean = fullFinal.StartsWith(uiRoot1 & IO.Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
                If Not String.IsNullOrEmpty(uiRoot2) Then
                    underUi = underUi OrElse fullFinal.StartsWith(uiRoot2 & IO.Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
                End If
                If Not underUi Then
                    System.Diagnostics.Debug.WriteLine("[MobileControlBridge] TryServeStaticFile 目录穿越拦截: " & path)
                    Return False
                End If

                ' Content-Type 映射
                Select Case ext
                    Case ".css" : outContentType = "text/css; charset=utf-8"
                    Case ".js", ".mjs" : outContentType = "application/javascript; charset=utf-8"
                    Case ".json" : outContentType = "application/json; charset=utf-8"
                    Case ".map" : outContentType = "application/json; charset=utf-8"
                    Case ".png" : outContentType = "image/png"
                    Case ".jpg", ".jpeg" : outContentType = "image/jpeg"
                    Case ".gif" : outContentType = "image/gif"
                    Case ".svg" : outContentType = "image/svg+xml"
                    Case ".ico" : outContentType = "image/x-icon"
                    Case ".webp" : outContentType = "image/webp"
                    Case ".bmp" : outContentType = "image/bmp"
                    Case ".woff" : outContentType = "font/woff"
                    Case ".woff2" : outContentType = "font/woff2"
                    Case ".ttf" : outContentType = "font/ttf"
                    Case ".otf" : outContentType = "font/otf"
                    Case ".eot" : outContentType = "application/vnd.ms-fontobject"
                    Case Else : outContentType = "application/octet-stream"
                End Select

                outBytes = IO.File.ReadAllBytes(finalPath)
                Return True
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[MobileControlBridge] TryServeStaticFile 异常: " & ex.Message)
                Return False
            End Try
        End Function

#End Region

#Region "路由处理"

        ''' <summary>
        ''' GET / - 返回移动端 HTML 页面
        ''' </summary>
        Private Function GetMobilePageHtml() As String
            Try
                Dim htmlPath As String = Nothing
                Dim candidates = {
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "UI", "mobile-ocr.html"),
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "mobile-ocr.html"),
                    Path.Combine(Path.GetDirectoryName(GetType(MobileControlBridge).Assembly.Location), "UI", "mobile-ocr.html")
                }
                For Each c In candidates
                    If File.Exists(c) Then
                        htmlPath = c
                        Exit For
                    End If
                Next

                If htmlPath Is Nothing Then
                    Return "<html><body><h1>mobile-ocr.html not found</h1></body></html>"
                End If

                Dim html = File.ReadAllText(htmlPath, Encoding.UTF8)
                html = html.Replace("{{SERVER_IP}}", _localIp)
                html = html.Replace("{{SERVER_PORT}}", _port.ToString())
                Return html
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[MobileControlBridge] GetMobilePageHtml 异常: " & ex.Message)
                Return $"<html><body><h1>Error: {ex.Message}</h1></body></html>"
            End Try
        End Function

        ''' <summary>
        ''' GET /api/status - 返回服务状态
        ''' 注意：此处**不得**懒加载任何 COM 型重依赖（如 OcrBridge.Instance），否则 HTTP 健康检查会死锁。
        ''' OCR 引擎名从配置文件直接读取。
        ''' </summary>
        Private Function GetApiStatus(request As HttpRequestInfo) As String
            ' 安全：sessionToken 只对非跨站请求 / 可信来源下发。
            ' 第三方网页（Origin 为其它站点）拿不到 token，也就无法调用需鉴权的 /api/*。
            Dim token = If(CanDiscloseSessionToken(request), _sessionToken, "")
            Return JsonConvert.SerializeObject(New With {
                .success = True,
                .isRunning = True,
                .ocrEngine = ReadOcrEngineNameSafe(),
                .serverVersion = "1.0",
                .sessionToken = token
            })
        End Function

        ''' <summary>
        ''' 安全读取 OCR 引擎名：只读配置文件，不触发 OcrBridge 单例初始化，避免 HTTP 线程在健康检查上挂住。
        ''' </summary>
        Private Function ReadOcrEngineNameSafe() As String
            Try
                Dim cfg = ConfigurationManager.Instance
                If cfg IsNot Nothing Then
                    Dim v = cfg.GetSetting("OCR_ENGINE", "rapidocr")
                    If Not String.IsNullOrEmpty(v) Then Return v.ToLowerInvariant()
                End If
            Catch : End Try
            Return "rapidocr"
        End Function

        ''' <summary>
        ''' POST /api/upload - 接收 Base64 图片，执行 OCR
        ''' </summary>
        Private Function ProcessApiUpload(request As HttpRequestInfo) As String
            Try
                Dim json As JObject
                Try
                    json = JObject.Parse(request.Body)
                Catch ex As Exception  ' 
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "无效的 JSON 数据"})
                End Try

                Dim base64Image = json.Value(Of String)("image")
                If String.IsNullOrEmpty(base64Image) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "缺少 image 字段"})
                End If

                ' 去除 Data URI 前缀
                Dim base64Data As String = base64Image
                If base64Data.Contains(",") Then
                    base64Data = base64Data.Substring(base64Data.IndexOf(",") + 1)
                End If

                ' 解码并保存为临时文件
                Dim bytes = Convert.FromBase64String(base64Data)
                Dim timestamp = DateTime.Now.ToString("yyyyMMddHHmmssfff")
                Dim fileName = $"mobile_{timestamp}.jpg"
                Dim filePath = Path.Combine(_tempImageDir, fileName)
                File.WriteAllBytes(filePath, bytes)

                System.Diagnostics.Debug.WriteLine($"[MobileControlBridge] 收到移动端图片: {fileName}, 大小: {bytes.Length} 字节")

                ' 获取识别模式
                Dim mode = If(json.Value(Of String)("mode"), "auto")

                ' 通知桌面打开 OCR 面板并加载图片
                RaiseEvent UploadReceived(fileName)

                ' 设置图片到桌面 OCR 面板（通过 OcrBridge）
                Dim ocrSvc = OcrBridge.Instance
                ocrSvc.SetImageForRecognition(filePath, base64Image)

                ' 在桌面端执行识别（调用桌面的识别流程）
                Dim desktopResult = ocrSvc.RecognizeCurrentImage(mode)

                If desktopResult.success Then
                    ' 触发事件，通知桌面 OCR 面板显示结果
                    RaiseEvent OcrCompleted(desktopResult.text, fileName, base64Image)

                    ' 返回结果给手机
                    Return JsonConvert.SerializeObject(New With {
                        .success = True,
                        .text = desktopResult.text,
                        .fileName = fileName,
                        .message = "识别完成"
                    })
                Else
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = desktopResult.error,
                        .fileName = fileName
                    })
                End If

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[MobileControlBridge] ProcessApiUpload 异常: " & ex.Message)
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' GET /api/ocr-result - 轮询 OCR 结果（备用接口）
        ''' </summary>
        Private Function GetApiOcrResult() As String
            Try
                Return OcrBridge.Instance.GetOcrResult()
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

#End Region

#Region "辅助方法"

        ''' <summary>
        ''' 获取本机可用于移动端连接的 IP 地址。
        ''' 优先选择状态为 Up 的 WLAN/无线网卡；若无线不可用，再选其他活动网卡。
        ''' 避免在同时插网线和连热点时错误地选到手机无法访问的以太网 IP。
        ''' </summary>
        Private Function GetLocalIPAddress() As String
            Try
                Dim candidates As New List(Of String)
                Dim wirelessIp As String = ""
                Dim ethernetIp As String = ""

                For Each ni In NetworkInterface.GetAllNetworkInterfaces()
                    ' 只考虑已启用的真实物理接口
                    If ni.OperationalStatus <> OperationalStatus.Up Then Continue For

                    Dim type = ni.NetworkInterfaceType
                    If type = NetworkInterfaceType.Loopback Then Continue For
                    If type = NetworkInterfaceType.Tunnel Then Continue For

                    ' 排除常见虚拟/虚拟化网卡
                    Dim desc = (ni.Description & " " & ni.Name).ToLowerInvariant()
                    If desc.Contains("virtual") OrElse desc.Contains("vmware") OrElse
                       desc.Contains("hyper-v") OrElse desc.Contains("vethernet") OrElse
                       desc.Contains("wan miniport") Then Continue For

                    For Each addr In ni.GetIPProperties().UnicastAddresses
                        If addr.Address.AddressFamily <> Net.Sockets.AddressFamily.InterNetwork Then Continue For

                        Dim ipStr = addr.Address.ToString()
                        If ipStr.StartsWith("127.") Then Continue For

                        ' 优先 WLAN（WiFi 或手机热点）
                        If type = NetworkInterfaceType.Wireless80211 Then
                            wirelessIp = ipStr
                        ElseIf type = NetworkInterfaceType.Ethernet Then
                            If String.IsNullOrEmpty(ethernetIp) Then ethernetIp = ipStr
                        Else
                            candidates.Add(ipStr)
                        End If
                    Next
                Next

                If Not String.IsNullOrEmpty(wirelessIp) Then Return wirelessIp
                If Not String.IsNullOrEmpty(ethernetIp) Then Return ethernetIp

                ' 兜底：返回第一个非 APIPA 的候选
                For Each ip In candidates
                    If Not ip.StartsWith("169.254.") Then Return ip
                Next

                If candidates.Count > 0 Then Return candidates(0)

                ' 兜底：通过主机名解析获取一个非回环 IPv4
                Try
                    Dim host = Dns.GetHostEntry(Dns.GetHostName())
                    For Each ip In host.AddressList
                        If ip.AddressFamily = Net.Sockets.AddressFamily.InterNetwork AndAlso Not ip.ToString().StartsWith("127.") Then
                            Return ip.ToString()
                        End If
                    Next
                Catch ex As Exception  ' 
                End Try

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[MobileControlBridge] 获取本机 IP 失败: " & ex.Message)
            End Try
            Return ""
        End Function

#End Region

        ''' <summary>
        ''' GET /task-planner - 返回独立任务规划页面 HTML
        ''' </summary>
        Private Function GetTaskPlannerHtml() As String
            Try
                Dim htmlPath As String = Nothing
                Dim candidates = {
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "UI", "task-planner.html"),
                    Path.Combine(Path.GetDirectoryName(GetType(MobileControlBridge).Assembly.Location), "UI", "task-planner.html")
                }
                For Each c In candidates
                    If File.Exists(c) Then
                        htmlPath = c
                        Exit For
                    End If
                Next

                If htmlPath Is Nothing Then
                    Return "<html><body><h1>task-planner.html not found</h1><p>请确保文件位于 UI/task-planner.html</p></body></html>"
                End If

                Return File.ReadAllText(htmlPath, Encoding.UTF8)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[MobileControlBridge] GetTaskPlannerHtml 异常: " & ex.Message)
                Return $"<html><body><h1>Error: {ex.Message}</h1></body></html>"
            End Try
        End Function

        ''' <summary>
        ''' GET /pomodoro - 返回独立番茄钟页面 HTML
        ''' </summary>
        Private Function GetPomodoroHtml() As String
            Try
                Dim htmlPath As String = Nothing
                Dim candidates = {
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "UI", "pomodoro.html"),
                    Path.Combine(Path.GetDirectoryName(GetType(MobileControlBridge).Assembly.Location), "UI", "pomodoro.html")
                }
                For Each c In candidates
                    If File.Exists(c) Then
                        htmlPath = c
                        Exit For
                    End If
                Next

                If htmlPath Is Nothing Then
                    Return "<html><body><h1>pomodoro.html not found</h1><p>请确保文件位于 UI/pomodoro.html</p></body></html>"
                End If

                Return File.ReadAllText(htmlPath, Encoding.UTF8)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[MobileControlBridge] GetPomodoroHtml 异常: " & ex.Message)
                Return "<html><body><h1>Error: " & ex.Message & "</h1></body></html>"
            End Try
        End Function

        ''' <summary>
        ''' 把共享番茄钟状态包装为通用响应 JSON
        ''' </summary>
        Private Shared Function WrapPomodoroSuccess(success As Boolean, Optional extra As JObject = Nothing) As JObject
            Dim resp As New JObject()
            resp("success") = If(success, True, False)
            resp("version") = _pomodoroVersion
            resp("lastWriterId") = If(_pomodoroLastWriterId, "")
            resp("data") = If(_pomodoroState, New JObject())
            If extra IsNot Nothing Then
                For Each p In extra.Properties()
                    resp(p.Name) = p.Value
                Next
            End If
            Return resp
        End Function

        ''' <summary>
        ''' GET /api/pomodoro/get - 获取共享番茄钟状态
        ''' </summary>
        Private Function PomodoroGet() As String
            SyncLock _pomodoroLock
                Try
                    Return WrapPomodoroSuccess(True).ToString(Formatting.None)
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("[MobileControlBridge] PomodoroGet 异常: " & ex.Message)
                    Return "{""success"":false,""error"":" & JsonConvert.ToString(ex.Message) & "}"
                End Try
            End SyncLock
        End Function

        ''' <summary>
        ''' POST /api/pomodoro/set - 写入共享番茄钟状态（fromVersion 冲突检测）
        ''' </summary>
        Private Function PomodoroSet(body As String) As String
            SyncLock _pomodoroLock
                Try
                    Dim payload As JObject = Nothing
                    Try
                        If Not String.IsNullOrWhiteSpace(body) Then payload = JObject.Parse(body)
                    Catch ex As Exception  ' 
                        payload = Nothing
                    End Try
                    If payload Is Nothing Then
                        Return "{""success"":false,""error"":""invalid body""}"
                    End If

                    Dim fromVersion As Long = -1
                    Dim tok = payload("fromVersion")
                    If tok IsNot Nothing AndAlso tok.Type <> JTokenType.Null Then
                        Int64.TryParse(tok.ToString(), fromVersion)
                    End If

                    If fromVersion >= 0 AndAlso fromVersion <> _pomodoroVersion Then
                        Dim resp = WrapPomodoroSuccess(False)
                        resp("conflict") = True
                        resp("error") = "version conflict"
                        resp("fromVersion") = fromVersion
                        Return resp.ToString(Formatting.None)
                    End If

                    Dim dataTok = payload("data")
                    If dataTok IsNot Nothing AndAlso dataTok.Type = JTokenType.Object Then
                        _pomodoroState = CType(dataTok, JObject)
                    End If

                    Dim writerId As String = ""
                    Dim wTok = payload("writerId")
                    If wTok IsNot Nothing AndAlso wTok.Type <> JTokenType.Null Then
                        writerId = wTok.ToString()
                    End If
                    If Not String.IsNullOrEmpty(writerId) Then _pomodoroLastWriterId = writerId

                    Dim evt As String = ""
                    Dim eTok = payload("event")
                    If eTok IsNot Nothing AndAlso eTok.Type <> JTokenType.Null Then
                        evt = eTok.ToString()
                    End If
                    If Not String.IsNullOrEmpty(evt) Then
                        Dim logEntry = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") & "|" & writerId & "|" & evt
                        _pomodoroEvents.Add(logEntry)
                        If _pomodoroEvents.Count > 100 Then _pomodoroEvents.RemoveAt(0)
                    End If

                    _pomodoroVersion += 1

                    Return WrapPomodoroSuccess(True).ToString(Formatting.None)
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("[MobileControlBridge] PomodoroSet 异常: " & ex.Message)
                    Return "{""success"":false,""error"":" & JsonConvert.ToString(ex.Message) & "}"
                End Try
            End SyncLock
        End Function

        ''' <summary>
        ''' GET /api/pomodoro/poll?sinceVersion=N - 轮询版本变化
        ''' </summary>
        Private Function PomodoroPoll(request As HttpRequestInfo) As String
            SyncLock _pomodoroLock
                Try
                    Dim sinceVersion As Long = -1
                    Dim rawPath As String = ""
                    If request IsNot Nothing Then
                        If Not String.IsNullOrEmpty(request.RawPath) Then
                            rawPath = request.RawPath
                        ElseIf Not String.IsNullOrEmpty(request.Path) Then
                            rawPath = request.Path
                        End If
                    End If
                    Dim qIdx = rawPath.IndexOf("?"c)
                    If qIdx >= 0 Then
                        Dim qs = rawPath.Substring(qIdx + 1)
                        For Each pair In qs.Split("&"c)
                            If String.IsNullOrEmpty(pair) Then Continue For
                            Dim kv = pair.Split("="c)
                            If kv.Length >= 2 AndAlso kv(0) = "sinceVersion" Then
                                Long.TryParse(Uri.UnescapeDataString(kv(1)), sinceVersion)
                            End If
                        Next
                    End If
                    Dim changed As Boolean = If(sinceVersion < 0, True, (sinceVersion <> _pomodoroVersion))
                    Dim resp = WrapPomodoroSuccess(If(changed, True, False))
                    resp("changed") = changed
                    resp("sinceVersion") = sinceVersion
                    Return resp.ToString(Formatting.None)
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("[MobileControlBridge] PomodoroPoll 异常: " & ex.Message)
                    Return "{""success"":false,""error"":" & JsonConvert.ToString(ex.Message) & "}"
                End Try
            End SyncLock
        End Function

#Region "独立服务（9529）进程管理"

        Private _standaloneProc As System.Diagnostics.Process

        ''' <summary>
        ''' 检测独立服务是否在运行（TCP 连接 9529）
        ''' </summary>
        Private Function IsStandaloneAgentRunning() As Boolean
            Try
                Using client As New TcpClient()
                    Dim ar = client.BeginConnect("127.0.0.1", STANDALONE_AGENT_PORT, Nothing, Nothing)
                    If Not ar.AsyncWaitHandle.WaitOne(800) Then Return False
                    client.EndConnect(ar)
                    Return True
                End Using
            Catch ex As Exception  ' 
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 解析独立服务脚本的绝对路径
        ''' 优先级：安装目录 → 开发源码目录
        ''' </summary>
        Private Function ResolveStandaloneAgentScript() As String
            Dim candidates As New List(Of String)
            candidates.Add(Path.Combine(PathHelper.PluginInstallPath, "StandaloneAgent", "standalone_agent.py"))
            candidates.Add(Path.Combine(PathHelper.PluginInstallPath, "..", "StandaloneAgent", "standalone_agent.py"))
            candidates.Add(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "StandaloneAgent", "standalone_agent.py"))
            For Each c In candidates
                Try
                    Dim p = Path.GetFullPath(c)
                    If File.Exists(p) Then Return p
                Catch ex As Exception  ' 
                End Try
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' GET /api/standalone/status - 独立服务运行状态
        ''' </summary>
        <ComVisible(True)>
        Public Function GetStandaloneAgentStatus() As String
            Try
                Dim running = IsStandaloneAgentRunning()
                Dim script = ResolveStandaloneAgentScript()
                Dim cfgOk = File.Exists(Path.Combine(PathHelper.ConfigDirectory, "llm_config.json"))
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .running = running,
                    .port = STANDALONE_AGENT_PORT,
                    .scriptFound = (script IsNot Nothing),
                    .configFound = cfgOk
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' POST /api/standalone/start - 启动独立服务（拉起 python 进程）
        ''' </summary>
        <ComVisible(True)>
        Public Function StartStandaloneAgent() As String
            Try
                If IsStandaloneAgentRunning() Then
                    ' 检查老进程是否缺失配置（通常是之前启动时用了老脚本、没传 --config）
                    ' 如果缺 llm_config 就视为 stale：停掉 + 重新启动（会强制重导出）
                    Dim stale As Boolean = False
                    Try
                        Using http As New System.Net.Http.HttpClient()
                            http.Timeout = TimeSpan.FromSeconds(2)
                            Try
                                Dim hjson = http.GetStringAsync("http://127.0.0.1:" & STANDALONE_AGENT_PORT.ToString() & "/health").Result
                                If Not String.IsNullOrEmpty(hjson) Then
                                    Dim hj = JObject.Parse(hjson)
                                    Dim cf = hj("config_found")
                                    If cf Is Nothing OrElse cf.ToObject(Of Boolean)() = False Then stale = True
                                End If
                            Catch : stale = True ' 读不到 health 视作 stale
                            End Try
                        End Using
                    Catch : End Try
                    If Not stale Then
                        Return JsonConvert.SerializeObject(New With {.success = True, .alreadyRunning = True, .message = "独立服务已在运行"})
                    End If
                    ' 停掉 stale 进程
                    Try : StopStandaloneAgent() : Catch : End Try
                    Thread.Sleep(400)
                End If

                Dim script = ResolveStandaloneAgentScript()
                If String.IsNullOrEmpty(script) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "未找到 standalone_agent.py（请确认 StandaloneAgent 目录存在）"})
                End If

                ' ① 启动前强制导出 LLM 配置：让 standalone_agent.py 启动时一定能找到 llm_config.json
                '    （即使之前没在 Excel 里保存过 AI 设置，也会把当前 ConfigurationManager 的值落盘）
                Dim configExportMsg As String = ""
                Try
                    Dim exportRes = ConfigBridge.Instance.ExportLLMConfigForStandalone()
                    If Not String.IsNullOrEmpty(exportRes) Then
                        Try
                            Dim j = JObject.Parse(exportRes)
                            If j("success") Is Nothing OrElse j("success").ToObject(Of Boolean)() = False Then
                                configExportMsg = "（警告：导出配置失败，请先在 Excel 里保存一次 AI 模型设置）"
                            End If
                        Catch : End Try
                    End If
                Catch : End Try

                Dim pythonExe As String = Nothing
                For Each p In PathHelper.GetPythonSearchPaths()
                    If File.Exists(p) Then pythonExe = p : Exit For
                Next
                If String.IsNullOrEmpty(pythonExe) Then pythonExe = "python"

                ' ② 给 python 明确传 --config 参数（指向 ConfigDirectory 下刚导出的 llm_config.json），
                '    消除 WorkingDirectory / 候选路径相对搜索失败的问题。
                Dim scriptDir = Path.GetDirectoryName(script)
                Dim explicitConfig As String = ""
                Try
                    Dim cfgDir = Utils.PathHelper.ConfigDirectory
                    If Not String.IsNullOrEmpty(cfgDir) Then
                        Dim cp = Path.Combine(cfgDir, "llm_config.json")
                        If File.Exists(cp) Then explicitConfig = cp
                    End If
                Catch : End Try
                ' 兜底：脚本目录的 ../Config/llm_config.json
                If String.IsNullOrEmpty(explicitConfig) Then
                    Try
                        Dim fallback = Path.GetFullPath(Path.Combine(scriptDir, "..", "Config", "llm_config.json"))
                        If File.Exists(fallback) Then explicitConfig = fallback
                    Catch : End Try
                End If

                Dim args As String
                If Not String.IsNullOrEmpty(explicitConfig) Then
                    args = """" & script & """ --config """ & explicitConfig & """"
                Else
                    args = """" & script & """"
                End If

                Dim psi As New System.Diagnostics.ProcessStartInfo(pythonExe, args) With {
                    .UseShellExecute = False,
                    .CreateNoWindow = True,
                    .WorkingDirectory = scriptDir,
                    .RedirectStandardOutput = False,
                    .RedirectStandardError = False
                }
                _standaloneProc = System.Diagnostics.Process.Start(psi)

                ' 等待端口就绪（最多 8 秒，比以前更长，避免刚启动被误判）
                Dim deadline = DateTime.UtcNow.AddSeconds(8)
                While DateTime.UtcNow < deadline
                    If IsStandaloneAgentRunning() Then
                        ' 服务刚起来 -> 一次性把所有任务推送过去（含 TargetFile）
                        Try : TaskSchedulerBridge.SyncAllToStandaloneAgent() : Catch : End Try
                        Dim msg As String = "独立服务已启动"
                        If Not String.IsNullOrEmpty(configExportMsg) Then msg &= configExportMsg
                        Return JsonConvert.SerializeObject(New With {
                            .success = True, .message = msg, .pid = _standaloneProc?.Id, .configPath = explicitConfig
                        })
                    End If
                    Thread.Sleep(300)
                End While
                Return JsonConvert.SerializeObject(New With {.success = False, .error = "独立服务启动超时（请确认 python 可用、已在 Excel 中保存过 AI 模型设置：设置 → AI 模型 → 保存）"})
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = "启动失败: " & ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' POST /api/standalone/stop - 停止独立服务
        ''' </summary>
        <ComVisible(True)>
        Public Function StopStandaloneAgent() As String
            Try
                ' 优先调用独立服务自身的 /shutdown（干净退出并保存任务）
                Try
                    Using client As New System.Net.Http.HttpClient
                        client.Timeout = TimeSpan.FromSeconds(5)
                        Dim resp = client.PostAsync("http://127.0.0.1:" & STANDALONE_AGENT_PORT.ToString() & "/shutdown", Nothing).GetAwaiter().GetResult()
                    End Using
                    ' 等待端口释放
                    For i As Integer = 0 To 9
                        If Not IsStandaloneAgentRunning() Then Exit For
                        Thread.Sleep(300)
                    Next
                Catch ex As Exception  ' 
                End Try

                ' 兜底：若仍存活，kill 掉持有的进程句柄
                If _standaloneProc IsNot Nothing AndAlso Not _standaloneProc.HasExited Then
                    Try : _standaloneProc.Kill() : Catch : End Try
                    _standaloneProc.Dispose()
                    _standaloneProc = Nothing
                End If

                Return JsonConvert.SerializeObject(New With {.success = True, .message = "独立服务已停止"})
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = "停止失败: " & ex.Message})
            End Try
        End Function

#End Region

#Region "鉴权辅助"

        ''' <summary>
        ''' 判断给定请求路径是否属于公开路径（无需 session token）。
        ''' 公开路径：HTML 页面、静态资源、CORS 预检、健康检查。
        ''' 其他所有 API（/api/* 等）均需鉴权。
        ''' </summary>
        Private Shared Function IsPublicPath(path As String, method As String) As Boolean
            ' CORS 预检永远放行
            If method = "OPTIONS" Then Return True

            ' HTML 页面公开
            If path = "" OrElse path = "/" OrElse path = "/task-planner" OrElse path = "/pomodoro" Then Return True

            ' 健康检查
            If path = "/api/status" Then Return True

            ' 静态文件公开（css/js/lib/images 等目录）
            If path.StartsWith("/css/") OrElse path.StartsWith("/js/") OrElse
               path.StartsWith("/lib/") OrElse path.StartsWith("/images/") Then Return True

            Return False
        End Function

        ''' <summary>
        ''' 判定请求来源是否可信（可读取响应体 / 可获取 session token）。
        ''' 可信 = 本机宿主页面（app.excelai.local 等虚拟主机名）、localhost/127.0.0.1/::1、
        ''' 以及服务自身对外暴露的局域网 IP（手机端页面就是从它加载的，属同源）。
        ''' </summary>
        Private Function IsTrustedOrigin(origin As String) As Boolean
            If String.IsNullOrWhiteSpace(origin) Then Return False
            Dim host As String = ""
            Try
                Dim u As New Uri(origin.Trim())
                host = u.Host
            Catch
                ' 非法 URL：按不可信处理
                Return False
            End Try
            If String.IsNullOrEmpty(host) Then Return False

            Dim trusted As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            trusted.Add("localhost")
            trusted.Add("127.0.0.1")
            trusted.Add("::1")
            trusted.Add("app.excelai.local")
            trusted.Add("screenassets.local")
            If Not String.IsNullOrEmpty(_localIp) Then trusted.Add(_localIp)

            ' [::1] 形式
            If host.StartsWith("[") AndAlso host.EndsWith("]") Then host = host.Substring(1, host.Length - 2)
            Return trusted.Contains(host)
        End Function

        ''' <summary>
        ''' 计算响应应输出的 Access-Control-Allow-Origin。
        ''' · 无 Origin 头（手机浏览器同源 GET / 原生客户端 / curl）→ "*"，保持原有可用性；
        ''' · Origin 可信（本机宿主页面、localhost、手机端同源的局域网 IP）→ 回显该 Origin；
        ''' · Origin 不可信（任意第三方网页）→ 空串，即**不输出 ACAO**，浏览器将阻断响应体。
        '''
        ''' 【安全修复 2026-10-09】此前无条件输出 "Access-Control-Allow-Origin: *"，
        ''' 而 /api/status 又是公开路径且直接返回 sessionToken —— 任意网页只要
        ''' fetch('http://127.0.0.1:9528/api/status') 就能拿到 token，进而调用
        ''' /api/upload 等接口读写本机的 Excel/文档。现在 token 与响应体都受来源约束。
        ''' </summary>
        Private Function ResolveCorsOrigin(request As HttpRequestInfo) As String
            If request Is Nothing OrElse request.Headers Is Nothing Then Return "*"
            Dim origin As String = ""
            If request.Headers.TryGetValue("origin", origin) AndAlso Not String.IsNullOrWhiteSpace(origin) Then
                origin = origin.Trim()
                Return If(IsTrustedOrigin(origin), origin, "")
            End If
            ' 无 Origin：非跨站场景（同源请求 / 原生 HTTP 客户端），保持兼容
            Return "*"
        End Function

        ''' <summary>
        ''' 判断当前请求是否「可被告知 session token」。
        ''' 只有非跨站请求或可信来源才能拿到 token，第三方网页拿不到。
        ''' </summary>
        Private Function CanDiscloseSessionToken(request As HttpRequestInfo) As Boolean
            Dim cors = ResolveCorsOrigin(request)
            Return Not String.IsNullOrEmpty(cors)
        End Function

        ''' <summary>
        ''' 从请求的 Header 或 QueryString 中提取并验证 session token。
        ''' 接受 X-Session-Token / X-Auth-Token header 或 ?token= / ?sessionToken= 查询参数。
        ''' </summary>
        Private Function ValidateRequestToken(request As HttpRequestInfo) As Boolean
            If String.IsNullOrEmpty(_sessionToken) Then
                ' token 尚未生成（服务未正常启动？），拒绝所有请求
                Return False
            End If

            Dim presented As String = ""

            ' 1) 从 Header 读取（大小写不敏感）
            If request.Headers IsNot Nothing Then
                If request.Headers.TryGetValue("x-session-token", presented) Then
                    presented = presented?.Trim()
                End If
                If String.IsNullOrEmpty(presented) AndAlso request.Headers.TryGetValue("x-auth-token", presented) Then
                    presented = presented?.Trim()
                End If
            End If

            ' 2) 从 QueryString 读取（RawPath 含查询串）
            If String.IsNullOrEmpty(presented) AndAlso Not String.IsNullOrEmpty(request.RawPath) Then
                Dim qIdx = request.RawPath.IndexOf("?"c)
                If qIdx >= 0 Then
                    Dim qs = request.RawPath.Substring(qIdx + 1)
                    For Each pair In qs.Split("&"c)
                        Dim eq = pair.IndexOf("="c)
                        If eq <= 0 Then Continue For
                        Dim k = pair.Substring(0, eq).ToLowerInvariant()
                        Dim v = pair.Substring(eq + 1)
                        If (k = "token" OrElse k = "sessiontoken") AndAlso Not String.IsNullOrEmpty(v) Then
                            presented = Uri.UnescapeDataString(v)
                            Exit For
                        End If
                    Next
                End If
            End If

            If String.IsNullOrEmpty(presented) Then Return False
            Return String.Equals(presented, _sessionToken, StringComparison.Ordinal)
        End Function

#End Region

        ''' <summary>
        ''' 解析 JSON 请求体为 Dictionary
        ''' </summary>
        Private Function ParseJsonBody(body As String) As Dictionary(Of String, Object)
            If String.IsNullOrWhiteSpace(body) Then Return Nothing
            Try
                Return JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(body)
            Catch ex As Exception  ' 
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' HTTP 请求信息
        ''' </summary>
        Private Class HttpRequestInfo
            Public Property Method As String
            Public Property Path As String
            ''' <summary>原始请求路径（含查询字符串，如 /api/pomodoro/poll?sinceVersion=0）</summary>
            Public Property RawPath As String
            Public Property Headers As Dictionary(Of String, String)
            Public Property Body As String
        End Class

    End Class

End Namespace