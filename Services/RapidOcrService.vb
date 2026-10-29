Imports System
Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Diagnostics
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services

    ''' <summary>
    ''' RapidOCR 推理引擎进程管理器
    ''' 管理 Python 子进程生命周期，通过 JSON stdin/stdout 协议与 Python 引擎通信
    ''' 【架构】长驻进程模式：启动后保持运行，避免每次识别的模型加载开销
    ''' </summary>
    Friend Class RapidOcrEngineProcess
        Implements IDisposable

        Private _pythonPath As String
        Private _engineScriptPath As String
        Private _process As Process
        Private _processLock As New Object()
        Private _stdinWriter As StreamWriter
        Private _stdoutReader As StreamReader
        Private _stderrReader As StreamReader
        Private _isRunning As Boolean = False
        Private _disposed As Boolean = False
        Private _readTimeoutMs As Integer = 600000  ' 10 分钟超时
        Private _startTimeoutMs As Integer = 60000   ' 1 分钟启动超时

        ''' <summary>
        ''' 最近一次 OCR 识别的原始 JSON 响应（含 raw_data）
        ''' </summary>
        Public Property LastRawResponse As String = Nothing

        ' 事件：引擎日志
        Public Event OnLog As Action(Of String)
        Public Event OnError As Action(Of String)

        Private _ioSemaphore As New System.Threading.SemaphoreSlim(1, 1)
        Private _startLock As New Object()

        ''' <summary>最近一次命令失败原因。成功（含识别到空文本）时为空。</summary>
        Public Property LastCommandError As String = Nothing

        ''' <summary>传给引擎的语言。RapidOCR 内置模型以中文为主，命令里仍带上以便引擎记录。</summary>
        Public Property Language As String = "ch"

        Public Sub New(pythonPath As String)
            _pythonPath = pythonPath
            _engineScriptPath = OcrScriptPaths.Resolve("rapidocr_inference_engine.py")
        End Sub

        ''' <summary>
        ''' 引擎是否在运行
        ''' </summary>
        Public ReadOnly Property IsRunning As Boolean
            Get
                SyncLock _processLock
                    Return _isRunning AndAlso _process IsNot Nothing AndAlso Not _process.HasExited
                End SyncLock
            End Get
        End Property

        ''' <summary>
        ''' 启动引擎进程（同步启动 + 超时，可从任何线程安全调用）
        ''' </summary>
        Public Function StartEngine() As Boolean
            SyncLock _startLock
                Return StartEngineCore()
            End SyncLock
        End Function

        Private Function StartEngineCore() As Boolean
            If IsRunning Then Return True
            ' 上次启动失败留下的进程不能继续占着管道
            If _process IsNot Nothing Then
                KillProcess()
            End If

            If Not File.Exists(_engineScriptPath) Then
                Dim msg = $"推理引擎脚本不存在: {_engineScriptPath}"
                LogMessage(msg)
                RaiseEvent OnError(msg)
                Return False
            End If

            _ioSemaphore.Wait()
            Try
                LogMessage($"检查脚本路径: {_engineScriptPath}")
                If Not File.Exists(_engineScriptPath) Then
                    Dim altPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "rapidocr_inference_engine.py")
                    LogMessage($"尝试替代路径: {altPath}")
                    If File.Exists(altPath) Then
                        _engineScriptPath = altPath
                        LogMessage($"使用替代脚本路径: {_engineScriptPath}")
                    Else
                        LogMessage($"脚本文件不存在: {_engineScriptPath}")
                        Return False
                    End If
                End If

                LogMessage($"启动 Python 引擎: {_pythonPath}")
                LogMessage($"脚本路径: {_engineScriptPath}")

                Dim psi As New ProcessStartInfo()
                psi.FileName = _pythonPath
                psi.Arguments = $"-u ""{_engineScriptPath}"""
                psi.UseShellExecute = False
                psi.RedirectStandardInput = True
                psi.RedirectStandardOutput = True
                psi.RedirectStandardError = True
                psi.CreateNoWindow = True
                psi.StandardOutputEncoding = Encoding.UTF8
                psi.StandardErrorEncoding = Encoding.UTF8
                ' 强制 Python 以 UTF-8 输出 stdout/stderr，否则在中文系统(GBK)下
                ' 重定向管道会输出 GBK 字节，被 VB 按 UTF-8 解码产生乱码(替换符)
                psi.EnvironmentVariables("PYTHONIOENCODING") = "utf-8"
                psi.EnvironmentVariables("PYTHONUTF8") = "1"

                _process = New Process With {.StartInfo = psi, .EnableRaisingEvents = False}
                ' 注意：stdout 由显式 ReadLine/ReadLineAsync 消费（见 WaitForReady/ReadResponseAsync），
                ' 不能再调用 BeginOutputReadLine，否则事件读取器会与显式读取竞争并吞掉响应数据。
                AddHandler _process.ErrorDataReceived, Sub(s, e)
                    If Not String.IsNullOrEmpty(e.Data) Then LogMessage($"[stderr] {e.Data}")
                End Sub

                _process.Start()

                _stdinWriter = New StreamWriter(_process.StandardInput.BaseStream, New UTF8Encoding(False)) With {.AutoFlush = True}
                _stdoutReader = New StreamReader(_process.StandardOutput.BaseStream, Encoding.UTF8)
                ' 注意：stderr 用 BeginErrorReadLine 异步事件读取（见下方 ErrorDataReceived handler），
                ' 不要再创建同步 StreamReader 读取 StandardError，否则抛 "Cannot mix synchronous and asynchronous operation on process stream"

                _process.BeginErrorReadLine()

                ' 等待引擎就绪（读取一行输出）。超时必须杀掉进程，避免下一次把残留输出当成响应。
                Dim ready = WaitForReady(_startTimeoutMs)
                If Not ready Then
                    LogMessage("引擎启动超时或未收到就绪信号")
                    KillProcess()
                    Return False
                End If

                SyncLock _processLock
                    _isRunning = True
                End SyncLock
                LogMessage("引擎已就绪")
                Return True

            Catch ex As Exception
                LogMessage($"启动异常: {ex.Message}")
                RaiseEvent OnError(ex.ToString())
                KillProcess()
                Return False
            Finally
                Try
                    _ioSemaphore.Release()
                Catch ex As Exception
                End Try
            End Try
        End Function

        Private Function WaitForReady(timeoutMs As Integer) As Boolean
            Dim deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs)
            Try
                While DateTime.UtcNow < deadline
                    If _process.HasExited Then Return False

                    ' 使用带剩余超时的异步读取（同步 ReadLine 会无限阻塞且不受超时控制）
                    Dim remaining = CInt((deadline - DateTime.UtcNow).TotalMilliseconds)
                    If remaining <= 0 Then Return False
                    Dim readTask = _stdoutReader.ReadLineAsync()
                    If Not readTask.Wait(remaining) Then
                        LogMessage("引擎就绪等待超时")
                        ' 结束进程，让挂起的 ReadLine 完成，避免和后续读取并发
                        Try
                            If _process IsNot Nothing AndAlso Not _process.HasExited Then _process.Kill()
                        Catch ex As Exception
                        End Try
                        Try
                            readTask.Wait(2000)
                        Catch ex As Exception
                        End Try
                        Return False
                    End If

                    Dim line = readTask.Result
                    If line Is Nothing Then
                        LogMessage("引擎输出流已关闭")
                        Return False
                    End If
                    LogMessage($"引擎响应: {line}")
                    ' RapidOCR 引擎就绪时会输出就绪信号（JSON），检查是否已就绪
                    If line.Contains("RapidOCR") OrElse line.Contains("初始化") OrElse line.Contains("ready") Then
                        Return True
                    End If
                    ' 如果收到 JSON 响应说明引擎已启动
                    Try
                        Dim resp = JsonConvert.DeserializeObject(Of JObject)(line)
                        If resp IsNot Nothing Then Return True
                    Catch ex As Exception
                        ' 忽略解析错误
                    End Try
                End While
                LogMessage("引擎就绪等待超时（无就绪信号）")
            Catch ex As Exception
                LogMessage($"WaitForReady 异常: {ex.Message}")
            End Try
            Return False
        End Function

        ''' <summary>
        ''' 发送命令并等待响应（线程安全）
        ''' </summary>
        Public Async Function SendCommandAsync(command As Dictionary(Of String, Object)) As Task(Of JObject)
            Await _ioSemaphore.WaitAsync()
            Try
                If Not IsRunning Then
                    Return New JObject(New JProperty("success", False), New JProperty("error", "引擎未运行"))
                End If

                Dim json = JsonConvert.SerializeObject(command)
                LogMessage($"发送命令: {json}")
                Await _stdinWriter.WriteLineAsync(json)
                Await _stdinWriter.FlushAsync()

                Dim response = Await ReadResponseAsync(_readTimeoutMs)
                LastRawResponse = response?.ToString()
                Return response

            Catch ex As Exception
                LogMessage($"SendCommand 异常: {ex.Message}")
                Return New JObject(New JProperty("success", False), New JProperty("error", ex.Message))
            Finally
                _ioSemaphore.Release()
            End Try
        End Function

        Private Shared Function ErrorObject(message As String) As JObject
            Return New JObject(New JProperty("success", False), New JProperty("error", message))
        End Function

        Private Sub MarkNotRunning()
            SyncLock _processLock
                _isRunning = False
            End SyncLock
        End Sub

        ''' <summary>
        ''' 读取一条 JSON 响应。超时后终止进程，避免挂起的 ReadLine 吃掉下一条命令的结果。
        ''' </summary>
        Private Async Function ReadResponseAsync(timeoutMs As Integer) As Task(Of JObject)
            Dim deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs)
            Try
                While True
                    Dim remaining = CInt((deadline - DateTime.UtcNow).TotalMilliseconds)
                    If remaining <= 0 Then
                        Await AbortPendingReadAsync(Nothing)
                        Return ErrorObject("引擎响应超时")
                    End If
                    If _process Is Nothing OrElse _process.HasExited Then
                        MarkNotRunning()
                        Return ErrorObject("引擎进程已退出")
                    End If

                    Dim readTask = _stdoutReader.ReadLineAsync()
                    Dim completed = Await Task.WhenAny(readTask, Task.Delay(remaining))
                    If completed IsNot readTask Then
                        LogMessage($"读取响应超时 ({timeoutMs}ms)")
                        Await AbortPendingReadAsync(readTask)
                        Return ErrorObject("引擎响应超时")
                    End If

                    Dim line As String = Nothing
                    Try
                        line = Await readTask
                    Catch ex As Exception
                        MarkNotRunning()
                        Return ErrorObject(ex.Message)
                    End Try

                    If line Is Nothing Then
                        MarkNotRunning()
                        Return ErrorObject("引擎输出流已关闭")
                    End If

                    LogMessage($"收到响应: {line}")
                    Try
                        Dim obj = JsonConvert.DeserializeObject(Of JObject)(line)
                        If obj IsNot Nothing Then Return obj
                    Catch ex As Exception
                        LogMessage($"忽略非 JSON 输出: {ex.Message}")
                    End Try
                End While
                Return ErrorObject("引擎响应超时")
            Catch ex As Exception
                LogMessage($"ReadResponse 异常: {ex.Message}")
                Return ErrorObject(ex.Message)
            End Try
        End Function

        Private Async Function AbortPendingReadAsync(readTask As Task(Of String)) As Task
            Try
                If _process IsNot Nothing AndAlso Not _process.HasExited Then _process.Kill()
            Catch ex As Exception
            End Try
            If readTask IsNot Nothing Then
                Try
                    Await Task.WhenAny(readTask, Task.Delay(2000))
                Catch ex As Exception
                End Try
            End If
            KillProcess()
        End Function

        ''' <summary>
        ''' 执行 OCR 识别（纯文本）
        ''' </summary>
        Private Shared Function IsSuccessResponse(resp As JObject) As Boolean
            If resp Is Nothing OrElse resp("success") Is Nothing Then Return False
            Dim flag As Boolean = False
            Return Boolean.TryParse(resp("success").ToString(), flag) AndAlso flag
        End Function

        Public Async Function RecognizeTextAsync(imagePath As String, Optional recognitionMode As String = "auto") As Task(Of String)
            LastCommandError = Nothing
            Dim cmd = New Dictionary(Of String, Object) From {
                {"cmd", "ocr"},
                {"image_path", imagePath},
                {"recognition_mode", If(String.IsNullOrEmpty(recognitionMode), "auto", recognitionMode)},
                {"lang", If(String.IsNullOrEmpty(Language), "ch", Language)}
            }
            Dim resp = Await SendCommandAsync(cmd)
            If IsSuccessResponse(resp) Then
                Dim textToken = resp("text")
                Return If(textToken IsNot Nothing, textToken.ToString(), "")
            End If
            Dim errToken = If(resp IsNot Nothing, resp("error"), Nothing)
            Dim errMsg = If(errToken IsNot Nothing, errToken.ToString(), "未知错误")
            LastCommandError = errMsg
            Return $"【错误：OCR识别失败 {errMsg}】"
        End Function

        ''' <summary>
        ''' 执行 OCR 识别（带坐标）
        ''' </summary>
        Public Async Function RecognizeTextWithPositionAsync(imagePath As String, Optional recognitionMode As String = "auto") As Task(Of List(Of PaddleOcrResult))
            Dim results = New List(Of PaddleOcrResult)()
            LastCommandError = Nothing
            Dim cmd = New Dictionary(Of String, Object) From {
                {"cmd", "ocr_pos"},
                {"image_path", imagePath},
                {"recognition_mode", If(String.IsNullOrEmpty(recognitionMode), "auto", recognitionMode)},
                {"lang", If(String.IsNullOrEmpty(Language), "ch", Language)}
            }
            Dim resp = Await SendCommandAsync(cmd)
            If IsSuccessResponse(resp) Then
                Dim items = resp("items")
                If items IsNot Nothing Then
                    For Each item In items
                        Dim points As New List(Of Point)()
                        Dim boxToken = item("box")
                        If boxToken IsNot Nothing AndAlso boxToken.Type = JTokenType.Array Then
                            Dim boxPoints = boxToken.ToObject(Of List(Of List(Of Double)))()
                            If boxPoints IsNot Nothing Then
                                For Each pt In boxPoints
                                    If pt IsNot Nothing AndAlso pt.Count >= 2 Then
                                        points.Add(New Point(CInt(pt(0)), CInt(pt(1))))
                                    End If
                                Next
                            End If
                        End If
                        Dim textToken = item("text")
                        Dim confidence As Double = 0
                        If item("confidence") IsNot Nothing Then
                            Double.TryParse(item("confidence").ToString(), confidence)
                        End If
                        results.Add(New PaddleOcrResult() With {
                            .Text = If(textToken IsNot Nothing, textToken.ToString(), ""),
                            .Confidence = confidence,
                            .Box = points
                        })
                    Next
                End If
                Return results
            End If
            Dim errToken = If(resp IsNot Nothing, resp("error"), Nothing)
            LastCommandError = If(errToken IsNot Nothing, errToken.ToString(), "未知错误")
            Return results
        End Function

        ''' <summary>
        ''' 文档结构化分析
        ''' </summary>
        Public Async Function AnalyzeDocumentAsync(imagePath As String, lang As String,
                                                   enableLayout As Boolean, enableTable As Boolean,
                                                   enableFormula As Boolean) As Task(Of String)
            ' RapidOCR 不支持版面分析，降级到 ocr_pos
            Dim cmd = New Dictionary(Of String, Object) From {
                {"cmd", "ocr_pos"},
                {"image_path", imagePath}
            }
            Dim resp = Await SendCommandAsync(cmd)
            Return resp?.ToString()
        End Function

        ''' <summary>
        ''' 设置配置
        ''' </summary>
        Public Async Function SetConfigAsync(config As Dictionary(Of String, Object)) As Task(Of Boolean)
            Dim cmd = New Dictionary(Of String, Object) From {
                {"cmd", "set_config"},
                {"config", config}
            }
            Dim resp = Await SendCommandAsync(cmd)
            Return resp IsNot Nothing AndAlso resp("success").Value(Of Boolean)()
        End Function

        ''' <summary>
        ''' 执行 PDF OCR（一步到位：转图 + 逐页识别 + 合并）
        ''' 返回的 List(Of PaddleOcrResult) 每条额外带 PageIndex 属性需要调用方自行分组。
        ''' 如果不需要分组，可直接拿所有 items 平铺。
        ''' </summary>
        Public Async Function RecognizePdfWithPositionAsync(pdfPath As String, Optional recognitionMode As String = "auto", Optional dpi As Integer = 300) As Task(Of OcrPdfResult)
            Dim result As New OcrPdfResult()
            LastCommandError = Nothing

            If Not File.Exists(pdfPath) Then
                LastCommandError = "PDF 文件不存在"
                result.Success = False
                result.ErrorMessage = LastCommandError
                Return result
            End If

            Dim cmd = New Dictionary(Of String, Object) From {
                {"cmd", "ocr_pdf"},
                {"pdf_path", pdfPath},
                {"dpi", dpi},
                {"mode", "pos"},
                {"recognition_mode", If(String.IsNullOrEmpty(recognitionMode), "auto", recognitionMode)},
                {"lang", If(String.IsNullOrEmpty(Language), "ch", Language)}
            }
            Dim resp = Await SendCommandAsync(cmd)
            If resp Is Nothing Then
                result.Success = False
                result.ErrorMessage = "引擎无响应"
                LastCommandError = "引擎无响应"
                Return result
            End If

            Dim ok As Boolean = False
            If resp("success") IsNot Nothing Then Boolean.TryParse(resp("success").ToString(), ok)
            result.Success = ok
            If Not ok Then
                Dim err = If(resp("error"), Nothing)
                result.ErrorMessage = If(err IsNot Nothing, err.ToString(), "PDF OCR 失败")
                LastCommandError = result.ErrorMessage
                Return result
            End If

            ' 解析顶层 items（已带 page_index / page_number）
            Dim allItems As New List(Of PaddleOcrResult)()
            Dim itemsToken = resp("items")
            If itemsToken IsNot Nothing AndAlso itemsToken.Type = JTokenType.Array Then
                For Each item In CType(itemsToken, JArray)
                    Dim points As New List(Of Point)()
                    Dim boxToken = item("box")
                    If boxToken IsNot Nothing AndAlso boxToken.Type = JTokenType.Array Then
                        For Each pt In CType(boxToken, JArray)
                            If pt IsNot Nothing AndAlso pt.Type = JTokenType.Array AndAlso CType(pt, JArray).Count >= 2 Then
                                Dim arr = CType(pt, JArray)
                                points.Add(New Point(CInt(Math.Round(CDbl(arr(0)))), CInt(Math.Round(CDbl(arr(1))))))
                            End If
                        Next
                    End If
                    Dim textToken = item("text")
                    Dim conf As Double = 0
                    If item("confidence") IsNot Nothing Then Double.TryParse(item("confidence").ToString(), conf)
                    Dim pageIdx As Integer = 0
                    If item("page_index") IsNot Nothing Then Integer.TryParse(item("page_index").ToString(), pageIdx)
                    Dim pageNum As Integer = pageIdx + 1
                    If item("page_number") IsNot Nothing Then Integer.TryParse(item("page_number").ToString(), pageNum)

                    allItems.Add(New PaddleOcrResult() With {
                        .Text = If(textToken IsNot Nothing, textToken.ToString(), ""),
                        .Confidence = conf,
                        .Box = points,
                        .PageIndex = pageIdx,
                        .PageNumber = pageNum
                    })
                Next
            End If

            result.AllItems = allItems

            ' 按页分组
            Dim pagesDict As New Dictionary(Of Integer, OcrPdfPage)()
            Dim pageCount As Integer = 0
            If resp("page_count") IsNot Nothing Then Integer.TryParse(resp("page_count").ToString(), pageCount)

            For Each it In allItems
                Dim idx = it.PageIndex
                If Not pagesDict.ContainsKey(idx) Then
                    pagesDict(idx) = New OcrPdfPage() With {.PageIndex = idx, .PageNumber = idx + 1}
                End If
                pagesDict(idx).Items.Add(it)
            Next

            result.Pages = pagesDict.Values.OrderBy(Function(p) p.PageIndex).ToList()
            result.PageCount = If(pageCount > 0, pageCount, result.Pages.Count)

            Dim elapsedToken = resp("total_elapse")
            Dim elapsed As Double = 0
            If elapsedToken IsNot Nothing Then Double.TryParse(elapsedToken.ToString(), elapsed)
            result.TotalElapseSeconds = elapsed

            LastCommandError = Nothing
            Return result
        End Function

        Public Sub Shutdown()
            Try
                If _process IsNot Nothing AndAlso Not _process.HasExited Then
                    If _isRunning AndAlso _stdinWriter IsNot Nothing Then
                        Try
                            Dim cmd = New Dictionary(Of String, Object) From {{"cmd", "shutdown"}}
                            _stdinWriter.WriteLine(JsonConvert.SerializeObject(cmd))
                            _stdinWriter.Flush()
                            If _process.WaitForExit(1500) Then
                                MarkNotRunning()
                                Return
                            End If
                        Catch ex As Exception
                        End Try
                    End If
                    Try
                        If Not _process.HasExited Then _process.Kill()
                    Catch ex As Exception
                    End Try
                End If
            Catch ex As Exception
            End Try
            MarkNotRunning()
        End Sub

        Private Sub KillProcess()
            Try
                If _process IsNot Nothing AndAlso Not _process.HasExited Then
                    _process.Kill()
                    _process.WaitForExit(2000)
                End If
            Catch ex As Exception
            End Try
            Try
                _stdinWriter?.Dispose()
            Catch ex As Exception
            End Try
            Try
                _stdoutReader?.Dispose()
            Catch ex As Exception
            End Try
            Try
                _stderrReader?.Dispose()
            Catch ex As Exception
            End Try
            Try
                _process?.Dispose()
            Catch ex As Exception
            End Try
            _stdinWriter = Nothing
            _stdoutReader = Nothing
            _stderrReader = Nothing
            _process = Nothing
            MarkNotRunning()
        End Sub

        Private Sub LogMessage(msg As String)
            System.Diagnostics.Debug.WriteLine($"[RapidOcrEngineProcess] {msg}")
            RaiseEvent OnLog(msg)
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            If _disposed Then Return
            If disposing Then
                Shutdown()
                SyncLock _processLock
                    _isRunning = False
                End SyncLock
                If _process IsNot Nothing Then
                    RemoveHandler _process.OutputDataReceived, Nothing
                    RemoveHandler _process.ErrorDataReceived, Nothing
                    _process.Dispose()
                    _process = Nothing
                End If
                _stdinWriter?.Dispose()
                _stdoutReader?.Dispose()
                _stderrReader?.Dispose()
            End If
            _disposed = True
        End Sub

    End Class

    ''' <summary>
    ''' RapidOCR 服务 - 使用 RapidOCR ONNX 引擎
    ''' 【架构】轻量级 OCR，模型体积小（约 30MB），启动快，随插件分发
    ''' </summary>
    Public Class RapidOcrService
        Implements IDisposable

        Private _pythonPath As String
        Private _engineProcess As RapidOcrEngineProcess
        Private _processLock As New Object()
        Private _isInitialized As Boolean = False
        Private _disposed As Boolean = False

        ' 缓存：IsAvailable() 结果
        Private _isAvailableCached As Boolean? = Nothing
        Private _isAvailableCacheTime As DateTime = DateTime.MinValue
        Private Shared ReadOnly _isAvailableCacheDuration As TimeSpan = TimeSpan.FromMinutes(5)

        ' 持久化配置引用，用于读写验证记录
        Private _config As OcrManagerConfig

        ' 后台重新校验控制（每个 Excel 会话只触发一次）
        Private _reverifyScheduled As Boolean = False
        Private _reverifyLock As New Object()

        ''' <summary>当前识别语言（默认中文）</summary>
        Public Property Language As String = "ch"

        ''' <summary>
        ''' 关联的 OCR 管理器配置，用于持久化验证结果
        ''' </summary>
        Public Property Config As OcrManagerConfig
            Get
                Return _config
            End Get
            Set(value As OcrManagerConfig)
                _config = value
            End Set
        End Property

        ''' <summary>
        ''' 暴露 Python 路径供外部调用
        ''' </summary>
        Public Function GetPythonPathPublic() As String
            If String.IsNullOrEmpty(_pythonPath) Then
                _pythonPath = FindPythonPath()
            End If
            Return _pythonPath
        End Function

        ''' <summary>
        ''' 设置已验证可用的 Python 路径
        ''' </summary>
        Public Property WorkingPythonPath As String
            Get
                Return _pythonPath
            End Get
            Set(value As String)
                If Not String.IsNullOrEmpty(value) AndAlso File.Exists(value) Then
                    _pythonPath = value
                End If
            End Set
        End Property

        Public Sub New()
            InitializePaths()
        End Sub

        ''' <summary>
        ''' 使用已验证的 Python 路径构造服务，避免重新探测 Python 环境
        ''' </summary>
        Public Sub New(initialPythonPath As String)
            If Not String.IsNullOrEmpty(initialPythonPath) AndAlso File.Exists(initialPythonPath) Then
                _pythonPath = initialPythonPath
            Else
                InitializePaths()
            End If
        End Sub

        ''' <summary>
        ''' 初始化路径 - 优先使用嵌入式 Python
        ''' </summary>
        Private Sub InitializePaths()
            ' 优先查找嵌入式 Python
            Dim embedPythonPath = FindEmbeddedPythonPath()
            If Not String.IsNullOrEmpty(embedPythonPath) AndAlso File.Exists(embedPythonPath) Then
                _pythonPath = embedPythonPath
                System.Diagnostics.Debug.WriteLine($"[RapidOcrService] 使用嵌入式 Python: {_pythonPath}")
                Return
            End If

            ' 回退到普通 Python 查找
            _pythonPath = FindPythonPath()
        End Sub

        ''' <summary>
        ''' 查找嵌入式 Python 路径（优先 Bootstrapper 按需下载位置，兼容旧路径）
        ''' </summary>
        Private Function FindEmbeddedPythonPath() As String
            ' 优先使用 Bootstrapper 按需下载的 Python
            Dim bootstrapperExe = PythonBootstrapper.EmbeddedPythonExe
            If Not String.IsNullOrEmpty(bootstrapperExe) AndAlso File.Exists(bootstrapperExe) Then
                If IsRapidOcrInstalled(bootstrapperExe) Then
                    Return bootstrapperExe
                End If
            End If

            ' 兼容旧路径：插件安装目录下的 PythonEmbed
            Dim legacyCandidates As String() = {
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PythonEmbed", "python.exe"),
                Path.Combine(Path.GetDirectoryName(GetType(RapidOcrService).Assembly.Location), "PythonEmbed", "python.exe")
            }
            For Each path In legacyCandidates
                If File.Exists(path) Then
                    If IsRapidOcrInstalled(path) Then
                        Return path
                    End If
                End If
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' 获取或创建推理引擎进程
        ''' </summary>
        Private ReadOnly Property EngineProcess As RapidOcrEngineProcess
            Get
                If _engineProcess Is Nothing OrElse Not _engineProcess.IsRunning Then
                    SyncLock _processLock
                        If _engineProcess Is Nothing OrElse Not _engineProcess.IsRunning Then
                            If _engineProcess IsNot Nothing Then
                                _engineProcess.Dispose()
                                _engineProcess = Nothing
                            End If
                            System.Diagnostics.Debug.WriteLine("[RapidOcrService] EngineProcess: 正在检查 Python 路径...")
                            Dim pyOk = Not String.IsNullOrEmpty(_pythonPath) AndAlso File.Exists(_pythonPath)
                            System.Diagnostics.Debug.WriteLine($"[RapidOcrService] EngineProcess: Python路径={_pythonPath}, 存在={pyOk}")
                            If pyOk Then
                                System.Diagnostics.Debug.WriteLine("[RapidOcrService] EngineProcess: 创建 RapidOcrEngineProcess...")
                                _engineProcess = New RapidOcrEngineProcess(_pythonPath)
                                Dim started As Boolean = False
                                Try
                                    System.Diagnostics.Debug.WriteLine("[RapidOcrService] EngineProcess: 调用 StartEngine...")
                                    started = _engineProcess.StartEngine()
                                    System.Diagnostics.Debug.WriteLine($"[RapidOcrService] EngineProcess: StartEngine 返回 {started}")
                                Catch ex As Exception
                                    System.Diagnostics.Debug.WriteLine("[RapidOcrService] 引擎启动异常: " & ex.Message)
                                End Try
                                If started Then
                                    System.Diagnostics.Debug.WriteLine("[RapidOcrService] 推理引擎进程已启动")
                                    ' 设置初始配置
                                    Task.Run(Async Function()
                                        Try
                                            Await _engineProcess.SetConfigAsync(New Dictionary(Of String, Object) From {
                                                {"lang", Language},
                                                {"cpu_threads", 4}
                                            })
                                        Catch ex As Exception  '
                                        End Try
                                    End Function)
                                Else
                                    System.Diagnostics.Debug.WriteLine("[RapidOcrService] 推理引擎启动失败")
                                End If
                            Else
                                System.Diagnostics.Debug.WriteLine("[RapidOcrService] 无可用 Python 路径")
                            End If
                        End If
                    End SyncLock
                End If
                Return _engineProcess
            End Get
        End Property

        ''' <summary>
        ''' 查找 Python 路径
        ''' </summary>
        Private Function FindPythonPath() As String
            Dim candidates As New List(Of String)()

            ' 嵌入式 Python 优先
            Dim embedPath = FindEmbeddedPythonPath()
            If Not String.IsNullOrEmpty(embedPath) Then candidates.Add(embedPath)

            ' PathHelper 常见路径
            Dim helperPath = Utils.PathHelper.FindFirstExistingPath(Utils.PathHelper.GetPythonSearchPaths())
            If Not String.IsNullOrEmpty(helperPath) AndAlso File.Exists(helperPath) AndAlso Not candidates.Contains(helperPath) Then candidates.Add(helperPath)

            ' 注册表查找
            Dim regPath = Utils.PythonLocator.FindFromRegistry()
            If Not String.IsNullOrEmpty(regPath) AndAlso File.Exists(regPath) AndAlso Not candidates.Contains(regPath) Then candidates.Add(regPath)

            ' 优先使用 Python 3.9（兼容性最好）
            Dim pythonVersions As String() = {"3.9", "3.10", "3.11", "3.12"}
            For Each version In pythonVersions
                Try
                    Dim psi As New ProcessStartInfo()
                    psi.FileName = "py"
                    psi.Arguments = $"-{version} -c ""import sys; print(sys.executable)"""
                    psi.RedirectStandardOutput = True
                    psi.RedirectStandardError = True
                    psi.UseShellExecute = False
                    psi.CreateNoWindow = True
                    Using process As Process = Process.Start(psi)
                        Dim output As String = process.StandardOutput.ReadToEnd()
                        process.WaitForExit(2000)
                        If process.ExitCode = 0 Then
                            Dim p As String = output.Trim()
                            If File.Exists(p) AndAlso Not candidates.Contains(p) Then candidates.Add(p)
                        End If
                    End Using
                Catch ex As Exception  ' 
                End Try
            Next

            ' 验证 RapidOCR 可用性
            For Each pyPath In candidates
                If IsRapidOcrInstalled(pyPath) Then
                    System.Diagnostics.Debug.WriteLine($"FindPythonPath: 找到可用 Python: {pyPath}")
                    Return pyPath
                End If
            Next

            System.Diagnostics.Debug.WriteLine("FindPythonPath: 未找到可用的 RapidOCR Python")
            Return If(candidates.Count > 0, candidates(0), Nothing)
        End Function

        Private Function IsRapidOcrInstalled(pythonPath As String) As Boolean
            Try
                Dim process As New System.Diagnostics.Process()
                process.StartInfo.FileName = pythonPath
                process.StartInfo.Arguments = "-c ""from rapidocr_onnxruntime import RapidOCR; print('OK')"""
                process.StartInfo.RedirectStandardOutput = True
                process.StartInfo.RedirectStandardError = True
                process.StartInfo.UseShellExecute = False
                process.StartInfo.CreateNoWindow = True
                process.Start()
                ' stdout / stderr 必须同时读，否则导入日志写满管道后 WaitForExit 进不去
                Dim outTask = process.StandardOutput.ReadToEndAsync()
                Dim errTask = process.StandardError.ReadToEndAsync()
                If Not process.WaitForExit(8000) Then
                    Try : process.Kill() : Catch : End Try
                    Try : Task.WaitAll(New Task() {outTask, errTask}, 2000) : Catch : End Try
                    Return False
                End If
                Try
                    Task.WaitAll(New Task() {outTask, errTask}, 2000)
                Catch ex As Exception
                End Try
                Dim output = If(outTask.IsCompleted, outTask.Result, "")
                Return process.ExitCode = 0 AndAlso output IsNot Nothing AndAlso output.Contains("OK")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("RapidOCR 安装检查失败: " & ex.Message)
                Return False
            End Try
        End Function

        Public Function IsAvailable() As Boolean
            ' 1. 检查内存缓存是否有效
            If _isAvailableCached.HasValue AndAlso
               (DateTime.UtcNow - _isAvailableCacheTime) < _isAvailableCacheDuration Then
                Return _isAvailableCached.Value
            End If

            ' 2. 如果引擎进程已在运行，直接返回 True
            If _engineProcess IsNot Nothing AndAlso _engineProcess.IsRunning AndAlso
               Not String.IsNullOrEmpty(_pythonPath) AndAlso File.Exists(_pythonPath) Then
                _isAvailableCached = True
                _isAvailableCacheTime = DateTime.UtcNow
                Return True
            End If

            ' 3. 持久化验证快速通道：重启 Excel 后如果上次验证仍在有效期内，直接信任并立即返回
            Dim persistent = If(_config IsNot Nothing, _config.RapidOcrVerification, Nothing)
            If IsPersistentVerificationValid(persistent) Then
                _pythonPath = persistent.PythonPath
                _isAvailableCached = True
                _isAvailableCacheTime = DateTime.UtcNow
                ScheduleBackgroundReverify()
                Return True
            End If

            ' 4. 执行完整校验（首次使用、验证过期或环境变化时）
            Return PerformFullVerification()
        End Function

        ''' <summary>
        ''' 判断持久化验证记录是否仍然可信（7 天有效期且 Python 路径存在）
        ''' </summary>
        Private Shared Function IsPersistentVerificationValid(info As OcrEngineVerificationInfo) As Boolean
            If info Is Nothing Then Return False
            If Not info.IsAvailable Then Return False
            If String.IsNullOrEmpty(info.PythonPath) Then Return False
            If Not File.Exists(info.PythonPath) Then Return False
            If Not info.VerifiedAt.HasValue Then Return False
            Return (DateTime.UtcNow - info.VerifiedAt.Value).TotalDays < 7.0
        End Function

        ''' <summary>
        ''' 执行完整的 RapidOCR 可用性校验（探测 Python、import 包）
        ''' </summary>
        Private Function PerformFullVerification() As Boolean
            System.Diagnostics.Debug.WriteLine("[RapidOcrService] 执行完整可用性校验...")

            ' 检查 Python 和 RapidOCR 是否可用
            If String.IsNullOrEmpty(_pythonPath) Then
                _pythonPath = FindPythonPath()
            End If

            If String.IsNullOrEmpty(_pythonPath) OrElse Not File.Exists(_pythonPath) Then
                SaveVerification(False, "")
                _isAvailableCached = False
                _isAvailableCacheTime = DateTime.UtcNow
                Return False
            End If

            If IsRapidOcrInstalled(_pythonPath) Then
                SaveVerification(True, _pythonPath)
                _isAvailableCached = True
                _isAvailableCacheTime = DateTime.UtcNow
                Return True
            End If

            SaveVerification(False, "")
            _isAvailableCached = False
            _isAvailableCacheTime = DateTime.UtcNow
            Return False
        End Function

        ''' <summary>
        ''' 将验证结果持久化到配置文件中
        ''' </summary>
        Private Sub SaveVerification(isAvailable As Boolean, pythonPath As String)
            If _config Is Nothing Then Return
            If _config.RapidOcrVerification Is Nothing Then _config.RapidOcrVerification = New OcrEngineVerificationInfo()
            _config.RapidOcrVerification.IsAvailable = isAvailable
            _config.RapidOcrVerification.PythonPath = pythonPath
            _config.RapidOcrVerification.VerifiedAt = If(isAvailable, DateTime.UtcNow, CType(Nothing, DateTime?))
            Try
                _config.Save()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[RapidOcrService] 保存验证状态失败: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' 后台异步重新校验可用性（用于持久化快速通道后的状态修正）
        ''' </summary>
        Private Sub ScheduleBackgroundReverify()
            SyncLock _reverifyLock
                If _reverifyScheduled Then Return
                _reverifyScheduled = True
            End SyncLock

            Task.Run(Sub()
                Try
                    System.Diagnostics.Debug.WriteLine("[RapidOcrService] 后台异步重新校验可用性...")
                    _isAvailableCached = Nothing
                    _isAvailableCacheTime = DateTime.MinValue
                    PerformFullVerification()
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"[RapidOcrService] 后台重新校验异常: {ex.Message}")
                End Try
            End Sub)
        End Sub

        ''' <summary>
        ''' 清除持久化验证状态（实际 OCR 调用失败或环境变化时调用）
        ''' </summary>
        Public Sub ClearVerificationState()
            _isAvailableCached = Nothing
            _isAvailableCacheTime = DateTime.MinValue
            If _config IsNot Nothing Then
                SaveVerification(False, "")
            End If
            System.Diagnostics.Debug.WriteLine("[RapidOcrService] 持久化验证状态已清除")
        End Sub

        ''' <summary>
        ''' 强制重新检查可用性（清除缓存）
        ''' </summary>
        Public Function ForceCheckAvailable() As Boolean
            _isAvailableCached = Nothing
            _isAvailableCacheTime = DateTime.MinValue
            ' 只清缓存，不重新初始化路径（避免子进程阻塞）
            Return IsAvailable()
        End Function

        ''' <summary>
        ''' 【修复】清除可用性缓存，强制下次重新检测
        ''' </summary>
        Public Sub ClearAvailabilityCache()
            _isAvailableCached = Nothing
            _isAvailableCacheTime = DateTime.MinValue
            System.Diagnostics.Debug.WriteLine("[RapidOcrService] 可用性缓存已清除")
        End Sub

        ''' <summary>最近一次识别失败原因。空文本成功时为空。</summary>
        Public Property LastError As String = Nothing

        ''' <summary>
        ''' 执行 OCR 识别（纯文本）
        ''' </summary>
        Public Async Function RecognizeText(imagePath As String, Optional recognitionMode As String = "auto") As Task(Of String)
            LastError = Nothing
            If Not File.Exists(imagePath) Then
                LastError = "【错误：图片文件不存在】"
                Return LastError
            End If
            If String.IsNullOrEmpty(_pythonPath) Then
                LastError = "【错误：未找到 Python，请先安装 Python 3.9+】"
                Return LastError
            End If
            If Not IsAvailable() Then
                LastError = "【错误：RapidOCR 未安装，请先在设置中安装】"
                Return LastError
            End If

            Try
                Dim engine = EngineProcess
                If engine Is Nothing OrElse Not engine.IsRunning Then
                    ClearVerificationState()
                    LastError = "【错误：推理引擎未就绪】"
                    If String.IsNullOrEmpty(_pythonPath) Then
                        LastError = "【错误：RapidOCR 无法启动 - 未找到 Python，请在设置中检查 Python 安装】"
                    ElseIf Not File.Exists(_pythonPath) Then
                        LastError = $"【错误：RapidOCR 无法启动 - Python 路径无效: {_pythonPath}】"
                    ElseIf Not IsAvailable() Then
                        LastError = "【错误：RapidOCR 无法启动 - rapidocr_onnxruntime 包未安装，请在设置中安装】"
                    End If
                    Return LastError
                End If
                engine.Language = Language
                Dim text = Await engine.RecognizeTextAsync(imagePath, recognitionMode)
                If OcrDocument.IsEngineErrorText(text) Then LastError = text
                Return text
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"RapidOCR 识别异常: {ex.Message}")
                LastError = $"【错误：RapidOCR 识别异常: {ex.Message}】"
                Return LastError
            End Try
        End Function

        ''' <summary>
        ''' 执行 OCR 识别（带坐标）
        ''' </summary>
        Public Async Function RecognizeTextWithPosition(imagePath As String, Optional recognitionMode As String = "auto") As Task(Of List(Of PaddleOcrResult))
            Dim results As New List(Of PaddleOcrResult)
            LastError = Nothing
            If Not File.Exists(imagePath) Then
                LastError = "【错误：图片文件不存在】"
                Return results
            End If
            If String.IsNullOrEmpty(_pythonPath) OrElse Not IsAvailable() Then
                LastError = "【错误：RapidOCR 未安装，请先在设置中安装】"
                Return results
            End If

            Try
                Dim engine = EngineProcess
                If engine Is Nothing OrElse Not engine.IsRunning Then
                    ClearVerificationState()
                    LastError = "【错误：推理引擎未就绪】"
                    Return results
                End If
                engine.Language = Language
                results = Await engine.RecognizeTextWithPositionAsync(imagePath, recognitionMode)
                If Not String.IsNullOrEmpty(engine.LastCommandError) Then
                    LastError = "【错误：" & engine.LastCommandError & "】"
                End If
                Return results
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"RapidOCR 位置识别异常: {ex.Message}")
                LastError = $"【错误：RapidOCR 位置识别异常: {ex.Message}】"
                Return results
            End Try
        End Function

        ''' <summary>
        ''' PDF OCR：RapidOCR 引擎侧完成 PDF 转图 + 逐页识别 + 按页分组
        ''' </summary>
        Public Async Function RecognizePdfWithPositionAsync(pdfPath As String, Optional recognitionMode As String = "auto", Optional dpi As Integer = 300) As Task(Of OcrPdfResult)
            LastError = Nothing
            Dim result As New OcrPdfResult()

            If Not File.Exists(pdfPath) Then
                LastError = "【错误：PDF 文件不存在】"
                result.Success = False
                result.ErrorMessage = LastError
                Return result
            End If
            If String.IsNullOrEmpty(_pythonPath) OrElse Not IsAvailable() Then
                LastError = "【错误：RapidOCR 未安装，请先在设置中安装】"
                result.Success = False
                result.ErrorMessage = LastError
                Return result
            End If

            Try
                Dim engine = EngineProcess
                If engine Is Nothing OrElse Not engine.IsRunning Then
                    ClearVerificationState()
                    LastError = "【错误：推理引擎未就绪】"
                    result.Success = False
                    result.ErrorMessage = LastError
                    Return result
                End If
                engine.Language = Language
                result = Await engine.RecognizePdfWithPositionAsync(pdfPath, recognitionMode, dpi)
                If Not String.IsNullOrEmpty(engine.LastCommandError) Then
                    LastError = "【错误：" & engine.LastCommandError & "】"
                End If
                Return result
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"RapidOCR PDF OCR 异常: {ex.Message}")
                LastError = $"【错误：RapidOCR PDF OCR 异常: {ex.Message}】"
                result.Success = False
                result.ErrorMessage = LastError
                Return result
            End Try
        End Function

        ''' <summary>
        ''' 获取最后一次识别的原始响应（含坐标、文字、置信度）
        ''' </summary>
        Public Function GetLastRawResponse() As String
            Try
                If _engineProcess IsNot Nothing Then
                    Return _engineProcess.LastRawResponse
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"RapidOCR 获取原始数据失败: {ex.Message}")
            End Try
            Return Nothing
        End Function

        ''' <summary>
        ''' 执行结构化文档识别
        ''' </summary>
        Public Async Function RecognizeDocumentAsync(imagePath As String, fileName As String, engineName As String) As Task(Of OcrDocument)
            Dim doc As New OcrDocument()
            doc.SourceFileName = fileName
            doc.SourceFilePath = imagePath
            doc.EngineName = engineName
            doc.TotalPages = 1
            doc.CurrentPage = 1

            Dim sw As New System.Diagnostics.Stopwatch()
            sw.Start()

            Try
                Dim engine = EngineProcess
                If engine Is Nothing OrElse Not engine.IsRunning Then
                    ClearVerificationState()
                    Dim rawText = Await RecognizeText(imagePath)
                    sw.Stop()
                    doc.ElapsedSeconds = sw.Elapsed.TotalSeconds
                    doc.RawText = If(rawText, "")
                    Return doc
                End If

                ' RapidOCR 不支持完整的版面分析，降级到带坐标 OCR
                Dim jsonResponse = Await engine.AnalyzeDocumentAsync(imagePath, Language, True, True, True)
                sw.Stop()
                doc.ElapsedSeconds = sw.Elapsed.TotalSeconds

                If jsonResponse IsNot Nothing Then
                    Try
                        Dim respObj = JObject.Parse(jsonResponse)
                        If respObj("success").Value(Of Boolean)() Then
                            Dim items = respObj("items")
                            If items IsNot Nothing Then
                                Dim textLines As New List(Of String)()
                                For Each item In items
                                    Dim text = item("text").ToString()
                                    textLines.Add(text)

                                    ' 解析位置框（如果引擎返回）
                                    Dim bbox As OcrBoundingBox = Nothing
                                    Dim boxToken = item("box")
                                    If boxToken IsNot Nothing AndAlso boxToken.Type = JTokenType.Array Then
                                        Dim boxArr = TryCast(boxToken, JArray)
                                        If boxArr IsNot Nothing AndAlso boxArr.Count >= 4 Then
                                            Dim xs As New List(Of Integer)()
                                            Dim ys As New List(Of Integer)()
                                            For Each pt In boxArr
                                                Dim ptArr = TryCast(pt, JArray)
                                                If ptArr IsNot Nothing AndAlso ptArr.Count >= 2 Then
                                                    xs.Add(CInt(Math.Round(CDbl(ptArr(0)))))
                                                    ys.Add(CInt(Math.Round(CDbl(ptArr(1)))))
                                                End If
                                            Next
                                            If xs.Count >= 4 Then
                                                bbox = New OcrBoundingBox(xs.Min(), ys.Min(), xs.Max(), ys.Max())
                                            End If
                                        End If
                                    End If

                                    doc.Elements.Add(New OcrElement() With {
                                        .Type = OcrElementType.Paragraph,
                                        .Text = text,
                                        .Confidence = item("confidence").Value(Of Single)(),
                                        .BoundingBox = bbox
                                    })
                                Next
                                doc.RawText = String.Join(vbCrLf, textLines)
                            End If
                        Else
                            Dim errToken = respObj("error")
                            doc.RawText = If(errToken IsNot Nothing, errToken.ToString(), "识别失败")
                        End If
                    Catch ex As Exception
                        doc.RawText = $"【解析错误】{ex.Message}"
                    End Try
                Else
                    doc.RawText = "【错误：引擎响应为空】"
                End If

                Return doc

            Catch ex As Exception
                sw.Stop()
                doc.ElapsedSeconds = sw.Elapsed.TotalSeconds
                doc.RawText = $"【识别异常】{ex.Message}"
                Return doc
            End Try
        End Function

        ''' <summary>
        ''' 获取状态信息
        ''' </summary>
        Public Function GetStatus() As Dictionary(Of String, String)
            Dim status As New Dictionary(Of String, String)
            status.Add("name", "RapidOCR")
            status.Add("available", IsAvailable().ToString())
            status.Add("python_path", If(_pythonPath, ""))
            status.Add("engine", "RapidOCR ONNX")
            Return status
        End Function

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            If _disposed Then Return
            If disposing Then
                If _engineProcess IsNot Nothing Then
                    _engineProcess.Dispose()
                    _engineProcess = Nothing
                End If
            End If
            _disposed = True
        End Sub

    End Class

End Namespace