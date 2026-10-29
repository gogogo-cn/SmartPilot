Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Diagnostics
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services

    ''' <summary>
    ''' PaddleOCR 推理引擎进程管理器
    ''' 管理 Python 子进程生命周期，通过 JSON stdin/stdout 协议与 Python 引擎通信
    ''' 【架构】长驻进程模式：启动后保持运行，避免每次识别的模型加载开销
    ''' </summary>
    Public Class OcrEngineProcess
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

        ''' <summary>最近一次命令失败原因。成功时为空。</summary>
        Public Property LastCommandError As String = Nothing

        ' 事件：引擎日志
        Public Event OnLog As Action(Of String)
        Public Event OnError As Action(Of String)

        ''' <summary>
        ''' 创建引擎进程管理器
        ''' </summary>
        Private _ioSemaphore As New System.Threading.SemaphoreSlim(1, 1)

        Public Sub New(pythonPath As String)
            _pythonPath = pythonPath
            _engineScriptPath = OcrScriptPaths.Resolve("paddle_inference_engine.py")
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
            If IsRunning Then Return True

            If Not File.Exists(_engineScriptPath) Then
                Dim msg = $"推理引擎脚本不存在: {_engineScriptPath}"
                LogMessage(msg)
                RaiseEvent OnError(msg)
                Return False
            End If

            ' 持有信号量直到读完就绪行，避免并发命令把就绪 JSON 当成识别结果
            _ioSemaphore.Wait()
            Try
                If IsRunning Then Return True
                If _process IsNot Nothing Then
                    KillProcess()
                End If

                LogMessage($"检查脚本路径: {_engineScriptPath}")
                If Not File.Exists(_engineScriptPath) Then
                    Dim altPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "paddle_inference_engine.py")
                    LogMessage($"尝试替代路径: {altPath}")
                    If File.Exists(altPath) Then
                        _engineScriptPath = altPath
                        LogMessage($"使用替代脚本路径: {_engineScriptPath}")
                    Else
                        LogMessage("错误: 脚本文件不存在于主路径和替代路径")
                        Return False
                    End If
                End If

                Dim psi As New ProcessStartInfo()
                psi.FileName = _pythonPath
                psi.Arguments = """" & _engineScriptPath & """"
                psi.RedirectStandardInput = True
                psi.RedirectStandardOutput = True
                psi.RedirectStandardError = True
                psi.UseShellExecute = False
                psi.CreateNoWindow = True
                psi.WorkingDirectory = Path.GetDirectoryName(_engineScriptPath)
                psi.StandardOutputEncoding = Encoding.UTF8
                psi.StandardErrorEncoding = Encoding.UTF8
                psi.EnvironmentVariables("PYTHONIOENCODING") = "utf-8"

                _process = New Process()
                _process.StartInfo = psi
                _process.EnableRaisingEvents = True

                LogMessage($"启动推理引擎: python={_pythonPath} script={_engineScriptPath}")

                If Not _process.Start() Then
                    LogMessage("Process.Start() 返回 false，进程启动失败")
                    Return False
                End If
                LogMessage("Python 进程已启动，PID: " & _process.Id)

                _stdinWriter = New StreamWriter(_process.StandardInput.BaseStream, Encoding.UTF8) With {
                    .AutoFlush = True
                }
                _stdoutReader = New StreamReader(_process.StandardOutput.BaseStream, Encoding.UTF8)
                _stderrReader = New StreamReader(_process.StandardError.BaseStream, Encoding.UTF8)

                Dim stderrTask = Task.Run(Sub() ReadStderrLoop())

                Dim readyLine As String = ""
                Dim readEnd = DateTime.UtcNow.AddMilliseconds(_startTimeoutMs)
                Dim readTask As Task(Of String) = Nothing
                Dim pollCount As Integer = 0
                LogMessage("等待引擎就绪信号（超时 " & (_startTimeoutMs / 1000) & " 秒）...")
                While DateTime.UtcNow < readEnd AndAlso String.IsNullOrEmpty(readyLine)
                    If _process Is Nothing OrElse _process.HasExited Then
                        LogMessage("引擎进程已提前退出")
                        If readTask IsNot Nothing Then
                            Try : readTask.Wait(1000) : Catch : End Try
                        End If
                        KillProcess()
                        Return False
                    End If
                    If readTask Is Nothing Then readTask = _stdoutReader.ReadLineAsync()
                    Dim remaining = CInt(Math.Max(1, (readEnd - DateTime.UtcNow).TotalMilliseconds))
                    If Not readTask.Wait(Math.Min(remaining, 500)) Then
                        pollCount += 1
                        If pollCount Mod 10 = 0 Then LogMessage($"等待就绪中 ({pollCount * 500}ms)")
                        Continue While
                    End If
                    Try
                        Dim line = readTask.Result
                        If line Is Nothing Then
                            LogMessage("引擎输出流已关闭")
                            KillProcess()
                            Return False
                        End If
                        readyLine = StripBom(line)
                    Catch ex As Exception
                        LogMessage($"读取就绪信号异常: {ex.Message}")
                        readyLine = ""
                    End Try
                    readTask = Nothing
                    If Not String.IsNullOrEmpty(readyLine) Then LogMessage($"收到就绪行: {readyLine}")
                End While

                If String.IsNullOrEmpty(readyLine) Then
                    LogMessage("引擎启动超时，未收到就绪信号")
                    Try
                        If _process IsNot Nothing AndAlso Not _process.HasExited Then _process.Kill()
                    Catch ex As Exception
                    End Try
                    If readTask IsNot Nothing Then
                        Try : readTask.Wait(2000) : Catch : End Try
                    End If
                    KillProcess()
                    Return False
                End If

                Try
                    Dim readyMsg = JObject.Parse(readyLine)
                    Dim readyEvent As String = ""
                    If readyMsg("event") IsNot Nothing Then readyEvent = readyMsg("event").ToString()
                    If readyEvent = "ready" OrElse readyMsg("status")?.ToString() = "ready" Then
                        _isRunning = True
                        LogMessage($"推理引擎就绪: {readyMsg("message")}")
                        Return True
                    End If
                Catch ex As Exception
                End Try

                _isRunning = True
                LogMessage("推理引擎已启动（非标准就绪）")
                Return True

            Catch ex As Exception
                LogMessage($"引擎启动失败: {ex.Message}")
                KillProcess()
                Return False
            Finally
                Try
                    _ioSemaphore.Release()
                Catch ex As Exception
                End Try
            End Try
        End Function

        ''' <summary>
        ''' 去除字符串开头的 UTF-8 BOM 字符
        ''' </summary>
        Private Shared Function StripBom(text As String) As String
            If String.IsNullOrEmpty(text) Then Return text
            If text.Length > 0 AndAlso AscW(text(0)) = &HFEFF Then
                Return text.Substring(1)
            End If
            Return text
        End Function

        ''' <summary>
        ''' 安装公式识别增强包（pix2tex）
        ''' </summary>
        Public Async Function InstallFormulaOcrAsync() As Task(Of String)
            Dim cmd = JsonConvert.SerializeObject(New With {
                Key .cmd = "install_formula"
            })
            Return Await SendCommandAsync(cmd, timeoutMs:=600000)
        End Function

        ''' <summary>
        ''' 发送命令并等待响应（同步等待）
        ''' </summary>
        Public Function SendCommand(command As String) As String
            _ioSemaphore.Wait()
            Try
                If Not IsRunning Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = "推理引擎未运行"
                    })
                End If

                _stdinWriter.WriteLine(command)
                _stdinWriter.Flush()

                Dim response = _stdoutReader.ReadLine()
                If response Is Nothing Then
                    _isRunning = False
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = "推理引擎连接已断开"
                    })
                End If

                Return StripBom(response)
            Catch ex As Exception
                _isRunning = False
                LogMessage($"命令发送失败: {ex.Message}")
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = $"通信失败: {ex.Message}"
                })
            Finally
                Try
                    _ioSemaphore.Release()
                Catch releaseEx As Exception
                End Try
            End Try
        End Function

        ''' <summary>
        ''' 发送命令（异步，带超时）
        ''' </summary>
        Public Async Function SendCommandAsync(command As String, Optional timeoutMs As Integer = 600000) As Task(Of String)
            If Not IsRunning Then
                Dim started = Await Task.Run(Function() StartEngine())
                If Not started Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = "推理引擎无法启动"
                    })
                End If
            End If

            Await _ioSemaphore.WaitAsync()
            Try
                _stdinWriter.WriteLine(command)
                _stdinWriter.Flush()

                Dim responseTask = _stdoutReader.ReadLineAsync()
                Dim completedTask = Await Task.WhenAny(responseTask, Task.Delay(timeoutMs))

                If completedTask IsNot responseTask Then
                    LogMessage($"命令超时 ({timeoutMs}ms)，终止引擎以避免响应串线")
                    Try
                        If _process IsNot Nothing AndAlso Not _process.HasExited Then _process.Kill()
                    Catch killEx As Exception
                    End Try
                    Try
                        Await Task.WhenAny(responseTask, Task.Delay(2000))
                    Catch waitEx As Exception
                    End Try
                    KillProcess()
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = $"命令执行超时 ({timeoutMs / 1000}秒)"
                    })
                End If

                Dim response = Await responseTask
                If response Is Nothing Then
                    _isRunning = False
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = "推理引擎连接已断开"
                    })
                End If

                ' 去除可能的 BOM 字符
                Return StripBom(response)
            Catch ex As Exception
                _isRunning = False
                LogMessage($"异步命令发送失败: {ex.Message}")
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = $"通信失败: {ex.Message}"
                })
            Finally
                _ioSemaphore.Release()
            End Try
        End Function

        ''' <summary>
        ''' 执行 OCR 纯文本识别
        ''' </summary>
        Public Async Function RecognizeTextAsync(imagePath As String, Optional recognitionMode As String = "auto") As Task(Of String)
            LastCommandError = Nothing
            Dim cmd = JsonConvert.SerializeObject(New With {
                Key .cmd = "ocr",
                Key .image_path = imagePath,
                Key .recognition_mode = recognitionMode
            })
            Dim response = Await SendCommandAsync(cmd)
            LastRawResponse = response

            Try
                Dim json = JObject.Parse(response)
                If json("success") IsNot Nothing AndAlso CBool(json("success")) Then
                    LastCommandError = Nothing
                    If json("markdown_text") IsNot Nothing Then
                        Return json("markdown_text").ToString()
                    End If
                    If json("text") IsNot Nothing Then
                        Return json("text").ToString()
                    End If
                    Return ""
                Else
                    Dim err = If(json("error") IsNot Nothing, json("error").ToString(), "未知错误")
                    LastCommandError = err
                    Return $"【错误：{err}】"
                End If
            Catch ex As Exception
                LastCommandError = "响应解析失败"
                Return "【错误：响应解析失败】"
            End Try
        End Function

        ''' <summary>
        ''' 执行带坐标的 OCR 识别
        ''' </summary>
        Public Async Function RecognizeTextWithPositionAsync(imagePath As String, Optional recognitionMode As String = "auto") As Task(Of List(Of PaddleOcrResult))
            Dim results As New List(Of PaddleOcrResult)()
            LastCommandError = Nothing

            Dim cmd = JsonConvert.SerializeObject(New With {
                Key .cmd = "ocr_pos",
                Key .image_path = imagePath,
                Key .recognition_mode = recognitionMode
            })
            Dim response = Await SendCommandAsync(cmd)

            Try
                Dim json = JObject.Parse(response)
                If json("success") Is Nothing OrElse Not CBool(json("success")) Then
                    LastCommandError = If(json("error") IsNot Nothing, json("error").ToString(), "识别失败")
                    Return results
                End If
                Dim items As JArray = TryCast(json("items"), JArray)
                If items IsNot Nothing Then
                    For Each item In items
                        Dim result As New PaddleOcrResult()
                        If item("text") IsNot Nothing Then result.Text = item("text").ToString()
                        If item("confidence") IsNot Nothing Then result.Confidence = CDbl(If(item("confidence"), 0.0))
                        If item("vdir") IsNot Nothing Then result.VerticalDir = item("vdir").ToString()

                        Dim box As JArray = TryCast(item("box"), JArray)
                        If box IsNot Nothing Then
                            result.Box = New List(Of Point)()
                            For Each p In box
                                Dim arr = TryCast(p, JArray)
                                If arr IsNot Nothing AndAlso arr.Count >= 2 Then
                                    result.Box.Add(New Point(
                                        CInt(Math.Round(CDbl(arr(0)))),
                                        CInt(Math.Round(CDbl(arr(1))))
                                    ))
                                End If
                            Next
                        End If

                        results.Add(result)
                    Next
                End If
            Catch ex As Exception
                LastCommandError = ex.Message
                LogMessage($"解析 OCR 位置结果失败: {ex.Message}")
            End Try

            Return results
        End Function

        ''' <summary>
        ''' 执行结构化文档分析（版面+表格+公式+OCR）
        ''' </summary>
        Public Async Function AnalyzeDocumentAsync(imagePath As String,
                                                   Optional lang As String = "ch",
                                                   Optional enableLayout As Boolean = True,
                                                   Optional enableTable As Boolean = True,
                                                   Optional enableFormula As Boolean = True) As Task(Of String)
            Dim cmd = JsonConvert.SerializeObject(New With {
                Key .cmd = "document",
                Key .image_path = imagePath,
                Key .lang = lang,
                Key .enable_layout = enableLayout,
                Key .enable_table = enableTable,
                Key .enable_formula = enableFormula
            })
            Return Await SendCommandAsync(cmd)
        End Function

        ''' <summary>
        ''' 执行批量 OCR
        ''' </summary>
        Public Async Function BatchOcrAsync(imagePaths As List(Of String), mode As String) As Task(Of String)
            Dim cmd = JsonConvert.SerializeObject(New With {
                Key .cmd = "batch_ocr",
                Key .image_paths = imagePaths,
                Key .mode = mode
            })
            Return Await SendCommandAsync(cmd)
        End Function

        ''' <summary>
        ''' 动态更新配置
        ''' </summary>
        Public Async Function SetConfigAsync(config As Dictionary(Of String, Object)) As Task(Of Boolean)
            Dim cmd = JsonConvert.SerializeObject(New With {
                Key .cmd = "set_config",
                Key .config = config
            })
            Dim response = Await SendCommandAsync(cmd)
            Try
                Dim json = JObject.Parse(response)
                Return CBool(json("success"))
            Catch ex As Exception  ' 
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 获取引擎状态
        ''' </summary>
        Public Async Function GetStatusAsync() As Task(Of String)
            Dim cmd = JsonConvert.SerializeObject(New With {Key .cmd = "status"})
            Return Await SendCommandAsync(cmd)
        End Function

        ''' <summary>
        ''' 关闭引擎进程
        ''' </summary>
        Public Sub Shutdown()
            SyncLock _processLock
                Try
                    If _isRunning AndAlso _process IsNot Nothing AndAlso Not _process.HasExited Then
                        Dim cmd = JsonConvert.SerializeObject(New With {Key .cmd = "shutdown"})
                        If _stdinWriter IsNot Nothing Then _stdinWriter.WriteLine(cmd)
                        If _stdinWriter IsNot Nothing Then _stdinWriter.Flush()
                        ' 等待优雅退出
                        If _process.WaitForExit(5000) Then
                            LogMessage("引擎已优雅关闭")
                        Else
                            KillProcess()
                        End If
                    End If
                Catch ex As Exception
                    LogMessage($"关闭引擎异常: {ex.Message}")
                    KillProcess()
                End Try
                _isRunning = False
            End SyncLock
        End Sub

        ''' <summary>
        ''' 强制终止进程
        ''' </summary>
        Private Sub KillProcess()
            Try
                If _process IsNot Nothing AndAlso Not _process.HasExited Then
                    _process.Kill()
                    _process.WaitForExit(3000)
                    LogMessage("引擎进程已强制终止")
                End If
            Catch ex As Exception  ' 
            End Try

            SafeDispose(_stdinWriter)
            SafeDispose(_stdoutReader)
            SafeDispose(_stderrReader)
            SafeDispose(_process)

            _stdinWriter = Nothing
            _stdoutReader = Nothing
            _stderrReader = Nothing
            _process = Nothing
            _isRunning = False
        End Sub

        ''' <summary>
        ''' 后台读取 stderr 日志
        ''' </summary>
        Private Sub ReadStderrLoop()
            Try
                While _process IsNot Nothing AndAlso Not _process.HasExited
                    Dim line As String = Nothing
                    If _stderrReader IsNot Nothing Then line = _stderrReader.ReadLine()
                    If line Is Nothing Then Exit While
                    If Not String.IsNullOrWhiteSpace(line) Then
                        LogMessage($"[ENGINE] {line}")
                    End If
                End While
            Catch ex As Exception  '  进程退出时正常
            End Try
        End Sub

        ''' <summary>
        ''' 重启引擎
        ''' </summary>
        Public Function Restart() As Boolean
            Shutdown()
            System.Threading.Thread.Sleep(500)
            Return StartEngine()
        End Function

        ''' <summary>
        ''' 日志记录
        ''' </summary>
        Private Sub LogMessage(msg As String)
            System.Diagnostics.Debug.WriteLine($"[OcrEngineProcess] {msg}")
            RaiseEvent OnLog(msg)
        End Sub

        ''' <summary>
        ''' 安全释放
        ''' </summary>
        Private Shared Sub SafeDispose(obj As IDisposable)
            If obj IsNot Nothing Then
                Try
                    obj.Dispose()
                Catch ex As Exception  ' 
                End Try
            End If
        End Sub

#Region "IDisposable"

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not _disposed Then
                If disposing Then
                    Shutdown()
                End If
                _disposed = True
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

#End Region

    End Class

End Namespace