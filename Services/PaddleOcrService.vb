Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Diagnostics
Imports System.Net.Http
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services

    ''' <summary>
    ''' PaddleOCR 服务 - 使用百度开源的 PaddleOCR 引擎
    ''' 【架构重构】使用统一推理引擎 OcrEngineProcess（长驻进程 + JSON 协议）
    ''' 支持：PP-OCRv6 / 版面分析 / 表格识别 / 公式识别 / 多语言 / OpenVINO 加速
    ''' </summary>
    Public Class PaddleOcrService
        Implements IDisposable

        Private _paddleOcrPath As String
        Private _pythonPath As String
        Private _engineProcess As OcrEngineProcess
        Private _configTask As Task = Nothing
        Private _appliedEngineLang As String = Nothing
        Private _processLock As New Object()
        Private _isInitialized As Boolean = False
        Private _disposed As Boolean = False

        ' 缓存：IsAvailable() 结果，避免每次 OCR 都启动多个 Python 进程检测（5分钟有效）
        Private _isAvailableCached As Boolean? = Nothing
        Private _isAvailableCacheTime As DateTime = DateTime.MinValue
        Private Shared ReadOnly _isAvailableCacheDuration As TimeSpan = TimeSpan.FromMinutes(5)

        ' 持久化配置引用，用于读写验证记录
        Private _config As OcrManagerConfig

        ' 后台重新校验控制（每个 Excel 会话只触发一次）
        Private _reverifyScheduled As Boolean = False
        Private _reverifyLock As New Object()

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

        ' 文档分析配置
        Private _enableLayoutAnalysis As Boolean = True
        Private _enableTableRecognition As Boolean = True
        Private _enableFormulaRecognition As Boolean = True

        ''' <summary>当前识别语言（默认中文）</summary>
        Public Property Language As String = "ch"

        ''' <summary>最近一次识别失败原因。空文本成功时为空。</summary>
        Public Property LastError As String = Nothing

        ''' <summary>启用版面分析</summary>
        Public Property EnableLayoutAnalysis As Boolean
            Get
                Return _enableLayoutAnalysis
            End Get
            Set(value As Boolean)
                _enableLayoutAnalysis = value
            End Set
        End Property

        ''' <summary>启用表格识别</summary>
        Public Property EnableTableRecognition As Boolean
            Get
                Return _enableTableRecognition
            End Get
            Set(value As Boolean)
                _enableTableRecognition = value
            End Set
        End Property

        ''' <summary>启用公式识别</summary>
        Public Property EnableFormulaRecognition As Boolean
            Get
                Return _enableFormulaRecognition
            End Get
            Set(value As Boolean)
                _enableFormulaRecognition = value
            End Set
        End Property

        ''' <summary>
        ''' 暴露 Python 路径供外部调用（如后台下载模型）
        ''' </summary>
        Public Function GetPythonPathPublic() As String
            If String.IsNullOrEmpty(_pythonPath) Then
                _pythonPath = FindPythonPath()
            End If
            Return _pythonPath
        End Function

        ''' <summary>
        ''' 获取最近一次 OCR 识别的原始 JSON 响应（含 raw_data）
        ''' </summary>
        Public Function GetLastRawResponse() As String
            If _engineProcess IsNot Nothing AndAlso _engineProcess.IsRunning Then
                Return _engineProcess.LastRawResponse
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' 【修复】清除可用性缓存，强制下次重新检测
        ''' </summary>
        Public Sub ClearAvailabilityCache()
            _isAvailableCached = Nothing
            _isAvailableCacheTime = DateTime.MinValue
            System.Diagnostics.Debug.WriteLine("[PaddleOcrService] 可用性缓存已清除")
        End Sub

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
        ''' 设置已验证可用的 Python 路径（由 OcrServiceManager 在检测成功后调用）
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

        ''' <summary>
        ''' 初始化路径
        ''' </summary>
        Private Sub InitializePaths()
            Dim basePath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "paddleocr")
            _paddleOcrPath = basePath
            _pythonPath = FindPythonPath()
        End Sub

        ''' <summary>
        ''' 获取或创建推理引擎进程
        ''' </summary>
        Private ReadOnly Property EngineProcess As OcrEngineProcess
            Get
                If _engineProcess Is Nothing OrElse Not _engineProcess.IsRunning Then
                    SyncLock _processLock
                        If _engineProcess Is Nothing OrElse Not _engineProcess.IsRunning Then
                            If _engineProcess IsNot Nothing Then
                                _engineProcess.Dispose()
                                _engineProcess = Nothing
                            End If
                            System.Diagnostics.Debug.WriteLine("[PaddleOcrService] EngineProcess: 正在检查 Python 路径...")
                            Dim pyOk = Not String.IsNullOrEmpty(_pythonPath) AndAlso File.Exists(_pythonPath)
                            System.Diagnostics.Debug.WriteLine($"[PaddleOcrService] EngineProcess: Python路径={_pythonPath}, 存在={pyOk}")
                            If pyOk Then
                                System.Diagnostics.Debug.WriteLine("[PaddleOcrService] EngineProcess: 创建 OcrEngineProcess...")
                                _engineProcess = New OcrEngineProcess(_pythonPath)
                                Dim started As Boolean = False
                                Try
                                    System.Diagnostics.Debug.WriteLine("[PaddleOcrService] EngineProcess: 调用 StartEngine...")
                                    started = _engineProcess.StartEngine()
                                    System.Diagnostics.Debug.WriteLine($"[PaddleOcrService] EngineProcess: StartEngine 返回 {started}")
                                Catch ex As Exception
                                    System.Diagnostics.Debug.WriteLine("[PaddleOcrService] 引擎启动异常: " & ex.Message)
                                    System.Diagnostics.Debug.WriteLine("[PaddleOcrService] 引擎启动堆栈: " & ex.StackTrace)
                                End Try
                                If started Then
                                    System.Diagnostics.Debug.WriteLine("[PaddleOcrService] 推理引擎进程已启动")
                                    ' 启动配置任务，存储引用以便后续等待
                                    _configTask = Task.Run(Async Function() As Task
                                        Try
                                            Await _engineProcess.SetConfigAsync(New Dictionary(Of String, Object) From {
                                                {"lang", Language},
                                                {"cpu_threads", 4},
                                                {"enable_mkldnn", True}
                                            })
                                            _appliedEngineLang = If(String.IsNullOrWhiteSpace(Language), "ch", Language)
                                            System.Diagnostics.Debug.WriteLine("[PaddleOcrService] 配置完成")
                                        Catch ex As Exception
                                            System.Diagnostics.Debug.WriteLine("[PaddleOcrService] 配置异常: " & ex.Message)
                                        End Try
                                    End Function)
                                Else
                                    System.Diagnostics.Debug.WriteLine("[PaddleOcrService] 推理引擎启动失败，将使用降级方案")
                                End If
                            Else
                                System.Diagnostics.Debug.WriteLine("[PaddleOcrService] 无可用 Python 路径")
                                System.Diagnostics.Debug.WriteLine("[PaddleOcrService] 检查 IsAvailable() 结果: " & IsAvailable().ToString())
                            End If
                        End If
                    End SyncLock
                End If
                Return _engineProcess
            End Get
        End Property

#Region "Python 路径查找（保持不变）"

        Private Function FindPythonPath() As String
            Dim candidates As New List(Of String)()
            Dim regPath = Utils.PythonLocator.FindFromRegistry()
            If Not String.IsNullOrEmpty(regPath) AndAlso File.Exists(regPath) Then candidates.Add(regPath)
            Dim helperPath = Utils.PathHelper.FindFirstExistingPath(Utils.PathHelper.GetPythonSearchPaths())
            If Not String.IsNullOrEmpty(helperPath) AndAlso File.Exists(helperPath) AndAlso Not candidates.Contains(helperPath) Then candidates.Add(helperPath)
            ' 优先使用 Python 3.9（paddleocr + pix2tex 兼容性最好）
            Dim pythonVersions As String() = {"3.9", "3.10", "3.11", "3.12"}
            For Each version In pythonVersions
                Try
                    Dim psi As New ProcessStartInfo()
                    psi.FileName = "py"
                    psi.Arguments = "-" & version & " -c ""import sys; print(sys.executable)"""
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
            For Each pyPath In candidates
                If IsPaddleOcrInstalled(pyPath) Then
                    System.Diagnostics.Debug.WriteLine($"FindPythonPath: 找到可用 Python: {pyPath}")
                    Return pyPath
                End If
            Next
            System.Diagnostics.Debug.WriteLine("FindPythonPath: 未找到可用的 PaddleOCR Python")
            Return If(candidates.Count > 0, candidates(0), Nothing)
        End Function

        Private Function IsPaddleOcrInstalled(pythonPath As String) As Boolean
            Try
                Dim process As New System.Diagnostics.Process()
                process.StartInfo.FileName = pythonPath
                process.StartInfo.Arguments = "-c ""import paddleocr; print('OK')"""
                process.StartInfo.RedirectStandardOutput = True
                process.StartInfo.RedirectStandardError = True
                process.StartInfo.UseShellExecute = False
                process.StartInfo.CreateNoWindow = True
                process.Start()
                Dim outTask = process.StandardOutput.ReadToEndAsync()
                Dim errTask = process.StandardError.ReadToEndAsync()
                If Not process.WaitForExit(15000) Then
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
                System.Diagnostics.Debug.WriteLine("PaddleOCR 安装检查失败: " & ex.Message)
                Return False
            End Try
        End Function

        Public Function IsAvailable() As Boolean
            ' 1. 检查内存缓存是否有效（5分钟内）
            If _isAvailableCached.HasValue AndAlso
               (DateTime.UtcNow - _isAvailableCacheTime) < _isAvailableCacheDuration Then
                Return _isAvailableCached.Value
            End If

            ' 2. 如果引擎进程已在运行且可用，直接返回 True（无需再启动 Python 进程检测）
            If _engineProcess IsNot Nothing AndAlso _engineProcess.IsRunning AndAlso
               Not String.IsNullOrEmpty(_pythonPath) AndAlso File.Exists(_pythonPath) Then
                _isAvailableCached = True
                _isAvailableCacheTime = DateTime.UtcNow
                Return True
            End If

            ' 3. 持久化验证快速通道：重启 Excel 后如果上次验证仍在有效期内，直接信任并立即返回
            Dim persistent = If(_config IsNot Nothing, _config.PaddleOcrVerification, Nothing)
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
        ''' 执行完整的 PaddleOCR 可用性校验（探测 Python、import 包、启动测试引擎）
        ''' </summary>
        Private Function PerformFullVerification() As Boolean
            System.Diagnostics.Debug.WriteLine("[PaddleOcrService] 执行完整可用性校验...")

            Dim candidates As New List(Of String)
            If Not String.IsNullOrEmpty(_pythonPath) Then candidates.Add(_pythonPath)
            Dim regPath = Utils.PythonLocator.FindFromRegistry()
            If Not String.IsNullOrEmpty(regPath) AndAlso Not candidates.Contains(regPath) Then candidates.Add(regPath)
            Dim foundPath = Utils.PathHelper.FindFirstExistingPath(Utils.PathHelper.GetPythonSearchPaths())
            If Not String.IsNullOrEmpty(foundPath) AndAlso Not candidates.Contains(foundPath) Then candidates.Add(foundPath)
            For Each version In {"3.12", "3.11", "3.10", "3.9"}
                Try
                    Dim psi As New ProcessStartInfo()
                    psi.FileName = "py"
                    psi.Arguments = "-" & version & " -c ""import sys; print(sys.executable)"""
                    psi.RedirectStandardOutput = True
                    psi.UseShellExecute = False
                    psi.CreateNoWindow = True
                    Using p As Process = Process.Start(psi)
                        Dim out = p.StandardOutput.ReadToEnd()
                        p.WaitForExit(2000)
                        If p.ExitCode = 0 Then
                            Dim pyPath = out.Trim()
                            If File.Exists(pyPath) AndAlso Not candidates.Contains(pyPath) Then candidates.Add(pyPath)
                        End If
                    End Using
                Catch ex As Exception  ' 
                End Try
            Next
            Try
                Dim psi As New ProcessStartInfo()
                psi.FileName = "where"
                psi.Arguments = "python.exe"
                psi.RedirectStandardOutput = True
                psi.UseShellExecute = False
                psi.CreateNoWindow = True
                Using p As Process = Process.Start(psi)
                    Dim out = p.StandardOutput.ReadToEnd()
                    p.WaitForExit(2000)
                    If p.ExitCode = 0 Then
                        For Each line In out.Split({Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
                            Dim wp = line.Trim()
                            If File.Exists(wp) AndAlso Not candidates.Contains(wp) Then candidates.Add(wp)
                        Next
                    End If
                End Using
            Catch ex As Exception  ' 
            End Try

            For Each pyPath In candidates
                Try
                    Dim psi As New ProcessStartInfo()
                    psi.FileName = pyPath
                    psi.Arguments = "-c ""import paddleocr; print('OK')"""
                    psi.RedirectStandardOutput = True
                    psi.RedirectStandardError = True
                    psi.UseShellExecute = False
                    psi.CreateNoWindow = True
                    Using process As Process = Process.Start(psi)
                        Dim output As String = process.StandardOutput.ReadToEnd()
                        process.WaitForExit(8000)
                        If process.ExitCode = 0 AndAlso output.Contains("OK") Then
                            _pythonPath = pyPath
                            ' 实际测试引擎能否启动，而不仅是包是否存在
                            Try
                                Dim testEngine = New OcrEngineProcess(pyPath)
                                Dim started = testEngine.StartEngine()
                                If started Then
                                    testEngine.Dispose()
                                    SaveVerification(True, pyPath)
                                    _isAvailableCached = True
                                    _isAvailableCacheTime = DateTime.UtcNow
                                    Return True
                                Else
                                    System.Diagnostics.Debug.WriteLine($"PaddleOCR 包存在但引擎启动失败: {pyPath}")
                                End If
                            Catch ex As Exception
                                System.Diagnostics.Debug.WriteLine($"PaddleOCR 引擎测试异常: {ex.Message}")
                            End Try
                        End If
                    End Using
                Catch ex As Exception
                End Try
            Next

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
            If _config.PaddleOcrVerification Is Nothing Then _config.PaddleOcrVerification = New OcrEngineVerificationInfo()
            _config.PaddleOcrVerification.IsAvailable = isAvailable
            _config.PaddleOcrVerification.PythonPath = pythonPath
            _config.PaddleOcrVerification.VerifiedAt = If(isAvailable, DateTime.UtcNow, CType(Nothing, DateTime?))
            Try
                _config.Save()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[PaddleOcrService] 保存验证状态失败: {ex.Message}")
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
                    System.Diagnostics.Debug.WriteLine("[PaddleOcrService] 后台异步重新校验可用性...")
                    _isAvailableCached = Nothing
                    _isAvailableCacheTime = DateTime.MinValue
                    PerformFullVerification()
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"[PaddleOcrService] 后台重新校验异常: {ex.Message}")
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
            System.Diagnostics.Debug.WriteLine("[PaddleOcrService] 持久化验证状态已清除")
        End Sub

#End Region

        Private Async Function EnsureEngineLanguageAsync(engine As OcrEngineProcess) As Task
            If engine Is Nothing Then Return
            If _configTask IsNot Nothing Then
                Await _configTask
                _configTask = Nothing
            End If
            Dim lang = If(String.IsNullOrWhiteSpace(Language), "ch", Language)
            If String.Equals(_appliedEngineLang, lang, StringComparison.OrdinalIgnoreCase) Then Return
            Await engine.SetConfigAsync(New Dictionary(Of String, Object) From {
                {"lang", lang}
            })
            _appliedEngineLang = lang
        End Function

#Region "OCR 识别核心方法（使用推理引擎）"

        ''' <summary>
        ''' 执行 OCR 识别（纯文本）
        ''' 使用统一推理引擎，模型常驻内存
        ''' </summary>
        ''' <param name="imagePath">图片路径</param>
        ''' <param name="recognitionMode">识别模式: auto/text/formula/number/date</param>
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
                LastError = "【错误：PaddleOCR 未安装，请先在设置中安装】"
                Return LastError
            End If

            Try
                Dim engine = EngineProcess
                If engine Is Nothing OrElse Not engine.IsRunning Then
                    _isAvailableCached = False
                    _isAvailableCacheTime = DateTime.UtcNow
                    ClearVerificationState()

                    If String.IsNullOrEmpty(_pythonPath) Then
                        LastError = "【错误：PaddleOCR 无法启动 - 未找到 Python，请在设置中检查 Python 安装】"
                    ElseIf Not File.Exists(_pythonPath) Then
                        LastError = $"【错误：PaddleOCR 无法启动 - Python 路径无效：{_pythonPath}】"
                    ElseIf Not IsAvailable() Then
                        LastError = "【错误：PaddleOCR 无法启动 - paddleocr Python 包未安装，请在设置中安装】"
                    Else
                        System.Diagnostics.Debug.WriteLine("【错误：PaddleOCR 推理引擎启动失败，尝试降级到临时脚本模式】")
                        Dim fallback = Await ExecuteFallbackOcr(imagePath)
                        If OcrDocument.IsEngineErrorText(fallback) Then LastError = fallback
                        Return fallback
                    End If
                    Return LastError
                End If
                Await EnsureEngineLanguageAsync(engine)
                Dim text = Await engine.RecognizeTextAsync(imagePath, recognitionMode)
                If OcrDocument.IsEngineErrorText(text) Then LastError = text
                Return text
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"PaddleOCR 识别异常: {ex.Message}")
                LastError = $"【错误：PaddleOCR 识别异常: {ex.Message}】"
                Return LastError
            End Try
        End Function

        ''' <summary>
        ''' 降级方案：使用临时 Python 脚本执行 OCR（引擎进程启动失败时的备用方案）
        ''' </summary>
        Private Async Function ExecuteFallbackOcr(imagePath As String) As Task(Of String)
            System.Diagnostics.Debug.WriteLine("[PaddleOcrService] 使用临时脚本降级方案")
            ' 调用旧的基于临时脚本的识别逻辑
            Return Await RunLegacyOcrScript(imagePath)
        End Function

        ''' <summary>
        ''' 传统临时脚本 OCR 方法
        ''' </summary>
        Private Async Function RunLegacyOcrScript(imagePath As String) As Task(Of String)
            ' BC42356：原实现整体为同步代码（写临时脚本 + Process.WaitForExit 最长 5 分钟）。
            ' 包一层 Task.Run 移到后台线程，避免阻塞调用方线程。
            Return Await Task.Run(Function() RunLegacyOcrScriptCore(imagePath))
        End Function

        ''' <summary>
        ''' 传统临时脚本 OCR 方法（同步核心，由 RunLegacyOcrScript 包 Task.Run 调用）
        ''' </summary>
        Private Function RunLegacyOcrScriptCore(imagePath As String) As String
            Dim scriptPath As String = ""
            Dim outputPath As String = ""
            Dim pathFile As String = ""
            Try
                Dim id = Guid.NewGuid().ToString("N")
                scriptPath = Path.Combine(Path.GetTempPath(), "paddleocr_fallback_" & id & ".py")
                outputPath = Path.Combine(Path.GetTempPath(), "paddleocr_fb_result_" & id & ".txt")
                pathFile = Path.Combine(Path.GetTempPath(), "paddleocr_fb_path_" & id & ".txt")
                File.WriteAllText(pathFile, imagePath, System.Text.Encoding.UTF8)

                Dim lang = If(String.IsNullOrWhiteSpace(Language), "ch", Language)
                If Not System.Text.RegularExpressions.Regex.IsMatch(lang, "^[A-Za-z0-9_\-]+$") Then lang = "ch"

                Dim sb As New StringBuilder()
                sb.AppendLine("# -*- coding: utf-8 -*-")
                sb.AppendLine("import os, sys, traceback")
                sb.AppendLine("os.environ['OMP_NUM_THREADS'] = '2'")
                sb.AppendLine("os.environ['KMP_DUPLICATE_LIB_OK'] = 'TRUE'")
                sb.AppendLine("path_file, output_file = sys.argv[1], sys.argv[2]")
                sb.AppendLine("with open(path_file, 'r', encoding='utf-8') as f:")
                sb.AppendLine("    image_path = f.read().strip()")
                sb.AppendLine("try:")
                sb.AppendLine("    from paddleocr import PaddleOCR")
                sb.AppendLine("    ocr = PaddleOCR(use_angle_cls=True, lang='" & lang & "', use_gpu=False, enable_mkldnn=False, cpu_threads=2, show_log=False, drop_score=0.5)")
                sb.AppendLine("    result = ocr.ocr(image_path, cls=True)")
                sb.AppendLine("    with open(output_file, 'w', encoding='utf-8') as f:")
                sb.AppendLine("        if result and result[0]:")
                sb.AppendLine("            for line in result[0]:")
                sb.AppendLine("                if line and len(line) > 1 and line[1]:")
                sb.AppendLine("                    text = line[1][0] if isinstance(line[1], (list, tuple)) else line[1]")
                sb.AppendLine("                    f.write(str(text) + chr(10))")
                sb.AppendLine("except Exception as e:")
                sb.AppendLine("    with open(output_file, 'w', encoding='utf-8') as f:")
                sb.AppendLine("        f.write('[OCR错误] ' + str(e))")
                File.WriteAllText(scriptPath, sb.ToString(), System.Text.Encoding.UTF8)

                Dim psi As New ProcessStartInfo()
                psi.FileName = _pythonPath
                psi.Arguments = """" & scriptPath & """ """ & pathFile & """ """ & outputPath & """"
                psi.RedirectStandardOutput = True
                psi.RedirectStandardError = True
                psi.StandardOutputEncoding = System.Text.Encoding.UTF8
                psi.StandardErrorEncoding = System.Text.Encoding.UTF8
                psi.UseShellExecute = False
                psi.CreateNoWindow = True
                psi.EnvironmentVariables("PYTHONIOENCODING") = "utf-8"

                Using process As Process = Process.Start(psi)
                    Dim outTask = process.StandardOutput.ReadToEndAsync()
                    Dim errTask = process.StandardError.ReadToEndAsync()
                    If Not process.WaitForExit(300000) Then
                        Try : process.Kill() : Catch : End Try
                        Try : Task.WaitAll(New Task() {outTask, errTask}, 2000) : Catch : End Try
                        Return "【错误：PaddleOCR 降级脚本超时】"
                    End If
                    Try
                        Task.WaitAll(New Task() {outTask, errTask}, 2000)
                    Catch ex As Exception
                    End Try
                    If File.Exists(outputPath) Then
                        Dim result = File.ReadAllText(outputPath, System.Text.Encoding.UTF8)
                        If result.StartsWith("[OCR错误]") Then
                            Return "【错误：PaddleOCR识别失败】" & result
                        End If
                        Return If(String.IsNullOrWhiteSpace(result), "【未识别到文字】", result.Trim())
                    End If
                    Return "【错误：PaddleOCR 降级脚本没有输出】"
                End Using
            Catch ex As Exception
                Return $"【错误：PaddleOCR 降级脚本异常: {ex.Message}】"
            Finally
                Dim leftovers As String() = {scriptPath, outputPath, pathFile}
                For Each leftover As String In leftovers
                    Try
                        If Not String.IsNullOrEmpty(leftover) AndAlso File.Exists(leftover) Then File.Delete(leftover)
                    Catch ex As Exception
                    End Try
                Next
            End Try
        End Function

        ''' <summary>
        ''' 执行 OCR 识别并返回结构化结果（带坐标）
        ''' 使用推理引擎的 ocr_pos 命令
        ''' </summary>
        ''' <param name="imagePath">图片路径</param>
        ''' <param name="recognitionMode">识别模式: auto/text/formula/number/date</param>
        Public Async Function RecognizeTextWithPosition(imagePath As String, Optional recognitionMode As String = "auto") As Task(Of List(Of PaddleOcrResult))
            Dim results As New List(Of PaddleOcrResult)
            LastError = Nothing
            If Not File.Exists(imagePath) Then
                LastError = "【错误：图片文件不存在】"
                Return results
            End If
            If String.IsNullOrEmpty(_pythonPath) OrElse Not IsAvailable() Then
                LastError = "【错误：PaddleOCR 未安装，请先在设置中安装】"
                Return results
            End If

            Try
                Dim engine = EngineProcess
                If engine Is Nothing OrElse Not engine.IsRunning Then
                    ClearVerificationState()
                    LastError = "【错误：推理引擎未就绪】"
                    Return results
                End If
                Await EnsureEngineLanguageAsync(engine)
                results = Await engine.RecognizeTextWithPositionAsync(imagePath, recognitionMode)
                If Not String.IsNullOrEmpty(engine.LastCommandError) Then
                    LastError = "【错误：" & engine.LastCommandError & "】"
                End If
                Return results
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"PaddleOCR 位置识别异常: {ex.Message}")
                LastError = $"【错误：PaddleOCR 位置识别异常: {ex.Message}】"
                Return results
            End Try
        End Function

        ''' <summary>
        ''' 执行结构化文档识别（版面分析 + 表格 + 公式 + OCR）
        ''' 使用 PP-StructureV3（不可用时降级到带坐标的 OCR）
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
                    ' 引擎不可用，清除持久化状态并降级到纯文本
                    ClearVerificationState()
                    Dim rawText = Await RecognizeText(imagePath)
                    sw.Stop()
                    doc.ElapsedSeconds = sw.Elapsed.TotalSeconds
                    doc.RawText = If(rawText, "")
                    Return doc
                End If

                Await EnsureEngineLanguageAsync(engine)

                ' 尝试文档分析（PP-StructureV3）
                Dim jsonResponse = Await engine.AnalyzeDocumentAsync(
                    imagePath,
                    Language,
                    _enableLayoutAnalysis,
                    _enableTableRecognition,
                    _enableFormulaRecognition
                )

                sw.Stop()
                doc.ElapsedSeconds = sw.Elapsed.TotalSeconds

                Dim parsed = ParseAnalysisResponse(jsonResponse, doc)
                If Not parsed Then
                    ' 解析失败：检查是否为 Python 引擎降级返回的位置数据（items 格式）
                    Dim fallbackItems = TryExtractFallbackItems(jsonResponse)
                    If fallbackItems IsNot Nothing AndAlso fallbackItems.Count > 0 Then
                        doc.RawText = String.Join(vbCrLf, fallbackItems.Select(Function(r) r.Text))
                        BuildDocumentFromPositionedResults(fallbackItems, doc)
                    Else
                        ' 纯文本降级
                        Dim rawText = Await RecognizeText(imagePath)
                        doc.RawText = If(rawText, "")
                    End If
                End If

            Catch ex As Exception
                sw.Stop()
                doc.ElapsedSeconds = sw.Elapsed.TotalSeconds
                System.Diagnostics.Debug.WriteLine($"PaddleOCR 文档分析异常: {ex.Message}")
                ' 兜底标记，在 Try-Catch 外执行
                doc.RawText = Nothing
            End Try

            ' Catch 内不能有 Await，此处兜底逻辑放在 Try-Catch 外
            If doc.RawText Is Nothing Then
                Try
                    Dim fallbackTask = RecognizeText(imagePath)
                    doc.RawText = If(fallbackTask.Result, "")
                Catch ex As Exception  ' 
                End Try
            End If

            Return doc
        End Function

        ''' <summary>
        ''' 解析 PP-StructureV3 的分析响应 JSON
        ''' </summary>
        Private Function ParseAnalysisResponse(jsonResponse As String, doc As OcrDocument) As Boolean
            Try
                Dim root = JObject.Parse(jsonResponse)
                If root Is Nothing OrElse Not CBool(If(root("success"), False)) Then
                    Return False
                End If

                ' 1. Markdown 文本
                Dim mdText As String = Nothing
                If root("markdown_text") IsNot Nothing Then mdText = root("markdown_text").ToString()
                If Not String.IsNullOrEmpty(mdText) Then
                    doc.MarkdownText = mdText
                End If

                ' 2. 整体 OCR 文本
                Dim overallText As String = Nothing
                If root("overall_text") IsNot Nothing Then overallText = root("overall_text").ToString()
                If Not String.IsNullOrEmpty(overallText) Then
                    doc.RawText = overallText
                ElseIf Not String.IsNullOrEmpty(mdText) Then
                    doc.RawText = mdText
                End If

                ' 3. 版面分析区域
                Dim layoutBoxes = TryCast(root("layout_boxes"), JArray)
                If layoutBoxes IsNot Nothing Then
                    For Each box In layoutBoxes
                        Dim region As New OcrLayoutRegion()
                        If box("label") IsNot Nothing Then region.Label = box("label").ToString()
                        If box("score") IsNot Nothing Then region.Score = CDbl(If(box("score"), 0.0))
                        Dim coords = TryCast(box("coordinate"), JArray)
                        If coords IsNot Nothing Then
                            For Each c In coords
                                region.Coordinate.Add(CInt(Math.Round(CDbl(c))))
                            Next
                        End If
                        doc.LayoutRegions.Add(region)
                    Next
                End If

                ' 4. 表格识别
                Dim tables = TryCast(root("tables"), JArray)
                If tables IsNot Nothing Then
                    For Each tbl In tables
                        Dim element As New OcrElement()
                        element.Type = OcrElementType.Table
                        element.TableData = New OcrTableData()
                        If tbl("pred_html") IsNot Nothing Then element.TableData.HtmlContent = tbl("pred_html").ToString()

                        ' 解析 HTML 表格为 Header + Rows
                        ParseTableHtmlToData(element.TableData)
                        doc.Elements.Add(element)
                    Next
                End If

                ' 5. 公式识别
                Dim formulas = TryCast(root("formulas"), JArray)
                If formulas IsNot Nothing Then
                    For Each frm In formulas
                        Dim element As New OcrElement()
                        element.Type = OcrElementType.Formula
                        element.FormulaData = New OcrFormulaData()
                        If frm("latex") IsNot Nothing Then element.FormulaData.Latex = frm("latex").ToString()
                        element.Text = element.FormulaData.Latex
                        If frm("confidence") IsNot Nothing Then element.Confidence = CDbl(If(frm("confidence"), 0.0))
                        doc.Elements.Add(element)
                    Next
                End If

                Return doc.Elements.Count > 0 OrElse Not String.IsNullOrEmpty(doc.RawText)

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"解析分析响应失败: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 解析表格 HTML 为结构化 Header + Rows
        ''' </summary>
        Private Sub ParseTableHtmlToData(tableData As OcrTableData)
            Try
                Dim html = tableData.HtmlContent
                If String.IsNullOrEmpty(html) Then Return

                ' 简单解析 <th> 和 <td> 标签
                Dim headerMatches = System.Text.RegularExpressions.Regex.Matches(html, "<th[^>]*>(.*?)</th>", System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.Singleline)
                For Each match As System.Text.RegularExpressions.Match In headerMatches
                    If match.Success Then
                        tableData.Header.Add(System.Net.WebUtility.HtmlDecode(match.Groups(1).Value.Trim()))
                    End If
                Next

                ' 解析数据行
                Dim rowMatches = System.Text.RegularExpressions.Regex.Matches(html, "<tr[^>]*>(.*?)</tr>", System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.Singleline)
                For Each rowMatch As System.Text.RegularExpressions.Match In rowMatches
                    If rowMatch.Success Then
                        Dim rowHtml = rowMatch.Groups(1).Value
                        Dim cellMatches = System.Text.RegularExpressions.Regex.Matches(rowHtml, "<t[hd][^>]*>(.*?)</t[hd]>", System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.Singleline)
                        If cellMatches.Count > 0 Then
                            Dim row As New List(Of String)()
                            For Each cellMatch As System.Text.RegularExpressions.Match In cellMatches
                                If cellMatch.Success Then
                                    row.Add(System.Net.WebUtility.HtmlDecode(cellMatch.Groups(1).Value.Trim()))
                                End If
                            Next
                            If row.Count > 0 Then
                                tableData.Rows.Add(row)
                            End If
                        End If
                    End If
                Next
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"解析表格 HTML 失败: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' 从带位置的结果构建结构化文档（降级方案，基于 Y 坐标分组）
        ''' 同一行内若存在较大水平间隙，则用制表符分隔，保留表格列结构
        ''' </summary>
        Public Shared Sub BuildDocumentFromPositionedResults(positionedResults As List(Of PaddleOcrResult), doc As OcrDocument)
            If positionedResults Is Nothing OrElse positionedResults.Count = 0 Then Return

            ' 保留原始单元格级元素（每个 OCR 文本框一个），供表格重建和 raw_ocr_data 使用
            For Each r In positionedResults
                If r Is Nothing OrElse String.IsNullOrWhiteSpace(r.Text) Then Continue For
                Dim cellElement As New OcrElement()
                cellElement.Type = OcrElementType.Paragraph
                cellElement.Text = r.Text.Trim()
                cellElement.Confidence = r.Confidence
                cellElement.VerticalDir = If(r.VerticalDir, "")
                If r.Box IsNot Nothing AndAlso r.Box.Count >= 2 Then
                    Dim allX = r.Box.Select(Function(p) p.X).ToList()
                    Dim allY = r.Box.Select(Function(p) p.Y).ToList()
                    cellElement.BoundingBox = New OcrBoundingBox(
                        If(allX.Count > 0, allX.Min(), 0),
                        If(allY.Count > 0, allY.Min(), 0),
                        If(allX.Count > 0, allX.Max(), 0),
                        If(allY.Count > 0, allY.Max(), 0)
                    )
                    ' 原始四边形也留住：竖排文字靠它画带朝向的盒子，否则 UI 只会画
                    ' 正上下左右对齐的矩形，看不出文字朝哪边读。
                    If r.Box.Count >= 4 Then
                        cellElement.Quad = New List(Of Point)(r.Box.Take(4))
                    End If
                End If
                doc.RawCellElements.Add(cellElement)
            Next

            Dim getCenterY = Function(r As PaddleOcrResult) As Integer
                If r.Box Is Nothing OrElse r.Box.Count = 0 Then Return 0
                Return CInt(r.Box.Select(Function(p) p.Y).Average())
            End Function
            Dim getCenterX = Function(r As PaddleOcrResult) As Integer
                If r.Box Is Nothing OrElse r.Box.Count = 0 Then Return 0
                Return CInt(r.Box.Select(Function(p) p.X).Average())
            End Function
            Dim getBoxHeight = Function(r As PaddleOcrResult) As Integer
                If r.Box Is Nothing OrElse r.Box.Count = 0 Then Return 0
                Return CInt(r.Box.Select(Function(p) p.Y).Max() - r.Box.Select(Function(p) p.Y).Min())
            End Function
            Dim getBoxWidth = Function(r As PaddleOcrResult) As Integer
                If r.Box Is Nothing OrElse r.Box.Count = 0 Then Return 0
                Return CInt(r.Box.Select(Function(p) p.X).Max() - r.Box.Select(Function(p) p.X).Min())
            End Function

            Dim sortedResults = positionedResults.OrderBy(Function(r) getCenterY(r)).ThenBy(Function(r) getCenterX(r)).ToList()

            ' 自适应行阈值：高分辨率/多行单元格图片下需要更宽松的阈值
            Dim avgHeight = If(sortedResults.Count > 0, sortedResults.Average(Function(r) getBoxHeight(r)), 12.0)
            Dim lineThreshold As Integer = Math.Max(15, CInt(avgHeight * 0.6))

            Dim groupedLines As New List(Of List(Of PaddleOcrResult))()
            Dim currentGroup As New List(Of PaddleOcrResult)()

            For Each r In sortedResults
                If currentGroup.Count = 0 Then
                    currentGroup.Add(r)
                Else
                    ' 用本行第一个框的 Y 做锚点。和上一个框比会把缓慢下移的多行收成一行。
                    Dim anchorY = getCenterY(currentGroup(0))
                    Dim thisAvgY = getCenterY(r)
                    If Math.Abs(thisAvgY - anchorY) <= lineThreshold Then
                        currentGroup.Add(r)
                    Else
                        groupedLines.Add(currentGroup)
                        currentGroup = New List(Of PaddleOcrResult)() From {r}
                    End If
                End If
            Next
            If currentGroup.Count > 0 Then groupedLines.Add(currentGroup)

            For Each group In groupedLines
                ' 按 X 排序，保持从左到右
                Dim orderedGroup = group.OrderBy(Function(r) getCenterX(r)).ToList()

                ' 检测是否为表格行：相邻单元格间存在明显水平间隙
                Dim useTabs As Boolean = False
                If orderedGroup.Count >= 2 Then
                    Dim avgWidth = orderedGroup.Average(Function(r) getBoxWidth(r))
                    For i As Integer = 1 To orderedGroup.Count - 1
                        If orderedGroup(i - 1).Box Is Nothing OrElse orderedGroup(i - 1).Box.Count = 0 Then Continue For
                        If orderedGroup(i).Box Is Nothing OrElse orderedGroup(i).Box.Count = 0 Then Continue For
                        Dim prevRight = orderedGroup(i - 1).Box.Select(Function(p) p.X).Max()
                        Dim currLeft = orderedGroup(i).Box.Select(Function(p) p.X).Min()
                        Dim gap = Math.Max(0, currLeft - prevRight)
                        If gap > Math.Max(avgWidth * 0.4, 8) Then
                            useTabs = True
                            Exit For
                        End If
                    Next
                End If

                Dim element As New OcrElement()
                If useTabs Then
                    element.Text = String.Join(vbTab, orderedGroup.Select(Function(r) r.Text.Trim()).Where(Function(t) Not String.IsNullOrEmpty(t)))
                Else
                    element.Text = String.Join(" ", orderedGroup.Select(Function(r) r.Text.Trim()).Where(Function(t) Not String.IsNullOrEmpty(t)))
                End If
                element.Confidence = If(group.Count > 0, group.Average(Function(r) r.Confidence), 0.0)

                If orderedGroup.Count = 1 AndAlso orderedGroup(0).Box IsNot Nothing AndAlso orderedGroup(0).Box.Count >= 2 Then
                    Dim boxHeight = getBoxHeight(orderedGroup(0))
                    element.BoundingBox = New OcrBoundingBox(
                        orderedGroup(0).Box.Select(Function(p) p.X).Min(),
                        orderedGroup(0).Box.Select(Function(p) p.Y).Min(),
                        orderedGroup(0).Box.Select(Function(p) p.X).Max(),
                        orderedGroup(0).Box.Select(Function(p) p.Y).Max()
                    )
                    element.VerticalDir = If(orderedGroup(0).VerticalDir, "")
                    If orderedGroup(0).Box.Count >= 4 Then
                        element.Quad = New List(Of Point)(orderedGroup(0).Box.Take(4))
                    End If
                    ' 只有明显高于平均行高的单个文本框才当标题，避免把表格行升成标题
                    If boxHeight > Math.Max(avgHeight * 1.8, 36) Then
                        element.Type = OcrElementType.Heading
                        element.HeadingLevel = If(boxHeight > avgHeight * 2.4, 1, 2)
                    End If
                ElseIf group.Count > 0 Then
                    Dim boxes = group.Where(Function(r) r.Box IsNot Nothing AndAlso r.Box.Count >= 2).SelectMany(Function(r) r.Box).ToList()
                    If boxes.Count > 0 Then
                        element.BoundingBox = New OcrBoundingBox(
                            boxes.Select(Function(p) p.X).Min(),
                            boxes.Select(Function(p) p.Y).Min(),
                            boxes.Select(Function(p) p.X).Max(),
                            boxes.Select(Function(p) p.Y).Max()
                        )
                    End If
                End If

                Dim trimmed = element.Text.Trim()
                If trimmed.Length > 1 AndAlso (trimmed(0) = "-"c OrElse trimmed(0) = "*"c OrElse
                   (Char.IsDigit(trimmed(0)) AndAlso trimmed.Length > 1 AndAlso (trimmed(1) = "."c OrElse trimmed(1) = ")"c))) Then
                    element.Type = OcrElementType.ListItem
                    element.ListMarker = trimmed(0).ToString()
                End If

                doc.Elements.Add(element)
            Next
        End Sub

        ''' <summary>
        ''' 从 Python 引擎降级返回的 JSON 中提取位置数据（items 格式）
        ''' 避免二次调用引擎造成性能浪费
        ''' </summary>
        Private Shared Function TryExtractFallbackItems(jsonResponse As String) As List(Of PaddleOcrResult)
            Try
                Dim root = JObject.Parse(jsonResponse)
                If root Is Nothing Then Return Nothing
                Dim success = CBool(If(root("success"), False))
                If Not success Then Return Nothing

                Dim items = TryCast(root("items"), JArray)
                If items Is Nothing OrElse items.Count = 0 Then Return Nothing

                Dim results As New List(Of PaddleOcrResult)()
                For Each item In items
                    If item Is Nothing Then Continue For
                    Dim text As String = If(item("text"), "").ToString()
                    Dim confidence As Double = CDbl(If(item("confidence"), 0.0))
                    Dim box = TryCast(item("box"), JArray)

                    Dim result As New PaddleOcrResult()
                    result.Text = text
                    result.Confidence = confidence
                    If item("vdir") IsNot Nothing Then result.VerticalDir = item("vdir").ToString()

                    If box IsNot Nothing Then
                        result.Box = New List(Of Point)()
                        For Each coord In box
                            Dim arr = TryCast(coord, JArray)
                            If arr IsNot Nothing AndAlso arr.Count >= 2 Then
                                result.Box.Add(New Point(CInt(CDbl(arr(0))), CInt(CDbl(arr(1)))))
                            End If
                        Next
                    End If

                    results.Add(result)
                Next

                If results.Count > 0 Then Return results
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"提取降级位置数据失败: {ex.Message}")
            End Try
            Return Nothing
        End Function

#End Region

#Region "安装和管理"

        Public Function GetStatus() As Dictionary(Of String, String)
            Dim status As New Dictionary(Of String, String)
            status("python_path") = If(_pythonPath, "未找到")
            status("paddleocr_available") = IsAvailable().ToString()
            status("python_found") = (Not String.IsNullOrEmpty(_pythonPath)).ToString()
            status("engine_running") = (_engineProcess IsNot Nothing AndAlso _engineProcess.IsRunning).ToString()
            status("language") = Language
            status("layout_analysis") = _enableLayoutAnalysis.ToString()
            status("table_recognition") = _enableTableRecognition.ToString()
            status("formula_recognition") = _enableFormulaRecognition.ToString()
            Return status
        End Function

        ''' <summary>
        ''' 检查公式识别（pix2tex）是否可用 — 返回 (可用性, Python原始状态JSON)
        ''' </summary>
        Private _pix2texReady As Boolean = False

        Public Async Function CheckFormulaOcrStatusAsync() As Task(Of Tuple(Of Boolean, String))
            Return Await Task.Run(Function() ProbeSharedFormulaPackage()).ConfigureAwait(False)
        End Function

        ''' <summary>
        ''' 用「已经能 import paddleocr」的那个 Python 检查 pix2tex。
        ''' 不要求推理进程已经启动，三个宿主因此看到同一套环境。
        ''' </summary>
        Private Function ProbeSharedFormulaPackage() As Tuple(Of Boolean, String)
            Dim py = EnsurePaddlePython()
            If String.IsNullOrEmpty(py) Then
                Return Tuple.Create(False, "{""error"":""没有找到已经装好 PaddleOCR 的 Python""}")
            End If
            If Pix2texImportable(py) Then
                _pix2texReady = True
                Return Tuple.Create(True, "{""initialized"":true,""formula_available"":true,""paddleocr_available"":true,""shared_env"":true}")
            End If
            Return Tuple.Create(False, "{""formula_available"":false,""paddleocr_available"":true,""shared_env"":true}")
        End Function

        ''' <summary>安装公式增强包。装进当前这台机器上已有 PaddleOCR 的那个 Python，不另起一套环境。</summary>
        Public Async Function InstallFormulaOcrAsync() As Task(Of String)
            Return Await Task.Run(Function() InstallFormulaOnSharedPython()).ConfigureAwait(False)
        End Function

        Private Function InstallFormulaOnSharedPython() As String
            Try
                Dim py = EnsurePaddlePython()
                If String.IsNullOrEmpty(py) Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .message = "没有找到已经装好 PaddleOCR 的 Python"
                    })
                End If
                If Pix2texImportable(py) Then
                    _pix2texReady = True
                    Return JsonConvert.SerializeObject(New With {
                        .success = True,
                        .installed = True,
                        .message = "公式识别已就绪"
                    })
                End If
                Dim commands As String() = {
                    "-m pip install pix2tex --no-deps -q",
                    "-m pip install pix2tex --no-deps -q -i https://pypi.tuna.tsinghua.edu.cn/simple"
                }
                Dim lastError = ""
                For Each args As String In commands
                    lastError = RunPython(py, args, 600000)
                    If Pix2texImportable(py) Then
                        _pix2texReady = True
                        Return JsonConvert.SerializeObject(New With {
                            .success = True,
                            .installed = True,
                            .message = "公式增强包已安装到共享 Python"
                        })
                    End If
                Next
                If lastError.Length > 400 Then lastError = lastError.Substring(lastError.Length - 400)
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .message = "公式增强包没有装进已有的 PaddleOCR 环境。" & lastError
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .message = ex.Message})
            End Try
        End Function

        Private Function EnsurePaddlePython() As String
            If Not String.IsNullOrEmpty(_pythonPath) AndAlso IsPaddleOcrInstalled(_pythonPath) Then
                Return _pythonPath
            End If
            Dim found = FindPythonPath()
            If Not String.IsNullOrEmpty(found) AndAlso IsPaddleOcrInstalled(found) Then
                _pythonPath = found
                Return found
            End If
            Return Nothing
        End Function

        Private Function Pix2texImportable(pythonPath As String) As Boolean
            If _pix2texReady Then Return True
            If String.IsNullOrEmpty(pythonPath) Then Return False
            Dim output = RunPython(pythonPath, "-c ""import importlib.util; print('yes' if importlib.util.find_spec('pix2tex') else 'no')""", 20000)
            Return output.IndexOf("yes", StringComparison.OrdinalIgnoreCase) >= 0
        End Function

        Private Function RunPython(pythonPath As String, arguments As String, timeoutMs As Integer) As String
            Try
                Dim psi As New ProcessStartInfo()
                psi.FileName = pythonPath
                psi.Arguments = arguments
                psi.RedirectStandardOutput = True
                psi.RedirectStandardError = True
                psi.UseShellExecute = False
                psi.CreateNoWindow = True
                Using process As Process = Process.Start(psi)
                    Dim outTask = process.StandardOutput.ReadToEndAsync()
                    Dim errTask = process.StandardError.ReadToEndAsync()
                    process.WaitForExit(timeoutMs)
                    Return If(outTask.Result, "") & If(errTask.Result, "")
                End Using
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function

        Public Async Function InstallPaddleOcrAsync(progress As IProgress(Of String)) As Task(Of Boolean)
            If String.IsNullOrEmpty(_pythonPath) Then
                progress?.Report("错误：未找到 Python，请先安装 Python 3.9 或更高版本")
                System.Diagnostics.Debug.WriteLine("PaddleOCR 安装失败: 未找到 Python")
                Return False
            End If

            System.Diagnostics.Debug.WriteLine("开始安装 PaddleOCR，Python 路径: " & _pythonPath)
            progress?.Report("正在安装 PaddleOCR（含 PP-StructureV3 版面分析模块），这可能需要 5-10 分钟...")

            Dim installScript As String = Path.Combine(Path.GetTempPath(), "install_paddleocr_v3.py")
            Dim outputLog As New StringBuilder()
            Dim errorLog As New StringBuilder()

            Try
                Dim sb As New StringBuilder()
                sb.AppendLine("import subprocess")
                sb.AppendLine("import sys")
                sb.AppendLine("import platform")
                sb.AppendLine("import os")
                sb.AppendLine("")
                sb.AppendLine("MIRRORS = [")
                sb.AppendLine("    'https://pypi.tuna.tsinghua.edu.cn/simple',")
                sb.AppendLine("    'https://mirrors.aliyun.com/pypi/simple/',")
                sb.AppendLine("    'https://pypi.mirrors.ustc.edu.cn/simple/',")
                sb.AppendLine("    'https://pypi.org/simple',")
                sb.AppendLine("]")
                sb.AppendLine("")
                sb.AppendLine("def log(msg):")
                sb.AppendLine("    print(msg)")
                sb.AppendLine("    sys.stdout.flush()")
                sb.AppendLine("")
                sb.AppendLine("python_version = sys.version_info")
                sb.AppendLine("log(f'Python 版本: {python_version.major}.{python_version.minor}.{python_version.micro}')")
                sb.AppendLine("")
                sb.AppendLine("def install_pkg(pkg, ver=None):")
                sb.AppendLine("    pkg_str = f'{pkg}=={ver}' if ver else pkg")
                sb.AppendLine("    for i, mirror in enumerate(MIRRORS):")
                sb.AppendLine("        log(f'[{i+1}/{len(MIRRORS)}] 尝试安装 {pkg_str} 从 {mirror}...')")
                sb.AppendLine("        try:")
                sb.AppendLine("            cmd = [sys.executable, '-m', 'pip', 'install', pkg_str, '-i', mirror, '--timeout', '120', '--retries', '3']")
                sb.AppendLine("            r = subprocess.run(cmd, capture_output=True, text=True, timeout=600)")
                sb.AppendLine("            if r.returncode == 0:")
                sb.AppendLine("                log(f'SUCCESS: {pkg} 安装成功')")
                sb.AppendLine("                return True")
                sb.AppendLine("            log(f'FAILED: {r.stderr[:200]}')")
                sb.AppendLine("        except Exception as e:")
                sb.AppendLine("            log(f'ERROR: {e}')")
                sb.AppendLine("    return False")
                sb.AppendLine("")
                sb.AppendLine("log('STEP 1: 升级 pip...')")
                sb.AppendLine("subprocess.run([sys.executable, '-m', 'pip', 'install', '--upgrade', 'pip', 'setuptools', 'wheel', '-q'], capture_output=True)")
                sb.AppendLine("")
                sb.AppendLine("log('STEP 2: 安装 PaddlePaddle 3.x (CPU 版，支持 MKLDNN中')")
                sb.AppendLine("if not install_pkg('paddlepaddle'):")
                sb.AppendLine("    log('FATAL: PaddlePaddle 安装失败')")
                sb.AppendLine("    sys.exit(1)")
                sb.AppendLine("")
                sb.AppendLine("log('STEP 3: 安装 PaddleOCR（含 PP-StructureV3 版面分析/表格/公式识别）...')")
                sb.AppendLine("if not install_pkg('paddleocr'):")
                sb.AppendLine("    log('FATAL: PaddleOCR 安装失败')")
                sb.AppendLine("    sys.exit(1)")
                sb.AppendLine("")
                sb.AppendLine("log('STEP 4: 下载 OCR 模型和版面分析模型...')")
                sb.AppendLine("try:")
                sb.AppendLine("    from paddleocr import PaddleOCR, PPStructureV3")
                sb.AppendLine("    log('初始化 PaddleOCR（触发模型下载）...')")
                sb.AppendLine("    ocr = PaddleOCR(use_angle_cls=True, lang='ch', ocr_version='PP-OCRv6')")
                sb.AppendLine("    del ocr")
                sb.AppendLine("    log('PaddleOCR 模型下载完成')")
                sb.AppendLine("    log('初始化 PP-StructureV3（触发版面分析/表格/公式模型下载）...')")
                sb.AppendLine("    pp = PPStructureV3()")
                sb.AppendLine("    del pp")
                sb.AppendLine("    log('PP-StructureV3 模型下载完成')")
                sb.AppendLine("except Exception as e:")
                sb.AppendLine("    log(f'WARNING: 模型下载部分失败: {e}')")
                sb.AppendLine("    log('提示: 模型文件将在首次使用时自动下载')")
                sb.AppendLine("")
                sb.AppendLine("log('SUCCESS: 所有组件安装完成！')")
                sb.AppendLine("log('功能: PP-OCRv6 / 版面分析 / 表格识别 / 公式识别')")

                File.WriteAllText(installScript, sb.ToString())

                Dim psi As New ProcessStartInfo()
                psi.FileName = _pythonPath
                psi.Arguments = """" & installScript & """"
                psi.RedirectStandardOutput = True
                psi.RedirectStandardError = True
                psi.UseShellExecute = False
                psi.CreateNoWindow = True
                psi.EnvironmentVariables("PYTHONIOENCODING") = "utf-8"

                Using process As Process = Process.Start(psi)
                    System.Diagnostics.Debug.WriteLine("Python 进程已启动")
                    progress?.Report("正在启动安装程序...")
                    Dim stdOutputTask = process.StandardOutput.ReadToEndAsync()
                    Dim stdErrorTask = process.StandardError.ReadToEndAsync()
                    Dim completed = Await Task.Run(Function() process.WaitForExit(900000))
                    If Not completed Then
                        progress?.Report("安装超时（超过15分钟），请检查网络连接")
                        Try : process.Kill() : Catch : End Try
                        Return False
                    End If
                    Dim stdOutput = Await stdOutputTask
                    Dim stdError = Await stdErrorTask
                    outputLog.AppendLine(stdOutput)
                    errorLog.AppendLine(stdError)
                    If Not String.IsNullOrEmpty(stdOutput) Then
                        For Each line In stdOutput.Split({Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
                            progress?.Report(line)
                        Next
                    End If
                    If Not String.IsNullOrEmpty(stdError) Then
                        For Each line In stdError.Split({Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
                            progress?.Report("错误: " & line)
                        Next
                    End If
                    Try : File.Delete(installScript) : Catch : End Try
                    If process.ExitCode = 0 Then
                        progress?.Report("✅ PaddleOCR + PP-StructureV3 安装成功！")
                        Return True
                    Else
                        progress?.Report("❌ 安装失败，退出码: " & process.ExitCode.ToString())
                        If errorLog.Length > 0 Then progress?.Report("错误详情: " & errorLog.ToString())
                        Return False
                    End If
                End Using

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("PaddleOCR 安装异常: " & ex.Message)
                progress?.Report("❌ 安装失败: " & ex.Message)
                Return False
            End Try
        End Function

#End Region

#Region "IDisposable"

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not _disposed Then
                If disposing Then
                    SyncLock _processLock
                        If _engineProcess IsNot Nothing Then
                            _engineProcess.Dispose()
                            _engineProcess = Nothing
                        End If
                    End SyncLock
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

    ''' <summary>
    ''' PaddleOCR 识别结果（带坐标）
    ''' </summary>
    Public Class PaddleOcrResult
        Public Property Text As String
        Public Property Confidence As Double
        Public Property Box As List(Of Point)
        ''' <summary>所在页索引（从 0 开始，PDF OCR 时填充）</summary>
        Public Property PageIndex As Integer = 0
        ''' <summary>所在页码（从 1 开始，PDF OCR 时填充）</summary>
        Public Property PageNumber As Integer = 1
        ''' <summary>
        ''' 竖排文字的阅读方向：up=从下往上读（逆时针放的文字），down=从上往下读。
        ''' 引擎对竖排框旋转重识别时定出；水平文字为空。
        ''' 下游按它把被切成两半的竖排标注按原阅读顺序接起来，UI 也靠它画朝向。
        ''' </summary>
        Public Property VerticalDir As String = ""
    End Class

    ''' <summary>
    ''' PDF OCR 按页分组的一页结果
    ''' </summary>
    Public Class OcrPdfPage
        Public Property PageIndex As Integer = 0
        Public Property PageNumber As Integer = 1
        Public Property Items As New List(Of PaddleOcrResult)()
        Public Property ElapsedSeconds As Double = 0.0
    End Class

    ''' <summary>
    ''' PDF OCR 完整结果：成功状态 + 所有页 + 平铺 items + 耗时
    ''' </summary>
    Public Class OcrPdfResult
        Public Property Success As Boolean = False
        Public Property ErrorMessage As String = ""
        Public Property PageCount As Integer = 0
        Public Property TotalElapseSeconds As Double = 0.0
        ''' <summary>按页分组</summary>
        Public Property Pages As New List(Of OcrPdfPage)()
        ''' <summary>所有页的 items 平铺，每条带 PageIndex/PageNumber</summary>
        Public Property AllItems As New List(Of PaddleOcrResult)()
    End Class

    ''' <summary>
    ''' 坐标点
    ''' </summary>
    Public Class Point
        Public Property X As Integer
        Public Property Y As Integer

        Public Sub New(x As Integer, y As Integer)
            Me.X = x
            Me.Y = y
        End Sub
    End Class

End Namespace