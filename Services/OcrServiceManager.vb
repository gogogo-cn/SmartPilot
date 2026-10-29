Imports System
Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Collections.Generic
Imports System.Diagnostics
Imports Newtonsoft.Json

Namespace Services

    ''' <summary>
    ''' OCR 引擎类型
    ''' </summary>
    Public Enum OcrEngineType
        PaddleOCR = 0
        CloudOCR = 1
        RapidOCR = 2  ' ONNX-based lightweight OCR，随插件分发
    End Enum

    ''' <summary>
    ''' OCR 服务管理器 - 统一管理所有 OCR 引擎
    ''' 【修复】提供 Dispose 方法以正确释放底层资源
    ''' </summary>
    Public Class OcrServiceManager

        Private _disposed As Boolean = False

        Private Shared ReadOnly _instance As New Lazy(Of OcrServiceManager)(Function() New OcrServiceManager(), System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)
        Private Shared ReadOnly _lock As New Object()

        ' 各个 OCR 服务实例（延迟初始化，由 ReadOnly Property 管理）
        Private _paddleOcrService As PaddleOcrService
        Private _cloudOcrService As CloudOcrService
        Private _rapidOcrService As RapidOcrService

        ' 当前配置的 OCR 引擎
        Private _currentEngine As OcrEngineType
        Private _config As OcrManagerConfig

        ' PaddleOCR 高级配置
        Private _paddleOcrLanguage As String = "ch"
        Private _enableLayoutAnalysis As Boolean = True
        Private _enableTableRecognition As Boolean = True
        Private _enableFormulaRecognition As Boolean = True

        ' 缓存：PaddleOCR 可用的 Python 路径（避免每次都重新检测）
        Private _workingPythonPath As String = ""
        ' 缓存：RapidOCR 可用的 Python 路径
        Private _rapidOcrPythonPath As String = ""

        ''' <summary>
        ''' 获取单例实例
        ''' </summary>
        Public Shared ReadOnly Property Instance As OcrServiceManager
            Get
                Return _instance.Value
            End Get
        End Property

        ''' <summary>
        ''' 私有构造函数 - 延迟初始化，不在构造时创建服务实例
        ''' </summary>
        Private Sub New()
            LoadConfig()
        End Sub

        ''' <summary>
        ''' 延迟获取 PaddleOCR 服务（使用已验证的 Python 路径）
        ''' </summary>
        Public ReadOnly Property PaddleOcrService As PaddleOcrService
            Get
                EnsureLocalEngine("PaddleOCR")
                If _paddleOcrService Is Nothing Then
                    _paddleOcrService = New PaddleOcrService()
                    ' 如果已有缓存的可用路径，直接设置
                    If Not String.IsNullOrEmpty(_workingPythonPath) Then
                        _paddleOcrService.WorkingPythonPath = _workingPythonPath
                    End If
                    _paddleOcrService.Config = _config
                End If
                Return _paddleOcrService
            End Get
        End Property

        ''' <summary>
        ''' 获取 Cloud OCR 服务（Public 以便外部重载配置）
        ''' </summary>
        Public ReadOnly Property CloudOcrService As CloudOcrService
            Get
                EnsureLocalEngine("云端 OCR")
                If _cloudOcrService Is Nothing Then _cloudOcrService = New CloudOcrService()
                Return _cloudOcrService
            End Get
        End Property

        ''' <summary>
        ''' 延迟获取 RapidOCR 服务（使用嵌入式 Python）
        ''' </summary>
        Public ReadOnly Property RapidOcrService As RapidOcrService
            Get
                EnsureLocalEngine("RapidOCR")
                If _rapidOcrService Is Nothing Then
                    _rapidOcrService = New RapidOcrService()
                    ' 如果已有缓存的可用路径，直接设置
                    If Not String.IsNullOrEmpty(_rapidOcrPythonPath) Then
                        _rapidOcrService.WorkingPythonPath = _rapidOcrPythonPath
                    End If
                    _rapidOcrService.Config = _config
                End If
                Return _rapidOcrService
            End Get
        End Property

        ''' <summary>
        ''' 加载配置
        ''' </summary>
        Private Sub LoadConfig()
            _config = OcrManagerConfig.Load()
            If _config Is Nothing Then _config = New OcrManagerConfig()
            _currentEngine = _config.SelectedEngine
            _paddleOcrLanguage = If(String.IsNullOrEmpty(_config.PaddleOcrLanguage), "ch", _config.PaddleOcrLanguage)
            _enableLayoutAnalysis = _config.EnableLayoutAnalysis
            _enableTableRecognition = _config.EnableTableRecognition
            _enableFormulaRecognition = _config.EnableFormulaRecognition

            ' 初始化持久化验证信息（兼容旧配置文件）
            If _config.PaddleOcrVerification Is Nothing Then _config.PaddleOcrVerification = New OcrEngineVerificationInfo()
            If _config.RapidOcrVerification Is Nothing Then _config.RapidOcrVerification = New OcrEngineVerificationInfo()

            ' 如果持久化验证仍有效，直接把上次验证成功的 Python 路径作为工作路径，
            ' 避免服务构造时重新探测 Python 环境。
            If IsPersistentVerificationValid(_config.PaddleOcrVerification) Then
                _workingPythonPath = _config.PaddleOcrVerification.PythonPath
            End If
            If IsPersistentVerificationValid(_config.RapidOcrVerification) Then
                _rapidOcrPythonPath = _config.RapidOcrVerification.PythonPath
            End If
        End Sub

        ''' <summary>
        ''' 判断持久化验证记录是否仍然可信
        ''' </summary>
        ''' <summary>
        ''' 插件进程不能在本地创建引擎。真正的 Python / 云端调用只发生在 OcrHost。
        ''' </summary>
        Private Shared Sub EnsureLocalEngine(engineName As String)
            If OcrHostGate.ShouldForward Then
                Throw New InvalidOperationException(engineName & " 在独立 OCR 服务中运行，不能在 Office 进程里启动")
            End If
        End Sub

        ''' <summary>
        ''' 识别前重新读磁盘配置，这样插件里切换引擎后，独立服务不用重启也能跟上。
        ''' </summary>
        Private Sub RefreshFromDisk()
            Try
                Dim cfg = OcrManagerConfig.Load()
                If cfg Is Nothing Then Return
                _config = cfg
                _currentEngine = cfg.SelectedEngine
                _paddleOcrLanguage = If(String.IsNullOrEmpty(cfg.PaddleOcrLanguage), "ch", cfg.PaddleOcrLanguage)
                _enableLayoutAnalysis = cfg.EnableLayoutAnalysis
                _enableTableRecognition = cfg.EnableTableRecognition
                _enableFormulaRecognition = cfg.EnableFormulaRecognition
                If _cloudOcrService IsNot Nothing Then _cloudOcrService.ReloadConfig()
            Catch
            End Try
        End Sub

        Private Function RemoteText(method As String, ParamArray args() As Object) As String
            Try
                Return OcrHostClient.Invoke(Of String)("OcrServiceManager." & method, args)
            Catch ex As Exception
                Return "【错误：OCR 服务不可用 " & ex.Message & "】"
            End Try
        End Function

        Private Function RemoteDocument(method As String, ParamArray args() As Object) As OcrDocument
            Try
                Dim doc = OcrHostClient.Invoke(Of OcrDocument)("OcrServiceManager." & method, args)
                If doc Is Nothing Then
                    doc = New OcrDocument()
                    doc.ErrorMessage = "【错误：OCR 服务没有返回结果】"
                    doc.RawText = doc.ErrorMessage
                End If
                Return doc
            Catch ex As Exception
                Dim doc As New OcrDocument()
                doc.ErrorMessage = "【错误：OCR 服务不可用 " & ex.Message & "】"
                doc.RawText = doc.ErrorMessage
                Return doc
            End Try
        End Function

        Private Shared Function IsPersistentVerificationValid(info As OcrEngineVerificationInfo) As Boolean
            If info Is Nothing Then Return False
            If Not info.IsAvailable Then Return False
            If String.IsNullOrEmpty(info.PythonPath) Then Return False
            If Not File.Exists(info.PythonPath) Then Return False
            If Not info.VerifiedAt.HasValue Then Return False
            ' 验证结果有效期 7 天
            Return (DateTime.UtcNow - info.VerifiedAt.Value).TotalDays < 7.0
        End Function

        ''' <summary>
        ''' 保存配置
        ''' </summary>
        Public Sub SaveConfig()
            _config.SelectedEngine = _currentEngine
            _config.PaddleOcrLanguage = _paddleOcrLanguage
            _config.EnableLayoutAnalysis = _enableLayoutAnalysis
            _config.EnableTableRecognition = _enableTableRecognition
            _config.EnableFormulaRecognition = _enableFormulaRecognition
            _config.Save()
        End Sub

        ''' <summary>PaddleOCR 识别语言（默认 "ch"）</summary>
        Public Property PaddleOcrLanguage As String
            Get
                Return _paddleOcrLanguage
            End Get
            Set(value As String)
                _paddleOcrLanguage = value
                If _paddleOcrService IsNot Nothing Then
                    _paddleOcrService.Language = value
                End If
                SaveConfig()
            End Set
        End Property

        ''' <summary>启用版面分析（PP-StructureV3）</summary>
        Public Property EnableLayoutAnalysis As Boolean
            Get
                Return _enableLayoutAnalysis
            End Get
            Set(value As Boolean)
                _enableLayoutAnalysis = value
                If _paddleOcrService IsNot Nothing Then
                    _paddleOcrService.EnableLayoutAnalysis = value
                End If
                SaveConfig()
            End Set
        End Property

        ''' <summary>启用表格识别</summary>
        Public Property EnableTableRecognition As Boolean
            Get
                Return _enableTableRecognition
            End Get
            Set(value As Boolean)
                _enableTableRecognition = value
                If _paddleOcrService IsNot Nothing Then
                    _paddleOcrService.EnableTableRecognition = value
                End If
                SaveConfig()
            End Set
        End Property

        ''' <summary>启用公式识别</summary>
        Public Property EnableFormulaRecognition As Boolean
            Get
                Return _enableFormulaRecognition
            End Get
            Set(value As Boolean)
                _enableFormulaRecognition = value
                If _paddleOcrService IsNot Nothing Then
                    _paddleOcrService.EnableFormulaRecognition = value
                End If
                SaveConfig()
            End Set
        End Property

        ''' <summary>
        ''' 获取/设置当前 OCR 引擎
        ''' </summary>
        Public Property CurrentEngine As OcrEngineType
            Get
                Return _currentEngine
            End Get
            Set(value As OcrEngineType)
                _currentEngine = value
                _config.SelectedEngine = value
                SaveConfig()
            End Set
        End Property

        ''' <summary>
        ''' 获取所有可用的 OCR 引擎
        ''' </summary>
        Public Function GetAvailableEngines() As Dictionary(Of OcrEngineType, String)
            If OcrHostGate.ShouldForward Then
                Try
                    Return OcrHostClient.Invoke(Of Dictionary(Of OcrEngineType, String))("OcrServiceManager.GetAvailableEngines")
                Catch
                    Return New Dictionary(Of OcrEngineType, String)
                End Try
            End If
            RefreshFromDisk()
            Dim engines As New Dictionary(Of OcrEngineType, String)

            ' RapidOCR（嵌入式，轻量级）
            If RapidOcrService.IsAvailable() Then
                engines.Add(OcrEngineType.RapidOCR, "RapidOCR ✅ (嵌入式，轻量)")
            Else
                engines.Add(OcrEngineType.RapidOCR, "RapidOCR ⚠️ (需安装 Python + rapidocr_onnxruntime)")
            End If

            ' PaddleOCR
            If PaddleOcrService.IsAvailable() Then
                engines.Add(OcrEngineType.PaddleOCR, "PaddleOCR ✅")
            Else
                engines.Add(OcrEngineType.PaddleOCR, "PaddleOCR ⚠️ (需安装 Python)")
            End If

            ' 云端 OCR
            If CloudOcrService.IsAvailable() Then
                engines.Add(OcrEngineType.CloudOCR, $"云端 OCR ({CloudOcrService.GetConfig().Provider}) ✅")
            Else
                engines.Add(OcrEngineType.CloudOCR, "云端 OCR ⚠️ (需配置 API Key)")
            End If

            Return engines
        End Function

        ''' <summary>
        ''' 检查指定的 OCR 引擎是否可用
        ''' PaddleOCR：优先使用 PaddleOcrService.IsAvailable()（带缓存），引擎已运行时几乎零开销
        ''' </summary>
        Public Function IsEngineAvailable(engine As OcrEngineType) As Boolean
            If OcrHostGate.ShouldForward Then
                Try
                    Return OcrHostClient.Invoke(Of Boolean)("OcrServiceManager.IsEngineAvailable", engine)
                Catch
                    Return False
                End Try
            End If
            RefreshFromDisk()
            Select Case engine
                Case OcrEngineType.RapidOCR
                    Return RapidOcrService.IsAvailable()
                Case OcrEngineType.PaddleOCR
                    ' 统一交给 PaddleOcrService 处理（优先使用持久化/内存缓存）
                    Return PaddleOcrService.IsAvailable()
                Case OcrEngineType.CloudOCR
                    Return CloudOcrService.IsAvailable()
                Case Else
                    Return False
            End Select
        End Function

        ''' <summary>
        ''' 检测 PaddleOCR 是否可用
        ''' 遍历所有 Python 环境，任一 import paddleocr 成功即可
        ''' 同时缓存可用的 Python 路径供后续识别使用
        ''' </summary>
        Private Function CheckPaddleOcrViaBridge() As Boolean
            Try
                ' 收集所有候选 Python 路径（去重）
                Dim allPaths As New List(Of String)
                Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                ' PathHelper 常见路径
                For Each p In Utils.PathHelper.GetPythonSearchPaths()
                    If File.Exists(p) AndAlso seen.Add(p) Then allPaths.Add(p)
                Next

                ' 注册表
                Dim regPath = Utils.PythonLocator.FindFromRegistry()
                If Not String.IsNullOrEmpty(regPath) AndAlso seen.Add(regPath) Then allPaths.Add(regPath)

                ' where python.exe
                For Each line In Utils.PythonLocator.FindAllOnPath()
                    If File.Exists(line) AndAlso seen.Add(line) Then allPaths.Add(line)
                Next

                ' py -3.x 命令
                For Each ver In {"3.12", "3.11", "3.10", "3.9", "3.14"}
                    Try
                        Using proc As New System.Diagnostics.Process()
                            proc.StartInfo.FileName = "py"
                            proc.StartInfo.Arguments = $"-{ver} -c ""import sys; print(sys.executable)"""
                            proc.StartInfo.RedirectStandardOutput = True
                            proc.StartInfo.UseShellExecute = False
                            proc.StartInfo.CreateNoWindow = True
                            proc.Start()
                            Dim out = proc.StandardOutput.ReadToEnd().Trim()
                            proc.WaitForExit(3000)
                            If Not String.IsNullOrEmpty(out) AndAlso File.Exists(out) AndAlso seen.Add(out) Then
                                allPaths.Add(out)
                            End If
                        End Using
                    Catch ex As Exception  ' 
                    End Try
                Next

                ' 逐个测试每个 Python 路径的 import paddleocr
                For Each pyPath In allPaths
                    If TestImportPaddleOcr(pyPath) Then
                        _workingPythonPath = pyPath  ' 缓存成功路径
                        ' 如果 PaddleOcrService 已创建，同步更新其 Python 路径
                        If _paddleOcrService IsNot Nothing Then
                            _paddleOcrService.WorkingPythonPath = pyPath
                        End If
                        System.Diagnostics.Debug.WriteLine($"PaddleOCR Bridge: ✅ 可用! path={pyPath}")
                        Return True
                    End If
                Next

                _workingPythonPath = ""
                System.Diagnostics.Debug.WriteLine("PaddleOCR Bridge: ❌ 所有 Python 均不可用")
                Return False

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("PaddleOCR Bridge Exception: " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 测试单个 Python 路径能否 import paddleocr（超时15秒）
        ''' </summary>
        Private Function TestImportPaddleOcr(pythonPath As String) As Boolean
            Try
                Using proc As New System.Diagnostics.Process()
                    proc.StartInfo.FileName = pythonPath
                    proc.StartInfo.Arguments = "-c ""import paddleocr; print('OK')"""
                    proc.StartInfo.RedirectStandardOutput = True
                    proc.StartInfo.RedirectStandardError = True
                    proc.StartInfo.UseShellExecute = False
                    proc.StartInfo.CreateNoWindow = True
                    proc.Start()

                    Dim stdOutTask = proc.StandardOutput.ReadToEndAsync()
                    Dim stdErrTask = proc.StandardError.ReadToEndAsync()
                    Dim finished = proc.WaitForExit(15000)

                    If Not finished Then
                        Try : proc.Kill() : Catch : End Try
                        Return False
                    End If

                    Dim output = stdOutTask.Result.Trim()
                    Return proc.ExitCode = 0 AndAlso output.Contains("OK")
                End Using
            Catch ex As Exception  ' 
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 执行 OCR 识别。当前引擎失败后按 RapidOCR、PaddleOCR、云端的顺序尝试其余可用引擎。
        ''' 空白识别结果会直接返回，不会被当成失败。
        ''' </summary>
        Public Async Function RecognizeText(imagePath As String, Optional recognitionMode As String = "auto") As Task(Of String)
            If OcrHostGate.ShouldForward Then
                Return RemoteText("RecognizeText", imagePath, recognitionMode)
            End If
            RefreshFromDisk()
            Dim lastError As String = Nothing
            For Each engine In BuildEngineAttemptOrder()
                Dim result As String = Nothing
                Try
                    result = Await RecognizeTextWithEngine(imagePath, engine, recognitionMode)
                Catch ex As Exception
                    result = $"【错误：{GetEngineDisplayName(engine)} 执行失败 {ex.Message}】"
                End Try

                If IsMissingImageError(result) Then Return result
                If Not OcrDocument.IsEngineErrorText(result) Then Return If(result, "")

                lastError = result
                Debug.WriteLine($"[OcrServiceManager] {engine} 失败：{result}")
                ClearEngineAvailability(engine)
            Next
            Return If(String.IsNullOrEmpty(lastError), "【错误：所有 OCR 引擎均不可用】", lastError)
        End Function

        ''' <summary>
        ''' 执行带位置信息的 OCR 识别。失败会写入 <see cref="OcrDocument.ErrorMessage"/> 并尝试下一个引擎。
        ''' </summary>
        Public Async Function RecognizeTextWithPositionAsync(imagePath As String, Optional fileName As String = "", Optional recognitionMode As String = "auto") As Task(Of OcrDocument)
            If OcrHostGate.ShouldForward Then
                Return RemoteDocument("RecognizeTextWithPositionAsync", imagePath, fileName, recognitionMode)
            End If
            RefreshFromDisk()
            Dim lastDoc As OcrDocument = Nothing
            For Each engine In BuildEngineAttemptOrder()
                Dim doc As OcrDocument = Nothing
                Try
                    doc = Await RecognizePositionWithEngine(imagePath, fileName, recognitionMode, engine)
                Catch ex As Exception
                    doc = NewDocumentShell(imagePath, fileName, engine)
                    doc.ErrorMessage = $"【错误：{GetEngineDisplayName(engine)} 执行失败 {ex.Message}】"
                    doc.RawText = doc.ErrorMessage
                End Try

                If doc Is Nothing Then Continue For
                lastDoc = doc
                If IsMissingImageError(doc.ErrorMessage) OrElse IsMissingImageError(doc.RawText) Then Return doc
                If String.IsNullOrEmpty(doc.ErrorMessage) Then
                    ApplyDomainPack(recognitionMode, doc)
                    Return doc
                End If

                Debug.WriteLine($"[OcrServiceManager] {engine} 位置识别失败：{doc.ErrorMessage}")
                ClearEngineAvailability(engine)
            Next

            If lastDoc Is Nothing Then
                lastDoc = NewDocumentShell(imagePath, fileName, _currentEngine)
                lastDoc.ErrorMessage = "【错误：所有 OCR 引擎均不可用】"
                lastDoc.RawText = lastDoc.ErrorMessage
            End If
            Return lastDoc
        End Function

        ''' <summary>
        ''' PDF OCR 统一入口：RapidOCR 引擎侧完成 PDF 转图 + 逐页识别 + 按页分组
        ''' 当前只委托给 RapidOCR（PDF 转图逻辑在 RapidOCR Python 引擎里实现），
        ''' 后续可以让 PaddleOCR / 云端也支持 PDF。
        ''' 机械图纸模式下每页都会跑 GdtSymbolRecognizer.Apply。
        ''' </summary>
        Public Async Function RecognizePdfWithPositionAsync(pdfPath As String,
                                                               Optional fileName As String = "",
                                                               Optional recognitionMode As String = "auto",
                                                               Optional dpi As Integer = 300) As Task(Of List(Of OcrDocument))
            If OcrHostGate.ShouldForward Then
                Try
                    Dim hostDocs = OcrHostClient.Invoke(Of List(Of OcrDocument))(
                        "OcrServiceManager.RecognizePdfWithPositionAsync", pdfPath, fileName, recognitionMode, dpi)
                    If hostDocs Is Nothing Then
                        Dim shell = NewDocumentShell(pdfPath, fileName, _currentEngine)
                        shell.ErrorMessage = "【错误：OCR 服务没有返回结果】"
                        shell.RawText = shell.ErrorMessage
                        Return New List(Of OcrDocument) From {shell}
                    End If
                    Return hostDocs
                Catch ex As Exception
                    Dim shell = NewDocumentShell(pdfPath, fileName, _currentEngine)
                    shell.ErrorMessage = "【错误：OCR 服务不可用 " & ex.Message & "】"
                    shell.RawText = shell.ErrorMessage
                    Return New List(Of OcrDocument) From {shell}
                End Try
            End If

            RefreshFromDisk()
            Dim docs As New List(Of OcrDocument)()
            Dim pack As Services.DomainPack.DomainPackManifest = ResolveDomainPack(recognitionMode)

            ' 优先用 RapidOCR 做 PDF OCR（Python 侧已实现完整链路）
            If IsEngineAvailable(OcrEngineType.RapidOCR) Then
                Try
                    Dim pdfResult = Await RapidOcrService.RecognizePdfWithPositionAsync(pdfPath, recognitionMode, dpi)
                    If pdfResult Is Nothing OrElse Not pdfResult.Success Then
                        Dim err = If(pdfResult IsNot Nothing, pdfResult.ErrorMessage, "RapidOCR PDF OCR 返回空")
                        Dim errDoc = NewDocumentShell(pdfPath, fileName, OcrEngineType.RapidOCR.ToString())
                        errDoc.ErrorMessage = "【错误：" & err & "】"
                        errDoc.RawText = errDoc.ErrorMessage
                        docs.Add(errDoc)
                        Return docs
                    End If

                    ' 每页构造一个 OcrDocument
                    Dim elapsedPerPage = If(pdfResult.PageCount > 0, pdfResult.TotalElapseSeconds / pdfResult.PageCount, pdfResult.TotalElapseSeconds)
                    For Each page In pdfResult.Pages
                        Dim doc As New OcrDocument()
                        doc.SourceFilePath = pdfPath
                        doc.SourceFileName = If(String.IsNullOrEmpty(fileName), Path.GetFileName(pdfPath), fileName) & $" (第 {page.PageNumber} 页)"
                        doc.EngineName = "RapidOCR"
                        doc.TotalPages = pdfResult.PageCount
                        doc.CurrentPage = page.PageNumber
                        doc.ElapsedSeconds = elapsedPerPage
                        If Not String.IsNullOrEmpty(page.ElapsedSeconds) Then doc.ElapsedSeconds = page.ElapsedSeconds

                        Dim sb As New System.Text.StringBuilder()
                        For Each it In page.Items
                            sb.AppendLine(it.Text)
                        Next
                        doc.RawText = sb.ToString()

                        ' 填充 Elements（复用 PaddleOcrService 的坐标建库逻辑）
                        PaddleOcrService.BuildDocumentFromPositionedResults(page.Items, doc)

                        ' 专项识别模块（机械制图等）按注册表分派
                        ApplyDomainPack(pack, doc)

                        docs.Add(doc)
                    Next
                    Return docs
                Catch ex As Exception
                    Debug.WriteLine($"[OcrServiceManager] RapidOCR PDF OCR 异常：{ex.Message}")
                End Try
            End If

            ' 降级：如果 RapidOCR 不可用（理论上 PDF OCR 依赖它）
            Dim fallback = NewDocumentShell(pdfPath, fileName, _currentEngine)
            fallback.ErrorMessage = "【错误：RapidOCR 引擎不可用，无法处理 PDF。请确保已安装 rapidocr_onnxruntime】"
            fallback.RawText = fallback.ErrorMessage
            docs.Add(fallback)
            Return docs
        End Function

        ''' <summary>
        ''' 按 recognitionMode 解析专项识别模块。
        ''' 向后兼容：mechanical / gdt / drawing → 内置机械制图（GdtSymbolRecognizer）。
        ''' 其他字符串按模块 id 精确匹配；解析失败或模块被停用时返回 Nothing（不打断识别）。
        ''' </summary>
        Private Shared Function ResolveDomainPack(recognitionMode As String) As Services.DomainPack.DomainPackManifest
            Try
                Dim p = Services.DomainPack.DomainPackRegistry.Instance.ResolveByMode(recognitionMode)
                If p Is Nothing OrElse Not p.Enabled OrElse Not String.IsNullOrEmpty(p.ErrorText) Then Return Nothing
                Return p
            Catch ex As Exception
                Diagnostics.Debug.WriteLine("[OcrServiceManager] 专项模块解析失败：" & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 识别完成后的专项后处理。
        ''' 内置机械制图仍走 GdtSymbolRecognizer。已勾选的 python 模块走识别改写入口，
        ''' 与当前 recognitionMode 是否点名该模块无关：勾选即生效。
        ''' </summary>
        Private Shared Sub ApplyDomainPack(modeOrPack As Object, doc As OcrDocument)
            If doc Is Nothing Then Return
            Dim pack As Services.DomainPack.DomainPackManifest = TryCast(modeOrPack, Services.DomainPack.DomainPackManifest)
            If pack Is Nothing AndAlso TypeOf modeOrPack Is String Then pack = ResolveDomainPack(DirectCast(modeOrPack, String))
            If pack IsNot Nothing AndAlso pack.Pipeline = Services.DomainPack.DomainPackPipeline.BuiltIn Then
                If String.Equals(pack.ModuleRef, "GdtSymbolRecognizer", StringComparison.OrdinalIgnoreCase) Then
                    GdtSymbolRecognizer.Apply(doc)
                End If
            End If
            Services.DomainPack.DomainPackPythonRunner.ApplyEnabled(doc)
        End Sub

        ''' <summary>兼容旧调用方：是否机械制图模式。内部已改为注册表判断。</summary>
        Private Shared Function IsMechanicalMode(recognitionMode As String) As Boolean
            Dim p = ResolveDomainPack(recognitionMode)
            Return p IsNot Nothing AndAlso p.Pipeline = Services.DomainPack.DomainPackPipeline.BuiltIn
        End Function

        Private Function BuildEngineAttemptOrder() As List(Of OcrEngineType)
            Dim order As New List(Of OcrEngineType)
            If IsEngineAvailable(_currentEngine) Then order.Add(_currentEngine)
            Dim priority = {OcrEngineType.RapidOCR, OcrEngineType.PaddleOCR, OcrEngineType.CloudOCR}
            For Each engine In priority
                If order.Contains(engine) Then Continue For
                If IsEngineAvailable(engine) Then order.Add(engine)
            Next
            If order.Count = 0 Then order.Add(_currentEngine)
            Return order
        End Function

        Private Sub ClearEngineAvailability(engine As OcrEngineType)
            Select Case engine
                Case OcrEngineType.PaddleOCR
                    PaddleOcrService.ClearAvailabilityCache()
                    PaddleOcrService.ClearVerificationState()
                Case OcrEngineType.RapidOCR
                    RapidOcrService.ClearAvailabilityCache()
                    RapidOcrService.ClearVerificationState()
            End Select
        End Sub

        Private Shared Function IsMissingImageError(text As String) As Boolean
            Return Not String.IsNullOrEmpty(text) AndAlso text.Contains("图片文件不存在")
        End Function

        Private Function NewDocumentShell(imagePath As String, fileName As String, engine As OcrEngineType) As OcrDocument
            Dim doc As New OcrDocument()
            doc.SourceFilePath = imagePath
            doc.SourceFileName = If(String.IsNullOrEmpty(fileName), Path.GetFileName(imagePath), fileName)
            doc.EngineName = GetEngineDisplayName(engine)
            doc.TotalPages = 1
            doc.CurrentPage = 1
            Return doc
        End Function

        Private Async Function RecognizePositionWithEngine(imagePath As String, fileName As String, recognitionMode As String, engine As OcrEngineType) As Task(Of OcrDocument)
            Dim doc = NewDocumentShell(imagePath, fileName, engine)
            Dim sw As New System.Diagnostics.Stopwatch()
            sw.Start()

            Select Case engine
                Case OcrEngineType.RapidOCR
                    Dim positionedResults = Await RapidOcrService.RecognizeTextWithPosition(imagePath, recognitionMode)
                    sw.Stop()
                    doc.ElapsedSeconds = sw.Elapsed.TotalSeconds
                    If Not String.IsNullOrEmpty(RapidOcrService.LastError) Then
                        doc.ErrorMessage = RapidOcrService.LastError
                        doc.RawText = doc.ErrorMessage
                        AttachRawEngineResponse(doc, engine)
                        Return doc
                    End If
                    ApplyPositionedResults(positionedResults, doc)

                Case OcrEngineType.PaddleOCR
                    If Not ArePaddleModelsReady() Then TriggerPaddleModelDownload()
                    Dim paddleSvc = PaddleOcrService
                    paddleSvc.Language = _paddleOcrLanguage
                    Dim positionedResults = Await paddleSvc.RecognizeTextWithPosition(imagePath, recognitionMode)
                    sw.Stop()
                    doc.ElapsedSeconds = sw.Elapsed.TotalSeconds
                    If Not String.IsNullOrEmpty(paddleSvc.LastError) Then
                        doc.ErrorMessage = paddleSvc.LastError
                        doc.RawText = doc.ErrorMessage
                        AttachRawEngineResponse(doc, engine)
                        Return doc
                    End If
                    ApplyPositionedResults(positionedResults, doc)

                Case OcrEngineType.CloudOCR
                    Dim rawText = Await CloudOcrService.RecognizeText(imagePath)
                    sw.Stop()
                    doc.ElapsedSeconds = sw.Elapsed.TotalSeconds
                    doc.RawText = If(rawText, "")
                    If OcrDocument.IsEngineErrorText(doc.RawText) Then
                        doc.ErrorMessage = doc.RawText
                        AttachRawEngineResponse(doc, engine)
                        Return doc
                    End If
                    Dim elements = HeuristicParse(doc.RawText)
                    If elements IsNot Nothing AndAlso elements.Count > 0 Then
                        doc.Elements.AddRange(elements)
                    End If

                Case Else
                    sw.Stop()
                    doc.ErrorMessage = "【错误：未知的 OCR 引擎类型】"
                    doc.RawText = doc.ErrorMessage
            End Select

            AttachRawEngineResponse(doc, engine)
            Return doc
        End Function

        Private Sub AttachRawEngineResponse(doc As OcrDocument, engine As OcrEngineType)
            If doc Is Nothing Then Return
            Try
                Dim raw As String = Nothing
                Select Case engine
                    Case OcrEngineType.RapidOCR
                        raw = RapidOcrService.GetLastRawResponse()
                    Case OcrEngineType.PaddleOCR
                        raw = PaddleOcrService.GetLastRawResponse()
                End Select
                If Not String.IsNullOrEmpty(raw) Then doc.RawEngineResponse = raw
            Catch
            End Try
        End Sub

        Private Shared Sub ApplyPositionedResults(positionedResults As List(Of PaddleOcrResult), doc As OcrDocument)
            If positionedResults Is Nothing OrElse positionedResults.Count = 0 Then
                doc.RawText = ""
                Return
            End If

            doc.RawText = String.Join(vbCrLf, positionedResults.OrderBy(Function(r)
                If r.Box Is Nothing OrElse r.Box.Count = 0 Then Return 0
                Return CInt(r.Box.Select(Function(p) p.Y).Average())
            End Function).ThenBy(Function(r)
                If r.Box Is Nothing OrElse r.Box.Count = 0 Then Return 0
                Return CInt(r.Box.Select(Function(p) p.X).Average())
            End Function).Select(Function(r) r.Text))

            Try
                PaddleOcrService.BuildDocumentFromPositionedResults(positionedResults, doc)
            Catch ex As Exception
                Debug.WriteLine($"[OcrServiceManager] 坐标建行失败：{ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' 执行 OCR 识别（指定引擎）
        ''' </summary>
        Public Async Function RecognizeTextWithEngine(imagePath As String, engine As OcrEngineType, Optional recognitionMode As String = "auto") As Task(Of String)
            If OcrHostGate.ShouldForward Then
                Return RemoteText("RecognizeTextWithEngine", imagePath, engine, recognitionMode)
            End If
            RefreshFromDisk()
            Select Case engine
                Case OcrEngineType.RapidOCR
                    Return Await RapidOcrService.RecognizeText(imagePath, recognitionMode)
                Case OcrEngineType.PaddleOCR
                    If Not ArePaddleModelsReady() Then
                        TriggerPaddleModelDownload()
                    End If
                    Dim paddleSvc = PaddleOcrService
                    paddleSvc.Language = _paddleOcrLanguage
                    Return Await paddleSvc.RecognizeText(imagePath, recognitionMode)
                Case OcrEngineType.CloudOCR
                    Return Await CloudOcrService.RecognizeText(imagePath)
                Case Else
                    Return "【错误：未知的 OCR 引擎类型】"
            End Select
        End Function

        ''' <summary>
        ''' 获取指定引擎的状态信息
        ''' </summary>
        Public Function GetEngineStatus(engine As OcrEngineType) As Dictionary(Of String, String)
            If OcrHostGate.ShouldForward Then
                Try
                    Return OcrHostClient.Invoke(Of Dictionary(Of String, String))("OcrServiceManager.GetEngineStatus", engine)
                Catch
                    Return New Dictionary(Of String, String)
                End Try
            End If
            RefreshFromDisk()
            Select Case engine
                Case OcrEngineType.RapidOCR
                    Return RapidOcrService.GetStatus()
                Case OcrEngineType.PaddleOCR
                    Return PaddleOcrService.GetStatus()
                Case OcrEngineType.CloudOCR
                    Dim status = New Dictionary(Of String, String)
                    status.Add("name", "云端 OCR")
                    status.Add("available", CloudOcrService.IsAvailable().ToString())
                    status.Add("provider", CloudOcrService.GetConfig().Provider)
                    Return status
                Case Else
                    Return New Dictionary(Of String, String)
            End Select
        End Function

        ''' <summary>
        ''' 安装 PaddleOCR
        ''' </summary>
        Public Async Function InstallPaddleOcrAsync(progress As IProgress(Of String)) As Task(Of Boolean)
            Return Await PaddleOcrService.InstallPaddleOcrAsync(progress)
        End Function

        ''' <summary>
        ''' 检查公式识别增强包（pix2tex）是否可用
        ''' </summary>
        Public Async Function CheckFormulaOcrStatusAsync() As Task(Of Tuple(Of Boolean, String))
            If OcrHostGate.ShouldForward Then
                Try
                    Return OcrHostClient.Invoke(Of Tuple(Of Boolean, String))("OcrServiceManager.CheckFormulaOcrStatusAsync")
                Catch ex As Exception
                    Return Tuple.Create(False, ex.Message)
                End Try
            End If
            Return Await PaddleOcrService.CheckFormulaOcrStatusAsync()
        End Function

        ' 公式增强包后台安装状态
        Private _formulaInstallTask As Task(Of String)
        Private _formulaInstallResult As String = ""

        ''' <summary>
        ''' 启动公式识别增强包后台安装（非阻塞）
        ''' </summary>
        Public Function StartFormulaOcrInstall() As String
            If OcrHostGate.ShouldForward Then
                Return RemoteText("StartFormulaOcrInstall")
            End If
            SyncLock _lock
                If _formulaInstallTask IsNot Nothing AndAlso Not _formulaInstallTask.IsCompleted Then
                    Return JsonConvert.SerializeObject(New With {.success = True, .message = "安装已在进行中"})
                End If
                _formulaInstallResult = ""
                _formulaInstallTask = Task.Run(Async Function()
                    Try
                        Dim result = Await PaddleOcrService.InstallFormulaOcrAsync()
                        _formulaInstallResult = result
                        Return result
                    Catch ex As Exception
                        _formulaInstallResult = JsonConvert.SerializeObject(New With {.success = False, .message = $"安装异常: {ex.Message}"})
                        Return _formulaInstallResult
                    End Try
                End Function)
            End SyncLock
            Return JsonConvert.SerializeObject(New With {.success = True, .message = "安装已启动"})
        End Function

        ''' <summary>
        ''' 获取公式增强包安装状态
        ''' </summary>
        Public Function GetFormulaInstallStatus() As String
            If OcrHostGate.ShouldForward Then
                Return RemoteText("GetFormulaInstallStatus")
            End If
            SyncLock _lock
                If _formulaInstallTask Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = True, .status = "idle", .message = "未开始安装", .result = ""})
                End If
                If Not _formulaInstallTask.IsCompleted Then
                    Return JsonConvert.SerializeObject(New With {.success = True, .status = "installing", .message = "正在安装 pix2tex（约 3-5 分钟）...", .result = ""})
                End If
                Return JsonConvert.SerializeObject(New With {.success = True, .status = "completed", .message = "安装完成", .result = _formulaInstallResult})
            End SyncLock
        End Function

        ''' <summary>
        ''' 安装公式识别增强包（pix2tex / LaTeX-OCR）——同步阻塞版本，供内部使用
        ''' </summary>
        Public Async Function InstallFormulaOcrAsync() As Task(Of String)
            If OcrHostGate.ShouldForward Then
                Return RemoteText("InstallFormulaOcrAsync")
            End If
            Return Await PaddleOcrService.InstallFormulaOcrAsync()
        End Function

        ''' <summary>
        ''' 获取云端 OCR 配置
        ''' </summary>
        Public Function GetCloudOcrConfig() As CloudOcrConfig
            If OcrHostGate.ShouldForward Then
                Try
                    Return OcrHostClient.Invoke(Of CloudOcrConfig)("OcrServiceManager.GetCloudOcrConfig")
                Catch
                    Return CloudOcrConfig.Load()
                End Try
            End If
            Return CloudOcrService.GetConfig()
        End Function

        ''' <summary>
        ''' 保存云端 OCR 配置
        ''' </summary>
        Public Sub SaveCloudOcrConfig(config As CloudOcrConfig)
            If OcrHostGate.ShouldForward Then
                OcrHostClient.Invoke(Of Object)("OcrServiceManager.SaveCloudOcrConfig", config)
                Return
            End If
            CloudOcrService.SaveConfig(config)
        End Sub

        ''' <summary>
        ''' 获取云端 OCR 提供商列表
        ''' </summary>
        Public Function GetCloudProviders() As Dictionary(Of String, String)
            If OcrHostGate.ShouldForward Then
                Try
                    Return OcrHostClient.Invoke(Of Dictionary(Of String, String))("OcrServiceManager.GetCloudProviders")
                Catch
                    Return New Dictionary(Of String, String)
                End Try
            End If
            Return CloudOcrService.GetProviders()
        End Function

        ''' <summary>
        ''' 获取引擎显示名称
        ''' </summary>
        ''' <summary>把浏览器里的图片落到临时文件，供独立 OCR 服务读取。</summary>
        Public Function SaveBase64Image(base64Image As String, fileName As String) As String
            Try
                Dim base64Data As String = If(base64Image, "")
                Dim comma = base64Data.IndexOf(","c)
                If comma >= 0 Then base64Data = base64Data.Substring(comma + 1)
                Dim bytes = Convert.FromBase64String(base64Data)
                Dim tempDir = Path.Combine(Path.GetTempPath(), "ExcelAIOCR")
                If Not Directory.Exists(tempDir) Then Directory.CreateDirectory(tempDir)
                Dim ext = Path.GetExtension(If(fileName, ""))
                If String.IsNullOrEmpty(ext) Then ext = ".png"
                Dim stem = Path.GetFileNameWithoutExtension(If(fileName, "image"))
                If String.IsNullOrEmpty(stem) Then stem = "image"
                For Each invalid In Path.GetInvalidFileNameChars()
                    stem = stem.Replace(invalid, "_"c)
                Next
                Dim filePath = Path.Combine(tempDir, stem & "_" & DateTime.Now.ToString("yyyyMMddHHmmssfff") & ext)
                File.WriteAllBytes(filePath, bytes)
                Return JsonConvert.SerializeObject(New With {.success = True, .path = filePath})
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>供独立服务探活，不启动识别引擎。</summary>
        Public Function HostEcho(value As String) As String
            Return If(value, "")
        End Function

        ''' <summary>清掉 RapidOCR 缓存后重新探测。插件侧会转发到 OcrHost。</summary>
        Public Function ForceCheckRapidAvailable() As Boolean
            If OcrHostGate.ShouldForward Then
                Try
                    Return OcrHostClient.Invoke(Of Boolean)("OcrServiceManager.ForceCheckRapidAvailable")
                Catch
                    Return False
                End Try
            End If
            Return RapidOcrService.ForceCheckAvailable()
        End Function

        ''' <summary>让独立服务重新读取云端 OCR 配置。</summary>
        Public Sub ReloadCloudOcrConfig()
            If OcrHostGate.ShouldForward Then
                Try
                    OcrHostClient.Invoke(Of Object)("OcrServiceManager.ReloadCloudOcrConfig")
                Catch
                End Try
                Return
            End If
            CloudOcrService.ReloadConfig()
        End Sub

        Public Shared Function GetEngineDisplayName(engine As OcrEngineType) As String
            Select Case engine
                Case OcrEngineType.RapidOCR
                    Return "RapidOCR (嵌入式，轻量快速)"
                Case OcrEngineType.PaddleOCR
                    Return "PaddleOCR (中文推荐)"
                Case OcrEngineType.CloudOCR
                    Return "云端 OCR API"
                Case Else
                    Return "未知"
            End Select
        End Function

        ''' <summary>
        ''' 获取引擎描述
        ''' </summary>
        Public Shared Function GetEngineDescription(engine As OcrEngineType) As String
            Select Case engine
                Case OcrEngineType.RapidOCR
                    Return "ONNX 轻量级 OCR，模型小（约 30MB），启动快，随插件分发"
                Case OcrEngineType.PaddleOCR
                    Return "百度开源 OCR，中文识别效果最佳，需要安装 Python"
                Case OcrEngineType.CloudOCR
                    Return "云端 OCR 服务，识别准确率高，需要 API Key 和联网"
                Case Else
                    Return ""
            End Select
        End Function

        ''' <summary>
        ''' 释放所有 OCR 服务实例持有的资源（请在应用关闭时调用）
        ''' 【增强】除 IDisposable 外，还清空所有服务引用以允许 GC 回收
        ''' </summary>
        Public Sub Dispose()
            If OcrHostGate.ShouldForward Then Return
            If _disposed Then Return
            _disposed = True
            ReleaseIfDisposable(_rapidOcrService)
            ReleaseIfDisposable(_paddleOcrService)
            ReleaseIfDisposable(_cloudOcrService)
            _rapidOcrService = Nothing
            _paddleOcrService = Nothing
            _cloudOcrService = Nothing
        End Sub

        ''' <summary>
        ''' 安全释放实现了 IDisposable 的对象
        ''' </summary>
        Private Shared Sub ReleaseIfDisposable(obj As Object)
            If obj Is Nothing Then Return
            Try
                Dim disposable = TryCast(obj, System.IDisposable)
                If disposable IsNot Nothing Then disposable.Dispose()
            Catch ex As Exception  ' 
            End Try
        End Sub

        ''' <summary>
        ''' 检测 PaddleOCR 模型文件是否已下载到本地
        ''' PaddleOCR 2.x: ~/.paddleocr/det|rec|cls
        ''' PaddleOCR 3.x: ~/.paddlex/official_models/
        ''' </summary>
        Private Function ArePaddleModelsReady() As Boolean
            Try
                Dim home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
                ' 检测 3.x 新路径
                Dim modelDirV3 = Path.Combine(home, ".paddlex", "official_models")
                If Directory.Exists(modelDirV3) AndAlso Directory.GetDirectories(modelDirV3).Length > 0 Then
                    Return True
                End If
                ' 兼容 2.x 旧路径
                Dim modelDirV2 = Path.Combine(home, ".paddleocr")
                If Directory.Exists(modelDirV2) Then
                    Dim subDirs = Directory.GetDirectories(modelDirV2)
                    For Each subDir In subDirs
                        If Directory.GetFiles(subDir).Length > 0 OrElse Directory.GetDirectories(subDir).Length > 0 Then
                            Return True
                        End If
                    Next
                End If
                Return False
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("检测 PaddleOCR 模型失败: " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 后台启动 PaddleOCR 模型下载（不阻塞当前操作）
        ''' </summary>
        Private Sub TriggerPaddleModelDownload()
            Task.Run(Sub()
                Try
                    System.Diagnostics.Debug.WriteLine("后台下载 PaddleOCR 模型开始...")
                    Dim pythonPath = Me.PaddleOcrService.GetPythonPathPublic()
                    If String.IsNullOrEmpty(pythonPath) Then
                        System.Diagnostics.Debug.WriteLine("后台下载模型：未找到 Python 路径")
                        Return
                    End If
                    Dim psi As New ProcessStartInfo()
                    psi.FileName = pythonPath
                    psi.Arguments = "-c ""from paddleocr import PaddleOCR; PaddleOCR(use_textline_orientation=True, lang='ch'); print('model download complete')"""
                    psi.UseShellExecute = False
                    psi.CreateNoWindow = True
                    psi.RedirectStandardOutput = True
                    psi.RedirectStandardError = True
                    Using p = Process.Start(psi)
                        Dim outTask = p.StandardOutput.ReadToEndAsync()
                        Dim errTask = p.StandardError.ReadToEndAsync()
                        If Not p.WaitForExit(600000) Then
                            Try : p.Kill() : Catch : End Try
                            System.Diagnostics.Debug.WriteLine("后台下载模型超时")
                            Return
                        End If
                        Dim err = ""
                        Try
                            err = errTask.Result
                        Catch ex As Exception
                        End Try
                        System.Diagnostics.Debug.WriteLine($"后台下载模型退出码: {p.ExitCode}")
                        If p.ExitCode = 0 Then
                            System.Diagnostics.Debug.WriteLine("PaddleOCR 模型下载成功，下次识别将使用 PaddleOCR")
                        Else
                            System.Diagnostics.Debug.WriteLine($"PaddleOCR 模型下载失败: {err.Substring(0, Math.Min(500, err.Length))}")
                        End If
                    End Using
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("后台下载 PaddleOCR 模型异常: " & ex.Message)
                End Try
            End Sub)
        End Sub

#Region "结构化识别与批量处理"

        ''' <summary>
        ''' 执行结构化 OCR 识别，返回 OcrDocument
        ''' 包含位置信息构建的结构化元素（段落、标题、列表等）
        ''' </summary>
        Public Async Function RecognizeDocumentAsync(imagePath As String, Optional fileName As String = "", Optional engine As OcrEngineType? = Nothing) As Task(Of OcrDocument)
            If OcrHostGate.ShouldForward Then
                Return RemoteDocument("RecognizeDocumentAsync", imagePath, fileName, engine)
            End If
            RefreshFromDisk()
            Dim engineToUse = If(engine.HasValue, engine.Value, _currentEngine)
            Dim engineName = GetEngineDisplayName(engineToUse)
            Dim doc As New OcrDocument()
            doc.SourceFilePath = imagePath
            doc.SourceFileName = If(String.IsNullOrEmpty(fileName), Path.GetFileName(imagePath), fileName)
            doc.EngineName = engineName

            If Not IsEngineAvailable(engineToUse) Then
                doc.ErrorMessage = $"【错误：{engineName} 当前不可用】"
                doc.RawText = doc.ErrorMessage
                Return doc
            End If

            Dim docResult As OcrDocument = doc
            Select Case engineToUse
                Case OcrEngineType.RapidOCR
                    docResult = Await RapidOcrService.RecognizeDocumentAsync(imagePath, doc.SourceFileName, engineName)

                Case OcrEngineType.PaddleOCR
                    ' PaddleOCR 有位置信息，做结构化输出
                    If Not ArePaddleModelsReady() Then
                        TriggerPaddleModelDownload()
                    End If
                    ' 同步当前高级配置到 PaddleOCR 服务
                    Dim paddleSvc = PaddleOcrService
                    paddleSvc.Language = _paddleOcrLanguage
                    paddleSvc.EnableLayoutAnalysis = _enableLayoutAnalysis
                    paddleSvc.EnableTableRecognition = _enableTableRecognition
                    paddleSvc.EnableFormulaRecognition = _enableFormulaRecognition
                    docResult = Await paddleSvc.RecognizeDocumentAsync(imagePath, doc.SourceFileName, engineName)

                Case OcrEngineType.CloudOCR
                    Dim rawText = Await CloudOcrService.RecognizeText(imagePath)
                    doc.RawText = If(rawText, "")
                    If OcrDocument.IsEngineErrorText(doc.RawText) Then
                        doc.ErrorMessage = doc.RawText
                    Else
                        doc.Elements = HeuristicParse(rawText)
                    End If
                    docResult = doc

                Case Else
                    doc.ErrorMessage = "【错误：未知的 OCR 引擎类型】"
                    doc.RawText = doc.ErrorMessage
                    docResult = doc
            End Select

            If docResult IsNot Nothing AndAlso String.IsNullOrEmpty(docResult.ErrorMessage) AndAlso OcrDocument.IsEngineErrorText(docResult.RawText) Then
                docResult.ErrorMessage = docResult.RawText
            End If
            Return docResult
        End Function

        ''' <summary>
        ''' 执行 OCR 识别并返回指定格式的结构化文本
        ''' </summary>
        Public Async Function RecognizeTextWithFormatAsync(imagePath As String, outputFormat As OcrOutputFormat,
                                                           Optional fileName As String = "",
                                                           Optional engine As OcrEngineType? = Nothing) As Task(Of String)
            If OcrHostGate.ShouldForward Then
                Return RemoteText("RecognizeTextWithFormatAsync", imagePath, outputFormat, fileName, engine)
            End If
            RefreshFromDisk()
            Dim doc = Await RecognizeDocumentAsync(imagePath, fileName, engine)
            Return OcrFormatFormatter.Format(doc, outputFormat)
        End Function

        ''' <summary>
        ''' 批量 OCR 识别
        ''' 处理文件列表，通过 progress 回调报告进度
        ''' </summary>
        Public Async Function BatchOcrAsync(filePaths As List(Of String),
                                            Optional outputFormat As OcrOutputFormat = OcrOutputFormat.PlainText,
                                            Optional engine As OcrEngineType? = Nothing,
                                            Optional progress As IProgress(Of OcrBatchProgress) = Nothing) As Task(Of List(Of OcrDocument))
            Dim results As New List(Of OcrDocument)()
            If filePaths Is Nothing OrElse filePaths.Count = 0 Then Return results

            Dim total As Integer = filePaths.Count
            Dim completed As Integer = 0
            Dim failed As Integer = 0

            For i As Integer = 0 To total - 1
                Dim filePath = filePaths(i)
                Dim fileName = Path.GetFileName(filePath)

                ' 报告进度
                If progress IsNot Nothing Then
                    progress.Report(New OcrBatchProgress() With {
                        .TotalTasks = total,
                        .CompletedTasks = completed,
                        .FailedTasks = failed,
                        .ProcessingTasks = 1,
                        .CurrentFileName = fileName,
                        .StatusMessage = $"正在处理 ({i + 1}/{total}): {fileName}"
                    })
                End If

                Try
                    Dim doc = Await RecognizeDocumentAsync(filePath, fileName, engine)
                    doc.SourceFilePath = filePath
                    results.Add(doc)
                    If Not String.IsNullOrEmpty(doc.ErrorMessage) OrElse OcrDocument.IsEngineErrorText(doc.RawText) Then
                        failed += 1
                    Else
                        completed += 1
                    End If
                Catch ex As Exception
                    Dim errorDoc As New OcrDocument()
                    errorDoc.SourceFileName = fileName
                    errorDoc.SourceFilePath = filePath
                    errorDoc.RawText = $"【处理失败】: {ex.Message}"
                    results.Add(errorDoc)
                    failed += 1
                End Try
            Next

            ' 最终进度
            If progress IsNot Nothing Then
                progress.Report(New OcrBatchProgress() With {
                    .TotalTasks = total,
                    .CompletedTasks = completed,
                    .FailedTasks = failed,
                    .ProcessingTasks = 0,
                    .StatusMessage = $"批量完成: {completed} 成功, {failed} 失败"
                })
            End If

            Return results
        End Function

        ''' <summary>
        ''' 对纯文本进行启发式结构化解析
        ''' 用于不支持位置信息的 OCR 引擎，也用于快速路径生成 Markdown/HTML/Excel 格式
        ''' </summary>
        Public Function HeuristicParse(text As String) As List(Of OcrElement)
            Dim elements As New List(Of OcrElement)()
            If String.IsNullOrEmpty(text) Then Return elements

            Dim lines = text.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
            Dim paraLines As New List(Of String)()

            For Each line In lines
                Dim trimmed As String = line.Trim()

                ' 空行 → 段落结束
                If String.IsNullOrEmpty(trimmed) Then
                    If paraLines.Count > 0 Then
                        elements.Add(New OcrElement() With {
                            .Type = OcrElementType.Paragraph,
                            .Text = String.Join(" ", paraLines).Trim()
                        })
                        paraLines.Clear()
                    End If
                    Continue For
                End If

                ' 分隔线
                If trimmed.All(Function(c) c = "-"c OrElse c = "="c OrElse c = "*"c) AndAlso trimmed.Length >= 3 Then
                    If paraLines.Count > 0 Then
                        elements.Add(New OcrElement() With {
                            .Type = OcrElementType.Paragraph,
                            .Text = String.Join(" ", paraLines).Trim()
                        })
                        paraLines.Clear()
                    End If
                    elements.Add(New OcrElement() With {.Type = OcrElementType.HorizontalRule})
                    Continue For
                End If

                ' 检测制表符表格
                If trimmed.Contains(vbTab) Then
                    If paraLines.Count > 0 Then
                        elements.Add(New OcrElement() With {
                            .Type = OcrElementType.Paragraph,
                            .Text = String.Join(" ", paraLines).Trim()
                        })
                        paraLines.Clear()
                    End If
                    elements.Add(New OcrElement() With {
                        .Type = OcrElementType.Table,
                        .Text = trimmed
                    })
                    Continue For
                End If

                ' 短行且以冒号结尾，视为标题。普通短句不再升成标题。
                If trimmed.Length < 60 AndAlso (trimmed.EndsWith("：") OrElse trimmed.EndsWith(":")) Then
                    If paraLines.Count > 0 Then
                        elements.Add(New OcrElement() With {
                            .Type = OcrElementType.Paragraph,
                            .Text = String.Join(" ", paraLines).Trim()
                        })
                        paraLines.Clear()
                    End If
                    elements.Add(New OcrElement() With {
                        .Type = OcrElementType.Heading,
                        .HeadingLevel = 2,
                        .Text = trimmed
                    })
                    Continue For
                End If

                ' 检测列表项
                If trimmed.StartsWith("- ") OrElse trimmed.StartsWith("* ") OrElse
                   (trimmed.Length > 2 AndAlso Char.IsDigit(trimmed(0)) AndAlso (trimmed(1) = "."c OrElse trimmed(1) = ")"c)) Then
                    If paraLines.Count > 0 Then
                        elements.Add(New OcrElement() With {
                            .Type = OcrElementType.Paragraph,
                            .Text = String.Join(" ", paraLines).Trim()
                        })
                        paraLines.Clear()
                    End If
                    elements.Add(New OcrElement() With {
                        .Type = OcrElementType.ListItem,
                        .ListMarker = trimmed(0).ToString(),
                        .Text = trimmed
                    })
                    Continue For
                End If

                ' 普通行，加入段落缓存
                paraLines.Add(trimmed)
            Next

            ' 处理最后缓存的段落
            If paraLines.Count > 0 Then
                elements.Add(New OcrElement() With {
                    .Type = OcrElementType.Paragraph,
                    .Text = String.Join(" ", paraLines).Trim()
                })
            End If

            Return elements
        End Function

#End Region

    End Class

    ''' <summary>
    ''' OCR 引擎验证信息（持久化到本地，避免每次重启 Excel 都重新验证）
    ''' </summary>
    Public Class OcrEngineVerificationInfo
        Public Property IsAvailable As Boolean = False
        Public Property PythonPath As String = ""
        Public Property VerifiedAt As DateTime? = Nothing
    End Class

    ''' <summary>
    ''' OCR 管理器配置
    ''' </summary>
    Public Class OcrManagerConfig

        ' 默认引擎：RapidOCR（嵌入式 ONNX 轻量，随插件分发零配置）；PaddleOCR 需 Python 环境，CloudOCR 需 API Key
        Public Property SelectedEngine As OcrEngineType = OcrEngineType.RapidOCR
        Public Property PaddleOcrLanguage As String = "ch"
        Public Property EnableLayoutAnalysis As Boolean = True
        Public Property EnableTableRecognition As Boolean = True
        Public Property EnableFormulaRecognition As Boolean = True

        ''' <summary>PaddleOCR 上次验证信息</summary>
        Public Property PaddleOcrVerification As OcrEngineVerificationInfo
        ''' <summary>RapidOCR 上次验证信息</summary>
        Public Property RapidOcrVerification As OcrEngineVerificationInfo

        Private Shared ReadOnly _saveLock As New Object()
        Private Shared ReadOnly ConfigPath As String = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelAIVSTOAddIn", "ocr_manager_config.json")

        ''' <summary>
        ''' 加载配置
        ''' </summary>
        Public Shared Function Load() As OcrManagerConfig
            Try
                If File.Exists(ConfigPath) Then
                    Dim json = File.ReadAllText(ConfigPath)
                    Dim loaded = JsonConvert.DeserializeObject(Of OcrManagerConfig)(json)
                    If loaded IsNot Nothing Then Return loaded
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("加载 OCR 管理器配置失败: " & ex.Message)
            End Try

            Return New OcrManagerConfig()
        End Function

        ''' <summary>
        ''' 保存配置
        ''' </summary>
        Public Sub Save()
            SyncLock _saveLock
                Try
                    Dim dir = Path.GetDirectoryName(ConfigPath)
                    If Not Directory.Exists(dir) Then
                        Directory.CreateDirectory(dir)
                    End If

                    Dim json = JsonConvert.SerializeObject(Me, Formatting.Indented)
                    File.WriteAllText(ConfigPath, json)
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("保存 OCR 管理器配置失败: " & ex.Message)
                End Try
            End SyncLock
        End Sub

    End Class

End Namespace
