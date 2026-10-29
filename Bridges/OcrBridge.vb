Option Strict Off
Option Explicit On
Option Infer On

Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Diagnostics
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json.Serialization
Imports ExcelAIVSTOAddIn.Utils
Imports ExcelAIVSTOAddIn.Services

Namespace Bridges

    ''' <summary>
    ''' OcrBridge - OCR 引擎管理桥梁类
    ''' 提供 COM 可见的方法供前端 JavaScript 调用，用于 OCR 引擎状态检测、配置管理等
    ''' </summary>
    <ComVisible(True)>
    <ClassInterface(ClassInterfaceType.AutoDispatch)>
    <ProgId("ExcelAIVSTOAddIn.OcrBridge")>
    Public Class OcrBridge
        Implements IBridge

        ' 单例实例 - 使用 Lazy(Of T) 确保线程安全
        Private Shared ReadOnly _lazyInstance As New Lazy(Of OcrBridge)(Function() New OcrBridge(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        ' 配置管理器
        Private _config As ConfigurationManager

        ' 安装检测会启动 Python 并 import paddleocr，一次要十几秒。
        ' 同一进程里切换面板或引擎时直接用上次的结果。
        Private Shared _paddleReadyStatus As Object = Nothing
        Private Shared _paddleReadyAt As DateTime = DateTime.MinValue
        Private Shared _formulaReadyJson As String = Nothing

        ' OCR 异步任务状态（从 ExcelBridge 迁移，消除 Facade 层状态耦合）
        Private _ocrGeneration As Integer = 0
        Private _ocrTask As Task(Of OcrDocument) = Nothing
        Private _ocrResult As String = Nothing
        Private _ocrDocument As OcrDocument = Nothing
        Private _ocrOutputFormat As OcrOutputFormat = OcrOutputFormat.PlainText
        Private _ocrStructuredResult As String = Nothing  ' 结构化结果（含 box 坐标）
        Private _ocrRawData As String = Nothing  ' PaddleOCR 原始数据（含坐标、文字、置信度）
        Private _ocrError As String = Nothing
        Private _ocrCompleted As Boolean = False
        Private _lastImagePath As String = Nothing
        Private _lastRecognitionMode As String = Nothing
        Private _ocrStructuredTask As Task = Nothing

        ' PDF OCR 异步任务状态机（独立于单页 OCR）
        Private _ocrPdfGeneration As Integer = 0
        Private _ocrPdfTask As Task(Of List(Of OcrDocument)) = Nothing
        Private _ocrPdfDocuments As List(Of OcrDocument) = Nothing
        Private _ocrPdfCompleted As Boolean = False
        Private _ocrPdfError As String = Nothing

        ' 批量处理共享实例（StartBatchOcr / GetBatchProgress / CancelBatchOcr 共用，
        ' 修复旧实现每次 New 一个局部处理器导致进度永远查不到的问题）
        Private Shared _batchProcessor As Services.OcrBatchProcessor = Nothing
        Private Shared ReadOnly _batchLock As New Object()

#Region "单例模式"

        ''' <summary>
        ''' 获取 OcrBridge 单例实例（线程安全 Lazy 初始化）
        ''' </summary>
        Public Shared ReadOnly Property Instance As OcrBridge
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        ''' <summary>
        ''' 私有构造函数（单例模式）
        ''' </summary>
        Private Sub New()
            ' 统一生命周期管理：实例化即登记，由 BridgeRegistry.ShutdownAll 统一释放
            BridgeRegistry.Instance.Register(Me)
            Try
                _config = ConfigurationManager.Instance
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("OcrBridge 构造函数异常: " & ex.Message)
            End Try
        End Sub

#End Region

#Region "OCR 异步识别 API（状态机封装于此）"

        ''' <summary>
        ''' 启动 OCR 识别（异步，立即返回，不阻塞 COM 线程）
        ''' </summary>
        ''' <param name="imagePath">图片路径</param>
        ''' <param name="recognitionMode">识别模式: auto/text/formula/number/date</param>
        <ComVisible(True)>
        Public Function StartOcrRecognition(imagePath As String, Optional recognitionMode As String = "auto") As String
            _ocrOutputFormat = OcrOutputFormat.PlainText
            Return StartOcrRecognitionCore(imagePath, recognitionMode)
        End Function

        ''' <summary>
        ''' 启动 OCR 识别并指定输出格式（异步，立即返回，不阻塞 COM 线程）
        ''' </summary>
        ''' <param name="imagePath">图片路径</param>
        ''' <param name="recognitionMode">识别模式: auto/text/formula/number/date</param>
        ''' <param name="outputFormat">输出格式: plaintext/markdown/html/excel/csv</param>
        <ComVisible(True)>
        Public Function StartOcrRecognitionWithFormat(imagePath As String, recognitionMode As String, outputFormat As String) As String
            Select Case outputFormat.ToLower()
                Case "markdown" : _ocrOutputFormat = OcrOutputFormat.Markdown
                Case "html" : _ocrOutputFormat = OcrOutputFormat.Html
                Case "excel", "csv" : _ocrOutputFormat = OcrOutputFormat.Excel
                Case Else : _ocrOutputFormat = OcrOutputFormat.PlainText
            End Select
            Return StartOcrRecognitionCore(imagePath, recognitionMode)
        End Function

        ''' <summary>
        ''' 启动 PDF OCR 识别（异步，立即返回，不阻塞 COM 线程）
        ''' 引擎侧完成 PDF → 高分辨率图片 → 逐页 OCR → 按页分组。
        ''' 机械图纸模式下每页自动跑 GDT 符号识别。
        ''' </summary>
        ''' <param name="pdfPath">PDF 文件路径</param>
        ''' <param name="recognitionMode">识别模式: auto / mechanical / gdt / drawing</param>
        ''' <param name="dpi">PDF 转图分辨率，默认 300（图纸推荐 300-400）</param>
        ''' <returns>JSON 状态 {"generation": n, "completed": false}</returns>
        <ComVisible(True)>
        Public Function StartPdfOcrRecognition(pdfPath As String,
                                                 Optional recognitionMode As String = "auto",
                                                 Optional dpi As Integer = 300) As String
            Debug.WriteLine($"[OcrBridge] StartPdfOcrRecognition path={pdfPath}, mode={recognitionMode}, dpi={dpi}")

            If String.IsNullOrEmpty(pdfPath) OrElse Not File.Exists(pdfPath) Then
                Return JsonConvert.SerializeObject(New With {
                    .generation = 0, .completed = True, .ok = False,
                    .error = "PDF 文件不存在: " & pdfPath
                })
            End If

            ' 重置 PDF OCR 状态
            Dim gen = Threading.Interlocked.Increment(_ocrPdfGeneration)
            _ocrPdfCompleted = False
            _ocrPdfDocuments = Nothing
            _ocrPdfError = Nothing

            _ocrPdfTask = Task.Run(Function()
                Try
                    Dim ocrManager = Services.OcrServiceManager.Instance
                    ocrManager.CurrentEngine = Services.OcrEngineType.RapidOCR

                    Debug.WriteLine($"[OcrBridge] PDF OCR 第 {gen} 代，开始识别...")
                    Dim docs = ocrManager.RecognizePdfWithPositionAsync(pdfPath, Path.GetFileName(pdfPath), recognitionMode, dpi).GetAwaiter().GetResult()

                    If gen <> _ocrPdfGeneration Then
                        Debug.WriteLine("[OcrBridge] PDF OCR 第 " & gen & " 代结果已过时（被新请求取代），丢弃")
                        Return docs
                    End If

                    _ocrPdfDocuments = docs
                    _ocrPdfCompleted = True
                    Debug.WriteLine($"[OcrBridge] PDF OCR 第 {gen} 代完成，共 {If(docs IsNot Nothing, docs.Count, 0)} 页")
                    Return docs
                Catch ex As Exception
                    Debug.WriteLine("[OcrBridge] PDF OCR 异常: " & ex.ToString())
                    If gen = _ocrPdfGeneration Then
                        _ocrPdfError = ex.Message
                        _ocrPdfCompleted = True
                    End If
                    Return Nothing
                End Try
            End Function)

            Return JsonConvert.SerializeObject(New With {.generation = gen, .completed = False})
        End Function

        ''' <summary>
        ''' 轮询 PDF OCR 结果（由前端定时调用）
        ''' </summary>
        ''' <returns>JSON 状态，未完成时 {"generation": n, "completed": false}；
        ''' 完成时 {"generation": n, "completed": true, "ok": true, "pages": [...]} 或错误</returns>
        <ComVisible(True)>
        Public Function GetPdfOcrResult() As String
            Try
                If Not _ocrPdfCompleted Then
                    Return JsonConvert.SerializeObject(New With {
                        .generation = _ocrPdfGeneration, .completed = False
                    })
                End If

                If Not String.IsNullOrEmpty(_ocrPdfError) Then
                    Return JsonConvert.SerializeObject(New With {
                        .generation = _ocrPdfGeneration,
                        .completed = True,
                        .ok = False,
                        .error = _ocrPdfError
                    })
                End If

                If _ocrPdfDocuments Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {
                        .generation = _ocrPdfGeneration,
                        .completed = True,
                        .ok = False,
                        .error = "PDF OCR 未返回结果"
                    })
                End If

                ' 按页序列化
                Dim pagesJson As New List(Of Object)()
                For Each doc In _ocrPdfDocuments
                    Dim pageObj = New With {
                        .page_number = doc.CurrentPage,
                        .page_index = doc.CurrentPage - 1,
                        .total_pages = doc.TotalPages,
                        .engine = doc.EngineName,
                        .elapsed_seconds = doc.ElapsedSeconds,
                        .raw_text = doc.RawText,
                        .error = doc.ErrorMessage,
                        .elements = doc.Elements.Select(Function(el) New With {
                            .type = el.Type.ToString(),
                            .kind = el.Kind,
                            .text = el.Text,
                            .nominal = el.Nominal,
                            .lower = el.Lower,
                            .upper = el.Upper,
                            .datum = el.Datum,
                            .confidence = el.Confidence,
                            .bbox = If(el.BoundingBox Is Nothing, Nothing,
                                CType(New Object() {el.BoundingBox.Left, el.BoundingBox.Top, el.BoundingBox.Right, el.BoundingBox.Bottom}, Object()))
                        }).ToList()
                    }
                    pagesJson.Add(pageObj)
                Next

                Dim ok = _ocrPdfDocuments.All(Function(d) String.IsNullOrEmpty(d.ErrorMessage))
                Return JsonConvert.SerializeObject(New With {
                    .generation = _ocrPdfGeneration,
                    .completed = True,
                    .ok = ok,
                    .page_count = _ocrPdfDocuments.Count,
                    .pages = pagesJson
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .generation = _ocrPdfGeneration,
                    .completed = True,
                    .ok = False,
                    .error = "GetPdfOcrResult 序列化异常: " & ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 启动 OCR 识别的公共实现
        ''' </summary>
        Private Function StartOcrRecognitionCore(imagePath As String, recognitionMode As String) As String
            System.Diagnostics.Debug.WriteLine($"[OcrBridge] StartOcrRecognitionCore 被调用，路径: {imagePath}, 模式: {recognitionMode}, 格式: {_ocrOutputFormat}")
            Try
                ' 取消之前的任务（如果还在运行）
                If _ocrTask IsNot Nothing AndAlso Not _ocrTask.IsCompleted Then
                    System.Diagnostics.Debug.WriteLine("[OcrBridge] 取消之前的 OCR 任务")
                    ' 注意：Task.Run 不支持直接取消，这里只是标记状态
                    ' 实际取消依赖于前端防重复点击机制
                End If

                ' 重置状态。旧任务不能真正取消，用代数丢弃它完成后的写回。
                Dim generation = Threading.Interlocked.Increment(_ocrGeneration)
                _ocrCompleted = False
                _ocrResult = Nothing
                _ocrDocument = Nothing
                _ocrStructuredResult = Nothing
                _ocrRawData = Nothing
                _ocrError = Nothing
                _ocrStructuredTask = Nothing
                _lastImagePath = imagePath
                _lastRecognitionMode = recognitionMode

                ' 启动后台 OCR 任务（不在 COM 线程上等待）
                _ocrTask = Task.Run(Function()
                    Dim sw = Stopwatch.StartNew()
                    System.Diagnostics.Debug.WriteLine($"[OcrBridge] Task.Run 开始，准备同步引擎配置")
                    Dim ocrManager = Services.OcrServiceManager.Instance

                    ' 从配置同步当前引擎
                    Dim engineName = GetCurrentOcrEngine()
                    System.Diagnostics.Debug.WriteLine($"[OcrBridge] 配置引擎: {engineName}, 识别模式: {recognitionMode}")
                    Dim engineType As Services.OcrEngineType
                    Select Case engineName.ToLower()
                        Case "paddleocr"
                            engineType = Services.OcrEngineType.PaddleOCR
                        Case "rapidocr"
                            engineType = Services.OcrEngineType.RapidOCR
                        Case "cloud"
                            engineType = Services.OcrEngineType.CloudOCR
                        Case Else
                            engineType = Services.OcrEngineType.RapidOCR
                    End Select
                    ocrManager.CurrentEngine = engineType

                    ' 统一走带位置的快速识别路径，所有格式共享同一套缓存结果
                    ' 避免格式切换时重复识别，也避免 Excel/Markdown 因缺坐标而无法重建表格
                    System.Diagnostics.Debug.WriteLine("[OcrBridge] 使用带位置的统一识别路径")
                    Dim doc = ocrManager.RecognizeTextWithPositionAsync(imagePath, Path.GetFileName(imagePath), recognitionMode).GetAwaiter().GetResult()
                    If doc Is Nothing Then
                        doc = New OcrDocument() With {
                            .SourceFileName = Path.GetFileName(imagePath),
                            .SourceFilePath = imagePath,
                            .EngineName = engineName,
                            .TotalPages = 1,
                            .CurrentPage = 1,
                            .RawText = ""
                        }
                    End If

                    ' 原始坐标在独立 OCR 进程里生成，随文档带回。只有识别仍在本进程时才读本地引擎。
                    Try
                        If doc IsNot Nothing AndAlso Not String.IsNullOrEmpty(doc.RawEngineResponse) Then
                            _ocrRawData = doc.RawEngineResponse
                        ElseIf Not Services.OcrHostGate.ShouldForward Then
                            If engineType = Services.OcrEngineType.PaddleOCR Then
                                Dim paddleSvc = ocrManager.PaddleOcrService
                                If paddleSvc IsNot Nothing Then
                                    _ocrRawData = paddleSvc.GetLastRawResponse()
                                End If
                            ElseIf engineType = Services.OcrEngineType.RapidOCR Then
                                Dim rapidSvc = ocrManager.RapidOcrService
                                If rapidSvc IsNot Nothing Then
                                    _ocrRawData = rapidSvc.GetLastRawResponse()
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine($"[OcrBridge] 获取原始数据失败（非致命）: {ex.Message}")
                    End Try

                    sw.Stop()
                    System.Diagnostics.Debug.WriteLine($"[OcrBridge] OCR 返回，RawText 长度: {doc?.RawText?.Length}，Elements: {doc?.Elements?.Count}，耗时: {sw.ElapsedMilliseconds}ms")
                    Return doc
                End Function)

                ' 任务完成后更新状态（ContinueWith 在后台线程执行）
                _ocrTask.ContinueWith(Sub(t)
                    If generation <> _ocrGeneration Then Return
                    If t.IsFaulted Then
                        Dim ex = t.Exception?.InnerException
                        If ex Is Nothing Then ex = t.Exception
                        _ocrError = If(ex?.Message, "未知错误")
                        System.Diagnostics.Debug.WriteLine($"[OcrBridge] OCR 失败: {_ocrError}")
                    Else
                        _ocrDocument = t.Result
                        If t.Result IsNot Nothing AndAlso Not String.IsNullOrEmpty(t.Result.ErrorMessage) Then
                            _ocrError = t.Result.ErrorMessage
                            _ocrResult = t.Result.ErrorMessage
                        ElseIf t.Result IsNot Nothing AndAlso OcrDocument.IsEngineErrorText(t.Result.RawText) AndAlso
                               (t.Result.Elements Is Nothing OrElse t.Result.Elements.Count = 0) Then
                            _ocrError = t.Result.RawText
                            _ocrResult = t.Result.RawText
                        Else
                            _ocrResult = If(t.Result?.RawText, "")
                        End If
                        System.Diagnostics.Debug.WriteLine($"[OcrBridge] OCR 结果已保存，长度: {_ocrResult?.Length}")
                    End If
                    _ocrCompleted = True
                End Sub)

                Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {.started = True})
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[OcrBridge] StartOcrRecognitionCore 异常: {ex.Message}")
                Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {.started = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 轮询 OCR 识别结果（非阻塞）
        ''' </summary>
        <ComVisible(True)>
        Public Function GetOcrResult() As String
            Return GetOcrResultCore(_ocrOutputFormat)
        End Function

        ''' <summary>
        ''' 轮询 OCR 识别结果（非阻塞，指定输出格式）
        ''' </summary>
        ''' <param name="format">输出格式: plaintext/markdown/html/excel/csv</param>
        <ComVisible(True)>
        Public Function GetOcrResultWithFormat(format As String) As String
            Dim selectedFormat As OcrOutputFormat
            Select Case format.ToLower()
                Case "markdown" : selectedFormat = OcrOutputFormat.Markdown
                Case "html" : selectedFormat = OcrOutputFormat.Html
                Case "excel", "csv" : selectedFormat = OcrOutputFormat.Excel
                Case Else : selectedFormat = OcrOutputFormat.PlainText
            End Select
            Return GetOcrResultCore(selectedFormat)
        End Function

        ''' <summary>
        ''' 轮询 OCR 识别结果的公共实现
        ''' </summary>
        Private Function GetOcrResultCore(selectedFormat As OcrOutputFormat) As String
            If Not _ocrCompleted Then
                Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {.ready = False})
            End If

            If _ocrError IsNot Nothing Then
                Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {.ready = True, .success = False, .error = _ocrError})
            End If

            ' 如果 _ocrDocument 还没赋值但任务已完成，直接从任务结果取（避免 ContinueWith 与轮询的竞态）
            Dim doc As OcrDocument = _ocrDocument
            If doc Is Nothing AndAlso _ocrTask IsNot Nothing AndAlso _ocrTask.IsCompleted AndAlso Not _ocrTask.IsFaulted Then
                doc = _ocrTask.Result
            End If

            ' 如果有结构化文档，按目标格式输出；否则回退到原始文本
            Dim outputText As String
            If doc IsNot Nothing Then
                outputText = OcrFormatFormatter.Format(doc, selectedFormat)
            Else
                outputText = _ocrResult
            End If

            Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                .ready = True,
                .success = True,
                .text = outputText,
                .rawText = _ocrResult,
                .format = selectedFormat.ToString()
            })
        End Function

        ''' <summary>
        ''' 获取 OCR 结构化结果（含 box 坐标，用于图片标注）
        ''' 返回格式: {"ready": bool, "success": bool, "items": [{"text": "...", "confidence": 0.99, "box": [{"x":1,"y":2},...]}]}
        ''' </summary>
        <ComVisible(True)>
        Private Function ElementStructItem(e As OcrElement) As Object
            Return New With {
                .text = e.Text,
                .confidence = e.Confidence,
                .type = e.Type.ToString(),
                .kind = e.Kind,
                .nominal = e.Nominal,
                .lower = e.Lower,
                .upper = e.Upper,
                .datum = e.Datum,
                .vdir = e.VerticalDir,
                .box = OrientedQuad(e)
            }
        End Function

        ''' <summary>
        ''' 按文字阅读方向排好序的四个角，供 UI 画盒子。
        ''' 竖排文字从下往上读时，起点角是左下角——盒子也从文字开始的那个角画起，
        ''' 不再不管横竖排永远左上角到右下角（那样完全看不出文字朝哪边读）。
        ''' 引擎给的原四边形（Quad）优先；被专项合并/改写过的用外框按方向重排。
        ''' </summary>
        Private Shared Function OrientedQuad(e As OcrElement) As List(Of Object)
            If e Is Nothing OrElse e.BoundingBox Is Nothing Then Return New List(Of Object)()
            Dim b = e.BoundingBox
            Dim pts As New List(Of Point)
            If e.Quad IsNot Nothing AndAlso e.Quad.Count >= 4 AndAlso String.IsNullOrEmpty(e.VerticalDir) Then
                pts.AddRange(e.Quad.Take(4))
            Else
                pts.Add(New Point(b.Left, b.Top))
                pts.Add(New Point(b.Right, b.Top))
                pts.Add(New Point(b.Right, b.Bottom))
                pts.Add(New Point(b.Left, b.Bottom))
                If String.Equals(e.VerticalDir, "up", StringComparison.OrdinalIgnoreCase) Then
                    ' 从下往上读：起点是左下角，逆时针转一圈
                    pts = New List(Of Point) From {
                        New Point(b.Left, b.Bottom),
                        New Point(b.Left, b.Top),
                        New Point(b.Right, b.Top),
                        New Point(b.Right, b.Bottom)
                    }
                End If
            End If
            Return pts.Select(Function(p) CType(New With {.x = p.X, .y = p.Y}, Object)).ToList()
        End Function

        Public Function GetOcrStructuredResult() As String
            If Not _ocrCompleted Then
                Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {.ready = False})
            End If
            If _ocrError IsNot Nothing Then
                Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {.ready = True, .success = False, .error = _ocrError})
            End If

            ' 已经有结果，直接返回
            If _ocrStructuredResult IsNot Nothing Then
                Dim baseObj = JObject.Parse(_ocrStructuredResult)
                baseObj("ready") = True
                Return baseObj.ToString()
            End If

            ' 非 PaddleOCR 引擎（如 RapidOCR）：从已识别文档的元素（含坐标）构建结构化结果
            If Not String.Equals(GetCurrentOcrEngine(), "paddleocr", StringComparison.OrdinalIgnoreCase) Then
                Try
                    Dim doc = _ocrDocument
                    If doc Is Nothing AndAlso _ocrTask IsNot Nothing AndAlso _ocrTask.IsCompleted AndAlso Not _ocrTask.IsFaulted Then
                        doc = _ocrTask.Result
                    End If
                    If doc IsNot Nothing AndAlso doc.Elements IsNot Nothing AndAlso doc.Elements.Count > 0 Then
                        Dim items = doc.Elements.Where(Function(e) e.BoundingBox IsNot Nothing).Select(Function(e) ElementStructItem(e)).ToList()
                        _ocrStructuredResult = Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                            .success = True,
                            .items = items
                        })
                        Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                            .ready = True,
                            .success = True,
                            .items = items
                        })
                    End If
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"[OcrBridge] 非 PaddleOCR 结构化结果构建失败: {ex.Message}")
                End Try
                Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                    .ready = True,
                    .success = True,
                    .items = New List(Of Object)()
                })
            End If

            ' 懒加载：首次请求标注时才执行位置识别（避免拖慢普通文本识别）
            If _ocrStructuredTask Is Nothing Then
                Dim imagePath = _lastImagePath
                Dim recognitionMode = _lastRecognitionMode
                Dim generation = _ocrGeneration
                _ocrStructuredTask = Task.Run(Sub()
                    Try
                        If generation <> _ocrGeneration Then Return
                        Dim ocrManager = Services.OcrServiceManager.Instance
                        Dim positioned = ocrManager.RecognizeTextWithPositionAsync(imagePath, IO.Path.GetFileName(imagePath), recognitionMode).GetAwaiter().GetResult()
                        If generation <> _ocrGeneration Then Return
                        If positioned IsNot Nothing AndAlso Not String.IsNullOrEmpty(positioned.RawEngineResponse) Then
                            _ocrRawData = positioned.RawEngineResponse
                        End If
                        If positioned IsNot Nothing AndAlso positioned.Elements IsNot Nothing AndAlso positioned.Elements.Count > 0 Then
                            Dim items = positioned.Elements.Where(Function(e) e.BoundingBox IsNot Nothing).Select(Function(e) ElementStructItem(e)).ToList()
                            _ocrStructuredResult = Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                                .success = True,
                                .items = items
                            })
                            System.Diagnostics.Debug.WriteLine($"[OcrBridge] 懒加载结构化结果完成，共 {items.Count} 个区域")
                        End If
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine($"[OcrBridge] 懒加载结构化结果失败: {ex.Message}")
                    End Try
                End Sub)
                Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {.ready = False})
            End If

            If Not _ocrStructuredTask.IsCompleted Then
                Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {.ready = False})
            End If

            ' 任务完成后仍未产生数据，返回空 items
            If _ocrStructuredResult IsNot Nothing Then
                Dim baseObj = JObject.Parse(_ocrStructuredResult)
                baseObj("ready") = True
                Return baseObj.ToString()
            Else
                Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                    .ready = True,
                    .success = True,
                    .items = New List(Of Object)()
                })
            End If
        End Function

        ''' <summary>
        ''' 获取 OCR 文档元素列表（含坐标，用于前端解析脚本）
        ''' 返回格式: {"ready": bool, "success": bool, "items": [{"text":"...","confidence":0.99,"box":{"left":1,"top":2,"right":3,"bottom":4}}]}
        ''' </summary>
        <ComVisible(True)>
        Public Function GetOcrElementsJson() As String
            If Not _ocrCompleted Then
                Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {.ready = False})
            End If
            If _ocrError IsNot Nothing Then
                Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {.ready = True, .success = False, .error = _ocrError})
            End If

            Dim doc = _ocrDocument
            If doc Is Nothing AndAlso _ocrTask IsNot Nothing AndAlso _ocrTask.IsCompleted AndAlso _ocrTask.Exception Is Nothing Then
                doc = _ocrTask.Result
            End If

            If doc Is Nothing OrElse doc.Elements Is Nothing Then
                Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                    .ready = True,
                    .success = True,
                    .items = New List(Of Object)()
                })
            End If

            Dim items = doc.Elements.Where(Function(e) e.BoundingBox IsNot Nothing).Select(Function(e) New With {
                .text = e.Text,
                .confidence = e.Confidence,
                .type = e.Type.ToString(),
                .kind = e.Kind,
                .nominal = e.Nominal,
                .lower = e.Lower,
                .upper = e.Upper,
                .datum = e.Datum,
                .box = New With {
                    .left = e.BoundingBox.Left,
                    .top = e.BoundingBox.Top,
                    .right = e.BoundingBox.Right,
                    .bottom = e.BoundingBox.Bottom
                }
            }).ToList()

            Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                .ready = True,
                .success = True,
                .count = items.Count,
                .items = items
            })
        End Function

        ''' <summary>
        ''' 获取 PaddleOCR 原始返回数据（含坐标、文字、置信度）
        ''' 返回格式: {"ready": bool, "success": bool, "raw_data": [...]}
        ''' </summary>
        <ComVisible(True)>
        Public Function GetOcrRawData() As String
            If Not _ocrCompleted Then
                Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {.ready = False})
            End If
            If _ocrError IsNot Nothing Then
                Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {.ready = True, .success = False, .error = _ocrError})
            End If
            If _ocrRawData IsNot Nothing Then
                Try
                    Dim rawObj = JObject.Parse(_ocrRawData)
                    If rawObj("raw_data") IsNot Nothing Then
                        Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                            .ready = True,
                            .success = True,
                            .raw_data = rawObj("raw_data")
                        })
                    End If
                Catch
                End Try
            End If
            Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                .ready = True,
                .success = True,
                .raw_data = New List(Of Object)()
            })
        End Function

#End Region

#Region "移动端遥控 API"

        ''' <summary>
        ''' 设置图片到桌面 OCR 面板（供移动端调用）
        ''' </summary>
        <ComVisible(True)>
        Public Sub SetImageForRecognition(filePath As String, base64Image As String)
            Try
                System.Diagnostics.Debug.WriteLine($"[OcrBridge] SetImageForRecognition: {filePath}")
                _pendingImagePath = filePath
                _pendingImageBase64 = base64Image
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[OcrBridge] SetImageForRecognition 异常: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' 执行当前图片的识别（同步，返回结果）
        ''' </summary>
        <ComVisible(True)>
        Public Function RecognizeCurrentImage(recognitionMode As String) As OcrRecognitionResult
            Dim result As New OcrRecognitionResult()
            Try
                System.Diagnostics.Debug.WriteLine($"[OcrBridge] RecognizeCurrentImage 开始，模式: {recognitionMode}")

                If String.IsNullOrEmpty(_pendingImagePath) Then
                    result.success = False
                    result.error = "没有待识别的图片"
                    Return result
                End If

                Dim startResult = StartOcrRecognition(_pendingImagePath, recognitionMode)
                System.Diagnostics.Debug.WriteLine($"[OcrBridge] StartOcrRecognition 返回: {startResult}")

                Dim maxWait As Integer = 120000
                Dim waited As Integer = 0
                Dim interval As Integer = 500

                While waited < maxWait
                    Dim resultJson = GetOcrResult()
                    Dim resultObj = JObject.Parse(resultJson)
                    If resultObj.Value(Of Boolean)("ready") Then
                        If resultObj.Value(Of Boolean)("success") Then
                            result.success = True
                            ' 【修复】移动端需要 Tab 分隔的表格文本，而不是 PlainText 格式默认输出的 raw_ocr_data JSON
                            ' 使用 OcrFormatFormatter.FormatAsMobileTable 获取纯 Tab 分隔表格（无 BOM、无元数据）
                            Dim tableText As String = ""
                            Try
                                Dim doc As OcrDocument = _ocrDocument
                                If doc Is Nothing AndAlso _ocrTask IsNot Nothing AndAlso _ocrTask.IsCompleted AndAlso Not _ocrTask.IsFaulted Then
                                    doc = _ocrTask.Result
                                End If
                                If doc IsNot Nothing Then
                                    tableText = Services.OcrFormatFormatter.FormatAsMobileTable(doc)
                                End If
                            Catch exFmt As Exception
                                System.Diagnostics.Debug.WriteLine($"[OcrBridge] FormatAsMobileTable 失败，兜底用 RawText: {exFmt.Message}")
                            End Try
                            ' 兜底：如果表格文本为空，使用 RawText
                            If String.IsNullOrEmpty(tableText) Then
                                tableText = If(_ocrDocument?.RawText, "")
                            End If
                            result.text = tableText
                            System.Diagnostics.Debug.WriteLine($"[OcrBridge] 识别完成，表格文本长度: {result.text?.Length}")
                        Else
                            result.success = False
                            result.error = resultObj.Value(Of String)("error")
                        End If
                        Exit While
                    End If
                    Thread.Sleep(interval)
                    waited += interval
                End While

                If waited >= maxWait Then
                    result.success = False
                    result.error = "识别超时"
                End If
            Catch ex As Exception
                result.success = False
                result.error = ex.Message
                System.Diagnostics.Debug.WriteLine($"[OcrBridge] RecognizeCurrentImage 异常: {ex.Message}")
            End Try
            Return result
        End Function

        Private _pendingImagePath As String = Nothing
        Private _pendingImageBase64 As String = Nothing

        Public Class OcrRecognitionResult
            Public success As Boolean
            Public text As String
            Public [error] As String
        End Class

#End Region

#Region "OCR API"

        ''' <summary>
        ''' 获取可用的 LLM 提供商列表
        ''' </summary>
        <ComVisible(True)>
        Public Function GetAvailableLLMProviders() As String
            Try
                Dim providers = New List(Of String) From {
                    "OpenAI", "Azure", "Ollama", "Moonshot", "通义千问", "智谱", "文心一言"
                }
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .providers = providers
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 获取可用的 OCR 引擎列表
        ''' </summary>
        <ComVisible(True)>
        Public Function GetAvailableOcrEngines() As String
            Try
                Dim engines = New List(Of String) From {"PaddleOCR", "Azure", "Google"}
                Return JsonConvert.SerializeObject(engines)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 获取 OCR 状态（根据当前选择的引擎类型检测）
        ''' </summary>
        <ComVisible(True)>
        Public Function GetOcrStatus() As String
            Try
                ' 获取当前选择的引擎类型
                Dim engineType = Services.OcrServiceManager.Instance.CurrentEngine
                Dim engineName = Services.OcrServiceManager.GetEngineDisplayName(engineType)
                Dim isAvailable As Boolean

                ' 根据引擎类型使用对应的检测方法（与各引擎专属检测逻辑一致）
                Select Case engineType
                    Case Services.OcrEngineType.PaddleOCR
                        ' 使用与 GetPaddleOcrStatus 完全相同的检测方法（遍历所有Python）
                        Dim paddleResult = CheckPaddleOcrInstallation()
                        isAvailable = CBool(paddleResult("available"))
                    Case Services.OcrEngineType.CloudOCR
                        Dim apiKey As String = If(_config IsNot Nothing, _config.GetSetting("OCR_CLOUD_API_KEY", ""), "")
                        isAvailable = Not String.IsNullOrEmpty(apiKey)
                    Case Else
                        isAvailable = False
                End Select

                Return JsonConvert.SerializeObject(New With {
                    .status = If(isAvailable, "ready", "not_installed"),
                    .engine = engineName,
                    .engineType = engineType.ToString(),
                    .isAvailable = isAvailable,
                    .currentVersion = If(isAvailable, "已安装", Nothing),
                    .needsUpgrade = False,
                    .recommendedVersion = "latest"
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .status = "error",
                    .engine = "未知",
                    .isAvailable = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 获取 OCR 引擎状态
        ''' </summary>
        <ComVisible(True)>
        Public Function GetOcrEngineStatus(engine As String) As String
            Try
                Select Case engine.ToLower()
                    Case "rapidocr"
                        ' 检查 RapidOCR 是否可用（随插件分发，优先使用）
                        Dim rapidOcrStatus = OcrServiceManager.Instance.GetEngineStatus(Services.OcrEngineType.RapidOCR)
                        Return JsonConvert.SerializeObject(New With {
                            .available = rapidOcrStatus.ContainsKey("available") AndAlso rapidOcrStatus("available").ToString().ToLower() = "true",
                            .engine = engine,
                            .message = If(rapidOcrStatus.ContainsKey("name"), rapidOcrStatus("name"), "RapidOCR 引擎"),
                            .embedded = True
                        })
                    Case "paddleocr"
                        ' 检查 PaddleOCR 是否安装
                        Dim paddleStatus = CheckPaddleOcrInstallation()
                        Return JsonConvert.SerializeObject(paddleStatus)
                    Case "cloud"
                        ' 检查云端 OCR 是否已配置 API Key（使用实际 OCR 服务检查，与识别逻辑一致）
                        Dim cloudStatus = OcrServiceManager.Instance.GetEngineStatus(Services.OcrEngineType.CloudOCR)
                        Return JsonConvert.SerializeObject(New With {
                            .available = cloudStatus.ContainsKey("available") AndAlso cloudStatus("available").ToString().ToLower() = "true",
                            .engine = engine,
                            .message = If(cloudStatus.ContainsKey("name"), cloudStatus("name"), "云端 OCR 引擎"),
                            .provider = If(cloudStatus.ContainsKey("provider"), cloudStatus("provider"), "")
                        })
                    Case Else
                        Return JsonConvert.SerializeObject(New With {
                            .available = False,
                            .engine = engine,
                            .message = "未知引擎类型"
                        })
                End Select
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 强制重新检查 RapidOCR 状态（清除缓存）
        ''' </summary>
        <ComVisible(True)>
        Public Function ForceCheckRapidOcrStatus() As String
            Try
                Dim available = OcrServiceManager.Instance.ForceCheckRapidAvailable()
                Return JsonConvert.SerializeObject(New With {
                    .available = available,
                    .engine = "rapidocr",
                    .message = If(available, "RapidOCR 已安装并可用", "RapidOCR 不可用，请检查 Python 环境")
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.available = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 测试云端 OCR 配置（通过真实 API 调用验证凭据有效性）
        ''' </summary>
        <ComVisible(True)>
        Public Function TestCloudOcrConfig(provider As String, configJson As String) As String
            Try
                If String.IsNullOrWhiteSpace(provider) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "未指定 OCR 提供商"})
                End If
                If String.IsNullOrWhiteSpace(configJson) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "未提供配置信息"})
                End If

                System.Diagnostics.Debug.WriteLine($"[TestCloudOcrConfig] 测试提供商: {provider}")
                Dim config = JsonConvert.DeserializeObject(Of Dictionary(Of String, String))(configJson)

                Dim apiKey As String = ""
                Dim secretKey As String = ""

                Select Case provider.ToLower()
                    Case "baidu"
                        If Not config.ContainsKey("apiKey") OrElse String.IsNullOrWhiteSpace(config("apiKey")) Then
                            Return JsonConvert.SerializeObject(New With {.success = False, .error = "百度 API Key 未配置或为空"})
                        End If
                        If Not config.ContainsKey("secretKey") OrElse String.IsNullOrWhiteSpace(config("secretKey")) Then
                            Return JsonConvert.SerializeObject(New With {.success = False, .error = "百度 Secret Key 未配置或为空"})
                        End If
                        apiKey = config("apiKey")
                        secretKey = config("secretKey")
                    Case "tencent"
                        If Not config.ContainsKey("secretId") OrElse String.IsNullOrWhiteSpace(config("secretId")) Then
                            Return JsonConvert.SerializeObject(New With {.success = False, .error = "腾讯云 SecretId 未配置或为空"})
                        End If
                        If Not config.ContainsKey("secretKey") OrElse String.IsNullOrWhiteSpace(config("secretKey")) Then
                            Return JsonConvert.SerializeObject(New With {.success = False, .error = "腾讯云 SecretKey 未配置或为空"})
                        End If
                        apiKey = config("secretId")
                        secretKey = config("secretKey")
                    Case Else
                        Return JsonConvert.SerializeObject(New With {.success = False, .error = "不支持的 OCR 提供商: " & provider})
                End Select

                Try
                    Dim cloudOcrService As New Services.CloudOcrService()
                    Dim isValid = cloudOcrService.ValidateCredentialsAsync(provider, apiKey, secretKey).Result
                    cloudOcrService.Dispose()

                    Return JsonConvert.SerializeObject(New With {
                        .success = isValid,
                        .message = If(isValid, $"与 {provider} 连接成功！", $"{provider} 连接不可用，请检查 API Key 和网络")
                    })
                Catch engineEx As Exception
                    System.Diagnostics.Debug.WriteLine($"[TestCloudOcrConfig] 验证失败: {engineEx.Message}")
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = $"配置验证失败: {engineEx.Message}"
                    })
                End Try
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[TestCloudOcrConfig] 错误: {ex.Message}")
                Return JsonConvert.SerializeObject(New With {.success = False, .error = $"测试失败: {ex.Message}"})
            End Try
        End Function

        ''' <summary>
        ''' 设置 OCR 引擎
        ''' </summary>
        <ComVisible(True)>
        Public Sub SetOcrEngine(engine As String)
            Try
                If _config IsNot Nothing Then
                    _config.SetSetting("OCR_ENGINE", engine)
                End If

                ' 同步到 OcrServiceManager，确保状态查询和实际识别使用同一配置
                Dim engineType As Services.OcrEngineType
                Select Case engine.ToLower()
                    Case "rapidocr"
                        engineType = Services.OcrEngineType.RapidOCR
                    Case "paddleocr"
                        engineType = Services.OcrEngineType.PaddleOCR
                    Case "cloud"
                        engineType = Services.OcrEngineType.CloudOCR
                    Case Else
                        engineType = Services.OcrEngineType.RapidOCR
                End Select
                Services.OcrServiceManager.Instance.CurrentEngine = engineType
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[SetOcrEngine] 错误: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' 保存云端 OCR 配置
        ''' </summary>
        <ComVisible(True)>
        Public Sub SaveCloudOcrConfig(configJson As String)
            Try
                If _config Is Nothing Then Return
                If String.IsNullOrWhiteSpace(configJson) Then Return

                System.Diagnostics.Debug.WriteLine("[SaveCloudOcrConfig] 保存云端 OCR 配置")
                Dim config = JsonConvert.DeserializeObject(Of Dictionary(Of String, String))(configJson)

                ' 同时保存到主配置（用于前端显示）和 CloudOcrService 配置（用于实际使用）
                If config.ContainsKey("provider") Then
                    _config.SetSetting("OCR_CLOUD_PROVIDER", config("provider"))
                End If
                If config.ContainsKey("apiKey") Then
                    _config.SetSetting("OCR_CLOUD_API_KEY", config("apiKey"))
                End If
                If config.ContainsKey("secretKey") Then
                    _config.SetSetting("OCR_CLOUD_SECRET_KEY", config("secretKey"))
                End If
                ' 腾讯云额外字段
                If config.ContainsKey("secretId") Then
                    _config.SetSetting("OCR_TENCENT_SECRET_ID", config("secretId"))
                End If
                _config.FlushSave()

                ' 同时保存到 CloudOcrService 配置文件（用于实际 OCR 调用）
                Dim cloudConfig = Services.CloudOcrConfig.Load()
                If config.ContainsKey("provider") Then cloudConfig.Provider = config("provider")
                If config.ContainsKey("apiKey") Then cloudConfig.BaiduApiKey = config("apiKey")
                If config.ContainsKey("secretKey") Then cloudConfig.BaiduSecretKey = config("secretKey")
                If config.ContainsKey("secretId") Then cloudConfig.TencentSecretId = config("secretId")
                cloudConfig.Save()

                ' 重载 CloudOcrService 的内存配置，使其立即生效
                Services.OcrServiceManager.Instance.ReloadCloudOcrConfig()

                System.Diagnostics.Debug.WriteLine("[SaveCloudOcrConfig] 配置已保存")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[SaveCloudOcrConfig] 保存失败: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' 获取当前 OCR 引擎
        ''' </summary>
        <ComVisible(True)>
        Public Function GetCurrentOcrEngine() As String
            Try
                ' 优先从 OcrServiceManager 读取（与 SetOcrEngine 保持一致）
                Dim engineType = Services.OcrServiceManager.Instance.CurrentEngine
                Select Case engineType
                    Case Services.OcrEngineType.RapidOCR
                        Return "rapidocr"
                    Case Services.OcrEngineType.PaddleOCR
                        Return "paddleocr"
                    Case Services.OcrEngineType.CloudOCR
                        Return "cloud"
                    Case Else
                        Return "rapidocr"
                End Select
            Catch ex As Exception
                ' 降级读取通用配置
                Try
                    If _config IsNot Nothing Then
                        Return _config.GetSetting("OCR_ENGINE", "rapidocr").ToLower()
                    End If
                Catch
                End Try
                Return "rapidocr"
            End Try
        End Function

        ''' <summary>
        ''' 获取云端 OCR 配置（返回完整配置，供前端直接调用 API）
        ''' </summary>
        <ComVisible(True)>
        Public Function GetCloudOcrConfig() As String
            Try
                Dim cloudConfig = Services.CloudOcrConfig.Load()
                
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .provider = cloudConfig.Provider,
                    .apiKey = cloudConfig.BaiduApiKey,
                    .apiSecret = cloudConfig.BaiduSecretKey,
                    .secretId = cloudConfig.TencentSecretId,
                    .secretKey = If(cloudConfig.Provider = "baidu", cloudConfig.BaiduSecretKey, cloudConfig.TencentSecretKey)
                })
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[GetCloudOcrConfig] 错误: {ex.Message}")
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 获取 OCR 配置（返回脱敏后的密钥给前端）
        ''' </summary>
        <ComVisible(True)>
        Public Function GetOcrConfig() As String
            Try
                If _config Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.error = "ConfigurationManager 未初始化"})
                End If

                Dim apiKey As String = _config.GetSetting("OCR_API_KEY", "")
                Dim secretKey As String = _config.GetSetting("OCR_SECRET_KEY", "")

                Dim config = New With {
                    .engine = _config.GetSetting("OCR_ENGINE", "rapidocr"),
                    .apiKey = MaskSensitiveValue(apiKey),
                    .secretKey = MaskSensitiveValue(secretKey),
                    .hasApiKey = Not String.IsNullOrEmpty(apiKey),
                    .hasSecretKey = Not String.IsNullOrEmpty(secretKey)
                }
                Return JsonConvert.SerializeObject(config)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 通过后端获取 Ollama 本地模型列表（绕过前端 CORS 限制）
        ''' 后端 .NET HttpClient 调用 Ollama API 没有 CORS 限制
        ''' </summary>
        <ComVisible(True)>
        Public Async Function GetOllamaModelsAsync() As Task(Of String)
            Try
                Dim models As New List(Of Services.OllamaImageService.LocalModelInfo)()
                Try
                    Using ollamaService As New Services.OllamaImageService()
                        models = Await ollamaService.GetAvailableModels()
                    End Using
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("[OcrBridge] Ollama 模型检测失败: " & ex.Message)
                    ' 返回空列表，前端会尝试直连或使用回退
                End Try

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .models = models.Select(Function(m) New With {.name = m.Name, .source = m.Source}).ToList(),
                    .count = models.Count
                })
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[OcrBridge] GetOllamaModelsAsync 异常: " & ex.Message)
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .models = New List(Of String)(),
                    .count = 0,
                    .error = ex.Message
                })
            End Try
        End Function

#End Region

#Region "私有辅助方法"

        ''' <summary>
        ''' 上次已经确认 PaddleOCR 可用时，不再启动 Python。
        ''' 内存结果给同一次 Excel 会话用；磁盘上的验证记录给下次打开用。
        ''' </summary>
        Private Function CachedPaddleStatus() As Object
            If _paddleReadyStatus IsNot Nothing AndAlso (DateTime.UtcNow - _paddleReadyAt).TotalHours < 12 Then
                Return _paddleReadyStatus
            End If
            Try
                Dim cfg = OcrManagerConfig.Load()
                Dim info = If(cfg Is Nothing, Nothing, cfg.PaddleOcrVerification)
                If info Is Nothing OrElse Not info.IsAvailable Then Return Nothing
                If String.IsNullOrEmpty(info.PythonPath) OrElse Not File.Exists(info.PythonPath) Then Return Nothing
                If Not info.VerifiedAt.HasValue OrElse (DateTime.UtcNow - info.VerifiedAt.Value).TotalDays > 7 Then Return Nothing
                Dim modelDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".paddleocr")
                If Not Directory.Exists(modelDir) OrElse Directory.GetDirectories(modelDir).Length = 0 Then Return Nothing
                Dim status = PaddleReadyStatus(info.PythonPath)
                _paddleReadyStatus = status
                _paddleReadyAt = DateTime.UtcNow
                Return status
            Catch
                Return Nothing
            End Try
        End Function

        Private Shared Function PaddleReadyStatus(pyPath As String) As Object
            Return New With {
                .available = True,
                .engine = "paddleocr",
                .python_found = "True",
                .message = "PaddleOCR 已安装 (" & Path.GetFileName(Path.GetDirectoryName(pyPath)) & ")"
            }
        End Function

        Private Shared Sub RememberPaddleStatus(status As Object, pyPath As String)
            _paddleReadyStatus = status
            _paddleReadyAt = DateTime.UtcNow
            If String.IsNullOrEmpty(pyPath) Then Return
            Try
                Dim cfg = OcrManagerConfig.Load()
                If cfg Is Nothing Then cfg = New OcrManagerConfig()
                If cfg.PaddleOcrVerification Is Nothing Then cfg.PaddleOcrVerification = New OcrEngineVerificationInfo()
                cfg.PaddleOcrVerification.IsAvailable = True
                cfg.PaddleOcrVerification.PythonPath = pyPath
                cfg.PaddleOcrVerification.VerifiedAt = DateTime.UtcNow
                cfg.Save()
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' 检查 PaddleOCR 安装状态（遍历所有 Python，任一成功即可）
        ''' </summary>
        Private Function CheckPaddleOcrInstallation() As Object
            Dim cached = CachedPaddleStatus()
            If cached IsNot Nothing Then Return cached
            Try
                ' 收集所有候选 Python 路径（去重）
                Dim allPaths As New List(Of String)
                Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                ' PathHelper 常见路径
                For Each p In Utils.PathHelper.GetPythonSearchPaths()
                    If File.Exists(p) AndAlso seen.Add(p) Then allPaths.Add(p)
                Next

                ' 注册表
                Dim regPath = PythonLocator.FindFromRegistry()
                If Not String.IsNullOrEmpty(regPath) AndAlso seen.Add(regPath) Then allPaths.Add(regPath)

                ' where python.exe
                For Each line In PythonLocator.FindAllOnPath()
                    If File.Exists(line) AndAlso seen.Add(line) Then allPaths.Add(line)
                Next

                ' py -3.x 命令
                For Each ver In {"3.12", "3.11", "3.10", "3.9", "3.14"}
                    Try
                        Using proc As New Process()
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
                    Catch
                    End Try
                Next

                ' 任一 Python 能 import paddleocr 即视为已安装
                Dim pythonFound = allPaths.Count > 0
                If Not pythonFound Then
                    Return New With {
                        .available = False,
                        .engine = "paddleocr",
                        .python_found = "False",
                        .message = "未检测到 Python"
                    }
                End If

                For Each pyPath In allPaths
                    If IsPaddleOcrInstalled(pyPath) Then
                        ' 检测模型文件是否已下载
                        Dim modelDir = Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                            ".paddleocr")
                        Dim modelsReady = Directory.Exists(modelDir) AndAlso
                            Directory.GetDirectories(modelDir).Length > 0

                        If modelsReady Then
                            Dim status = PaddleReadyStatus(pyPath)
                            RememberPaddleStatus(status, pyPath)
                            Return status
                        Else
                            Return New With {
                                .available = False,
                                .engine = "paddleocr",
                                .python_found = "True",
                                .message = "PaddleOCR 已安装，但模型未下载（请重新安装以触发模型下载）"
                            }
                        End If
                    End If
                Next

                Return New With {
                    .available = False,
                    .engine = "paddleocr",
                    .python_found = "True",
                    .message = "PaddleOCR 未安装（已找到 " & allPaths.Count & " 个 Python，均不可用）"
                }

            Catch ex As Exception
                Return New With {
                    .available = False,
                    .engine = "paddleocr",
                    .python_found = "False",
                    .message = "检测失败: " & ex.Message
                }
            End Try
        End Function

        ''' <summary>
        ''' 检查 PaddleOCR 是否已安装
        ''' </summary>
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
                Dim output = process.StandardOutput.ReadToEnd()
                process.WaitForExit(5000)

                Return process.ExitCode = 0 AndAlso output.Contains("OK")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("PaddleOCR安装检查失败: " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 敏感值脱敏 — 只显示前3后4字符，中间用***替代
        ''' 用于返回给前端的 API Key/Secret Key 等敏感信息
        ''' </summary>
        ''' <param name="value">原始值</param>
        ''' <returns>脱敏后的值，如 "sk-***f456"；空值返回空字符串</returns>
        Private Shared Function MaskSensitiveValue(value As String) As String
            If String.IsNullOrEmpty(value) Then
                Return String.Empty
            End If

            If value.Length <= 8 Then
                ' 太短的值：只显示首字符
                Return value.Substring(0, 1) & "***"
            End If

            ' 显示前3字符 + *** + 后4字符
            Return value.Substring(0, 3) & "***" & value.Substring(value.Length - 4)
        End Function

#End Region

#Region "结构化识别 & 批量处理 API"

        ''' <summary>
        ''' 执行结构化 OCR 识别，返回指定格式的结构化文本
        ''' </summary>
        <ComVisible(True)>
        Public Function RecognizeTextWithFormat(imagePath As String, format As String, Optional fileName As String = "") As String
            Try
                Dim outputFormat As OcrOutputFormat
                Select Case format.ToLower()
                    Case "markdown" : outputFormat = OcrOutputFormat.Markdown
                    Case "html" : outputFormat = OcrOutputFormat.Html
                    Case "excel", "csv" : outputFormat = OcrOutputFormat.Excel
                    Case Else : outputFormat = OcrOutputFormat.PlainText
                End Select

                If String.IsNullOrEmpty(fileName) Then
                    fileName = Path.GetFileName(imagePath)
                End If

                Dim result = Services.OcrServiceManager.Instance.RecognizeTextWithFormatAsync(
                    imagePath, outputFormat, fileName).GetAwaiter().GetResult()

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .text = result,
                    .format = format
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 扫描文件夹中的图片文件
        ''' </summary>
        <ComVisible(True)>
        Public Function ScanOcrFolder(folderPath As String, Optional recursive As Boolean = False) As String
            Try
                Dim processor As New Services.OcrBatchProcessor()
                Dim files = processor.ScanFolder(folderPath, recursive)

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .folder = folderPath,
                    .files = files,
                    .count = files.Count
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 启动批量 OCR 处理
        ''' </summary>
        <ComVisible(True)>
        Public Function StartBatchOcr(filePathsJson As String, Optional outputFormat As String = "plaintext") As String
            Try
                Dim filePaths = JsonConvert.DeserializeObject(Of List(Of String))(filePathsJson)
                If filePaths Is Nothing OrElse filePaths.Count = 0 Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = "文件列表为空"
                    })
                End If

                Dim outFmt As OcrOutputFormat
                Select Case outputFormat.ToLower()
                    Case "markdown" : outFmt = OcrOutputFormat.Markdown
                    Case "html" : outFmt = OcrOutputFormat.Html
                    Case "excel", "csv" : outFmt = OcrOutputFormat.Excel
                    Case Else : outFmt = OcrOutputFormat.PlainText
                End Select

                Dim processor As Services.OcrBatchProcessor = Nothing
                SyncLock _batchLock
                    If _batchProcessor IsNot Nothing AndAlso _batchProcessor.IsProcessing Then
                        Return JsonConvert.SerializeObject(New With {
                            .success = False,
                            .error = "已有批量任务在进行中，请先取消或等待完成"
                        })
                    End If
                    processor = New Services.OcrBatchProcessor()
                    _batchProcessor = processor
                End SyncLock

                processor.AddFiles(filePaths)

                ' 启动异步处理但不等待（前端轮询进度）
                Dim task = processor.StartAsync(Nothing)

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .total = filePaths.Count,
                    .message = "批量处理已启动"
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 获取批量处理进度（真实进度，由前端轮询调用）
        ''' </summary>
        <ComVisible(True)>
        Public Function GetBatchProgress() As String
            Try
                Dim processor As Services.OcrBatchProcessor = Nothing
                SyncLock _batchLock
                    processor = _batchProcessor
                End SyncLock

                If processor Is Nothing OrElse processor.TotalTasks = 0 Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = True,
                        .total = 0,
                        .completed = 0,
                        .failed = 0,
                        .percent = 0,
                        .currentFile = "",
                        .status = "idle"
                    })
                End If

                Dim p = processor.GetProgress()
                Dim status As String
                If processor.IsProcessing Then
                    status = If(processor.IsPaused, "paused", "processing")
                ElseIf p.CompletedTasks + p.FailedTasks >= p.TotalTasks Then
                    status = "completed"
                Else
                    status = "idle"
                End If

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .total = p.TotalTasks,
                    .completed = p.CompletedTasks,
                    .failed = p.FailedTasks,
                    .percent = p.Percent,
                    .currentFile = p.CurrentFileName,
                    .status = status
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 取消当前批量处理
        ''' </summary>
        <ComVisible(True)>
        Public Function CancelBatchOcr() As String
            Try
                Dim processor As Services.OcrBatchProcessor = Nothing
                SyncLock _batchLock
                    processor = _batchProcessor
                End SyncLock

                If processor Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = "没有批量任务实例"
                    })
                End If

                processor.Cancel()
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .message = "已请求取消批量处理"
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 获取 PaddleOCR 高级配置（语言、版面分析/表格/公式开关）
        ''' </summary>
        <ComVisible(True)>
        Public Function GetPaddleOcrAdvancedConfig() As String
            Try
                Dim mgr = Services.OcrServiceManager.Instance
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .language = mgr.PaddleOcrLanguage,
                    .enableLayoutAnalysis = mgr.EnableLayoutAnalysis,
                    .enableTableRecognition = mgr.EnableTableRecognition,
                    .enableFormulaRecognition = mgr.EnableFormulaRecognition,
                    .supportedLanguages = New List(Of Object) From {
                        New With {Key .id = "ch", Key .name = "中文"},
                        New With {Key .id = "en", Key .name = "英文"},
                        New With {Key .id = "japan", Key .name = "日语"},
                        New With {Key .id = "korean", Key .name = "韩语"},
                        New With {Key .id = "fr", Key .name = "法语"},
                        New With {Key .id = "ger", Key .name = "德语"},
                        New With {Key .id = "es", Key .name = "西班牙语"},
                        New With {Key .id = "ru", Key .name = "俄语"},
                        New With {Key .id = "arabic", Key .name = "阿拉伯语"},
                        New With {Key .id = "structure", Key .name = "自动检测"}
                    }
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 设置 PaddleOCR 高级配置
        ''' </summary>
        <ComVisible(True)>
        Public Sub SetPaddleOcrAdvancedConfig(configJson As String)
            Try
                If String.IsNullOrWhiteSpace(configJson) Then Return
                Dim config = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(configJson)
                Dim mgr = Services.OcrServiceManager.Instance

                If config.ContainsKey("language") AndAlso config("language") IsNot Nothing Then
                    Dim langObj = config("language")
                    mgr.PaddleOcrLanguage = If(langObj IsNot Nothing, langObj.ToString(), "ch")
                End If
                If config.ContainsKey("enableLayoutAnalysis") Then
                    mgr.EnableLayoutAnalysis = CBool(config("enableLayoutAnalysis"))
                End If
                If config.ContainsKey("enableTableRecognition") Then
                    mgr.EnableTableRecognition = CBool(config("enableTableRecognition"))
                End If
                If config.ContainsKey("enableFormulaRecognition") Then
                    mgr.EnableFormulaRecognition = CBool(config("enableFormulaRecognition"))
                End If

                System.Diagnostics.Debug.WriteLine("PaddleOCR 高级配置已保存")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("保存 PaddleOCR 高级配置失败: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' 检查公式识别增强包（pix2tex）是否已安装
        ''' </summary>
        <ComVisible(True)>
        Public Function CheckFormulaOcrStatus() As String
            If Not String.IsNullOrEmpty(_formulaReadyJson) Then Return _formulaReadyJson
            Try
                Dim result = Services.OcrServiceManager.Instance.CheckFormulaOcrStatusAsync().GetAwaiter().GetResult()
                Dim available = result.Item1
                Dim rawStatus = result.Item2
                Dim json = JsonConvert.SerializeObject(New With {
                    .success = True,
                    .available = available,
                    .message = If(available, "公式识别增强包已就绪", "公式识别增强包未安装"),
                    .raw_status = rawStatus
                })
                If available Then _formulaReadyJson = json
                Return json
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[OcrBridge] 检查公式状态失败: {ex.Message}")
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .available = False,
                    .message = $"检测失败: {ex.Message}",
                    .raw_status = ""
                })
            End Try
        End Function

        ''' <summary>
        ''' 启动公式识别增强包后台安装（非阻塞，立即返回）
        ''' </summary>
        <ComVisible(True)>
        Public Function StartFormulaOcrInstall() As String
            Try
                Return Services.OcrServiceManager.Instance.StartFormulaOcrInstall()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[OcrBridge] 启动公式安装失败: {ex.Message}")
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .message = $"启动安装失败: {ex.Message}"
                })
            End Try
        End Function

        ''' <summary>
        ''' 获取公式增强包安装状态（轮询用）
        ''' </summary>
        <ComVisible(True)>
        Public Function GetFormulaInstallStatus() As String
            Try
                Return Services.OcrServiceManager.Instance.GetFormulaInstallStatus()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[OcrBridge] 获取安装状态失败: {ex.Message}")
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .status = "error",
                    .message = $"获取状态失败: {ex.Message}",
                    .result = ""
                })
            End Try
        End Function

        ''' <summary>
        ''' 安装公式识别增强包（pix2tex / LaTeX-OCR）——同步阻塞版本，仅供内部使用
        ''' </summary>
        <ComVisible(False)>
        Public Function InstallFormulaOcr() As String
            Try
                Dim result = Services.OcrServiceManager.Instance.InstallFormulaOcrAsync().GetAwaiter().GetResult()
                Return result
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[OcrBridge] 安装公式增强包失败: {ex.Message}")
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .message = $"安装失败: {ex.Message}"
                })
            End Try
        End Function

        ''' <summary>
        ''' 获取支持的输出格式列表
        ''' </summary>
        <ComVisible(True)>
        Public Function GetSupportedOutputFormats() As String
            Try
                Dim formats = New List(Of Object) From {
                    New With {Key .id = "plaintext", Key .name = "纯文本", Key .desc = "纯文本格式"},
                    New With {Key .id = "markdown", Key .name = "Markdown", Key .desc = "Markdown 格式（含标题、列表、表格结构）"},
                    New With {Key .id = "html", Key .name = "HTML", Key .desc = "HTML 格式（保留排版布局）"},
                    New With {Key .id = "excel", Key .name = "Excel 表格", Key .desc = "制表符分隔格式，可直接粘贴到 Excel"}
                }
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .formats = formats
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

#End Region

#Region "OCR解析脚本 API"

        ' 共享 camelCase 序列化设置（前端 JavaScript 期望 camelCase 属性名）
        Private Shared ReadOnly _camelCaseSettings As New JsonSerializerSettings With {
            .ContractResolver = New CamelCasePropertyNamesContractResolver(),
            .NullValueHandling = NullValueHandling.Ignore
        }

        ''' <summary>
        ''' 清空解析脚本缓存并重新加载
        ''' </summary>
        <ComVisible(True)>
        Public Function ClearParseScriptCache() As String
            Try
                Dim success = Services.ParseScriptManager.Instance.ClearCache()
                Return JsonConvert.SerializeObject(New With {.success = success})
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 获取专项识别模块列表（内置 + 已安装）。
        ''' 前端据此动态渲染勾选项，不再硬编码 chk-mechanical-drawing。
        ''' </summary>
        <ComVisible(True)>
        Public Function GetDomainPacks() As String
            Try
                Dim packs = Services.DomainPack.DomainPackRegistry.Instance.All()
                Dim items As New List(Of Object)()
                For Each p In packs
                    items.Add(New With {
                        .id = p.Id,
                        .name = If(String.IsNullOrEmpty(p.Name), p.Id, p.Name),
                        .description = p.Description,
                        .icon = p.Icon,
                        .category = If(String.IsNullOrEmpty(p.Category), "专项", p.Category),
                        .version = p.Version,
                        .pipeline = p.PipelineName(),
                        .moduleRef = p.ModuleRef,
                        .requires = p.Requires,
                        .isBuiltIn = p.IsBuiltIn,
                        .enabled = p.Enabled,
                        .error = p.ErrorText,
                        .options = p.Options.Select(Function(o) New With {
                            .key = o.Key,
                            .label = If(String.IsNullOrEmpty(o.Label), o.Key, o.Label),
                            .type = If(String.IsNullOrEmpty(o.Type), "bool", o.Type),
                            .default = o.DefaultValue,
                            .values = o.Values
                        }).ToList()
                    })
                Next
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .packs = items,
                    .count = items.Count,
                    .rootDir = Services.DomainPack.DomainPackRegistry.Instance.RootDirectory
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 安装专项识别模块：传入 .md 清单内容，落盘到用户目录后重新加载。
        ''' </summary>
        <ComVisible(True)>
        Public Function InstallDomainPackFromMarkdown(markdown As String) As String
            Try
                Dim res = Services.DomainPack.DomainPackRegistry.Instance.InstallFromMarkdown(markdown)
                If res.Success Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = True,
                        .id = res.Id,
                        .name = res.Name,
                        .path = res.Path,
                        .scriptPath = res.ScriptPath
                    }, _camelCaseSettings)
                End If
                Return JsonConvert.SerializeObject(New With {.success = False, .error = res.Error})
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 启用 / 停用专项识别模块（停用只是不再参与识别，文件保留）。
        ''' </summary>
        <ComVisible(True)>
        Public Function SetDomainPackEnabled(id As String, enabled As Boolean) As String
            Try
                If String.IsNullOrWhiteSpace(id) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "缺少模块 id"})
                End If
                Services.DomainPack.DomainPackRegistry.Instance.SetEnabled(id, enabled)
                Dim p = Services.DomainPack.DomainPackRegistry.Instance.[Get](id)
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .id = id,
                    .enabled = (p IsNot Nothing AndAlso p.Enabled)
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>卸载专项模块（内置模块不可卸载）。</summary>
        <ComVisible(True)>
        Public Function UninstallDomainPack(id As String) As String
            Try
                If String.IsNullOrWhiteSpace(id) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "缺少模块 id"})
                End If
                Dim ok = Services.DomainPack.DomainPackRegistry.Instance.Uninstall(id)
                If Not ok Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "无法卸载：模块不存在或是内置模块"})
                End If
                Return JsonConvert.SerializeObject(New With {.success = True, .id = id}, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>设置专项模块自带开关的取值。</summary>
        <ComVisible(True)>
        Public Function SetDomainPackOption(id As String, key As String, value As String) As String
            Try
                Services.DomainPack.DomainPackRegistry.Instance.SetOptionValue(id, key, value)
                Return JsonConvert.SerializeObject(New With {.success = True, .value = value}, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>保存机械制图矫正配置到用户目录。下次识别时覆盖内置默认。</summary>
        <ComVisible(True)>
        Public Function SaveMechanicalCorrection(json As String) As String
            Try
                Dim err = Services.DomainPack.CorrectionProfile.SaveMechanical(json)
                If Not String.IsNullOrEmpty(err) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = err})
                End If
                Return JsonConvert.SerializeObject(New With {.success = True}, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 获取所有解析脚本列表
        ''' </summary>
        <ComVisible(True)>
        Public Function GetParseScripts() As String
            Try
                Dim mgr = Services.ParseScriptManager.Instance
                Dim scripts = mgr.GetAllScripts()
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .scripts = scripts,
                    .count = scripts.Count
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 获取单个解析脚本详情
        ''' </summary>
        <ComVisible(True)>
        Public Function GetParseScript(scriptId As String) As String
            Try
                Dim mgr = Services.ParseScriptManager.Instance
                Dim script = mgr.GetScriptById(scriptId)
                If script Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "找不到指定的解析脚本"})
                End If
                Return JsonConvert.SerializeObject(New With {.success = True, .script = script}, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 保存解析脚本（新增或更新）
        ''' </summary>
        <ComVisible(True)>
        Public Function SaveParseScript(scriptJson As String) As String
            Try
                Dim mgr = Services.ParseScriptManager.Instance
                Dim success = mgr.SaveScriptFromJson(scriptJson)
                Return JsonConvert.SerializeObject(New With {.success = success, .message = If(success, "保存成功", "保存失败")})
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 删除解析脚本（内置模板不可删除）
        ''' </summary>
        <ComVisible(True)>
        Public Function DeleteParseScript(scriptId As String) As String
            Try
                Dim mgr = Services.ParseScriptManager.Instance
                Dim success = mgr.DeleteScript(scriptId)
                If Not success Then
                    ' 检查是否是内置模板
                    Dim script = mgr.GetScriptById(scriptId)
                    If script IsNot Nothing AndAlso script.IsBuiltIn Then
                        Return JsonConvert.SerializeObject(New With {.success = False, .error = "内置模板不可删除"})
                    End If
                End If
                Return JsonConvert.SerializeObject(New With {.success = success, .message = If(success, "删除成功", "删除失败")})
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 还原内置脚本到默认配置（仅对 isBuiltIn=true 的脚本有效）
        ''' </summary>
        <ComVisible(True)>
        Public Function ResetBuiltinScript(scriptId As String) As String
            Try
                Dim mgr = Services.ParseScriptManager.Instance
                Dim resetScript = mgr.ResetBuiltinScript(scriptId)
                If resetScript IsNot Nothing Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = True,
                        .message = "已还原为默认配置",
                        .script = resetScript
                    }, _camelCaseSettings)
                Else
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = "还原失败：脚本不存在或非内置脚本"
                    })
                End If
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 复制解析脚本
        ''' </summary>
        <ComVisible(True)>
        Public Function DuplicateParseScript(sourceId As String, newName As String) As String
            Try
                Dim mgr = Services.ParseScriptManager.Instance
                Dim newScript = mgr.DuplicateScript(sourceId, newName)
                If newScript Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "复制失败，源脚本不存在"})
                End If
                Return JsonConvert.SerializeObject(New With {.success = True, .script = newScript, .message = "复制成功"}, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 执行解析脚本（返回脚本信息和上下文，前端执行JS）
        ''' </summary>
        <ComVisible(True)>
        Public Function ExecuteParseScript(scriptId As String, ocrText As String) As String
            Try
                Dim mgr = Services.ParseScriptManager.Instance
                Return mgr.ExecuteScript(scriptId, ocrText)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 获取所有解析脚本的分类列表
        ''' </summary>
        <ComVisible(True)>
        Public Function GetParseScriptCategories() As String
            Try
                Dim mgr = Services.ParseScriptManager.Instance
                Dim categories = mgr.GetCategories()
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .categories = categories
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 搜索解析脚本
        ''' </summary>
        <ComVisible(True)>
        Public Function SearchParseScripts(keyword As String) As String
            Try
                Dim mgr = Services.ParseScriptManager.Instance
                Dim scripts = mgr.SearchScripts(keyword)
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .scripts = scripts,
                    .count = scripts.Count
                }, _camelCaseSettings)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 从 Markdown 文件内容导入解析脚本
        ''' </summary>
        <ComVisible(True)>
        Public Function ImportParseScriptFromMarkdown(markdownContent As String) As String
            Try
                If String.IsNullOrWhiteSpace(markdownContent) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "文件内容为空"})
                End If

                Dim mgr = Services.ParseScriptManager.Instance
                Dim script = mgr.ImportScriptFromMarkdown(markdownContent)
                
                If script IsNot Nothing Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = True,
                        .script = script,
                        .message = "导入成功"
                    }, _camelCaseSettings)
                Else
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "解析失败：文件格式不正确"})
                End If
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

#End Region

#Region "IBridge 生命周期"

        Public ReadOnly Property BridgeName As String Implements IBridge.BridgeName
            Get
                Return "OcrBridge"
            End Get
        End Property

        Private _disposed As Boolean = False
        Public Sub Dispose() Implements IBridge.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not _disposed Then
                _disposed = True
                If disposing Then
                    ' 【关键修复】释放 OcrServiceManager 以终止 OCR 引擎 Python 进程
                    ' 否则 PaddleOCR/RapidOCR 的 Python 子进程会残留，占用 DLL/端口锁，
                    ' 导致下次插件启动失败
                    Try
                        Services.OcrServiceManager.Instance.Dispose()
                        System.Diagnostics.Debug.WriteLine("[OcrBridge] OcrServiceManager 已释放")
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine($"[OcrBridge] 释放 OcrServiceManager 失败: {ex.Message}")
                    End Try
                End If
            End If
        End Sub

#End Region

    End Class

End Namespace