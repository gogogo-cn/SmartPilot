Option Strict On
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Reflection
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ExcelAIVSTOAddIn.Utils

Namespace Services.Mcp

    Partial Public Class ExcelMcpTools

        '#Region "工具实现 - 文件与工作表"

        Private Function ExecuteOcrRecognize(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                ' 获取图片路径或 Base64 数据
                Dim imagePath As String = GetArg(Of String)(args, "imagePath", "")
                Dim imageData As String = GetArg(Of String)(args, "imageData", "")

                If String.IsNullOrEmpty(imagePath) AndAlso String.IsNullOrEmpty(imageData) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "必须提供 imagePath 或 imageData 参数"})
                    Return result
                End If

                ' 如果传入的是 Base64 数据，先保存为临时文件
                Dim actualImagePath As String = imagePath
                If Not String.IsNullOrEmpty(imageData) Then
                    ' 移除 data:image/xxx;base64, 前缀
                    Dim base64Data As String = imageData
                    If base64Data.Contains(",") Then
                        base64Data = base64Data.Substring(base64Data.IndexOf(",") + 1)
                    End If

                    actualImagePath = Path.Combine(Path.GetTempPath(), "mcp_ocr_" & Guid.NewGuid().ToString("N") & ".png")
                    Try
                        Dim imageBytes = Convert.FromBase64String(base64Data)
                        File.WriteAllBytes(actualImagePath, imageBytes)
                    Catch ex As Exception
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = $"Base64 图片数据解析失败: {ex.Message}"})
                        Return result
                    End Try
                End If

                ' 验证文件存在
                If Not File.Exists(actualImagePath) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"图片文件不存在: {actualImagePath}"})
                    Return result
                End If

                ' 预检 OCR 引擎可用性，提前给出友好提示
                Dim ocrManager = Services.OcrServiceManager.Instance
                Dim availableEngines = ocrManager.GetAvailableEngines()
                Dim hasAnyEngine As Boolean = False
                Dim engineHint As String = ""

                For Each kvp In availableEngines
                    If ocrManager.IsEngineAvailable(kvp.Key) Then
                        hasAnyEngine = True
                        Exit For
                    End If
                Next

                If Not hasAnyEngine Then
                    ' 清理临时文件
                    If Not String.IsNullOrEmpty(imageData) AndAlso File.Exists(actualImagePath) Then
                        Try : File.Delete(actualImagePath) : Catch : End Try
                    End If

                    result.IsError = True

                    ' 构建友好的引擎缺失提示
                    Dim hintBuilder As New System.Text.StringBuilder()
                    hintBuilder.AppendLine("当前系统未检测到可用的 OCR 引擎，无法进行图片文字识别。")
                    hintBuilder.AppendLine()
                    hintBuilder.AppendLine("建议安装以下任一引擎：")
                    hintBuilder.AppendLine("  1. PaddleOCR: 需要 Python 环境，运行 pip install paddleocr")
                    hintBuilder.AppendLine()
                    hintBuilder.AppendLine("安装后可在 SmartPilot 智航 → 设置 → OCR 配置 中选择默认引擎。")

                    result.Content.Add(New McpTextContent With {.Text = hintBuilder.ToString()})
                    Return result
                End If

                ' 获取可选的引擎参数
                Dim engineStr As String = GetArg(Of String)(args, "engine", "").ToLower()

                ' 获取语言参数
                Dim languageStr As String = GetArg(Of String)(args, "language", "").ToLower()
                If Not String.IsNullOrEmpty(languageStr) AndAlso (languageStr = "ch" OrElse languageStr = "en") Then
                    ocrManager.PaddleOcrLanguage = languageStr
                End If

                ' 执行 OCR 识别（同步等待异步结果）
                Dim ocrText As String = ""
                Dim ocrEngineName As String = ""

                Try
                    ' 如果指定了引擎，使用指定引擎；否则使用默认引擎
                    If Not String.IsNullOrEmpty(engineStr) Then
                        Dim targetEngine As Services.OcrEngineType
                        Select Case engineStr
                            Case "paddleocr"
                                targetEngine = Services.OcrEngineType.PaddleOCR
                            Case "rapidocr"
                                targetEngine = Services.OcrEngineType.RapidOCR
                            Case "cloudocr"
                                targetEngine = Services.OcrEngineType.CloudOCR
                            Case Else
                                targetEngine = ocrManager.CurrentEngine
                        End Select

                        ' 检查指定的引擎是否可用
                        If Not ocrManager.IsEngineAvailable(targetEngine) Then
                            ' 清理临时文件
                            If Not String.IsNullOrEmpty(imageData) AndAlso File.Exists(actualImagePath) Then
                                Try : File.Delete(actualImagePath) : Catch : End Try
                            End If
                            result.IsError = True
                            result.Content.Add(New McpTextContent With {
                                .Text = $"指定的 OCR 引擎 ({engineStr}) 当前不可用。" &
                                           $"请改用其他引擎或前往设置页面安装。已可用引擎: " &
                                           String.Join(", ", availableEngines.Where(Function(kvp) ocrManager.IsEngineAvailable(kvp.Key)).Select(Function(kvp) kvp.Value))
                            })
                            Return result
                        End If

                        ocrEngineName = Services.OcrServiceManager.GetEngineDisplayName(targetEngine)
                        ocrText = ocrManager.RecognizeTextWithEngine(actualImagePath, targetEngine).GetAwaiter().GetResult()
                    Else
                        ocrEngineName = Services.OcrServiceManager.GetEngineDisplayName(ocrManager.CurrentEngine)
                        ocrText = ocrManager.RecognizeText(actualImagePath).GetAwaiter().GetResult()
                    End If
                Catch ex As Exception
                    ' 清理临时文件
                    If Not String.IsNullOrEmpty(imageData) AndAlso File.Exists(actualImagePath) Then
                        Try : File.Delete(actualImagePath) : Catch : End Try
                    End If
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"OCR 识别失败: {ex.Message}"})
                    Return result
                End Try

                ' 清理临时文件
                If Not String.IsNullOrEmpty(imageData) AndAlso File.Exists(actualImagePath) Then
                    Try : File.Delete(actualImagePath) : Catch : End Try
                End If

                ' 检测 OCR 结果是否为错误信息（以【错误开头说明是原始 OCR 错误文本）
                If ocrText.StartsWith("【错误") OrElse ocrText.StartsWith("[错误") Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"OCR 识别失败: {ocrText}"})
                    Return result
                End If

                ' 构建结果 JSON
                Dim outputJson = JsonConvert.SerializeObject(New With {
                    .text = ocrText,
                    .engine = ocrEngineName,
                    .charCount = ocrText.Length,
                    .imageSource = If(String.IsNullOrEmpty(imageData), imagePath, "base64_data")
                }, Formatting.Indented)

                result.Content.Add(New McpTextContent With {
                    .Text = outputJson
                })

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"OCR 工具执行异常: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteReadFile(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim filePath As String = GetArg(Of String)(args, "filePath", "")
                Dim maxChars As Integer = GetArg(Of Integer)(args, "maxChars", 50000)

                If String.IsNullOrEmpty(filePath) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "文件路径不能为空"})
                    Return result
                End If

                If Not File.Exists(filePath) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"文件不存在: {filePath}"})
                    Return result
                End If

                Dim ext = Path.GetExtension(filePath).ToLower().TrimStart("."c)
                Dim fullText As String = ""

                Select Case ext
                    Case "txt", "md", "json", "xml", "csv"
                        fullText = File.ReadAllText(filePath, Encoding.UTF8)
                    Case "docx"
                        Dim base64Data = Convert.ToBase64String(File.ReadAllBytes(filePath))
                        Dim parser = New Services.Parsers.DocxDocumentParser()
                        Dim parseResult = parser.ExtractFullText(base64Data).GetAwaiter().GetResult()
                        If parseResult.Success Then
                            fullText = parseResult.FullText
                        Else
                            result.IsError = True
                            result.Content.Add(New McpTextContent With {.Text = $"DOCX 解析失败: {parseResult.ErrorMessage}"})
                            Return result
                        End If
                    Case "pdf"
                        Dim base64Data = Convert.ToBase64String(File.ReadAllBytes(filePath))
                        Dim parser = New Services.Parsers.PdfDocumentParser()
                        Dim parseResult = parser.ExtractFullText(base64Data).GetAwaiter().GetResult()
                        If parseResult.Success Then
                            fullText = parseResult.FullText
                        Else
                            result.IsError = True
                            result.Content.Add(New McpTextContent With {.Text = $"PDF 解析失败: {parseResult.ErrorMessage}"})
                            Return result
                        End If
                    Case "xlsx", "xls"
                        Dim base64Data = Convert.ToBase64String(File.ReadAllBytes(filePath))
                        Dim parser = New Services.Parsers.ExcelFileParser()
                        Dim parseResult = parser.ExtractFullText(base64Data).GetAwaiter().GetResult()
                        If parseResult.Success Then
                            fullText = parseResult.FullText
                        Else
                            result.IsError = True
                            result.Content.Add(New McpTextContent With {.Text = $"Excel 解析失败: {parseResult.ErrorMessage}"})
                            Return result
                        End If
                    Case Else
                        Try
                            fullText = File.ReadAllText(filePath, Encoding.UTF8)
                        Catch
                            fullText = $"不支持的文件格式: .{ext}。文件大小: {New FileInfo(filePath).Length} 字节"
                        End Try
                End Select

                If fullText.Length > maxChars Then
                    fullText = fullText.Substring(0, maxChars) & $"...(已截断，剩余 {fullText.Length - maxChars} 字符)"
                End If

                Dim outputJson = JsonConvert.SerializeObject(New With {
                    .filePath = filePath,
                    .fileSize = New FileInfo(filePath).Length,
                    .charCount = fullText.Length,
                    .content = fullText
                }, Formatting.Indented)

                result.Content.Add(New McpTextContent With {.Text = outputJson})

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"读取文件失败: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteListFiles(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim directoryPath As String = GetArg(Of String)(args, "directoryPath", "")

                If String.IsNullOrEmpty(directoryPath) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "目录路径不能为空"})
                    Return result
                End If

                If Not Directory.Exists(directoryPath) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"目录不存在: {directoryPath}"})
                    Return result
                End If

                Dim files = New List(Of Dictionary(Of String, Object))()
                For Each filePath In Directory.GetFiles(directoryPath)
                    Dim fi = New FileInfo(filePath)
                    Dim fileDict = New Dictionary(Of String, Object)()
                    fileDict.Add("name", fi.Name)
                    fileDict.Add("fullPath", filePath)
                    fileDict.Add("size", fi.Length)
                    fileDict.Add("extension", fi.Extension)
                    fileDict.Add("lastModified", fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"))
                    files.Add(fileDict)
                Next

                For Each dirPath In Directory.GetDirectories(directoryPath)
                    Dim di = New DirectoryInfo(dirPath)
                    Dim dirDict = New Dictionary(Of String, Object)()
                    dirDict.Add("name", di.Name)
                    dirDict.Add("fullPath", dirPath)
                    dirDict.Add("type", "directory")
                    dirDict.Add("lastModified", di.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"))
                    files.Add(dirDict)
                Next

                Dim outputJson = JsonConvert.SerializeObject(New With {
                    .directory = directoryPath,
                    .count = files.Count,
                    .items = files
                }, Formatting.Indented)

                result.Content.Add(New McpTextContent With {.Text = outputJson})

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"列出文件失败: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteBrowseFiles(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim filter As String = GetArg(Of String)(args, "filter", "all")

                ' 映射过滤器到 OpenFileDialog 格式
                Dim dlgFilter As String = "所有文件|*.*"
                Dim dlgTitle As String = "选择文件"
                Select Case filter.ToLower()
                    Case "excel"
                        dlgFilter = "Excel 文件|*.xlsx;*.xls;*.xlsm|所有文件|*.*"
                        dlgTitle = "选择 Excel 文件"
                    Case "doc"
                        dlgFilter = "文档文件|*.docx;*.doc;*.pdf|所有文件|*.*"
                        dlgTitle = "选择文档文件"
                    Case "image"
                        dlgFilter = "图片文件|*.jpg;*.jpeg;*.png;*.bmp;*.gif|所有文件|*.*"
                        dlgTitle = "选择图片文件"
                    Case "csv"
                        dlgFilter = "CSV 文件|*.csv|所有文件|*.*"
                        dlgTitle = "选择 CSV 文件"
                End Select

                ' COM 调用从 WebView2 前端发起时已在 UI 线程，直接显示对话框
                Dim selectedPath As String = ""
                Using dlg As New OpenFileDialog()
                    dlg.Title = dlgTitle
                    dlg.Filter = dlgFilter
                    dlg.Multiselect = False
                    dlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                    If dlg.ShowDialog() = DialogResult.OK Then
                        selectedPath = dlg.FileName
                    End If
                End Using

                If String.IsNullOrEmpty(selectedPath) Then
                    result.Content.Add(New McpTextContent With {.Text = JsonConvert.SerializeObject(New With {
                        .cancelled = True,
                        .message = "用户取消了文件选择"
                    }, Formatting.Indented)})
                    Return result
                End If

                result.Content.Add(New McpTextContent With {.Text = JsonConvert.SerializeObject(New With {
                    .filePath = selectedPath,
                    .fileName = Path.GetFileName(selectedPath),
                    .fileSize = New FileInfo(selectedPath).Length,
                    .extension = Path.GetExtension(selectedPath),
                    .message = "文件已选择，请根据文件类型调用合适的工具读取内容（图片文件请用 excel_ocr_recognize，文本文件请用 excel_read_file）"
                }, Formatting.Indented)})

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"文件浏览失败: {ex.Message}"})
            End Try
            Return result
        End Function

        '#End Region

    End Class

End Namespace
