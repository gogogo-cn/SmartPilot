' TECHNICAL DEBT: 本文件大量使用 Object 类型进行 Excel COM Interop（Workbook.Sheets/UsedRange/Cells 的返回值均为 Object）
' 并大量使用匿名类（New With {}）构造 JSON 输出；Option Strict On 需大量 DirectCast，收益低风险高。
Option Strict Off
Option Explicit On
Option Infer On

Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading.Tasks
Imports System.Diagnostics
Imports System.Text
Imports Newtonsoft.Json
Imports ExcelAIVSTOAddIn.Services
Imports ExcelAIVSTOAddIn.Utils

Namespace Bridges

    ''' <summary>
    ''' DocumentBridge - 文档解析 API 桥梁类
    ''' 提供 PDF/DOCX/XLSX 文档解析、语义分块、RAG 模式查询等功能
    ''' </summary>
    <ComVisible(True)>
    <ClassInterface(ClassInterfaceType.AutoDispatch)>
    <ProgId("ExcelAIVSTOAddIn.DocumentBridge")>
    Public Class DocumentBridge
        Implements IBridge

        ' 单例实例 - 使用 Lazy(Of T) 确保线程安全
        Private Shared ReadOnly _lazyInstance As New Lazy(Of DocumentBridge)(Function() New DocumentBridge(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

#Region "单例模式"

        ''' <summary>
        ''' 获取 DocumentBridge 单例实例（线程安全 Lazy 初始化）
        ''' </summary>
        Public Shared ReadOnly Property Instance As DocumentBridge
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        ''' <summary>
        ''' 私有构造函数（单例模式）
        ''' </summary>
        Private Sub New()
        End Sub

#End Region

#Region "文件保存 API"

        ''' <summary>
        ''' 保存上传文件到临时目录，返回文件路径
        ''' 前端调用: await documentBridge.SaveUploadedFile(fileName, base64Data)
        ''' </summary>
        ''' <param name="fileName">原始文件名</param>
        ''' <param name="base64Data">Base64 编码的文件内容</param>
        ''' <returns>保存后的文件路径</returns>
        <ComVisible(True)>
        Public Function SaveUploadedFile(fileName As String, base64Data As String) As String
            Try
                ' 剥离 DataURL 前缀
                Dim cleanBase64 = base64Data
                Dim commaIdx = cleanBase64.IndexOf(","c)
                If commaIdx > -1 AndAlso cleanBase64.Substring(0, Math.Min(commaIdx, 30)).Contains("base64") Then
                    cleanBase64 = cleanBase64.Substring(commaIdx + 1)
                End If

                ' 创建临时目录
                Dim tempDir = Path.Combine(Path.GetTempPath(), "ExcelAI_Uploads")
                If Not Directory.Exists(tempDir) Then
                    Directory.CreateDirectory(tempDir)
                End If

                ' 生成唯一文件名
                Dim ext = Path.GetExtension(fileName)
                Dim uniqueName = $"{Guid.NewGuid()}{ext}"
                Dim filePath = Path.Combine(tempDir, uniqueName)

                ' 写入文件
                Dim fileBytes = Convert.FromBase64String(cleanBase64)
                File.WriteAllBytes(filePath, fileBytes)

                System.Diagnostics.Debug.WriteLine($"[DocumentBridge] 文件已保存: {filePath} ({fileBytes.Length} bytes)")
                Return filePath

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[DocumentBridge] SaveUploadedFile 异常: {ex.Message}")
                Return ""
            End Try
        End Function

#End Region

#Region "文档解析 API (PDF/DOCX/XLSX 统一处理)"

        ''' <summary>
        ''' 解析文档文件（PDF/DOC/DOCX/XLS/XLSX）- 提取全文和结构信息
        ''' 前端调用: await documentBridge.ParseDocumentAsync(fileName, base64Data)
        ''' </summary>
        ''' <param name="fileName">文件名（含扩展名）</param>
        ''' <param name="base64Data">Base64 编码的文件内容</param>
        ''' <returns>解析结果JSON，包含全文、结构、元数据等</returns>
        <ComVisible(True)>
        Public Async Function ParseDocumentAsync(fileName As String, base64Data As String) As Task(Of String)
            Try
                ' 剥离可能的 DataURL 前缀（data:...;base64,），只保留纯 Base64 数据
                Dim cleanBase64 = base64Data
                Dim commaIdx = cleanBase64.IndexOf(","c)
                If commaIdx > -1 AndAlso cleanBase64.Substring(0, Math.Min(commaIdx, 30)).Contains("base64") Then
                    cleanBase64 = cleanBase64.Substring(commaIdx + 1)
                End If

                ' 文件大小限制检查 (200MB)
                Dim rawBytes = Convert.FromBase64String(cleanBase64)
                If rawBytes.Length > 200 * 1024 * 1024 Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = "文件过大，最大支持 200MB"
                    })
                End If

                System.Diagnostics.Debug.WriteLine($"[DocumentBridge] ParseDocumentAsync 开始: {fileName}, 原始Base64长度={base64Data.Length}, 清理后长度={cleanBase64.Length}, 字节大小={rawBytes.Length} bytes")

                ' 调用 DocumentRouter 自动路由到对应 Parser
                Dim result = Await DocumentRouter.Instance.ParseDocumentAsync(fileName, cleanBase64)

                Return JsonConvert.SerializeObject(New With {
                    .success = result.Success,
                    .fileType = result.FileType,
                    .fullText = result.FullText,
                    .totalChars = result.TotalChars,
                    .totalPages = result.TotalPages,
                    .isScanned = result.IsScanned,
                    .structure = result.Structure,
                    .metadata = result.Metadata,
                    .pages = result.Pages,
                    .parseTimeMs = result.ParseTimeMs,
                    .error = If(result.Success, Nothing, result.ErrorMessage)
                })
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[DocumentBridge] ParseDocumentAsync 异常: {ex.Message}")
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 从文件路径直接解析 Excel 文档（使用 Excel 对象模型，无需 Base64）
        ''' 前端调用: await documentBridge.ParseDocumentFromFile(filePath)
        ''' </summary>
        ''' <param name="filePath">Excel 文件的完整路径</param>
        ''' <returns>解析结果JSON</returns>
        <ComVisible(True)>
        Public Function ParseDocumentFromFile(filePath As String) As String
            Dim excelApp As Object = Nothing
            Dim workbook As Object = Nothing
            Dim sheets As Object = Nothing

            Try
                System.Diagnostics.Debug.WriteLine($"[DocumentBridge] ParseDocumentFromFile: {filePath}")

                If Not File.Exists(filePath) Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = $"文件不存在: {filePath}"
                    })
                End If

                Dim ext = Path.GetExtension(filePath).ToLower()
                If ext <> ".xls" AndAlso ext <> ".xlsx" Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = $"不支持的文件格式: {ext}"
                    })
                End If

                ' 使用 Excel 对象模型直接读取
                excelApp = Globals.ThisAddIn.Application
                workbook = excelApp.Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
                sheets = workbook.Sheets
                Dim sheetCount As Integer = sheets.Count

                Dim fullText = New StringBuilder()
                Dim pages = New List(Of Object)()
                Dim headings = New List(Of Object)()
                Dim tables = New List(Of Object)()
                Dim sheetNames = New List(Of String)()
                Dim totalChars = 0
                Dim globalPosition As Long = 0

                For i As Integer = 1 To sheetCount
                    Dim sheet As Object = Nothing
                    Dim usedRange As Object = Nothing
                    Dim rows As Object = Nothing
                    Dim columns As Object = Nothing

                    Try
                        sheet = sheets(i)
                        Dim sheetName = sheet.Name
                        Dim sheetTextBuilder = New StringBuilder()
                        sheetNames.Add(sheetName)

                        ' Sheet 标题标记（与 NPOI 路径格式一致）
                        sheetTextBuilder.AppendLine($"【工作表: {sheetName}】")
                        sheetTextBuilder.AppendLine()

                        ' 获取已用范围
                        usedRange = sheet.UsedRange
                        Dim rowCount = 0
                        Dim colCount = 0
                        Dim headerRowValues As List(Of String) = Nothing

                        If usedRange IsNot Nothing Then
                            rows = usedRange.Rows
                            columns = usedRange.Columns
                            rowCount = rows.Count
                            colCount = columns.Count

                            For row As Integer = 1 To rowCount
                                Dim rowValues = New List(Of String)()
                                For col As Integer = 1 To colCount
                                    Dim cell As Object = Nothing
                                    Try
                                        cell = usedRange.Cells(row, col)
                                        Dim cellValue = ""

                                        ' 【P1-1修复】优先使用 Text（格式化显示值），避免日期变数字、大数变科学计数法
                                        Try
                                            cellValue = cell.Text
                                        Catch ex As Exception  '  Text 不可用时回退到 Value2
                                            Dim value2 = cell.Value2
                                            If value2 IsNot Nothing Then
                                                cellValue = value2.ToString()
                                                Utils.ComReleaseHelper.Release(value2)
                                            End If
                                        End Try

                                        rowValues.Add(cellValue)
                                    Finally
                                        Utils.ComReleaseHelper.Release(cell)
                                    End Try
                                Next

                                ' 过滤掉完全空的行
                                Dim hasContent = rowValues.Any(Function(v) Not String.IsNullOrEmpty(v.Trim()))
                                If Not hasContent Then Continue For

                                Dim rowText = String.Join(vbTab, rowValues)

                                ' 第一行作为表头，添加特殊标记（与 NPOI 路径格式一致）
                                If row = 1 Then
                                    headerRowValues = rowValues
                                    sheetTextBuilder.AppendLine($"【表头】{rowText}")
                                Else
                                    sheetTextBuilder.AppendLine(rowText)
                                End If
                            Next
                        End If

                        sheetTextBuilder.AppendLine()

                        ' 构建 PageContent（与 NPOI 路径 PageContent 结构一致）
                        Dim sheetText = sheetTextBuilder.ToString()
                        Dim pageContent = New With {
                            .index = i,
                            .title = sheetName,
                            .text = sheetText,
                            .charCount = sheetText.Length,
                            .startPosition = globalPosition
                        }
                        pages.Add(pageContent)

                        ' 【P1-2】构建 Headings（与 NPOI BuildStructureFromWorkbook 对齐）
                        Dim headingTitle = $"{sheetName} ({rowCount}行 x {colCount}列)"
                        headings.Add(New With {
                            .Level = 1,
                            .Title = headingTitle,
                            .StartPosition = globalPosition,
                            .EndPosition = 0,
                            .PageIndex = i,
                            .Children = New List(Of Object)()
                        })

                        ' 【P1-2】构建 Tables（与 NPOI BuildStructureFromWorkbook 对齐）
                        tables.Add(New With {
                            .StartPosition = globalPosition,
                            .RowCount = rowCount,
                            .ColCount = colCount,
                            .HeaderRow = If(headerRowValues, New List(Of String)()),
                            .Summary = $"工作表: {sheetName}, 行数: {rowCount}, 列数: {colCount}"
                        })

                        ' 追加到全文
                        fullText.Append(sheetText)
                        globalPosition += sheetText.Length
                        totalChars += sheetText.Length
                    Finally
                        Utils.ComReleaseHelper.Release(columns)
                        Utils.ComReleaseHelper.Release(rows)
                        Utils.ComReleaseHelper.Release(usedRange)
                        Utils.ComReleaseHelper.Release(sheet)
                    End Try
                Next

                ' 【P1-3】构建 metadata（与 NPOI BuildMetadataFromWorkbook 对齐）
                Dim metadata = New With {
                    .fileFormat = ext.TrimStart("."c),
                    .fileSizeBytes = New FileInfo(filePath).Length,
                    .properties = New Dictionary(Of String, String) From {
                        {"工作表列表", String.Join(", ", sheetNames)},
                        {"格式", ext.TrimStart("."c).ToUpper()}
                    }
                }

                ' 【P1-2】构建 structure（structure 是 VB 关键字，用方括号转义）
                Dim docStructure = New With {
                    .hasTableOfContents = sheetCount > 1,
                    .sectionCount = sheetCount,
                    .headings = headings,
                    .tables = tables,
                    .images = New List(Of Object)()
                }

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .fileType = ext.TrimStart("."c),
                    .fullText = fullText.ToString(),
                    .totalChars = totalChars,
                    .totalPages = sheetCount,
                    .isScanned = False,
                    .structure = docStructure,
                    .metadata = metadata,
                    .pages = pages,
                    .parseTimeMs = 0
                })

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[DocumentBridge] ParseDocumentFromFile 异常: {ex.Message}")
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            Finally
                If workbook IsNot Nothing Then
                    Try
                        workbook.Close(SaveChanges:=False)
                    Catch ex As Exception  ' 
                    End Try
                End If
                Utils.ComReleaseHelper.Release(workbook)
                Utils.ComReleaseHelper.Release(sheets)
            End Try
        End Function

        ''' <summary>
        ''' 对长文档进行语义分块并生成摘要
        ''' 前端调用: await documentBridge.ChunkDocumentAsync(fullText, structureJson, optionsJson)
        ''' </summary>
        ''' <param name="fullText">完整文档文本</param>
        ''' <param name="structureJson">结构化标题树 JSON（可选）</param>
        ''' <param name="optionsJson">分块选项 JSON（可选）</param>
        ''' <returns>分块结果JSON，包含块列表和全局摘要</returns>
        <ComVisible(True)>
        Public Async Function ChunkDocumentAsync(
            fullText As String,
            Optional structureJson As String = "",
            Optional optionsJson As String = ""
        ) As Task(Of String)
            ' 超时保护：30秒内完成，否则返回失败让前端降级
            Dim timeoutMs As Integer = 30000
            Dim cts As New System.Threading.CancellationTokenSource(timeoutMs)

            Try
                Dim chunkTask As Task(Of String) = Task.Run(Of String)(Function()
                    ' 解析选项
                    Dim chunkSize As Integer = 1500
                    Dim overlap As Integer = 200
                    Dim mode As String = "auto"

                    If Not String.IsNullOrWhiteSpace(optionsJson) Then
                        Try
                            Dim opts = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(optionsJson)
                            If opts.ContainsKey("chunkSize") Then chunkSize = CInt(opts("chunkSize"))
                            If opts.ContainsKey("overlap") Then overlap = CInt(opts("overlap"))
                            If opts.ContainsKey("mode") Then mode = opts("mode").ToString()
                        Catch ex As Exception  '  使用默认值
                        End Try
                    End If

                    ' 解析结构化标题树（如有）
                    Dim headings As List(Of Services.Parsers.HeadingNode) = Nothing
                    If Not String.IsNullOrWhiteSpace(structureJson) Then
                        Try
                            headings = JsonConvert.DeserializeObject(Of List(Of Services.Parsers.HeadingNode))(structureJson)
                        Catch ex As Exception  ' 
                        End Try
                    End If

                    System.Diagnostics.Debug.WriteLine($"[DocumentBridge] ChunkDocumentAsync 开始: 文本长度={fullText.Length}, 模式={mode}")

                    ' 调用语义分块器
                    Dim chunker = Utils.SemanticChunker.Instance
                    Dim chunks = chunker.Chunk(fullText, headings, mode)

                    ' 构建全局摘要
                    Dim summaryText = chunker.BuildChunkSummary(chunks)

                    ' 序列化结果
                    Dim chunksData = chunks.Select(Function(c) New With {
                        .id = c.Id,
                        .text = c.Text,
                        .charCount = c.CharCount,
                        .startPosition = c.StartPosition,
                        .endPosition = c.EndPosition,
                        .sourceSection = c.SourceSection,
                        .level = c.Level
                    }).ToList()

                    Return JsonConvert.SerializeObject(New With {
                        .success = True,
                        .totalChunks = chunks.Count,
                        .chunks = chunksData,
                        .summary = summaryText,
                        .totalChars = fullText.Length
                    })
                End Function, cts.Token)

                Return Await chunkTask
            Catch ex As System.Threading.Tasks.TaskCanceledException
                System.Diagnostics.Debug.WriteLine("[DocumentBridge] ChunkDocumentAsync 超时(30s)")
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = "分块超时，请使用前端降级分块"
                })
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[DocumentBridge] ChunkDocumentAsync 异常: {ex.Message}")
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            Finally
                cts.Dispose()
            End Try
        End Function

        ''' <summary>
        ''' 基于用户问题从文档块中检索相关内容并回答 (RAG 模式)
        ''' 前端调用: await documentBridge.QueryDocumentAsync(question, chunksJson)
        ''' </summary>
        ''' <param name="question">用户问题</param>
        ''' <param name="chunksJson">文档块列表 JSON</param>
        ''' <param name="topK">返回的相关块数量（默认5）</param>
        ''' <returns>查询结果JSON，包含上下文和相关块列表</returns>
        <ComVisible(True)>
        Public Async Function QueryDocumentAsync(
            question As String,
            chunksJson As String,
            Optional topK As Integer = 5
        ) As Task(Of String)
            Return Await Task.Run(Of String)(Function()
                Try
                    ' 反序列化 chunks
                    Dim chunksInput = JsonConvert.DeserializeObject(Of List(Of Dictionary(Of String, Object)))(chunksJson)
                    Dim textChunks As New List(Of Utils.TextChunk)()

                    For Each item In chunksInput
                        textChunks.Add(New Utils.TextChunk With {
                            .Id = CInt(item("id")),
                            .Text = item("text").ToString(),
                            .CharCount = CInt(item("charCount")),
                            .StartPosition = CLng(item("startPosition")),
                            .EndPosition = CLng(item("endPosition")),
                            .SourceSection = If(item.ContainsKey("sourceSection"), item("sourceSection").ToString(), ""),
                            .Level = If(item.ContainsKey("level"), CInt(item("level")), 1)
                        })
                    Next

                    System.Diagnostics.Debug.WriteLine($"[DocumentBridge] QueryDocumentAsync 开始: 问题={question}, 块数={textChunks.Count}")

                    ' 关键词检索 Top-K 相关块
                    Dim chunker = Utils.SemanticChunker.Instance
                    Dim relevantChunks = chunker.RetrieveRelevantChunks(question, textChunks, topK)

                    ' 组装上下文文本
                    Dim contextBuilder As New StringBuilder()
                    contextBuilder.AppendLine("[相关文档片段]")
                    For i = 0 To relevantChunks.Count - 1
                        contextBuilder.AppendLine($"--- 片段 {i + 1} (来源: {relevantChunks(i).SourceSection}) ---")
                        contextBuilder.AppendLine(relevantChunks(i).Text)
                        contextBuilder.AppendLine()
                    Next

                    Return JsonConvert.SerializeObject(New With {
                        .success = True,
                        .context = contextBuilder.ToString(),
                        .relevantChunks = relevantChunks.Select(Function(c) New With {
                            .id = c.Id,
                            .sourceSection = c.SourceSection,
                            .charPreview = c.Text.Substring(0, Math.Min(200, c.Text.Length)) & "..."
                        }).ToList()
                    })
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"[DocumentBridge] QueryDocumentAsync 异常: {ex.Message}")
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = ex.Message
                    })
                End Try
            End Function)
        End Function

        ''' <summary>
        ''' 获取文档解析器支持的文件格式列表
        ''' 前端调用: documentBridge.GetSupportedDocFormats()
        ''' </summary>
        ''' <returns>支持的文件格式列表JSON</returns>
        <ComVisible(True)>
        Public Function GetSupportedDocFormats() As String
            Try
                Dim formats = DocumentRouter.Instance.GetSupportedExtensions()
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .formats = formats,
                    .count = formats.Count
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

#End Region

#Region "MinerU 高级解析 API"

        ''' <summary>
        ''' 检查 MinerU 服务状态
        ''' 前端调用: await documentBridge.CheckMinerUStatusAsync()
        ''' </summary>
        <ComVisible(True)>
        Public Async Function CheckMinerUStatusAsync() As Task(Of String)
            Try
                Dim mineruService As MinerUService = MinerUService.Instance
                Dim available = Await mineruService.IsAvailableAsync()
                Dim installStatus = Await mineruService.GetInstallStatusAsync()

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .available = available,
                    .details = installStatus
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .available = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 使用 MinerU 解析 PDF（支持图片/表格/公式）
        ''' 前端调用: await documentBridge.ParsePdfWithMineruAsync(base64Data)
        ''' </summary>
        <ComVisible(True)>
        Public Async Function ParsePdfWithMineruAsync(base64Data As String) As Task(Of String)
            Try
                Dim mineruService As MinerUService = MinerUService.Instance
                Dim result = Await mineruService.ParsePdfAsync(base64Data, extractImages:=True, fallback:=True)

                Return JsonConvert.SerializeObject(New With {
                    .success = result.Success,
                    .fullText = result.FullText,
                    .markdown = result.Markdown,
                    .imageCount = If(result.Images?.Count, 0),
                    .tableCount = If(result.Tables?.Count, 0),
                    .formulaCount = If(result.Formulas?.Count, 0),
                    .pageCount = result.PageCount,
                    .parseTimeMs = result.ParseTimeMs,
                    .engine = result.Engine,
                    .error = result.Error
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

#End Region

#Region "IBridge 生命周期"

        Public ReadOnly Property BridgeName As String Implements IBridge.BridgeName
            Get
                Return "DocumentBridge"
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
            End If
        End Sub

#End Region

    End Class

End Namespace