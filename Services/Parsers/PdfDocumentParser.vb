Option Strict On
Option Explicit On
Option Infer On
Imports System
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Diagnostics
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.IO

Namespace Services.Parsers

    ''' <summary>
    ''' PDF 文档解析器（PdfSharp/MinerU 双引擎）
    ''' P5 加固：硬文件字节上限、严格类型开关、单页字符截断叠加、MemoryStream容量声明。
    ''' </summary>
    Public Class PdfDocumentParser
        Implements IDocumentParser

        ' ==================== 常量定义 ====================

        ' —— P5：硬性安全上限（防超大文件 / 大对象堆LOH爆炸）——
        ''' <summary>P5：单PDF文件字节硬上限（超过直接拒绝，防 FromBase64 + MemoryStream 爆炸）</summary>
        Private Const MAX_FILE_SIZE_BYTES As Long = 2L * 1024L * 1024L * 1024L  ' 2 GB（PDF通常比xlsx更能容纳图）

        ''' <summary>扫描版判定阈值：每页平均字符数低于此值则视为扫描版</summary>
        Private Const SCANNED_THRESHOLD_CHARS_PER_PAGE As Integer = 10

        ''' <summary>最大允许的 PDF 大小（字节）- 超过此大小建议使用分页提取</summary>
        Private Const MAX_RECOMMENDED_SIZE_BYTES As Long = 50 * 1024 * 1024 ' 50MB

        ''' <summary>P2 防护：最大解析页数（防止超大PDF耗尽内存）</summary>
        Private Const MAX_PAGES_TO_PARSE As Integer = 2000

        ''' <summary>P2 防护：单页最大提取字符数（防止单页超大内容）</summary>
        Private Const MAX_PAGE_TEXT_LENGTH As Integer = 500000

        ''' <summary>P2 防护：全文最大字符上限（超过则截断）</summary>
        Private Const MAX_FULLTEXT_LENGTH As Integer = 20 * 1024 * 1024 ' 20MB 文本

        ' ==================== IDocumentParser 接口实现 ====================

        ''' <summary>支持的文件扩展名</summary>
        Public ReadOnly Property SupportedExtensions As List(Of String) Implements IDocumentParser.SupportedExtensions
            Get
                Return New List(Of String) From {"pdf"}
            End Get
        End Property

        ''' <summary>解析器显示名称</summary>
        Public ReadOnly Property DisplayName As String Implements IDocumentParser.DisplayName
            Get
                Return "PDF 文档解析器"
            End Get
        End Property

        Public Async Function ExtractFullText(base64Data As String) As Task(Of UnifiedDocumentResult) Implements IDocumentParser.ExtractFullText
            If String.IsNullOrWhiteSpace(base64Data) Then
                Return New UnifiedDocumentResult() With {
                    .Success = False,
                    .ErrorMessage = "传入的PDF数据为空",
                    .FileType = "pdf"
                }
            End If

            Dim bytes As Byte() = Nothing
            Try
                bytes = Convert.FromBase64String(base64Data.Trim())
            Catch exFmt As FormatException
                Return New UnifiedDocumentResult() With {
                    .Success = False,
                    .ErrorMessage = "传入的文件不是合法的 Base64 编码: " & exFmt.Message,
                    .FileType = "pdf"
                }
            End Try

            If bytes Is Nothing OrElse bytes.Length < 5 Then
                Return New UnifiedDocumentResult() With {
                    .Success = False,
                    .ErrorMessage = "PDF 文件数据太小，不是有效文档",
                    .FileType = "pdf"
                }
            End If

            ' —— P5：PDF单文件字节硬上限（MinerU/PdfSharp走之前都必须先过；拒绝超1GB/2GB的超大文档）——
            If CLng(bytes.LongLength) > MAX_FILE_SIZE_BYTES Then
                Dim sizeMB As Long = CLng(bytes.LongLength) \ (1024L * 1024L)
                Debug.WriteLine("[PdfDocumentParser] PDF超过MAX_FILE_SIZE_BYTES, 大小约=" & sizeMB.ToString() & "MB，拒绝解析")
                Return New UnifiedDocumentResult() With {
                    .Success = False,
                    .ErrorMessage = "PDF 文件大小超过安全上限（" & (MAX_FILE_SIZE_BYTES \ (1024L * 1024L)).ToString() & "MB）。当前约 " & sizeMB.ToString() & "MB。请使用分页提取或减小文件。",
                    .FileType = "pdf"
                }
            End If

            ' 尝试 MinerU 增强解析（服务健康时）
            Dim mineruResult = Await TryMinerUParseAsync(base64Data)
            If mineruResult IsNot Nothing AndAlso mineruResult.Success AndAlso mineruResult.TotalChars > 0 Then
                Return mineruResult
            End If

            ' 降级到 PdfSharp 基础解析
            Dim pdfSharpResult = Await ExtractWithPdfSharpAsync(base64Data)

            ' 如果 PdfSharp 未提取到文本，强制尝试 MinerU（扫描件 OCR fallback）
            If pdfSharpResult Is Nothing OrElse pdfSharpResult.TotalChars = 0 Then
                Debug.WriteLine("[PdfDocumentParser] PdfSharp 未提取到文本，强制尝试 MinerU OCR...")
                Dim forcedMineruResult = Await ForceMinerUParseAsync(base64Data)
                If forcedMineruResult IsNot Nothing AndAlso forcedMineruResult.Success AndAlso forcedMineruResult.TotalChars > 0 Then
                    Return forcedMineruResult
                End If
            End If

            Return pdfSharpResult
        End Function

        Private Async Function ExtractWithPdfSharpAsync(base64Data As String) As Task(Of UnifiedDocumentResult)
            Return Await Task.Run(Function()
                Dim stopwatch As New Stopwatch()
                stopwatch.Start()

                Dim result As New UnifiedDocumentResult()
                result.FileType = "pdf"

                Debug.WriteLine("[PdfDocumentParser] 开始提取 PDF 全文...")

                Try
                    Dim pdfBytes As Byte() = Convert.FromBase64String(base64Data)
                    result.Metadata = New DocumentMetadata()
                    result.Metadata.FileSizeBytes = pdfBytes.LongLength

                    Debug.WriteLine("[PdfDocumentParser] PDF 文件大小: " & pdfBytes.Length.ToString() & " 字节")

                    If pdfBytes.LongLength > MAX_RECOMMENDED_SIZE_BYTES Then
                        Debug.WriteLine("[PdfDocumentParser] ⚠️ PDF 超过推荐大小 " & (MAX_RECOMMENDED_SIZE_BYTES \ 1024 \ 1024).ToString() & "MB，可能影响性能")
                    End If

                    ' P5：MemoryStream 包装使用(byte[],index,count,writable,visible)避免重新分配合并byte[]
                    Using memoryStream As New MemoryStream(pdfBytes, 0, pdfBytes.Length, False, True)
                        Dim document As PdfDocument = Nothing
                        Try
                            document = PdfReader.Open(memoryStream, PdfDocumentOpenMode.InformationOnly)
                        Catch exOpen As Exception
                            result.Success = False
                            result.ErrorMessage = "无法打开PDF文件(可能已损坏或加密): " & exOpen.Message
                            Return result
                        End Try

                        Try
                            result.Metadata = ExtractMetadataInternal(document)
                        Catch exMeta As Exception
                            Debug.WriteLine("[PdfDocumentParser] 提取Metadata失败: " & exMeta.Message)
                            result.Metadata = If(result.Metadata, New DocumentMetadata() With {.FileSizeBytes = pdfBytes.LongLength})
                        End Try
                        result.TotalPages = document.PageCount

                        Debug.WriteLine("[PdfDocumentParser] PDF 页数: " & document.PageCount.ToString())

                        Dim effectivePages = Math.Min(document.PageCount, MAX_PAGES_TO_PARSE)
                        If document.PageCount > MAX_PAGES_TO_PARSE Then
                            Debug.WriteLine("[PdfDocumentParser] 页数超过上限 " & MAX_PAGES_TO_PARSE.ToString() & "，仅处理前 " & effectivePages.ToString() & " 页")
                        End If

                        memoryStream.Position = 0
                        Dim contentDoc As PdfDocument = Nothing
                        Try
                            contentDoc = PdfReader.Open(memoryStream, PdfDocumentOpenMode.Import)
                        Catch exOpen2 As Exception
                            result.Success = False
                            result.ErrorMessage = "无法打开PDF内容流: " & exOpen2.Message
                            Return result
                        End Try

                        ' P5：StringBuilder初始容量按文件大小估（文本通常≈文件20-30%字符量）
                        Dim estimateInitial As Integer = CInt(Math.Min(CLng(pdfBytes.Length) \ 2L + 1024L, CLng(MAX_FULLTEXT_LENGTH)))
                        If estimateInitial < 4096 Then estimateInitial = 4096
                        Dim fullTextBuilder As New StringBuilder(estimateInitial)
                        result.Pages = New List(Of PageContent)()
                        Dim totalExtractedChars As Long = 0
                        Dim currentStartPosition As Long = 0
                        Dim fullTextTruncated As Boolean = False

                        For pageIndex As Integer = 0 To effectivePages - 1
                            Dim pageNumberSafe As Integer = pageIndex + 1
                            Dim pageText As String = ""
                            Try
                                Dim page As PdfPage = contentDoc.Pages(pageIndex)
                                pageText = If(ExtractTextFromPage(page, pageNumberSafe), "")
                            Catch exPage As Exception
                                Debug.WriteLine("[PdfDocumentParser] 提取第 " & pageNumberSafe.ToString() & " 页失败: " & exPage.Message)
                                pageText = "[该页解析失败: " & exPage.Message & "]"
                            End Try

                            If Not String.IsNullOrEmpty(pageText) AndAlso pageText.Length > MAX_PAGE_TEXT_LENGTH Then
                                pageText = pageText.Substring(0, MAX_PAGE_TEXT_LENGTH) & vbCrLf & "[...该页内容已截断...]"
                            End If

                            Dim pageContent As New PageContent()
                            pageContent.Index = pageNumberSafe
                            pageContent.Title = "第 " & pageNumberSafe.ToString() & " 页"
                            pageContent.Text = pageText
                            pageContent.CharCount = pageText.Length
                            pageContent.StartPosition = currentStartPosition
                            result.Pages.Add(pageContent)

                            If Not fullTextTruncated Then
                                If fullTextBuilder.Length + pageText.Length + 4 > MAX_FULLTEXT_LENGTH Then
                                    Dim remaining = Math.Max(0, MAX_FULLTEXT_LENGTH - fullTextBuilder.Length - 4)
                                    If remaining > 0 AndAlso pageText.Length > 0 Then
                                        fullTextBuilder.AppendLine(pageText.Substring(0, Math.Min(remaining, pageText.Length)))
                                    End If
                                    fullTextBuilder.AppendLine()
                                    fullTextBuilder.AppendLine("[...全文过长已截断，剩余内容未提取...]")
                                    fullTextTruncated = True
                                Else
                                    fullTextBuilder.AppendLine(pageText)
                                    fullTextBuilder.AppendLine()
                                End If
                            End If

                            totalExtractedChars += pageText.Length
                            currentStartPosition = fullTextBuilder.Length
                        Next

                        result.FullText = fullTextBuilder.ToString().Trim()
                        result.TotalChars = result.FullText.Length
                        result.Success = True

                        Dim pageCountForAvg As Integer = If(contentDoc.PageCount > 0, contentDoc.PageCount, 1)
                        Dim avgCharsPerPage As Double = CDbl(totalExtractedChars) / CDbl(pageCountForAvg)
                        result.IsScanned = (avgCharsPerPage < SCANNED_THRESHOLD_CHARS_PER_PAGE) AndAlso contentDoc.PageCount > 0

                        If result.IsScanned Then
                            result.ErrorMessage = "检测到该 PDF 可能为扫描版图片文档，无法直接提取文本内容。建议使用 OCR 功能进行文字识别。"
                            Debug.WriteLine("[PdfDocumentParser] ⚠️ 检测到扫描版 PDF")
                        End If

                        Try
                            If contentDoc.Outlines IsNot Nothing Then
                                result.Structure = ExtractBookmarksFromOutline(contentDoc.Outlines)
                            Else
                                result.Structure = New DocumentStructure() With {
                                    .Headings = New List(Of HeadingNode)(),
                                    .Tables = New List(Of TableInfo)(),
                                    .Images = New List(Of ImageInfo)()
                                }
                            End If
                        Catch exOutline As Exception
                            Debug.WriteLine("[PdfDocumentParser] 提取书签失败: " & exOutline.Message)
                            result.Structure = New DocumentStructure() With {
                                .Headings = New List(Of HeadingNode)(),
                                .Tables = New List(Of TableInfo)(),
                                .Images = New List(Of ImageInfo)()
                            }
                        End Try

                        Try
                            contentDoc.Close()
                        Catch ex As Exception  ' 
                        End Try
                    End Using

                    stopwatch.Stop()
                    result.ParseTimeMs = stopwatch.ElapsedMilliseconds

                    Debug.WriteLine("[PdfDocumentParser] ✅ PDF 解析完成 - 耗时: " & result.ParseTimeMs.ToString() & "ms, 总字符数: " & result.TotalChars.ToString() & ", 是否扫描版: " & result.IsScanned.ToString())

                Catch ex As Exception
                    result.Success = False
                    result.ErrorMessage = "PDF 解析失败: " & ex.Message
                    result.IsScanned = False

                    Debug.WriteLine("[PdfDocumentParser] ❌ PDF 解析异常: " & ex.ToString())
                End Try

                Return result
            End Function)
        End Function

        ''' <summary>
        ''' 提取 PDF 文档结构信息（书签/目录/标题层级）
        ''' </summary>
        Public Async Function ExtractStructure(base64Data As String) As Task(Of DocumentStructure) Implements IDocumentParser.ExtractStructure
            Return Await Task.Run(Of DocumentStructure)(Function()
                Debug.WriteLine("[PdfDocumentParser] 开始提取 PDF 结构...")

                Dim structureInfo As New DocumentStructure()

                Try
                    Dim pdfBytes As Byte() = Convert.FromBase64String(base64Data)

                    Using memoryStream As New MemoryStream(pdfBytes)
                        Dim document As PdfDocument = PdfReader.Open(memoryStream, PdfDocumentOpenMode.InformationOnly)

                        ' 从 PDF 书签中提取目录结构
                        structureInfo = ExtractBookmarksFromOutline(document.Outlines)

                        structureInfo.SectionCount = If(structureInfo.Headings IsNot Nothing, structureInfo.Headings.Count, 0)

                        document.Close()
                    End Using

                    Debug.WriteLine(String.Format("[PdfDocumentParser] ✅ PDF 结构提取完成 - 章节数: {0}", structureInfo.SectionCount))

                Catch ex As Exception
                    Debug.WriteLine(String.Format("[PdfDocumentParser] ❌ PDF 结构提取异常: {0}", ex.Message))
                    structureInfo.Headings = New List(Of HeadingNode)()
                End Try

                Return structureInfo
            End Function)
        End Function

        ''' <summary>
        ''' 提取 PDF 元数据信息（标题、作者、创建日期等）
        ''' </summary>
        Public Function ExtractMetadata(base64Data As String) As DocumentMetadata Implements IDocumentParser.ExtractMetadata
            Debug.WriteLine("[PdfDocumentParser] 开始提取 PDF 元数据...")

            Dim metadata As New DocumentMetadata()

            Try
                Dim pdfBytes As Byte() = Convert.FromBase64String(base64Data)
                metadata.FileSizeBytes = pdfBytes.LongLength

                Using memoryStream As New MemoryStream(pdfBytes)
                    Dim document As PdfDocument = PdfReader.Open(memoryStream, PdfDocumentOpenMode.InformationOnly)
                    metadata = ExtractMetadataInternal(document)
                    document.Close()
                End Using

                Debug.WriteLine("[PdfDocumentParser] ✅ PDF 元数据提取完成")

            Catch ex As Exception
                Debug.WriteLine(String.Format("[PdfDocumentParser] ❌ PDF 元数据提取异常: {0}", ex.Message))
            End Try

            Return metadata
        End Function

        ''' <summary>
        ''' 提取指定范围的 PDF 内容（用于大文件分页加载）
        ''' 参数：startPosition - 起始页码（从1开始），maxLength - 要提取的最大页数
        ''' 返回：指定页码范围内的文本内容
        ''' </summary>
        Public Async Function ExtractRange(base64Data As String, startPosition As Long, maxLength As Integer) As Task(Of String) Implements IDocumentParser.ExtractRange
            Return Await Task.Run(Of String)(Function()
                Debug.WriteLine(String.Format("[PdfDocumentParser] 开始提取 PDF 范围内容 - 起始页: {0}, 最大长度: {1} 页", startPosition, maxLength))

                Dim rangeText As String = ""

                Try
                    Dim pdfBytes As Byte() = Convert.FromBase64String(base64Data)

                    Using memoryStream As New MemoryStream(pdfBytes)
                        Dim document As PdfDocument = PdfReader.Open(memoryStream, PdfDocumentOpenMode.Import)

                        ' 参数校验
                        If startPosition < 1 Then startPosition = 1
                        If maxLength < 1 Then maxLength = 1

                        Dim startPageIndex As Integer = CInt(startPosition) - 1 ' 转换为0-based索引
                        Dim endPageIndex As Integer = Math.Min(startPageIndex + maxLength - 1, document.PageCount - 1)

                        ' 边界检查
                        If startPageIndex >= document.PageCount Then
                            Debug.WriteLine(String.Format("[PdfDocumentParser] ⚠️ 起始页超出范围（总页数: {0}）", document.PageCount))
                            Return ""
                        End If

                        ' 提取指定范围的页面文本
                        Dim textBuilder As New StringBuilder()

                        For pageIndex As Integer = startPageIndex To endPageIndex
                            Dim page As PdfPage = document.Pages(pageIndex)
                            Dim pageText As String = ExtractTextFromPage(page, pageIndex + 1)

                            textBuilder.AppendFormat("--- 第 {0} 页 ---{1}{2}{1}",
                                                      pageIndex + 1,
                                                      Environment.NewLine,
                                                      pageText)
                            textBuilder.AppendLine()
                        Next

                        rangeText = textBuilder.ToString().Trim()

                        Debug.WriteLine(String.Format("[PdfDocumentParser] ✅ 范围提取完成 - 实际提取页数: {0}, 字符数: {1}",
                                                          (endPageIndex - startPageIndex + 1), rangeText.Length))

                        document.Close()
                    End Using

                Catch ex As Exception
                    Debug.WriteLine(String.Format("[PdfDocumentParser] ❌ PDF 范围提取异常: {0}", ex.Message))
                    rangeText = String.Format("提取失败: {0}", ex.Message)
                End Try

                Return rangeText
            End Function)
        End Function

        ' ==================== 私有辅助方法 ====================

        ''' <summary>
        ''' 从单个 PDF 页面提取文本内容
        ''' 策略1（主）：PdfTextExtractor —— PdfSharp ContentReader 语法树遍历，
        '''   按字体（ToUnicode CMap / Type0 CID / WinAnsi 编码表）解码 Tj/TJ/'/" 算子，
        '''   跟踪 Td/Tm/TD/T* 文本定位保留行结构，递归进入 Form XObject（标题栏）。
        ''' 策略2（兜底）：旧的内容流括号启发式，仅当策略1无产出时尝试。
        ''' </summary>
        Private Function ExtractTextFromPage(page As PdfPage, pageNumber As Integer) As String
            Dim pageText As String = ""

            Try
                ' 策略1：ContentReader + 字体解码提取
                pageText = PdfTextExtractor.ExtractPageText(page)

                ' 策略2：旧启发式兜底（新提取器无产出时再试，成本极低）
                If String.IsNullOrEmpty(pageText) Then
                    pageText = ExtractTextFromContents(page)
                End If

                ' 清理：仅折叠行内空白（空格/制表符），保留换行结构
                ' —— 图纸的尺寸标注/技术要求按行分布，\s+ 整体折叠会让行结构全部丢失
                If Not String.IsNullOrEmpty(pageText) Then
                    pageText = System.Text.RegularExpressions.Regex.Replace(pageText, "[ \t]+", " ")
                    pageText = pageText.Trim()
                End If

            Catch ex As Exception
                Debug.WriteLine(String.Format("[PdfDocumentParser] 第 {0} 页文本提取异常: {1}", pageNumber, ex.Message))
                pageText = ""
            End Try

            Return pageText
        End Function

        ''' <summary>
        ''' 从 PDF 页面的 Contents 中提取文本
        ''' 解析 PDF 内容流中的文本操作符（Tj, TJ, Tm 等）
        ''' </summary>
        Private Function ExtractTextFromContents(page As PdfPage) As String
            Dim extractedText As New StringBuilder()

            Try
                ' 获取页面的内容流
                If page.Contents Is Nothing OrElse page.Contents.Elements.Count = 0 Then
                    Return ""
                End If

                ' 遍历内容流元素
                For Each element In page.Contents.Elements
                    Try
                        ' 将元素转换为字符串进行分析
                        Dim elementStr As String = element.ToString()

                        ' 尝试提取文本操作符中的内容
                        ' PDF 文本操作符包括：Tj (字符串), TJ (字符串数组), ' (下一行), " (下一行带间距)
                        If elementStr.Contains("(") AndAlso elementStr.Contains(")") Then
                            ' 简单提取括号内的文本内容
                            Dim textContent As String = ExtractTextFromPdfOperator(elementStr)
                            If Not String.IsNullOrEmpty(textContent) Then
                                extractedText.Append(textContent).Append(" ")
                            End If
                        End If
                    Catch ex As Exception  '  忽略单个元素的解析错误，继续处理下一个
                    End Try
                Next

            Catch ex As Exception
                Debug.WriteLine(String.Format("[PdfDocumentParser] 内容流解析异常: {0}", ex.Message))
            End Try

            Return extractedText.ToString()
        End Function

        ''' <summary>
        ''' 从 PDF 文本操作符字符串中提取文本内容
        ''' 处理常见的 PDF 文本表示格式
        ''' </summary>
        Private Function ExtractTextFromPdfOperator(operatorStr As String) As String
            Dim result As String = ""

            Try
                ' 提取圆括号内的文本（PDF 字符串格式）
                Dim startIndex As Integer = operatorStr.IndexOf("("c)
                If startIndex >= 0 Then
                    Dim endIndex As Integer = operatorStr.LastIndexOf(")"c)
                    If endIndex > startIndex Then
                        result = operatorStr.Substring(startIndex + 1, endIndex - startIndex - 1)

                        ' 解码 PDF 字符串转义序列
                        result = DecodePdfString(result)
                    End If
                End If

            Catch ex As Exception  ' 
                result = ""
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 解码 PDF 字符串中的转义序列
        ''' 处理常见的转义字符如 \n, \r, \t, \(, \), \\
        ''' </summary>
        Private Function DecodePdfString(encoded As String) As String
            If String.IsNullOrEmpty(encoded) Then Return ""

            Dim decoded As String = encoded

            ' 替换常见转义序列
            decoded = decoded.Replace("\n", Environment.NewLine)
            decoded = decoded.Replace("\r", "")
            decoded = decoded.Replace("\t", vbTab)
            decoded = decoded.Replace("\(", "(")
            decoded = decoded.Replace("\)", ")")
            decoded = decoded.Replace("\\", "\")

            Return decoded
        End Function

        ''' <summary>
        ''' 从 PDF 文档内部提取完整的元数据信息
        ''' </summary>
        Private Function ExtractMetadataInternal(document As PdfDocument) As DocumentMetadata
            Dim metadata As New DocumentMetadata()
            metadata.Properties = New Dictionary(Of String, String)()

            Try
                ' 基本信息
                metadata.Title = If(document.Info.Title, "")
                metadata.Author = If(document.Info.Author, "")
                metadata.Subject = If(document.Info.Subject, "")

                ' 日期信息：P5 Option Strict On 关键修复
                ' 注意：PdfSharp 的 PdfDocumentInformation.CreationDate/ModificationDate 本身就是 Date(DateTime) 类型！
                ' 之前错误地把它们当 String 用（String.IsNullOrEmpty/DateTime.TryParse），造成 6 处 Date→String 隐式错误。
                ' 正确做法：直接判断 Date != DateTime.MinValue（默认值=未设置），同时写 DateTime? 属性 + Properties字符串字典
                Dim minDate As DateTime = DateTime.MinValue

                If document.Info.CreationDate <> minDate Then
                    ' 显式 CType 转 DateTime?（防 Strict On 隐式）
                    metadata.CreatedDate = CType(document.Info.CreationDate, DateTime?)
                    metadata.Properties("CreatedDate") = document.Info.CreationDate.ToString("yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture)
                ElseIf Not String.IsNullOrEmpty(document.Info.CreationDate.ToString()) AndAlso document.Info.CreationDate.ToString().Length > 4 Then
                    ' 某些旧PDF CreationDate返回String但VB暴露为Date；兜底：直接把格式化后的值写字典
                    metadata.Properties("CreatedDate") = document.Info.CreationDate.ToString("yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture)
                End If

                If document.Info.ModificationDate <> minDate Then
                    metadata.ModifiedDate = CType(document.Info.ModificationDate, DateTime?)
                    metadata.Properties("ModifiedDate") = document.Info.ModificationDate.ToString("yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture)
                ElseIf Not String.IsNullOrEmpty(document.Info.ModificationDate.ToString()) AndAlso document.Info.ModificationDate.ToString().Length > 4 Then
                    metadata.Properties("ModifiedDate") = document.Info.ModificationDate.ToString("yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture)
                End If

                ' 页数
                metadata.PageCount = document.PageCount

                ' 其他可用属性
                If Not String.IsNullOrEmpty(document.Info.Keywords) Then
                    metadata.Properties.Add("Keywords", document.Info.Keywords)
                End If
                If Not String.IsNullOrEmpty(document.Info.Producer) Then
                    metadata.Properties.Add("Producer", document.Info.Producer)
                End If
                If Not String.IsNullOrEmpty(document.Info.Creator) Then
                    metadata.Properties.Add("Creator", document.Info.Creator)
                End If

                Debug.WriteLine(String.Format("[PdfDocumentParser] 元数据 - 标题: {0}, 作者: {1}, 页数: {2}",
                                              metadata.Title, metadata.Author, metadata.PageCount))

            Catch ex As Exception
                Debug.WriteLine(String.Format("[PdfDocumentParser] 元数据提取异常: {0}", ex.Message))
            End Try

            Return metadata
        End Function

        ''' <summary>
        ''' 从 PDF 书签大纲（Outlines）中提取文档结构
        ''' 将书签转换为层级化的 HeadingNode 列表
        ''' </summary>
        Private Function ExtractBookmarksFromOutline(outlines As PdfOutlineCollection) As DocumentStructure
            Dim structureInfo As New DocumentStructure()
            structureInfo.Headings = New List(Of HeadingNode)()
            structureInfo.Tables = New List(Of TableInfo)()
            structureInfo.Images = New List(Of ImageInfo)()

            Try
                If outlines Is Nothing OrElse outlines.Count = 0 Then
                    Debug.WriteLine("[PdfDocumentParser] PDF 无书签/目录信息")
                    Return structureInfo
                End If

                structureInfo.HasTableOfContents = True

                ' 递归提取书签节点
                Dim position As Long = 0
                ExtractOutlineNodesRecursive(outlines, structureInfo.Headings, 1, position)

                Debug.WriteLine(String.Format("[PdfDocumentParser] 提取到 {0} 个书签节点", structureInfo.Headings.Count))

            Catch ex As Exception
                Debug.WriteLine(String.Format("[PdfDocumentParser] 书签提取异常: {0}", ex.Message))
            End Try

            Return structureInfo
        End Function

        ''' <summary>
        ''' 递归提取书签大纲节点
        ''' </summary>
        Private Sub ExtractOutlineNodesRecursive(outlines As PdfOutlineCollection,
                                                 headings As List(Of HeadingNode),
                                                 level As Integer,
                                                 ByRef currentPosition As Long)
            If outlines Is Nothing Then Return

            For Each outline In outlines
                Try
                    Dim heading As New HeadingNode()
                    heading.Level = level
                    heading.Title = If(outline.Title, String.Format("书签 {0}", headings.Count + 1))
                    heading.StartPosition = currentPosition
                    heading.EndPosition = currentPosition + heading.Title.Length
                    heading.Children = New List(Of HeadingNode)()

                    headings.Add(heading)
                    currentPosition = heading.EndPosition + 1

                    ' 递归处理子书签
                    If outline.Outlines IsNot Nothing AndAlso outline.Outlines.Count > 0 Then
                        ExtractOutlineNodesRecursive(outline.Outlines, heading.Children, level + 1, currentPosition)
                    End If

                Catch ex As Exception
                    Debug.WriteLine(String.Format("[PdfDocumentParser] 单个书签处理异常: {0}", ex.Message))
                End Try
            Next
        End Sub

        ' ==================== MinerU 集成 ====================

        ''' <summary>
        ''' 尝试使用 MinerU 进行 PDF 高级解析
        ''' 返回 Nothing 表示 MinerU 不可用，应降级到 PdfSharp
        ''' </summary>
        Private Async Function TryMinerUParseAsync(base64Data As String) As Task(Of UnifiedDocumentResult)
            Try
                Dim mineruService As Services.MinerUService = Services.MinerUService.Instance
                Dim available = Await mineruService.IsAvailableAsync()
                If Not available Then
                    Debug.WriteLine("[PdfDocumentParser] MinerU 不可用，使用 PdfSharp 降级")
                    Return Nothing
                End If

                Debug.WriteLine("[PdfDocumentParser] 使用 MinerU 进行高级解析...")
                Dim result = Await mineruService.ParsePdfAsync(base64Data, extractImages:=True, fallback:=True)

                If Not result.Success Then
                    Debug.WriteLine($"[PdfDocumentParser] MinerU 解析失败: {result.Error}，降级到 PdfSharp")
                    Return Nothing
                End If

                ' 将 MinerU 结果转换为 UnifiedDocumentResult
                Dim docResult As New UnifiedDocumentResult()
                docResult.FileType = "pdf"
                docResult.Success = True
                docResult.FullText = If(result.Markdown, result.FullText)
                docResult.TotalChars = CLng(docResult.FullText.Length)
                docResult.TotalPages = result.PageCount
                docResult.ParseTimeMs = result.ParseTimeMs
                docResult.IsScanned = False

                ' 构建分页内容
                docResult.Pages = New List(Of PageContent)()
                If result.Pages IsNot Nothing Then
                    Dim currentPos As Long = 0
                    For Each pg In result.Pages
                        Dim pc As New PageContent()
                        Dim maybeIdx As Integer? = pg("page_number")?.ToObject(Of Integer)()
                        pc.Index = If(maybeIdx.HasValue, maybeIdx.Value, 0)
                        pc.Title = $"第 {pc.Index} 页"
                        pc.Text = pg("text")?.ToString()
                        Dim cc As Integer = If(pc.Text?.Length, 0)
                        If cc < 0 Then cc = 0
                        pc.CharCount = cc
                        pc.StartPosition = currentPos
                        docResult.Pages.Add(pc)
                        currentPos += CLng(cc)
                    Next
                End If

                ' 构建文档结构（表格、图片信息）
                docResult.Structure = New DocumentStructure()
                docResult.Structure.Headings = New List(Of HeadingNode)()
                docResult.Structure.Tables = New List(Of TableInfo)()
                docResult.Structure.Images = New List(Of ImageInfo)()

                If result.Tables IsNot Nothing Then
                    For Each tbl In result.Tables
                        Dim ti As New TableInfo()
                        ti.RowCount = tbl.RowCount
                        ti.ColCount = tbl.ColCount
                        ti.Summary = If(tbl.Markdown, tbl.Html)
                        docResult.Structure.Tables.Add(ti)
                    Next
                End If

                If result.Images IsNot Nothing Then
                    For Each img In result.Images
                        Dim ii As New ImageInfo()
                        ii.AltText = img.AltText
                        ii.Width = img.Width
                        ii.Height = img.Height
                        docResult.Structure.Images.Add(ii)
                    Next
                End If

                docResult.Structure.SectionCount = 0
                docResult.Metadata = New DocumentMetadata()
                docResult.Metadata.PageCount = result.PageCount
                docResult.Metadata.FileSizeBytes = CLng(Math.Truncate(CDbl(base64Data.Length) * 3.0 / 4.0))

                ' 将公式信息存入 Metadata.Properties
                If result.Formulas IsNot Nothing AndAlso result.Formulas.Count > 0 Then
                    docResult.Metadata.Properties("formula_count") = result.Formulas.Count.ToString()
                    docResult.Metadata.Properties("engine") = result.Engine
                    Debug.WriteLine($"[PdfDocumentParser] MinerU 解析完成 - 表格: {result.Tables?.Count}, 图片: {result.Images?.Count}, 公式: {result.Formulas?.Count}")
                End If

                Return docResult

            Catch ex As Exception
                Debug.WriteLine($"[PdfDocumentParser] MinerU 集成异常: {ex.Message}，降级到 PdfSharp")
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 强制使用 MinerU 解析 PDF（不检查服务健康状态，用于扫描件 OCR fallback）
        ''' </summary>
        Private Async Function ForceMinerUParseAsync(base64Data As String) As Task(Of UnifiedDocumentResult)
            Try
                Dim mineruService As Services.MinerUService = Services.MinerUService.Instance
                Debug.WriteLine("[PdfDocumentParser] 强制调用 MinerU 进行扫描件 OCR...")
                Dim result = Await mineruService.ParsePdfAsync(base64Data, extractImages:=True, fallback:=True)

                If result Is Nothing OrElse Not result.Success OrElse String.IsNullOrWhiteSpace(result.FullText) Then
                    Debug.WriteLine($"[PdfDocumentParser] 强制 MinerU OCR 未获得文本: {result?.Error}")
                    Return Nothing
                End If

                Debug.WriteLine($"[PdfDocumentParser] 强制 MinerU OCR 成功，字符数: {result.FullText.Length}")
                Return ConvertMinerUResultToUnifiedResult(result, base64Data)
            Catch ex As Exception
                Debug.WriteLine($"[PdfDocumentParser] 强制 MinerU OCR 异常: {ex.Message}")
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 将 MinerUParseResult 转换为 UnifiedDocumentResult
        ''' </summary>
        Private Function ConvertMinerUResultToUnifiedResult(result As Services.MinerUParseResult, base64Data As String) As UnifiedDocumentResult
            Dim docResult As New UnifiedDocumentResult()
            docResult.FileType = "pdf"
            docResult.FullText = result.FullText
            docResult.TotalChars = CLng(result.FullText.Length)
            docResult.Success = True
            docResult.IsScanned = False
            docResult.TotalPages = result.PageCount
            docResult.ParseTimeMs = result.ParseTimeMs
            docResult.Pages = New List(Of PageContent)()

            If result.Pages IsNot Nothing Then
                For Each page In result.Pages
                    Dim pc As New PageContent()
                    Dim maybeIdx As Integer? = page("index")?.ToObject(Of Integer)()
                    pc.Index = If(maybeIdx.HasValue, maybeIdx.Value, 0)
                    pc.Title = $"第 {pc.Index} 页"
                    pc.Text = page("text")?.ToString()
                    Dim cc As Integer = If(pc.Text?.Length, 0)
                    If cc < 0 Then cc = 0
                    pc.CharCount = cc
                    docResult.Pages.Add(pc)
                Next
            End If

            docResult.Structure = New DocumentStructure()
            docResult.Structure.Tables = New List(Of TableInfo)()
            docResult.Structure.Images = New List(Of ImageInfo)()

            If result.Tables IsNot Nothing Then
                For Each tbl In result.Tables
                    Dim ti As New TableInfo()
                    ti.RowCount = tbl.RowCount
                    ti.ColCount = tbl.ColCount
                    ti.Summary = If(tbl.Markdown, tbl.Html)
                    docResult.Structure.Tables.Add(ti)
                Next
            End If

            If result.Images IsNot Nothing Then
                For Each img In result.Images
                    Dim ii As New ImageInfo()
                    ii.AltText = img.AltText
                    ii.Width = img.Width
                    ii.Height = img.Height
                    docResult.Structure.Images.Add(ii)
                Next
            End If

            docResult.Metadata = New DocumentMetadata()
            docResult.Metadata.PageCount = result.PageCount
            docResult.Metadata.FileSizeBytes = CLng(Math.Truncate(CDbl(base64Data.Length) * 3.0 / 4.0))
            If result.Formulas IsNot Nothing AndAlso result.Formulas.Count > 0 Then
                docResult.Metadata.Properties("formula_count") = result.Formulas.Count.ToString()
            End If
            docResult.Metadata.Properties("engine") = result.Engine

            Return docResult
        End Function

    End Class

End Namespace