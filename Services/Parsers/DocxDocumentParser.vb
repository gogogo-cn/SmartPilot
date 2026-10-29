Option Strict On
Option Explicit On
Option Infer On
Imports System
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports NPOI.XWPF.UserModel
Imports System.Diagnostics

Namespace Services.Parsers

    ''' <summary>
    ''' Word 文档解析器 - 支持 .docx 格式（.doc 格式降级返回提示）
    ''' 使用 NPOI XWPF 解析。P5 加固：硬文件字节上限；MemoryStream 容量ctor；Strict On。
    ''' </summary>
    Public Class DocxDocumentParser
        Implements IDocumentParser

        ' —— P5：硬性安全上限（防OOM / 大对象堆LOH）——
        Private Const MAX_FILE_SIZE_BYTES As Long = 512L * 1024L * 1024L   ' 单docx 512MB 硬上限
        Private Const MAX_PARAGRAPHS As Integer = 200000                   ' 段落数上限
        Private Const MAX_FULLTEXT_CHARS As Long = 30L * 1024L * 1024L     ' 输出文本总字符 30MB
        Private Const MAX_PARAGRAPH_CHARS As Integer = 200000              ' 单段字符上限（防段内爆炸）

#Region "属性实现"

        ''' <summary>支持的文件扩展名</summary>
        Public ReadOnly Property SupportedExtensions As List(Of String) Implements IDocumentParser.SupportedExtensions
            Get
                Return New List(Of String) From {"docx", "doc"}
            End Get
        End Property

        ''' <summary>解析器显示名称</summary>
        Public ReadOnly Property DisplayName As String Implements IDocumentParser.DisplayName
            Get
                Return "Word 文档解析器"
            End Get
        End Property

#End Region

#Region "公共方法"

        ''' <summary>提取全文文本 + 基础结构信息</summary>
        Public Async Function ExtractFullText(base64Data As String) As Task(Of UnifiedDocumentResult) Implements IDocumentParser.ExtractFullText
            Return Await Task.Run(Function()
                Dim sw = Stopwatch.StartNew()
                Debug.WriteLine("[DocxDocumentParser] 开始提取全文...")

                Dim result = New UnifiedDocumentResult()

                Try
                    Dim fileBytes = Convert.FromBase64String(base64Data)

                    ' —— P5：文件大小硬上限 （防止FromBase64和内存流爆GC大对象堆）——
                    If CLng(fileBytes.LongLength) > MAX_FILE_SIZE_BYTES Then
                        Dim sizeMB As Long = CLng(fileBytes.LongLength) \ (1024L * 1024L)
                        result.Success = False
                        result.ErrorMessage = "DOCX文件大小超过安全上限（" & (MAX_FILE_SIZE_BYTES \ (1024L * 1024L)).ToString() & "MB）。当前约 " & sizeMB.ToString() & "MB。"
                        Debug.WriteLine("[DocxDocumentParser] 超过MAX_FILE_SIZE_BYTES，拒绝解析")
                        sw.Stop()
                        result.ParseTimeMs = sw.ElapsedMilliseconds
                        Return result
                    End If

                    ' P5：用(byte[],idx,count,writable,visible)包装，避免重分配内部buf
                    Using memoryStream As New MemoryStream(fileBytes, 0, fileBytes.Length, False, True)

                        If IsDocxFormat(fileBytes) Then
                            result = ParseDocxFile(memoryStream, fileBytes.Length)
                        Else
                            result = CreateDocFallbackResult(fileBytes.Length)
                        End If
                    End Using

                    result.Success = True
                    Debug.WriteLine($"[DocxDocumentParser] 全文提取成功，总字符数: {result.TotalChars}")

                Catch ex As Exception
                    result.Success = False
                    result.ErrorMessage = $"Word文档解析失败: {ex.Message}"
                    Debug.WriteLine($"[DocxDocumentParser] 全文提取失败: {ex.Message}")
                Finally
                    sw.Stop()
                    result.ParseTimeMs = sw.ElapsedMilliseconds
                End Try

                Return result
            End Function)
        End Function

        ''' <summary>提取文档结构 - 识别标题样式构建标题树</summary>
        Public Async Function ExtractStructure(base64Data As String) As Task(Of DocumentStructure) Implements IDocumentParser.ExtractStructure
            Return Await Task.Run(Of DocumentStructure)(Function()
                Debug.WriteLine("[DocxDocumentParser] 开始提取文档结构...")
                Dim structureInfo = New DocumentStructure()

                Try
                    Dim fileBytes = Convert.FromBase64String(base64Data)
                    ' P5：大小硬上限（结构提取对超大docx没必要走）
                    If CLng(fileBytes.LongLength) > MAX_FILE_SIZE_BYTES Then
                        Debug.WriteLine("[DocxDocumentParser] ExtractStructure: 文件过大，直接返回空结构")
                        Return New DocumentStructure()
                    End If
                    Using memoryStream As New MemoryStream(fileBytes, 0, fileBytes.Length, False, True)

                        If IsDocxFormat(fileBytes) Then
                            structureInfo = ExtractDocxStructure(memoryStream)
                        Else
                            ' .doc 返回空结构
                            structureInfo = New DocumentStructure()
                        End If
                    End Using

                    Debug.WriteLine($"[DocxDocumentParser] 结构提取完成，章节数: {structureInfo.SectionCount}")

                Catch ex As Exception
                    Debug.WriteLine($"[DocxDocumentParser] 结构提取失败: {ex.Message}")
                End Try

                Return structureInfo
            End Function)
        End Function

        ''' <summary>提取文档元数据</summary>
        Public Function ExtractMetadata(base64Data As String) As DocumentMetadata Implements IDocumentParser.ExtractMetadata
            Debug.WriteLine("[DocxDocumentParser] 开始提取元数据...")
            Dim metadata = New DocumentMetadata()

            Try
                Dim fileBytes = Convert.FromBase64String(base64Data)
                metadata.FileSizeBytes = fileBytes.Length

                ' P5：元数据阶段也过大小上限；超上限则不再打开docx
                If CLng(fileBytes.LongLength) > MAX_FILE_SIZE_BYTES Then
                    Debug.WriteLine("[DocxDocumentParser] ExtractMetadata: 文件过大，仅返回大小元")
                    Return metadata
                End If

                Using memoryStream As New MemoryStream(fileBytes, 0, fileBytes.Length, False, True)

                    If IsDocxFormat(fileBytes) Then
                        metadata = ExtractDocxMetadata(memoryStream, metadata)
                    Else
                        metadata.Properties.Add("Format", "DOC (Word 97-2003) - 有限支持")
                    End If
                End Using

                Debug.WriteLine($"[DocxDocumentParser] 元数据提取完成，标题: {metadata.Title}")

            Catch ex As Exception
                Debug.WriteLine($"[DocxDocumentParser] 元数据提取失败: {ex.Message}")
            End Try

            Return metadata
        End Function

        ''' <summary>提取指定范围内容</summary>
        Public Async Function ExtractRange(base64Data As String, startPosition As Long, maxLength As Integer) As Task(Of String) Implements IDocumentParser.ExtractRange
            Debug.WriteLine($"[DocxDocumentParser] 提取范围内容，起始位置: {startPosition}, 长度: {maxLength}")
            Dim rangeText = ""

            Try
                Dim fullResult = Await ExtractFullText(base64Data)

                If fullResult.Success AndAlso fullResult.FullText.Length > startPosition Then
                    Dim actualLength = Math.Min(maxLength, fullResult.FullText.Length - startPosition)
                    rangeText = fullResult.FullText.Substring(CInt(startPosition), CInt(actualLength))
                End If

            Catch ex As Exception
                Debug.WriteLine($"[DocxDocumentParser] 范围提取失败: {ex.Message}")
                rangeText = ""
            End Try

            Return rangeText
        End Function

#End Region

#Region "DOCX 格式解析"

        ''' <summary>解析 DOCX 文件</summary>
        Private Function ParseDocxFile(stream As Stream, originalSize As Integer) As UnifiedDocumentResult
            Dim result = New UnifiedDocumentResult()
            result.FileType = "docx"

            Using document = New XWPFDocument(stream)
                ' P5：按估计初始容量，减少 StringBuilder 多次 realloc
                Dim initialCap As Integer = CInt(Math.Min(CLng(originalSize) * 3L, MAX_FULLTEXT_CHARS))
                If initialCap < 4096 Then initialCap = 4096
                Dim fullTextBuilder As New StringBuilder(initialCap)
                Dim pages = New List(Of PageContent)()
                Dim fullTextTruncated As Boolean = False

                ' 提取所有段落文本（P5：段落数 + 单段字符 + 全文字符三重保护）
                Dim paragraphCount As Integer = 0
                For Each paragraph As XWPFParagraph In document.Paragraphs
                    If fullTextTruncated Then Exit For
                    paragraphCount += 1
                    If paragraphCount > MAX_PARAGRAPHS Then
                        AppendLineSafe(fullTextBuilder, vbCrLf & "[警告: 段落超过上限，后续段落未提取]", CInt(MAX_FULLTEXT_CHARS), fullTextTruncated)
                        Exit For
                    End If

                    Dim paragraphText = If(paragraph.Text, "").Trim()
                    If String.IsNullOrEmpty(paragraphText) Then Continue For

                    ' P5：单段字符上限
                    If paragraphText.Length > MAX_PARAGRAPH_CHARS Then
                        paragraphText = paragraphText.Substring(0, MAX_PARAGRAPH_CHARS) & "...[单段内容过长已截断]"
                    End If

                    Dim styleName = If(paragraph.Style, "")
                    Dim decorated As String
                    If Not String.IsNullOrEmpty(styleName) AndAlso styleName.ToLower().Contains("heading") Then
                        decorated = $"【{styleName}】{paragraphText}"
                    Else
                        decorated = paragraphText
                    End If

                    ' P5：全文字符上限（AppendLineSafe帮做截断）
                    AppendLineSafe(fullTextBuilder, decorated, CInt(MAX_FULLTEXT_CHARS), fullTextTruncated)
                Next

                ' 提取表格文本（P5：段落/表格总上限叠加）
                Dim tableIndex As Integer = 0
                For Each table As XWPFTable In document.Tables
                    If fullTextTruncated Then Exit For
                    tableIndex += 1
                    AppendLineSafe(fullTextBuilder, vbCrLf & $"【表格 {tableIndex}】", CInt(MAX_FULLTEXT_CHARS), fullTextTruncated)

                    For rowIndex As Integer = 0 To table.Rows.Count - 1
                        If fullTextTruncated Then Exit For
                        Dim row = table.Rows(rowIndex)
                        Dim cellTexts = New List(Of String)()

                        For Each cell As XWPFTableCell In row.GetTableCells()
                            Dim txt = If(cell.GetText(), "").Trim()
                            If txt.Length > MAX_PARAGRAPH_CHARS \ 2 Then txt = txt.Substring(0, MAX_PARAGRAPH_CHARS \ 2) & "..."
                            cellTexts.Add(txt)
                        Next

                        AppendLineSafe(fullTextBuilder, String.Join(" | ", cellTexts), CInt(MAX_FULLTEXT_CHARS), fullTextTruncated)
                    Next

                    AppendLineSafe(fullTextBuilder, "", CInt(MAX_FULLTEXT_CHARS), fullTextTruncated)
                Next

                If fullTextTruncated Then
                    Debug.WriteLine("[DocxDocumentParser] 已触发 MAX_FULLTEXT_CHARS 上限 = " & MAX_FULLTEXT_CHARS.ToString())
                End If

                ' 构建结果
                Dim fullText = fullTextBuilder.ToString()
                result.FullText = fullText
                result.TotalChars = fullText.Length

                pages.Add(New PageContent With {
                    .Index = 1,
                    .Title = "Word 文档",
                    .Text = fullText,
                    .CharCount = fullText.Length,
                    .StartPosition = 0
                })
                result.Pages = pages
                result.TotalPages = pages.Count
                result.Structure = ExtractDocxStructureFromDoc(document)
                result.Metadata = ExtractDocxMetadataFromDoc(document)
                result.Metadata.FileSizeBytes = originalSize
            End Using

            Return result
        End Function

        ''' <summary>从 XWPFDocument 提取结构</summary>
        Private Function ExtractDocxStructure(stream As Stream) As DocumentStructure
            Using document = New XWPFDocument(stream)
                Return ExtractDocxStructureFromDoc(document)
            End Using
        End Function

        ''' <summary>从 XWPFDocument 提取标题树</summary>
        Private Function ExtractDocxStructureFromDoc(document As XWPFDocument) As DocumentStructure
            Dim structureInfo = New DocumentStructure()
            structureInfo.Headings = New List(Of HeadingNode)()
            structureInfo.Tables = New List(Of TableInfo)()
            structureInfo.Images = New List(Of ImageInfo)()

            Dim currentPosition As Long = 0

            For Each paragraph As XWPFParagraph In document.Paragraphs
                Dim paragraphText = paragraph.Text.Trim()
                Dim styleName = If(paragraph.Style, "").ToLower()

                Dim headingLevel As Integer = 0
                If styleName.Contains("heading1") OrElse styleName.Contains("title") Then
                    headingLevel = 1
                ElseIf styleName.Contains("heading2") Then
                    headingLevel = 2
                ElseIf styleName.Contains("heading3") Then
                    headingLevel = 3
                ElseIf styleName.Contains("heading4") Then
                    headingLevel = 4
                End If

                If headingLevel > 0 AndAlso Not String.IsNullOrEmpty(paragraphText) Then
                    structureInfo.Headings.Add(New HeadingNode With {
                        .Level = headingLevel,
                        .Title = paragraphText,
                        .StartPosition = currentPosition,
                        .Children = New List(Of HeadingNode)()
                    })
                    structureInfo.SectionCount += 1
                End If

                currentPosition += paragraphText.Length + Environment.NewLine.Length
            Next

            ' 表格信息
            For Each table As XWPFTable In document.Tables
                Dim colCount = If(table.Rows.Count > 0, table.Rows(0).GetTableCells().Count, 0)
                structureInfo.Tables.Add(New TableInfo With {
                    .RowCount = table.Rows.Count,
                    .ColCount = colCount,
                    .Summary = $"表格 ({table.Rows.Count}行 x {colCount}列)"
                })
            Next

            structureInfo.HasTableOfContents = structureInfo.Headings.Count > 3
            Return structureInfo
        End Function

        ''' <summary>从 XWPFDocument 提取元数据</summary>
        Private Function ExtractDocxMetadata(stream As Stream, existingMetadata As DocumentMetadata) As DocumentMetadata
            Using document = New XWPFDocument(stream)
                Return ExtractDocxMetadataFromDoc(document, existingMetadata)
            End Using
        End Function

        ''' <summary>从 XWPFDocument 对象提取属性</summary>
        Private Function ExtractDocxMetadataFromDoc(document As XWPFDocument, Optional existingMetadata As DocumentMetadata = Nothing) As DocumentMetadata
            Dim metadata = If(existingMetadata, New DocumentMetadata())

            Try
                ' NPOI XWPFDocument 属性 API 版本差异大，暂时跳过
                Debug.WriteLine("[DocxDocumentParser] 文档属性提取已跳过（NPOI 版本兼容）")

                metadata.Properties.Add("Format", "DOCX (Office Open XML)")
                metadata.Properties.Add("ParagraphCount", document.Paragraphs.Count.ToString())
                metadata.Properties.Add("TableCount", document.Tables.Count.ToString())

            Catch ex As Exception
                Debug.WriteLine($"[DocxDocumentParser] 读取 DOCX 属性时出错: {ex.Message}")
            End Try

            Return metadata
        End Function

#End Region

#Region "DOC 格式降级处理"

        ''' <summary>.doc 格式不支持时的降级返回</summary>
        Private Function CreateDocFallbackResult(fileSize As Integer) As UnifiedDocumentResult
            Dim result = New UnifiedDocumentResult With {
                .Success = True,
                .FileType = "doc",
                .FullText = "[DOC 格式] 此为旧版 Word 97-2003 格式。建议将文件另存为 .docx 格式后重新上传以获得完整支持。",
                .TotalChars = 80,
                .TotalPages = 1,
                .Pages = New List(Of PageContent) From {
                    New PageContent With {.Index = 1, .Title = "Word 97-2003", .Text = "格式不支持", .CharCount = 4}
                },
                .Structure = New DocumentStructure(),
                .Metadata = New DocumentMetadata() With {
                    .FileSizeBytes = fileSize,
                    .Properties = New Dictionary(Of String, String) From {
                        {"Format", "DOC (Word 97-2003)"},
                        {"Note", "HWPF 未安装或未引用"}
                    }
                }
            }
            Return result
        End Function

#End Region

#Region "辅助方法"

        ''' <summary>判断是否为 DOCX 格式（PK 头 = ZIP）</summary>
        Private Function IsDocxFormat(fileBytes As Byte()) As Boolean
            If fileBytes Is Nothing OrElse fileBytes.Length < 4 Then Return False
            Return fileBytes(0) = &H50 AndAlso fileBytes(1) = &H4B
        End Function

        ''' <summary>
        ''' P5：安全的 AppendLine 包装（到字符上限后只写最后一次可容纳部分，capHit后立即静默丢弃）
        ''' </summary>
        Private Shared Sub AppendLineSafe(sb As StringBuilder, line As String, maxChars As Integer, ByRef capHit As Boolean)
            If sb Is Nothing Then Return
            If capHit Then Return

            Dim addChars As Integer = If(line, "").Length + Environment.NewLine.Length
            If sb.Length + addChars > maxChars Then
                capHit = True
                Dim remaining As Integer = Math.Max(0, maxChars - sb.Length)
                If remaining > 0 AndAlso line IsNot Nothing AndAlso remaining <= line.Length Then
                    sb.Append(line.Substring(0, remaining))
                ElseIf remaining > Environment.NewLine.Length Then
                    sb.Append(If(line, ""))
                    sb.AppendLine()
                End If
                Return
            End If
            sb.AppendLine(line)
        End Sub

#End Region

    End Class

End Namespace