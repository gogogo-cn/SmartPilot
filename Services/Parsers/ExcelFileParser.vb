Option Strict On
Option Explicit On
Option Infer On
Imports System
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports NPOI.XSSF.UserModel
Imports NPOI.HSSF.UserModel
Imports NPOI.SS.UserModel
Imports System.Diagnostics

Namespace Services.Parsers

    ''' <summary>
    ''' Excel 文件解析器 - 支持 .xlsx 和 .xls 格式（NPOI）
    ''' P5 加固：硬文件大小上限、全文/单Sheet StringBuilder字符上限、MemoryStream容量声明。
    ''' </summary>
    Public Class ExcelFileParser
        Implements IDocumentParser

        ' —— P5：硬性安全上限（防 OOM/大对象堆爆炸）——
        Private Const MAX_FILE_SIZE_BYTES As Long = 1024L * 1024L * 1024L        ' 1 GB：单文件绝对硬上限
        Private Const MAX_TOTAL_BUILDER_CHARS As Long = 50L * 1024L * 1024L      ' 50 MB 约 1亿汉字（全文总输出）
        Private Const MAX_SHEET_BUILDER_CHARS As Integer = 10 * 1024 * 1024      ' 单个 Sheet 10MB 单独上限

        Private Const MAX_ROWS_PER_SHEET As Integer = 50000
        Private Const MAX_CELLS_PER_SHEET As Integer = 500000
        Private Const MAX_SHEETS As Integer = 200
        Private Const MAX_CELL_TEXT_LENGTH As Integer = 32767

#Region "属性实现"

        ''' <summary>
        ''' 支持的文件扩展名
        ''' </summary>
        Public ReadOnly Property SupportedExtensions As List(Of String) Implements IDocumentParser.SupportedExtensions
            Get
                Return New List(Of String) From {"xlsx", "xls"}
            End Get
        End Property

        ''' <summary>
        ''' 解析器显示名称
        ''' </summary>
        Public ReadOnly Property DisplayName As String Implements IDocumentParser.DisplayName
            Get
                Return "Excel 文档解析器"
            End Get
        End Property

#End Region

#region "公共方法"

        ''' <summary>
        ''' 提取全文文本 - 遍历所有 Sheet 的单元格内容
        ''' 每个 Sheet 作为独立的 PageContent
        ''' </summary>
        Public Async Function ExtractFullText(base64Data As String) As Task(Of UnifiedDocumentResult) Implements IDocumentParser.ExtractFullText
            Dim sw = Stopwatch.StartNew()
            Debug.WriteLine("[ExcelFileParser] 开始提取全文...")

            Dim result = New UnifiedDocumentResult()

            If String.IsNullOrWhiteSpace(base64Data) Then
                result.Success = False
                result.ErrorMessage = "传入的文件数据为空"
                sw.Stop()
                result.ParseTimeMs = sw.ElapsedMilliseconds
                Return result
            End If

            Dim fileBytes As Byte() = Nothing
            Try
                fileBytes = Convert.FromBase64String(base64Data.Trim())
            Catch exFmt As FormatException
                result.Success = False
                result.ErrorMessage = "传入的文件不是合法的 Base64 编码: " & exFmt.Message
                Debug.WriteLine("[ExcelFileParser] Base64 解码失败: " & exFmt.Message)
                sw.Stop()
                result.ParseTimeMs = sw.ElapsedMilliseconds
                Return result
            End Try

            Try
                Debug.WriteLine("[ExcelFileParser] Base64长度=" & base64Data.Length.ToString() & ", 字节数组长度=" & fileBytes.Length.ToString())
                If fileBytes Is Nothing OrElse fileBytes.Length < 4 Then
                    result.Success = False
                    result.ErrorMessage = "文件数据太小，不是有效的 Excel 文件"
                    sw.Stop()
                    result.ParseTimeMs = sw.ElapsedMilliseconds
                    Return result
                End If

                ' —— P5：单文件字节硬上限检查（防超大base64直接撑爆GC大对象堆）——
                If CLng(fileBytes.LongLength) > MAX_FILE_SIZE_BYTES Then
                    Dim sizeMB As Long = CLng(fileBytes.LongLength) \ (1024L * 1024L)
                    result.Success = False
                    result.ErrorMessage = "文件大小超过安全上限（" & (MAX_FILE_SIZE_BYTES \ (1024L * 1024L)).ToString() & "MB）。当前约 " & sizeMB.ToString() & "MB。"
                    Debug.WriteLine("[ExcelFileParser] 超过MAX_FILE_SIZE_BYTES上限, 拒绝解析")
                    sw.Stop()
                    result.ParseTimeMs = sw.ElapsedMilliseconds
                    Return result
                End If

                Debug.WriteLine("[ExcelFileParser] 文件头字节: " & fileBytes(0).ToString("X2") & " " & fileBytes(1).ToString("X2") & " " & fileBytes(2).ToString("X2") & " " & fileBytes(3).ToString("X2"))
                ' P5：使用带 capacity 参数的 MemoryStream 构造器（避免多次扩容；再包一层保护上限）
                Dim safeCapacity As Integer = CInt(Math.Min(CLng(fileBytes.Length) + 1024L, MAX_FILE_SIZE_BYTES))
                Using memoryStream As New MemoryStream(fileBytes, 0, fileBytes.Length, False, True)
                    ' （上述 capacity 参数等价写法：MemoryStream 用 byte[] 包装，不再重新分配）
                    Debug.WriteLine("[ExcelFileParser] Stream位置=" & memoryStream.Position.ToString() & ", 长度=" & memoryStream.Length.ToString())

                    If IsXlsxFormat(fileBytes) Then
                        Debug.WriteLine("[ExcelFileParser] 检测为 XLSX 格式")
                        result = Await ParseXlsxFile(memoryStream, fileBytes.Length)
                    Else
                        Debug.WriteLine("[ExcelFileParser] 检测为 XLS 格式")
                        result = Await ParseXlsFile(memoryStream, fileBytes.Length)
                    End If
                End Using

                result.Success = True
                Debug.WriteLine("[ExcelFileParser] 全文提取成功，总字符数: " & result.TotalChars.ToString() & ", Sheet数: " & result.TotalPages.ToString())

            Catch ex As Exception
                result.Success = False
                result.ErrorMessage = "Excel文档解析失败: " & ex.Message
                Debug.WriteLine("[ExcelFileParser] 全文提取失败: " & ex.Message)
            Finally
                sw.Stop()
                result.ParseTimeMs = sw.ElapsedMilliseconds
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 提取文档结构 - 列出所有 Sheet 名称和行列数
        ''' </summary>
        Public Async Function ExtractStructure(base64Data As String) As Task(Of DocumentStructure) Implements IDocumentParser.ExtractStructure
            Return Await Task.Run(Of DocumentStructure)(Function()
                Debug.WriteLine("[ExcelFileParser] 开始提取文档结构...")
                Dim structureInfo = New DocumentStructure()

                Try
                    Dim fileBytes = Convert.FromBase64String(base64Data)
                    Using memoryStream = New MemoryStream(fileBytes)

                        If IsXlsxFormat(fileBytes) Then
                            structureInfo = ExtractXlsxStructure(memoryStream)
                        Else
                            structureInfo = ExtractXlsStructure(memoryStream)
                        End If
                    End Using

                    Debug.WriteLine($"[ExcelFileParser] 结构提取完成，Sheet数: {structureInfo.SectionCount}")

                Catch ex As Exception
                    Debug.WriteLine($"[ExcelFileParser] 结构提取失败: {ex.Message}")
                End Try

                Return structureInfo
            End Function)
        End Function

        ''' <summary>
        ''' 提取文档元数据 - 工作簿属性信息
        ''' </summary>
        Public Function ExtractMetadata(base64Data As String) As DocumentMetadata Implements IDocumentParser.ExtractMetadata
            Debug.WriteLine("[ExcelFileParser] 开始提取元数据...")
            Dim metadata = New DocumentMetadata()

            Try
                Dim fileBytes = Convert.FromBase64String(base64Data)
                metadata.FileSizeBytes = fileBytes.Length

                Using memoryStream = New MemoryStream(fileBytes)

                    If IsXlsxFormat(fileBytes) Then
                        metadata = ExtractXlsxMetadata(memoryStream, metadata)
                    Else
                        metadata = ExtractXlsMetadata(memoryStream, metadata)
                    End If
                End Using

                Debug.WriteLine($"[ExcelFileParser] 元数据提取完成，标题: {metadata.Title}")

            Catch ex As Exception
                Debug.WriteLine($"[ExcelFileParser] 元数据提取失败: {ex.Message}")
            End Try

            Return metadata
        End Function

        ''' <summary>
        ''' 提取指定范围内容 - 用于大文件分页加载
        ''' </summary>
        Public Async Function ExtractRange(base64Data As String, startPosition As Long, maxLength As Integer) As Task(Of String) Implements IDocumentParser.ExtractRange
            Debug.WriteLine($"[ExcelFileParser] 提取范围内容，起始位置: {startPosition}, 长度: {maxLength}")
            Dim rangeText = ""

            Try
                ' 先提取全文，再截取指定范围
                Dim fullResult = Await ExtractFullText(base64Data)

                If fullResult.Success AndAlso fullResult.FullText.Length > startPosition Then
                    Dim actualLength = Math.Min(maxLength, fullResult.FullText.Length - startPosition)
                    rangeText = fullResult.FullText.Substring(CInt(startPosition), CInt(actualLength))
                End If

            Catch ex As Exception
                Debug.WriteLine($"[ExcelFileParser] 范围提取失败: {ex.Message}")
                rangeText = ""
            End Try

            Return rangeText
        End Function

#End Region

#Region "XLSX 格式解析 (XSSFWorkbook)"

        ''' <summary>
        ''' 解析 XLSX 格式文件 (Excel 2007+)
        ''' </summary>
        Private Async Function ParseXlsxFile(stream As Stream, originalSize As Integer) As Task(Of UnifiedDocumentResult)
            Return Await Task.Run(Function()
                Dim result = New UnifiedDocumentResult()
                result.FileType = "xlsx"

                Using workbook = New XSSFWorkbook(stream)
                    result = BuildResultFromWorkbook(workbook, originalSize, "xlsx")
                End Using

                Return result
            End Function)
        End Function

        ''' <summary>
        ''' 从 XSSFWorkbook 提取文档结构
        ''' </summary>
        Private Function ExtractXlsxStructure(stream As Stream) As DocumentStructure
            Dim structureInfo = New DocumentStructure()

            Using workbook = New XSSFWorkbook(stream)
                structureInfo = BuildStructureFromWorkbook(workbook)
            End Using

            Return structureInfo
        End Function

        ''' <summary>
        ''' 从 XSSFWorkbook 提取元数据
        ''' </summary>
        Private Function ExtractXlsxMetadata(stream As Stream, existingMetadata As DocumentMetadata) As DocumentMetadata
            Using workbook = New XSSFWorkbook(stream)
                Return BuildMetadataFromWorkbook(workbook, existingMetadata, "xlsx")
            End Using
        End Function

#End Region

#region "XLS 格式解析 (HSSFWorkbook)"

        ''' <summary>
        ''' 解析 XLS 格式文件 (Excel 97-2003)
        ''' </summary>
        Private Async Function ParseXlsFile(stream As Stream, originalSize As Integer) As Task(Of UnifiedDocumentResult)
            Return Await Task.Run(Function()
                Dim result = New UnifiedDocumentResult()
                result.FileType = "xls"

                Using workbook = New HSSFWorkbook(stream)
                    result = BuildResultFromWorkbook(workbook, originalSize, "xls")
                End Using

                Return result
            End Function)
        End Function

        ''' <summary>
        ''' 从 HSSFWorkbook 提取文档结构
        ''' </summary>
        Private Function ExtractXlsStructure(stream As Stream) As DocumentStructure
            Dim structureInfo = New DocumentStructure()

            Using workbook = New HSSFWorkbook(stream)
                structureInfo = BuildStructureFromWorkbook(workbook)
            End Using

            Return structureInfo
        End Function

        ''' <summary>
        ''' 从 HSSFWorkbook 提取元数据
        ''' </summary>
        Private Function ExtractXlsMetadata(stream As Stream, existingMetadata As DocumentMetadata) As DocumentMetadata
            Using workbook = New HSSFWorkbook(stream)
                Return BuildMetadataFromWorkbook(workbook, existingMetadata, "xls")
            End Using
        End Function

#End Region

#region "通用处理方法（基于 IWorkbook 接口）"

        ''' <summary>
        ''' 从 IWorkbook 构建统一结果对象
        ''' 遍历所有 Sheet，每个 Sheet 提取为 PageContent
        ''' </summary>
        Private Function BuildResultFromWorkbook(workbook As IWorkbook, originalSize As Integer, fileType As String) As UnifiedDocumentResult
            Dim result = New UnifiedDocumentResult()
            result.FileType = fileType

            If workbook Is Nothing Then
                result.Success = False
                result.ErrorMessage = "Workbook 对象为空"
                Return result
            End If

            Debug.WriteLine("[ExcelFileParser] BuildResult: NumberOfSheets=" & workbook.NumberOfSheets.ToString() & ", originalSize=" & originalSize.ToString())

            ' P5：全文 StringBuilder 提前给初始容量（减少 realloc）
            Dim initialCapacity As Integer = CInt(Math.Min(CLng(originalSize) * 3L, MAX_TOTAL_BUILDER_CHARS \ 2L))
            If initialCapacity < 4096 Then initialCapacity = 4096
            Dim fullTextBuilder As New StringBuilder(initialCapacity)
            Dim pages = New List(Of PageContent)()
            Dim globalPosition As Long = 0
            Dim totalCharsHardCapHit As Boolean = False
            Dim sheetCountToProcess = Math.Min(workbook.NumberOfSheets, MAX_SHEETS)

            If workbook.NumberOfSheets > MAX_SHEETS Then
                Debug.WriteLine("[ExcelFileParser] 工作表数超过上限 " & MAX_SHEETS.ToString() & "，仅处理前 " & sheetCountToProcess.ToString() & " 个")
            End If

            For sheetIndex As Integer = 0 To sheetCountToProcess - 1
                If totalCharsHardCapHit Then Exit For
                Dim sheetNameSafe = $"Sheet{sheetIndex + 1}"
                Try
                    Dim sheet = workbook.GetSheetAt(sheetIndex)
                    If sheet Is Nothing Then
                        Debug.WriteLine("[ExcelFileParser] 跳过空 Sheet: 索引=" & sheetIndex.ToString())
                        Continue For
                    End If
                    Dim sheetName = If(sheet.SheetName, sheetNameSafe)
                    sheetNameSafe = sheetName

                    ' P5：单 Sheet StringBuilder 初始化 + 单 Sheet 独立字符上限
                    Dim sheetTextBuilder As New StringBuilder(Math.Min(originalSize \ 2 + 1024, MAX_SHEET_BUILDER_CHARS))
                    Dim sheetCapHit As Boolean = False
                    Dim cellsProcessedInSheet As Integer = 0

                    Debug.WriteLine("[ExcelFileParser] 正在处理 Sheet: " & sheetName)

                    AppendLineSafe(sheetTextBuilder, "【工作表: " & sheetName & "】", MAX_SHEET_BUILDER_CHARS, sheetCapHit)
                    AppendLineSafe(sheetTextBuilder, "", MAX_SHEET_BUILDER_CHARS, sheetCapHit)

                    Dim firstRowNum As Integer = 0
                    Dim lastRowNum As Integer = 0
                    Try
                        firstRowNum = sheet.FirstRowNum
                        lastRowNum = sheet.LastRowNum
                    Catch ex As Exception
                        Debug.WriteLine("[ExcelFileParser] 获取 Sheet 行范围失败: " & sheetName & " -> " & ex.Message)
                        AppendLineSafe(sheetTextBuilder, "[警告: 该工作表无法读取，可能已损坏]", MAX_SHEET_BUILDER_CHARS, sheetCapHit)
                        GoTo FinalizeSheetPage
                    End Try

                    If lastRowNum < firstRowNum Then
                        Debug.WriteLine("[ExcelFileParser] Sheet 无有效行: " & sheetName)
                        GoTo FinalizeSheetPage
                    End If

                    Dim effectiveLastRow = Math.Min(lastRowNum, firstRowNum + MAX_ROWS_PER_SHEET - 1)
                    If lastRowNum > effectiveLastRow Then
                        Debug.WriteLine("[ExcelFileParser] Sheet """ & sheetName & """ 行数 " & (lastRowNum - firstRowNum + 1).ToString() & " 超限，仅处理前 " & MAX_ROWS_PER_SHEET.ToString() & " 行")
                        AppendLineSafe(sheetTextBuilder, "[注意: 该工作表超过 " & MAX_ROWS_PER_SHEET.ToString() & " 行，仅提取前 " & MAX_ROWS_PER_SHEET.ToString() & " 行]", MAX_SHEET_BUILDER_CHARS, sheetCapHit)
                        AppendLineSafe(sheetTextBuilder, "", MAX_SHEET_BUILDER_CHARS, sheetCapHit)
                    End If

                    Dim headerRowTexts As List(Of String) = Nothing

                    For rowIndex As Integer = firstRowNum To effectiveLastRow
                        If sheetCapHit Then Exit For
                        Dim row As IRow = Nothing
                        Try
                            row = sheet.GetRow(rowIndex)
                        Catch ex As Exception
                            Debug.WriteLine("[ExcelFileParser] 读取行失败: " & sheetName & " R" & rowIndex.ToString() & " -> " & ex.Message)
                            Continue For
                        End Try
                        If row Is Nothing Then Continue For

                        Dim cellTexts = New List(Of String)()
                        Dim lastCellIdx As Integer = -1
                        Try
                            lastCellIdx = row.LastCellNum - 1
                        Catch
                            Continue For
                        End Try
                        If lastCellIdx < 0 Then Continue For

                        For colIndex As Integer = 0 To lastCellIdx
                            If cellsProcessedInSheet >= MAX_CELLS_PER_SHEET Then Exit For

                            Dim cell As ICell = Nothing
                            Try
                                cell = row.GetCell(colIndex)
                            Catch
                                cellTexts.Add("")
                                cellsProcessedInSheet += 1
                                Continue For
                            End Try

                            If cell Is Nothing Then
                                cellTexts.Add("")
                                cellsProcessedInSheet += 1
                                Continue For
                            End If

                            Dim cellValue As String = ""
                            Try
                                cellValue = GetCellStringValue(cell)
                            Catch ex As Exception
                                Debug.WriteLine("[ExcelFileParser] 单元格读取失败: " & sheetName & " R" & rowIndex.ToString() & "C" & colIndex.ToString() & " -> " & ex.Message)
                                cellValue = "[读取失败]"
                            End Try

                            If Not String.IsNullOrEmpty(cellValue) AndAlso cellValue.Length > MAX_CELL_TEXT_LENGTH Then
                                cellValue = cellValue.Substring(0, MAX_CELL_TEXT_LENGTH) & "..."
                            End If
                            cellTexts.Add(cellValue)
                            cellsProcessedInSheet += 1
                        Next

                        If cellsProcessedInSheet >= MAX_CELLS_PER_SHEET Then
                            AppendLineSafe(sheetTextBuilder, "", MAX_SHEET_BUILDER_CHARS, sheetCapHit)
                            AppendLineSafe(sheetTextBuilder, "[注意: 该工作表超过 " & MAX_CELLS_PER_SHEET.ToString() & " 单元格，剩余内容未提取]", MAX_SHEET_BUILDER_CHARS, sheetCapHit)
                            Exit For
                        End If

                        Dim hasContent As Boolean = False
                        Try
                            hasContent = cellTexts.Any(Function(ct) Not String.IsNullOrEmpty(If(ct, "").Trim()))
                        Catch
                            hasContent = False
                        End Try
                        If Not hasContent Then Continue For

                        Dim rowText = String.Join(vbTab, cellTexts)

                        If rowIndex = firstRowNum Then
                            Try
                                headerRowTexts = cellTexts.Where(Function(c) Not String.IsNullOrEmpty(If(c, "").Trim())).ToList()
                            Catch
                                headerRowTexts = New List(Of String)()
                            End Try
                            AppendLineSafe(sheetTextBuilder, "【表头】" & rowText, MAX_SHEET_BUILDER_CHARS, sheetCapHit)
                        Else
                            AppendLineSafe(sheetTextBuilder, rowText, MAX_SHEET_BUILDER_CHARS, sheetCapHit)
                        End If
                    Next

FinalizeSheetPage:
                    AppendLineSafe(sheetTextBuilder, "", MAX_SHEET_BUILDER_CHARS, sheetCapHit)
                    If sheetCapHit Then
                        AppendLineSafe(sheetTextBuilder, "[警告: 该工作表文本内容超过单Sheet上限，已截断]", MAX_SHEET_BUILDER_CHARS, sheetCapHit)
                    End If

                    Dim sheetText = sheetTextBuilder.ToString()

                    ' —— P5：全局 StringBuilder 全文字符硬上限 ——
                    Dim remainingGlobalCapacity As Long = MAX_TOTAL_BUILDER_CHARS - CLng(fullTextBuilder.Length)
                    If remainingGlobalCapacity <= 0 Then
                        totalCharsHardCapHit = True
                    Else
                        If sheetText.Length > remainingGlobalCapacity Then
                            sheetText = sheetText.Substring(0, CInt(remainingGlobalCapacity))
                            totalCharsHardCapHit = True
                            sheetText = sheetText & vbCrLf & vbCrLf & "[警告: 全文输出超过全局上限，后续内容已全部截断]"
                        End If
                        fullTextBuilder.Append(sheetText)
                    End If

                    Dim pageContent = New PageContent With {
                        .Index = sheetIndex + 1,
                        .Title = sheetNameSafe,
                        .Text = sheetText,
                        .CharCount = sheetText.Length,
                        .StartPosition = globalPosition
                    }
                    pages.Add(pageContent)
                    globalPosition += sheetText.Length

                Catch exSheet As Exception
                    Debug.WriteLine("[ExcelFileParser] 处理 Sheet 失败: 索引" & sheetIndex.ToString() & " -> " & exSheet.Message)
                    Dim errText = "【工作表: " & sheetNameSafe & "】" & vbCrLf & vbCrLf & "[该工作表解析失败: " & exSheet.Message & "]"
                    Dim remainingGlobalCapacity As Long = MAX_TOTAL_BUILDER_CHARS - CLng(fullTextBuilder.Length)
                    If remainingGlobalCapacity > 0 AndAlso errText.Length > remainingGlobalCapacity Then
                        errText = errText.Substring(0, CInt(remainingGlobalCapacity))
                        totalCharsHardCapHit = True
                    End If
                    If remainingGlobalCapacity > 0 Then fullTextBuilder.Append(errText)
                    Dim errPage = New PageContent With {
                        .Index = sheetIndex + 1,
                        .Title = sheetNameSafe,
                        .Text = errText,
                        .CharCount = errText.Length,
                        .StartPosition = globalPosition
                    }
                    pages.Add(errPage)
                    globalPosition += errText.Length
                End Try
            Next

            If totalCharsHardCapHit Then
                Debug.WriteLine("[ExcelFileParser] 已触发全文字符全局上限 MAX_TOTAL_BUILDER_CHARS = " & MAX_TOTAL_BUILDER_CHARS.ToString())
            End If

            result.FullText = If(fullTextBuilder IsNot Nothing, fullTextBuilder.ToString(), "")
            result.TotalChars = If(result.FullText IsNot Nothing, result.FullText.Length, 0)
            result.Pages = pages
            result.TotalPages = pages.Count

            Try
                result.Structure = BuildStructureFromWorkbook(workbook)
            Catch exStr As Exception
                Debug.WriteLine("[ExcelFileParser] 构建 Structure 失败: " & exStr.Message)
                result.Structure = New DocumentStructure() With {
                    .Headings = New List(Of HeadingNode)(),
                    .Tables = New List(Of TableInfo)(),
                    .Images = New List(Of ImageInfo)()
                }
            End Try

            Try
                result.Metadata = BuildMetadataFromWorkbook(workbook, Nothing, fileType)
                result.Metadata.FileSizeBytes = originalSize
            Catch exMeta As Exception
                Debug.WriteLine("[ExcelFileParser] 构建 Metadata 失败: " & exMeta.Message)
                result.Metadata = If(result.Metadata, New DocumentMetadata() With {.FileSizeBytes = originalSize})
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 从 IWorkbook 构建文档结构信息
        ''' 列出所有 Sheet 名称、行列数等
        ''' </summary>
        Private Function BuildStructureFromWorkbook(workbook As IWorkbook) As DocumentStructure
            Dim structureInfo = New DocumentStructure()
            structureInfo.Headings = New List(Of HeadingNode)()
            structureInfo.Tables = New List(Of TableInfo)()
            structureInfo.Images = New List(Of ImageInfo)()

            Dim currentPosition As Long = 0

            ' 为每个 Sheet 创建一个标题节点
            For sheetIndex As Integer = 0 To workbook.NumberOfSheets - 1
                Dim sheet = workbook.GetSheetAt(sheetIndex)
                Dim sheetName = If(sheet.SheetName, $"Sheet{sheetIndex + 1}")

                ' 获取行列数
                Dim rowCount = sheet.LastRowNum - sheet.FirstRowNum + 1
                Dim maxColCount As Integer = 0

                ' 计算最大列数
                For rowIndex As Integer = sheet.FirstRowNum To sheet.LastRowNum
                    Dim row = sheet.GetRow(rowIndex)
                    If row IsNot Nothing AndAlso row.LastCellNum > maxColCount Then
                        maxColCount = row.LastCellNum
                    End If
                Next

                ' 创建 Sheet 标题节点
                Dim headingNode = New HeadingNode With {
                    .Level = 1,
                    .Title = $"{sheetName} ({rowCount}行 x {maxColCount}列)",
                    .StartPosition = currentPosition,
                    .PageIndex = sheetIndex + 1,
                    .Children = New List(Of HeadingNode)()
                }
                structureInfo.Headings.Add(headingNode)

                ' 创建表格信息（将每个 Sheet 视为一个表格）
                Dim tableInfo = New TableInfo With {
                    .StartPosition = currentPosition,
                    .RowCount = rowCount,
                    .ColCount = maxColCount,
                    .HeaderRow = New List(Of String)(),
                    .Summary = $"工作表: {sheetName}, 行数: {rowCount}, 列数: {maxColCount}"
                }

                ' 尝试获取表头行
                Dim firstRow = sheet.GetRow(sheet.FirstRowNum)
                If firstRow IsNot Nothing Then
                    For colIndex As Integer = 0 To firstRow.LastCellNum - 1
                        Dim cell = firstRow.GetCell(colIndex)
                        If cell IsNot Nothing Then
                            tableInfo.HeaderRow.Add(GetCellStringValue(cell))
                        End If
                    Next
                End If

                tableInfo.EndPosition = currentPosition + 1000 ' 简化估算
                structureInfo.Tables.Add(tableInfo)

                structureInfo.SectionCount += 1
                currentPosition += 1000
            Next

            structureInfo.HasTableOfContents = workbook.NumberOfSheets > 1

            Return structureInfo
        End Function

        ''' <summary>
        ''' 从 IWorkbook 构建元数据信息
        ''' </summary>
        Private Function BuildMetadataFromWorkbook(workbook As IWorkbook, Optional existingMetadata As DocumentMetadata = Nothing, Optional fileType As String = "") As DocumentMetadata
            Dim metadata = If(existingMetadata, New DocumentMetadata())
            metadata.Properties = New Dictionary(Of String, String)()

            Try
                ' NPOI 属性 API 版本差异大，暂时跳过元数据读取
                ' 核心功能（文本提取/分块）不依赖文档属性
                Debug.WriteLine("[ExcelFileParser] 文档属性提取已跳过（NPOI 版本兼容）")

                ' 添加 Excel 特有属性
                metadata.Properties("格式") = If(fileType = "xlsx", "XLSX (Excel 2007+)", "XLS (Excel 97-2003)")
                metadata.Properties("工作表数量") = workbook.NumberOfSheets.ToString()

                ' 列出所有工作表名称
                Dim sheetNames = New List(Of String)()
                For i As Integer = 0 To workbook.NumberOfSheets - 1
                    sheetNames.Add(workbook.GetSheetName(i))
                Next
                metadata.Properties("工作表列表") = String.Join(", ", sheetNames)

            Catch ex As Exception
                Debug.WriteLine($"[ExcelFileParser] 读取工作簿属性时出错: {ex.Message}")
            End Try

            Return metadata
        End Function

        ''' <summary>
        ''' 将单元格值转换为字符串表示
        ''' </summary>
        ''' <remarks>处理各种数据类型：文本、数字、日期、布尔值、公式等</remarks>
        Private Function GetCellStringValue(cell As ICell) As String
            If cell Is Nothing Then Return ""

            Try
                Select Case cell.CellType
                    Case CellType.String
                        Try
                            Dim s = cell.StringCellValue
                            Return If(String.IsNullOrEmpty(s), "", s.Trim())
                        Catch
                            Return ""
                        End Try

                    Case CellType.Numeric
                        Try
                            Dim numericValue = cell.NumericCellValue
                            If [Double].IsNaN(numericValue) OrElse [Double].IsInfinity(numericValue) Then
                                Return ""
                            End If
                        Catch
                            Return ""
                        End Try

                        Dim isDate As Boolean = False
                        Try
                            isDate = DateUtil.IsCellDateFormatted(cell)
                        Catch
                            isDate = False
                        End Try

                        If isDate Then
                            Try
                                Dim dNullable = cell.DateCellValue
                                If Not dNullable.HasValue Then Return ""
                                Dim d = dNullable.Value
                                If d = DateTime.MinValue OrElse d.Year < 1900 OrElse d.Year > 9999 Then
                                    Return ""
                                End If
                                Return d.ToString("yyyy-MM-dd HH:mm:ss")
                            Catch
                                Try
                                    Return cell.NumericCellValue.ToString()
                                Catch
                                    Return ""
                                End Try
                            End Try
                        Else
                            Try
                                Dim numericValue = cell.NumericCellValue
                                If numericValue = Math.Floor(numericValue) AndAlso Math.Abs(numericValue) < Long.MaxValue AndAlso Not [Double].IsNaN(numericValue) AndAlso Not [Double].IsInfinity(numericValue) Then
                                    Return CLng(numericValue).ToString()
                                Else
                                    Return numericValue.ToString("F4").TrimEnd("0"c).TrimEnd("."c)
                                End If
                            Catch
                                Return ""
                            End Try
                        End If

                    Case CellType.Boolean
                        Try
                            Return cell.BooleanCellValue.ToString()
                        Catch
                            Return ""
                        End Try

                    Case CellType.Formula
                        Try
                            Select Case cell.CachedFormulaResultType
                                Case CellType.String
                                    Try
                                        Return cell.StringCellValue.Trim()
                                    Catch
                                        Return ""
                                    End Try
                                Case CellType.Numeric
                                    Try
                                        Dim n = cell.NumericCellValue
                                        If [Double].IsNaN(n) OrElse [Double].IsInfinity(n) Then Return ""
                                        If DateUtil.IsCellDateFormatted(cell) Then
                                            Try
                                                Dim dNullable2 = cell.DateCellValue
                                                If dNullable2.HasValue Then
                                                    Dim d2 = dNullable2.Value
                                                    If d2.Year >= 1900 AndAlso d2.Year <= 9999 Then
                                                        Return d2.ToString("yyyy-MM-dd HH:mm:ss")
                                                    End If
                                                End If
                                                Return n.ToString()
                                            Catch
                                                Return n.ToString()
                                            End Try
                                        Else
                                            Return n.ToString()
                                        End If
                                    Catch
                                        Return ""
                                    End Try
                                Case CellType.Boolean
                                    Try
                                        Return cell.BooleanCellValue.ToString()
                                    Catch
                                        Return ""
                                    End Try
                                Case Else
                                    Try
                                        Return cell.ToString()
                                    Catch
                                        Try
                                            Return cell.CellFormula
                                        Catch
                                            Return ""
                                        End Try
                                    End Try
                            End Select
                        Catch
                            Try
                                Return cell.CellFormula
                            Catch
                                Return ""
                            End Try
                        End Try

                    Case CellType.Blank
                        Return ""

                    Case Else
                        Try
                            Return cell.ToString()
                        Catch
                            Return ""
                        End Try
                End Select

            Catch ex As Exception
                Debug.WriteLine("[ExcelFileParser] 单元格转换异常: " & ex.Message)
                Return ""
            End Try
        End Function

#End Region

#region "辅助方法"

        ''' <summary>
        ''' 判断是否为 XLSX 格式（通过文件头魔数判断）
        ''' XLSX: PK (ZIP 压缩包)
        ''' XLS: D0 CF (OLE 复合文档)
        ''' </summary>
        Private Function IsXlsxFormat(fileBytes As Byte()) As Boolean
            If fileBytes Is Nothing OrElse fileBytes.Length < 4 Then
                Return False
            End If

            ' PK 头标识 ZIP/XLSX 格式
            Return fileBytes(0) = &H50 AndAlso fileBytes(1) = &H4B
        End Function

        ''' <summary>
        ''' P5：安全的 AppendLine 包装（达到单Sheet字符上限后静默丢弃，后续不再追加；capHit置true一次后不变）
        ''' </summary>
        Private Shared Sub AppendLineSafe(sb As StringBuilder, line As String, maxChars As Integer, ByRef capHit As Boolean)
            If sb Is Nothing Then Return
            If capHit Then Return

            Dim addChars As Integer = If(line, "").Length + Environment.NewLine.Length
            If sb.Length + addChars > maxChars Then
                capHit = True
                ' 尽可能把当前 line 能塞的部分塞进去（最后一次）
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