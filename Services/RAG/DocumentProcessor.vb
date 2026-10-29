Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Net.Http
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports ExcelAIVSTOAddIn.Services.Parsers
Imports ExcelAIVSTOAddIn.Utils

Namespace Services.RAG

    ''' <summary>
    ''' 文档处理流水线：下载(OSS URL) → 解析 → 分块 → 返回 Chunk 列表
    ''' 复用现有 IDocumentParser 和 SemanticChunker
    ''' </summary>
    Public Class DocumentProcessor
        Implements IDisposable

        Private _chunker As SemanticChunker
        Private _httpClient As HttpClient
        Private _isDisposed As Boolean = False

        ' 支持的导入格式（与 RagBridge.BrowseLocalFile 的过滤器保持一致）
        Private Shared ReadOnly SupportedExtensions As String() = {".pdf", ".doc", ".docx", ".xls", ".xlsx", ".csv", ".txt", ".md"}

        Public Sub New()
            ' 复用已有单例
            _chunker = SemanticChunker.Instance
            _httpClient = New HttpClient()
            _httpClient.Timeout = TimeSpan.FromMinutes(5)  ' 文件下载可能较慢
        End Sub

#Region "主入口方法"

        ''' <summary>
        ''' 处理文档主入口
        ''' 流程: 下载OSS文件 → 临时保存 → 选择解析器 → 解析文本 → 语义分块
        ''' </summary>
        ''' <param name="ossUrl">OSS 文件 URL</param>
        ''' <param name="kbId">所属知识库 ID</param>
        ''' <param name="docId">文档 ID</param>
        ''' <param name="chunkSize">分块大小，默认 1500</param>
        ''' <param name="overlap">重叠字符数，默认 200</param>
        ''' <returns>KnowledgeChunk 列表（向量尚未计算）</returns>
        Public Async Function ProcessDocumentAsync(ossUrl As String,
                                                    kbId As String,
                                                    docId As String,
                                                    Optional chunkSize As Integer = 1500,
                                                    Optional overlap As Integer = 200) As Task(Of List(Of KnowledgeChunk))

            Dim chunks As New List(Of KnowledgeChunk)()

            Try
                System.Diagnostics.Debug.WriteLine($"[DocProcessor] 开始处理文档: {ossUrl}")
                System.Diagnostics.Debug.WriteLine($"[DocProcessor] kbId={kbId}, docId={docId}, chunkSize={chunkSize}, overlap={overlap}")

                ' Step 1: 从 OSS URL 下载文件到临时目录
                Dim localPath = Await DownloadFromOssAsync(ossUrl)
                If String.IsNullOrEmpty(localPath) OrElse Not File.Exists(localPath) Then
                    Throw New Exception("文件下载失败或本地路径无效")
                End If

                Try
                    ' 校验扩展名（URL 去掉 query 后取扩展名）
                    Dim urlExt = Path.GetExtension(ossUrl.Split({"?"c, "#"c}, StringSplitOptions.None)(0)).ToLower()
                    If Not SupportedExtensions.Contains(urlExt) Then
                        Dim extDisplay = If(String.IsNullOrEmpty(urlExt), "未知", urlExt)
                        Throw New Exception($"不支持的文件格式: {extDisplay}（支持: PDF/Word/Excel/CSV/TXT/MD）")
                    End If

                    ' Step 2: 根据文件类型选择合适的解析器并提取纯文本
                    Dim parsedText = Await ParseToTextAsync(localPath)

                    If String.IsNullOrWhiteSpace(parsedText) Then
                        Throw New Exception("文档解析结果为空")
                    End If

                    System.Diagnostics.Debug.WriteLine($"[DocProcessor] 文本解析完成, 长度: {parsedText.Length}")

                    ' Step 3: 使用 SemanticChunker 进行语义分块
                    Dim textChunks = _chunker.Chunk(parsedText, Nothing, "auto")

                    If textChunks Is Nothing OrElse textChunks.Count = 0 Then
                        Throw New Exception("分块结果为空")
                    End If

                    System.Diagnostics.Debug.WriteLine($"[DocProcessor] 分块完成, 块数: {textChunks.Count}")

                    ' Step 4: 将 TextChunk 转换为 KnowledgeChunk（附加元数据）
                    chunks = ConvertToKnowledgeChunks(textChunks, docId, kbId)

                    System.Diagnostics.Debug.WriteLine($"[DocProcessor] 文档处理完成, 生成 {chunks.Count} 个 KnowledgeChunk")

                Finally
                    ' 清理临时文件
                    Try
                        If File.Exists(localPath) Then
                            File.Delete(localPath)
                            System.Diagnostics.Debug.WriteLine($"[DocProcessor] 临时文件已清理: {localPath}")
                        End If
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine($"[DocProcessor] 清理临时文件失败: {ex.Message}")
                    End Try
                End Try

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[DocProcessor] ProcessDocumentAsync 异常: {ex.Message}")
                Throw
            End Try

            Return chunks
        End Function

        ''' <summary>
        ''' 处理本地文件（跳过下载，直接解析）
        ''' 流程: 读取本地文件 → 解析文本 → 语义分块
        ''' </summary>
        ''' <param name="localPath">本地文件路径</param>
        ''' <param name="kbId">所属知识库 ID</param>
        ''' <param name="docId">文档 ID</param>
        ''' <param name="chunkSize">分块大小，默认 1500</param>
        ''' <param name="overlap">重叠字符数，默认 200</param>
        ''' <returns>KnowledgeChunk 列表（向量尚未计算）</returns>
        Public Async Function ProcessLocalDocumentAsync(localPath As String,
                                                         kbId As String,
                                                         docId As String,
                                                         Optional chunkSize As Integer = 1500,
                                                         Optional overlap As Integer = 200) As Task(Of List(Of KnowledgeChunk))
            Dim chunks As New List(Of KnowledgeChunk)()

            Try
                System.Diagnostics.Debug.WriteLine($"[DocProcessor] 开始处理本地文件: {localPath}")
                System.Diagnostics.Debug.WriteLine($"[DocProcessor] kbId={kbId}, docId={docId}, chunkSize={chunkSize}, overlap={overlap}")

                ' 验证文件存在
                If Not File.Exists(localPath) Then
                    Throw New Exception($"本地文件不存在: {localPath}")
                End If

                ' 校验扩展名（提前失败，给出明确错误而不是解析期报错）
                Dim extCheck = Path.GetExtension(localPath).ToLower()
                If Not SupportedExtensions.Contains(extCheck) Then
                    Dim extDisplay2 = If(String.IsNullOrEmpty(extCheck), "未知", extCheck)
                    Throw New Exception($"不支持的文件格式: {extDisplay2}（支持: PDF/Word/Excel/CSV/TXT/MD）")
                End If

                ' Step 1: 解析文本
                Dim parsedText = Await ParseToTextAsync(localPath)

                If String.IsNullOrWhiteSpace(parsedText) Then
                    Throw New Exception("文档解析结果为空")
                End If

                System.Diagnostics.Debug.WriteLine($"[DocProcessor] 本地文件解析完成, 长度: {parsedText.Length}")

                ' Step 2: 语义分块
                Dim textChunks = _chunker.Chunk(parsedText, Nothing, "auto")

                If textChunks Is Nothing OrElse textChunks.Count = 0 Then
                    Throw New Exception("分块结果为空")
                End If

                System.Diagnostics.Debug.WriteLine($"[DocProcessor] 分块完成, 块数: {textChunks.Count}")

                ' Step 3: 转换为 KnowledgeChunk
                chunks = ConvertToKnowledgeChunks(textChunks, docId, kbId)

                System.Diagnostics.Debug.WriteLine($"[DocProcessor] 本地文件处理完成, 生成 {chunks.Count} 个 KnowledgeChunk")

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[DocProcessor] ProcessLocalDocumentAsync 异常: {ex.Message}")
                Throw
            End Try

            Return chunks
        End Function

#End Region

#Region "内部方法"

        ''' <summary>
        ''' 从 OSS URL 下载文件到临时目录
        ''' </summary>
        Private Async Function DownloadFromOssAsync(ossUrl As String) As Task(Of String)
            Try
                Dim tempDir = Path.Combine(Path.GetTempPath(), "SmartPilot_RAG_Docs")

                If Not Directory.Exists(tempDir) Then
                    Directory.CreateDirectory(tempDir)
                End If

                ' 从 URL 提取文件名
                Dim uri = New Uri(ossUrl)
                Dim fileName = Path.GetFileName(uri.LocalPath)

                If String.IsNullOrWhiteSpace(fileName) Then
                    fileName = $"document_{Guid.NewGuid():N}.bin"
                End If

                Dim localPath = Path.Combine(tempDir, fileName)

                System.Diagnostics.Debug.WriteLine($"[DocProcessor] 正在下载: {ossUrl}")
                System.Diagnostics.Debug.WriteLine($"[DocProcessor] 保存到: {localPath}")

                Using response = Await _httpClient.GetAsync(ossUrl, HttpCompletionOption.ResponseHeadersRead)
                    response.EnsureSuccessStatusCode()

                    Using fs As New FileStream(localPath, FileMode.Create, FileAccess.Write, FileShare.None)
                        Await response.Content.CopyToAsync(fs)
                    End Using
                End Using

                Dim fileInfo As New FileInfo(localPath)
                System.Diagnostics.Debug.WriteLine($"[DocProcessor] 下载完成, 大小: {fileInfo.Length} 字节")

                Return localPath

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[DocProcessor] 下载失败: {ex.Message}")
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' 根据文件类型选择合适的解析器并提取纯文本
        ''' </summary>
        Private Async Function ParseToTextAsync(filePath As String) As Task(Of String)
            Try
                Dim extension = Path.GetExtension(filePath).ToLower()
                Dim parser As IDocumentParser = Nothing

                Select Case extension
                    Case ".pdf"
                        parser = New PdfDocumentParser()
                    Case ".docx", ".doc"
                        parser = New DocxDocumentParser()
                    Case ".xlsx", ".xls", ".csv"
                        parser = New ExcelFileParser()
                    Case Else
                        ' 尝试作为纯文本读取
                        If extension = ".txt" OrElse extension = ".md" Then
                            Return ReadTextWithEncodingFallback(filePath)
                        Else
                            Throw New Exception($"不支持的文件格式: {extension}")
                        End If
                End Select

                ' 使用解析器解析文档（读取文件为 Base64 后调用 ExtractFullText）
                Dim fileBytes = File.ReadAllBytes(filePath)
                Dim base64Data = Convert.ToBase64String(fileBytes)
                Dim parseResult = Await parser.ExtractFullText(base64Data)
                Dim result = If(parseResult?.Success = True, parseResult.FullText, "")

                If Not String.IsNullOrWhiteSpace(result) Then
                    ' 清理文本：移除多余空白字符
                    result = CleanText(result)
                End If

                Return result

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[DocProcessor] 解析文档失败: {ex.Message}")

                ' 如果专用解析器失败，尝试回退到纯文本读取
                Try
                    System.Diagnostics.Debug.WriteLine("[DocProcessor] 尝试回退到纯文本读取...")
                    Return ReadTextWithEncodingFallback(filePath)
                Catch ex2 As Exception
                    System.Diagnostics.Debug.WriteLine($"[DocProcessor] 回退也失败: {ex2.Message}")
                    Throw
                End Try
            End Try
        End Function

        ''' <summary>
        ''' 读取文本文件：BOM 优先，UTF-8 严格解析失败时回退 GBK（中文 Windows 常见编码）。
        ''' 直接用 File.ReadAllText(UTF8) 读 GBK 文件会得到乱码，检索质量归零。
        ''' </summary>
        Private Shared Function ReadTextWithEncodingFallback(filePath As String) As String
            Dim bytes = File.ReadAllBytes(filePath)
            If bytes.Length = 0 Then Return ""

            ' BOM 检测
            If bytes.Length >= 3 AndAlso bytes(0) = &HEF AndAlso bytes(1) = &HBB AndAlso bytes(2) = &HBF Then
                Return New System.Text.UTF8Encoding(False).GetString(bytes, 3, bytes.Length - 3)
            End If
            If bytes.Length >= 2 AndAlso bytes(0) = &HFF AndAlso bytes(1) = &HFE Then
                Return New System.Text.UnicodeEncoding(False, False).GetString(bytes, 2, bytes.Length - 2)   ' UTF-16 LE
            End If
            If bytes.Length >= 2 AndAlso bytes(0) = &HFE AndAlso bytes(1) = &HFF Then
                Return New System.Text.UnicodeEncoding(True, False).GetString(bytes, 2, bytes.Length - 2)   ' UTF-16 BE
            End If

            ' 无 BOM：先按 UTF-8 严格解析，失败回退 GBK
            Try
                Return New System.Text.UTF8Encoding(False, True).GetString(bytes)
            Catch
                Try
                    Return System.Text.Encoding.GetEncoding(936).GetString(bytes)   ' GBK/GB2312
                Catch
                    Return System.Text.Encoding.Default.GetString(bytes)
                End Try
            End Try
        End Function

        ''' <summary>
        ''' 清理文本内容
        ''' </summary>
        Private Function CleanText(text As String) As String
            If String.IsNullOrWhiteSpace(text) Then Return ""

            ' 移除多余空行但保留段落结构
            text = Regex.Replace(text, "[\r\n]{3,}", vbCrLf & vbCrLf)

            ' 移除首尾空白
            text = text.Trim()

            Return text
        End Function

        ''' <summary>
        ''' 将 TextChunk 转换为 KnowledgeChunk（附加元数据）
        ''' </summary>
        Private Function ConvertToKnowledgeChunks(textChunks As List(Of TextChunk),
                                                     docId As String,
                                                     kbId As String) As List(Of KnowledgeChunk)
            Dim knowledgeChunks As New List(Of KnowledgeChunk)()

            For Each tc In textChunks
                Dim kc As New KnowledgeChunk With {
                    .Id = Guid.NewGuid().ToString("N"),
                    .DocId = docId,
                    .KbId = kbId,
                    .Text = tc.Text,
                    .TextHash = ComputeHash(tc.Text),
                    .ChunkIndex = tc.Id,
                    .SourceSection = tc.SourceSection,
                    .StartPosition = tc.StartPosition,
                    .EndPosition = tc.EndPosition,
                    .Metadata = New Dictionary(Of String, String) From {
                        {"chunk_index", tc.Id.ToString()},
                        {"char_count", tc.CharCount.ToString()},
                        {"level", tc.Level.ToString()}
                    }
                }

                If Not String.IsNullOrWhiteSpace(tc.SourceSection) Then
                    kc.Metadata("source_section") = tc.SourceSection
                End If

                knowledgeChunks.Add(kc)
            Next

            Return knowledgeChunks
        End Function

        ''' <summary>
        ''' 计算文本哈希（用于去重）
        ''' </summary>
        Private Function ComputeHash(text As String) As String
            If String.IsNullOrWhiteSpace(text) Then Return ""

            Using md5Hash As MD5 = MD5.Create()
                Dim inputBytes = Encoding.UTF8.GetBytes(text)
                Dim hashBytes = md5Hash.ComputeHash(inputBytes)
                Dim sb As New StringBuilder()

                For Each b In hashBytes
                    sb.Append(b.ToString("x2"))
                Next

                Return sb.ToString()
            End Using
        End Function

#End Region

#Region "IDisposable Support"

	Protected Overloads Sub Dispose(disposing As Boolean)
		If Not _isDisposed Then
			If disposing Then
				If _httpClient IsNot Nothing Then
					_httpClient.Dispose()
					_httpClient = Nothing
				End If
			End If
			_isDisposed = True
		End If
	End Sub

	Public Overloads Sub Dispose() Implements IDisposable.Dispose
		Dispose(True)
		GC.SuppressFinalize(Me)
	End Sub

#End Region

    End Class

End Namespace
