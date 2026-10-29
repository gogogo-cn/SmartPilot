Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Diagnostics
Imports ExcelAIVSTOAddIn.Services.Parsers

Namespace Services

    ''' <summary>
    ''' 文档路由器 - 根据文件类型自动分发到对应 Parser
    ''' 所有格式的统一入口，前端只需调用一个方法
    ''' </summary>
    Public Class DocumentRouter

        Private Shared ReadOnly _instance As New Lazy(Of DocumentRouter)(Function()
            Dim router = New DocumentRouter()
            router.InitializeParsers()
            Return router
        End Function, System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        Private _parsers As New Dictionary(Of String, IDocumentParser)(StringComparer.OrdinalIgnoreCase)

        ' ========== 单例模式 ==========
        Public Shared ReadOnly Property Instance As DocumentRouter
            Get
                Return _instance.Value
            End Get
        End Property

        Private Sub New()
        End Sub

        ''' <summary>
        ''' 初始化所有可用的文档解析器（P0: PDF + DOCX + XLSX）
        ''' </summary>
        Private Sub InitializeParsers()
            Try
                ' PDF 解析器
                RegisterParser(New PdfDocumentParser())
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[DocumentRouter] PDF 解析器初始化失败: " & ex.Message)
            End Try

            Try
                ' Word 解析器 (DOC/DOCX)
                RegisterParser(New DocxDocumentParser())
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[DocumentRouter] Word 解析器初始化失败: " & ex.Message)
            End Try

            Try
                ' Excel 解析器 (XLS/XLSX)
                RegisterParser(New ExcelFileParser())
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[DocumentRouter] Excel 解析器初始化失败: " & ex.Message)
            End Try

            System.Diagnostics.Debug.WriteLine($"[DocumentRouter] 已注册 {_parsers.Count} 个解析器")
        End Sub

        ''' <summary>注册解析器</summary>
        Public Sub RegisterParser(parser As IDocumentParser)
            For Each ext In parser.SupportedExtensions
                If Not _parsers.ContainsKey(ext) Then
                    _parsers(ext) = parser
                    System.Diagnostics.Debug.WriteLine($"[DocumentRouter] 注册 [{parser.DisplayName}] → .{ext}")
                End If
            Next
        End Sub

        ''' <summary>获取已支持的扩展名列表</summary>
        Public Function GetSupportedExtensions() As List(Of String)
            Return _parsers.Keys.ToList()
        End Function

        ''' <summary>
        ''' 根据文件名获取对应的解析器
        ''' </summary>
        Public Function GetParser(fileName As String) As IDocumentParser
            Dim ext = Path.GetExtension(fileName).ToLower().TrimStart("."c)
            If _parsers.ContainsKey(ext) Then Return _parsers(ext)
            Return Nothing
        End Function

        ''' <summary>
        ''' 检查是否支持指定格式
        ''' </summary>
        Public Function IsSupportedFormat(fileName As String) As Boolean
            Dim ext = Path.GetExtension(fileName).ToLower().TrimStart("."c)
            Return _parsers.ContainsKey(ext)
        End Function

        ''' <summary>
        ''' 一站式解析入口 - 自动路由到对应 Parser
        ''' 前端只需调用此方法，无需关心具体格式
        ''' </summary>
        Public Async Function ParseDocumentAsync(
            fileName As String,
            base64Data As String
        ) As Task(Of UnifiedDocumentResult)

            Dim sw = Stopwatch.StartNew()

            ' 参数校验
            If String.IsNullOrWhiteSpace(base64Data) Then
                Return CreateErrorResult("文件数据为空")
            End If

            ' 路由到对应 Parser
            Dim parser = Me.GetParser(fileName)
            If parser Is Nothing Then
                Return CreateErrorResult($"不支持的文件格式: {fileName}")
            End If

            System.Diagnostics.Debug.WriteLine($"[DocumentRouter] 开始解析: {fileName} → [{parser.DisplayName}]")

            Try
                Dim result = Await parser.ExtractFullText(base64Data)
                sw.Stop()
                result.ParseTimeMs = sw.ElapsedMilliseconds

                If result.Success Then
                    System.Diagnostics.Debug.WriteLine(
                        $"[DocumentRouter] 解析完成: {fileName}, " &
                        $"字符数={result.TotalChars}, 耗时={result.ParseTimeMs}ms"
                    )
                Else
                    System.Diagnostics.Debug.WriteLine(
                        $"[DocumentRouter] 解析失败: {fileName}, {result.ErrorMessage}"
                    )
                End If

                Return result

            Catch ex As Exception
                sw.Stop()
                System.Diagnostics.Debug.WriteLine($"[DocumentRouter] 解析异常: {ex.Message}")
                Return CreateErrorResult("解析异常: " & ex.Message)
            End Try

        End Function

        ''' <summary>
        ''' 仅提取元数据（轻量操作，用于快速预览）
        ''' </summary>
        Public Async Function GetMetadataAsync(
            fileName As String,
            base64Data As String
        ) As Task(Of DocumentMetadata)

            Return Await Task.Run(Function()
                Dim parser = Me.GetParser(fileName)
                If parser Is Nothing Then Return New DocumentMetadata()

                Try
                    Return parser.ExtractMetadata(base64Data)
                Catch ex As Exception
                    Return New DocumentMetadata()
                End Try
            End Function)

        End Function

        ''' <summary>
        ''' 提取文档结构信息（用于语义分块）
        ''' </summary>
        Public Async Function GetStructureAsync(
            fileName As String,
            base64Data As String
        ) As Task(Of DocumentStructure)

            Dim parser = Me.GetParser(fileName)
            If parser Is Nothing Then Return New DocumentStructure()

            Try
                Return Await parser.ExtractStructure(base64Data)
            Catch ex As Exception
                Return New DocumentStructure()
            End Try

        End Function

        ''' <summary>
        ''' 创建错误结果
        ''' </summary>
        Private Function CreateErrorResult(msg As String) As UnifiedDocumentResult
            Return New UnifiedDocumentResult With {
                .Success = False,
                .ErrorMessage = msg
            }
        End Function

    End Class

End Namespace
