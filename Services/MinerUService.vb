Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Net.Http
Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services

    ''' <summary>
    ''' MinerU 文档解析服务客户端
    ''' 调用 mineru_server.py HTTP API 实现 PDF 高级解析
    ''' 支持：文本提取、图片提取、表格识别、公式识别（LaTeX）
    '''
    ''' 当 MinerU 服务不可用时，自动降级到 PdfSharp 基础提取
    ''' </summary>
    Public Class MinerUService

        Private Shared ReadOnly _lazyInstance As New Lazy(Of MinerUService)(
            Function() New MinerUService(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        Public Shared ReadOnly Property Instance As MinerUService
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        Private _httpClient As HttpClient
        Private _endpoint As String = "http://127.0.0.1:8036"
        Private _available As Boolean? = Nothing
        Private _lastCheckTime As DateTime? = Nothing
        Private _checkInterval As TimeSpan = TimeSpan.FromSeconds(30)

        Private Sub New()
            _httpClient = New HttpClient()
            _httpClient.Timeout = TimeSpan.FromMinutes(5)  ' PDF 解析可能很慢
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "SmartPilot/1.0")
            LoadConfiguration()
        End Sub

        ''' <summary>
        ''' 从配置加载 MinerU 端点
        ''' </summary>
        Private Sub LoadConfiguration()
            Try
                Dim configPath = Path.Combine(Utils.PathHelper.ConfigDirectory, "rag_config.json")
                If File.Exists(configPath) Then
                    Dim json = File.ReadAllText(configPath)
                    Dim configObj = JObject.Parse(json)
                    Dim mineruConfig = configObj("mineru")
                    If mineruConfig IsNot Nothing Then
                        Dim ep = mineruConfig("endpoint")?.ToString()
                        If Not String.IsNullOrWhiteSpace(ep) Then
                            _endpoint = ep.TrimEnd("/"c)
                        End If
                    End If
                End If
            Catch ex As Exception
                Debug.WriteLine($"[MinerUService] 配置加载失败: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' 检查 MinerU 服务是否可用
        ''' </summary>
        Public Async Function IsAvailableAsync() As Task(Of Boolean)
            ' 缓存检查结果，30秒内不重复检查
            If _available.HasValue AndAlso _lastCheckTime.HasValue Then
                If (DateTime.Now - _lastCheckTime.Value) < _checkInterval Then
                    Return _available.Value
                End If
            End If

            Try
                Dim response = Await _httpClient.GetAsync($"{_endpoint}/health")
                If response.IsSuccessStatusCode Then
                    Dim json = Await response.Content.ReadAsStringAsync()
                    Dim obj = JObject.Parse(json)
                    _available = obj("mineru_available")?.ToObject(Of Boolean)()
                    _lastCheckTime = DateTime.Now

                    Debug.WriteLine($"[MinerUService] 服务可用={_available.Value}, " &
                                    $"版本={obj("mineru_version")?.ToString()}")
                    Return _available.Value
                Else
                    _available = False
                    _lastCheckTime = DateTime.Now
                    Return False
                End If
            Catch ex As Exception
                _available = False
                _lastCheckTime = DateTime.Now
                Debug.WriteLine($"[MinerUService] 服务不可用: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 获取 MinerU 安装状态详情
        ''' </summary>
        Public Async Function GetInstallStatusAsync() As Task(Of JObject)
            Try
                Dim response = Await _httpClient.GetAsync($"{_endpoint}/check-install")
                If response.IsSuccessStatusCode Then
                    Dim json = Await response.Content.ReadAsStringAsync()
                    Return JObject.Parse(json)
                End If
            Catch ex As Exception
                Debug.WriteLine($"[MinerUService] 获取安装状态失败: {ex.Message}")
            End Try

            Return JObject.FromObject(New With {
                .mineru_installed = False,
                .error = "MinerU 服务未启动或不可达"
            })
        End Function

        ''' <summary>
        ''' 解析 PDF 文件（通过 Base64 传输）
        ''' </summary>
        ''' <param name="pdfBase64">PDF 文件的 Base64 编码</param>
        ''' <param name="extractImages">是否提取图片</param>
        ''' <param name="fallback">MinerU 不可用时是否降级到 PyMuPDF</param>
        ''' <returns>解析结果 JSON</returns>
        Public Async Function ParsePdfAsync(pdfBase64 As String,
                                             Optional extractImages As Boolean = True,
                                             Optional fallback As Boolean = True) As Task(Of MinerUParseResult)
            Dim result As New MinerUParseResult()

            Try
                Dim payload = JsonConvert.SerializeObject(New With {
                    .base64 = pdfBase64,
                    .extract_images = extractImages,
                    .fallback = fallback
                })

                Dim content = New StringContent(payload, Encoding.UTF8, "application/json")
                Dim response = Await _httpClient.PostAsync($"{_endpoint}/parse-base64", content)

                If response.IsSuccessStatusCode Then
                    Dim json = Await response.Content.ReadAsStringAsync()
                    Dim obj = JObject.Parse(json)

                    result.Success = If(obj("success")?.ToObject(Of Boolean)(), False)
                    result.FullText = obj("full_text")?.ToString()
                    result.Markdown = obj("markdown")?.ToString()
                    result.Error = obj("error")?.ToString()

                    ' 解析图片列表
                    If obj("images") IsNot Nothing Then
                        result.Images = New List(Of MinerUImage)()
                        For Each img In obj("images")
                            result.Images.Add(New MinerUImage With {
                                .Index = If(img("index")?.ToObject(Of Integer)(), 0),
                                .Base64 = img("base64")?.ToString(),
                                .AltText = img("alt_text")?.ToString(),
                                .Width = If(img("width")?.ToObject(Of Integer)(), 0),
                                .Height = If(img("height")?.ToObject(Of Integer)(), 0),
                                .Page = If(img("page")?.ToObject(Of Integer)(), 0)
                            })
                        Next
                    End If

                    ' 解析表格列表
                    If obj("tables") IsNot Nothing Then
                        result.Tables = New List(Of MinerUTable)()
                        For Each tbl In obj("tables")
                            result.Tables.Add(New MinerUTable With {
                                .Index = If(tbl("index")?.ToObject(Of Integer)(), 0),
                                .Html = tbl("html")?.ToString(),
                                .Markdown = tbl("markdown")?.ToString(),
                                .RowCount = If(tbl("row_count")?.ToObject(Of Integer)(), 0),
                                .ColCount = If(tbl("col_count")?.ToObject(Of Integer)(), 0),
                                .Page = If(tbl("page")?.ToObject(Of Integer)(), 0)
                            })
                        Next
                    End If

                    ' 解析公式列表
                    If obj("formulas") IsNot Nothing Then
                        result.Formulas = New List(Of MinerUFormula)()
                        For Each fml In obj("formulas")
                            result.Formulas.Add(New MinerUFormula With {
                                .Index = If(fml("index")?.ToObject(Of Integer)(), 0),
                                .Latex = fml("latex")?.ToString(),
                                .Type = fml("type")?.ToString(),
                                .Page = If(fml("page")?.ToObject(Of Integer)(), 0)
                            })
                        Next
                    End If

                    ' 解析元数据
                    If obj("metadata") IsNot Nothing Then
                        result.PageCount = If(obj("metadata")("page_count")?.ToObject(Of Integer)(), 0)
                        result.ParseTimeMs = If(obj("metadata")("parse_time_ms")?.ToObject(Of Long)(), 0L)
                        result.Engine = obj("metadata")("engine")?.ToString()
                    End If

                    ' 保存分页原始数据
                    result.Pages = TryCast(obj("pages"), JArray)

                    Debug.WriteLine($"[MinerUService] 解析完成 - 成功={result.Success}, " &
                                    $"文本={result.FullText?.Length}字符, " &
                                    $"图片={If(result.Images?.Count, 0)}, " &
                                    $"表格={If(result.Tables?.Count, 0)}, " &
                                    $"公式={If(result.Formulas?.Count, 0)}")
                Else
                    result.Success = False
                    Dim statusCode = CInt(response.StatusCode)
                    Dim errorBody = Await response.Content.ReadAsStringAsync()
                    result.Error = $"HTTP {statusCode}: {errorBody}"
                End If

            Catch ex As Exception
                result.Success = False
                result.Error = $"MinerU 调用失败: {ex.Message}"
                Debug.WriteLine($"[MinerUService] 调用异常: {ex.ToString()}")
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 解析 PDF 文件（通过本地文件路径）
        ''' </summary>
        Public Async Function ParsePdfFileAsync(filePath As String,
                                                  Optional extractImages As Boolean = True,
                                                  Optional fallback As Boolean = True) As Task(Of MinerUParseResult)
            If Not File.Exists(filePath) Then
                Dim errResult As New MinerUParseResult()
                errResult.Success = False
                errResult.Error = $"文件不存在: {filePath}"
                Return errResult
            End If

            Dim fileBytes = File.ReadAllBytes(filePath)
            Dim base64 = Convert.ToBase64String(fileBytes)
            Return Await ParsePdfAsync(base64, extractImages, fallback)
        End Function

        ''' <summary>
        ''' 释放资源
        ''' </summary>
        Public Sub Dispose()
            If _httpClient IsNot Nothing Then
                _httpClient.Dispose()
                _httpClient = Nothing
            End If
        End Sub

    End Class

#Region "数据模型"

    ''' <summary>
    ''' MinerU 解析结果
    ''' </summary>
    Public Class MinerUParseResult
        Public Property Success As Boolean = False
        Public Property FullText As String = ""
        Public Property Markdown As String = ""
        Public Property Images As List(Of MinerUImage)
        Public Property Tables As List(Of MinerUTable)
        Public Property Formulas As List(Of MinerUFormula)
        Public Property PageCount As Integer = 0
        Public Property ParseTimeMs As Long = 0
        Public Property Engine As String = ""
        Public Property Pages As JArray
        Public Property [Error] As String = ""
    End Class

    ''' <summary>
    ''' MinerU 提取的图片
    ''' </summary>
    Public Class MinerUImage
        Public Property Index As Integer = 0
        Public Property Base64 As String = ""
        Public Property AltText As String = ""
        Public Property Width As Integer = 0
        Public Property Height As Integer = 0
        Public Property Page As Integer = 0
    End Class

    ''' <summary>
    ''' MinerU 识别的表格
    ''' </summary>
    Public Class MinerUTable
        Public Property Index As Integer = 0
        Public Property Html As String = ""
        Public Property Markdown As String = ""
        Public Property RowCount As Integer = 0
        Public Property ColCount As Integer = 0
        Public Property Page As Integer = 0
    End Class

    ''' <summary>
    ''' MinerU 识别的公式
    ''' </summary>
    Public Class MinerUFormula
        Public Property Index As Integer = 0
        Public Property Latex As String = ""
        Public Property Type As String = ""      ' "inline" / "display"
        Public Property Page As Integer = 0
    End Class

#End Region

End Namespace