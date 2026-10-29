Option Strict On
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Net.Http
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services.Mcp

    ''' <summary>
    ''' 进程内网页搜索工具集（Bing 直连，免 API Key，国内网络可达）
    ''' 独立于 Excel/PPT/Computer 工具集，单一职责：联网搜索知识。
    ''' 全程只使用 HttpClient 在网络层做请求，不拉起任何子进程，Excel 关闭时零额外清理开销。
    ''' 数据源说明：优先选 Bing（未选 DuckDuckGo —— 该网站在多数国内网络环境不可达，实测连不上）。
    ''' </summary>
    Partial Public Class WebSearchTools

        Private Shared ReadOnly _httpClient As HttpClient = New HttpClient()

        Shared Sub New()
            Try
                System.Net.ServicePointManager.SecurityProtocol = System.Net.ServicePointManager.SecurityProtocol Or
                    CType(3072, System.Net.SecurityProtocolType) ' 3072 = SecurityProtocolType.Tls12
            Catch
            End Try
            Try
                _httpClient.Timeout = TimeSpan.FromSeconds(20)
                _httpClient.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent",
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
                _httpClient.DefaultRequestHeaders.TryAddWithoutValidation("Accept", "text/html,application/xhtml+xml")
                _httpClient.DefaultRequestHeaders.TryAddWithoutValidation("Accept-Language", "zh-CN,zh;q=0.9")
            Catch
            End Try
        End Sub

        ' ── 工具定义（真源，与注册统一） ──

        Public Shared Function GetToolDefinitions() As List(Of McpTool)
            Return New List(Of McpTool) From {
                New McpTool("web_search", "联网搜索网页知识并返回结构化结果标题/网址/摘要，适合查资料、查资讯、陌生概念等时效性信息",
                    New McpToolInputSchema() With {
                        .Properties = New Dictionary(Of String, McpToolProperty) From {
                            {"query", New McpToolProperty("string", "搜索关键词（必填）")},
                            {"maxResults", New McpToolProperty("integer", "返回结果数，默认 5，范围 1~10")},
                            {"region", New McpToolProperty("string", "搜索地域代码，可选，如 cn-zh（中文）/ us-en（英文）")}
                        },
                        .Required = New List(Of String) From {"query"}
                    })
            }
        End Function

        Public Shared Sub RegisterAllTools()
            For Each tool In GetToolDefinitions()
                McpCapabilityRegistry.Instance.RegisterMcpTool(tool,
                    displayName:="联网搜索",
                    category:=CapabilityCategory.Utility,
                    keywords:=New List(Of String) From {
                        "web_search", "搜索", "查询", "网上", "网络", "联网", "搜索知识",
                        "查一下", "搜一下", "资料", "百科", "资讯", "最新", "实时", "消息"
                    },
                    isWriteOperation:=False,
                    isDangerous:=False,
                    source:="web-local")
            Next
        End Sub

        ' ── 统一执行入口 ──

        Public Shared Function ExecuteTool(toolName As String, args As Dictionary(Of String, Object)) As ToolsCallResult
            Select Case toolName
                Case "web_search"
                    Return Search(args)
                Case Else
                    Return Err("未找到执行方法: " & toolName)
            End Select
        End Function

        ' ── 搜索实现 ──

        Private Shared Function Search(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim query = G(Of String)(args, "query", "")
            If String.IsNullOrWhiteSpace(query) Then
                Return Err("缺少必要参数 query（搜索关键词）")
            End If
            Dim maxResults = Math.Max(1, Math.Min(10, G(args, "maxResults", 5)))
            Dim region = G(Of String)(args, "region", "")

            Dim sb As New StringBuilder()
            sb.Append("https://www.bing.com/search?q=").Append(System.Uri.EscapeDataString(query))
            sb.Append("&num=").Append(maxResults)
            If Not String.IsNullOrWhiteSpace(region) Then
                ' 地域参数：如 cn-zh → zh-CN；us-en → en-US（Bing setmkt 用语言-地区）
                sb.Append("&setmkt=").Append(System.Uri.EscapeDataString(region.Trim()))
            End If

            Try
                Dim html = AwaitSearchAsync(sb.ToString()).GetAwaiter().GetResult()
                Dim results = ParseResults(html, maxResults)

                Dim arr = New JArray()
                For Each r In results
                    arr.Add(New JObject() From {
                        {"title", CStr(r("title"))},
                        {"url", CStr(r("url"))},
                        {"snippet", CStr(r("snippet"))}
                    })
                Next

                Dim payload = New JObject() From {
                    {"query", query},
                    {"total", results.Count},
                    {"results", arr}
                }

                If results.Count = 0 Then
                    payload("note") = JToken.FromObject("未解析到网页结果（可能触发反爬、或确实无与该关键词匹配的独立网页），可尝试更换关键词。")
                End If

                Return OK(payload.ToString(Formatting.Indented))
            Catch ex As Exception
                Debug.WriteLine($"[WebSearchTools] 搜索异常: {ex.Message}")
                Return Err("搜索失败: " & ex.Message & "（Bing 偶尔拦截自动请求，可稍后重试，或尝试更换关键词）")
            End Try
        End Function

        Private Shared Async Function AwaitSearchAsync(url As String) As Task(Of String)
            Using req As New HttpRequestMessage(HttpMethod.Get, url)
                Dim resp = Await _httpClient.SendAsync(req, HttpCompletionOption.ResponseContentRead).ConfigureAwait(False)
                If Not resp.IsSuccessStatusCode Then
                    Throw New Exception("搜索引擎返回 HTTP " & CInt(resp.StatusCode))
                End If
                Return Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
            End Using
        End Function

        ' ── HTML 解析：按 b_algo 结果块拆分，取 h2 内标题+链接 与 p 摘要 ──

        Private Shared Function ParseResults(html As String, maxResults As Integer) As List(Of Dictionary(Of String, String))
            Dim results As New List(Of Dictionary(Of String, String))()

            Dim anchorRe = New Regex("<h2[^>]*>\s*<a[^>]*href=""([^""]+)""[^>]*>(.*?)</a>",
                RegexOptions.IgnoreCase Or RegexOptions.Singleline)
            Dim snippetRe = New Regex("<p[^>]*>(.*?)</p>",
                RegexOptions.IgnoreCase Or RegexOptions.Singleline)

            For Each block In Regex.Split(html, "(?=class=""b_algo"")")
                If Not block.Contains("b_algo") Then Continue For

                Dim a = anchorRe.Match(block)
                If Not a.Success Then Continue For

                Dim rawUrl = a.Groups(1).Value.Trim()
                If Not IsExternalUrl(rawUrl) Then Continue For

                Dim title = StripHtml(a.Groups(2).Value)
                Dim snippet As String = ""
                Dim p = snippetRe.Match(block)
                If p.Success Then snippet = StripHtml(p.Groups(1).Value)

                results.Add(New Dictionary(Of String, String) From {
                    {"title", title},
                    {"url", rawUrl},
                    {"snippet", snippet}
                })

                If results.Count >= maxResults Then Exit For
            Next

            Return results
        End Function

        ' 只保留真实外部网页链接；过滤 Bing 站内功能页（搜索/图片/视频/新闻等）与空链接
        Private Shared Function IsExternalUrl(link As String) As Boolean
            If String.IsNullOrWhiteSpace(link) Then Return False
            If Not link.StartsWith("http", StringComparison.OrdinalIgnoreCase) Then Return False
            Dim low = link.ToLowerInvariant()
            If low.Contains("bing.com/search") OrElse low.Contains("bing.com/speca") Then Return False
            If low.Contains("/images") OrElse low.Contains("/videos") OrElse low.Contains("/news") _
                OrElse low.Contains("/maps") OrElse low.Contains("/shop") Then Return False
            Return True
        End Function

        Private Shared Function StripHtml(input As String) As String
            If String.IsNullOrEmpty(input) Then Return input
            Dim s = Regex.Replace(input, "<[^>]+>", " ")
            s = System.Net.WebUtility.HtmlDecode(s)
            s = Regex.Replace(s, "\s+", " ").Trim()
            Return s
        End Function

        ' ── 共享辅助（与 ComputerAgentTools 一致的取参/返回约定） ──

        Private Shared Function G(Of T)(a As Dictionary(Of String, Object), n As String, d As T) As T
            If a Is Nothing Then Return d
            Dim v As Object = Nothing
            If a.TryGetValue(n, v) AndAlso v IsNot Nothing Then
                Try
                    Return CType(v, T)
                Catch
                    Try
                        Return DirectCast(Convert.ChangeType(v, GetType(T)), T)
                    Catch
                    End Try
                End Try
            End If
            Return d
        End Function

        Private Shared Function OK(s As String) As ToolsCallResult
            Return New ToolsCallResult() With {
                .Content = New List(Of McpContentItem) From {New McpTextContent With {.Text = s}}
            }
        End Function

        Private Shared Function Err(s As String) As ToolsCallResult
            Return New ToolsCallResult() With {
                .IsError = True,
                .Content = New List(Of McpContentItem) From {New McpTextContent With {.Text = s}}
            }
        End Function

    End Class

End Namespace