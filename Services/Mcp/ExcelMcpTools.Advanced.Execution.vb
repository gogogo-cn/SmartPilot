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

        Private Function ExecuteRunPythonScript(arguments As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim script = GetArg(Of String)(arguments, "script", "")
                Dim timeout = GetArg(Of Integer)(arguments, "timeout", 120)
                If String.IsNullOrEmpty(script) Then
                    result.Content.Add(New McpTextContent With {.Text = "必须指定 script 参数"})
                    result.IsError = True
                    Return result
                End If
                ' 提取引用区域数据，注入 Python 环境（包含完整 rows）
                ' 提前提取：当前选区为空/读取失败时，可回退到引用区域数据
                Dim referenceData As List(Of Object) = Nothing
                Try
                    Dim refToken = arguments("referenceRegions")
                    If refToken IsNot Nothing Then
                        referenceData = DirectCast(refToken, Newtonsoft.Json.Linq.JToken).ToObject(Of List(Of Object))()
                        System.Diagnostics.Debug.WriteLine("[ExecuteRunPythonScript] 注入引用区域数量: " & referenceData.Count)
                    End If
                Catch refEx As Exception
                    System.Diagnostics.Debug.WriteLine("[ExecuteRunPythonScript] 读取引用区域失败: " & refEx.Message)
                End Try
                ' 读取当前选中区域作为 input_data 注入脚本
                ' 注意：当前选区可能是单个单元格/空，导致读取失败，此时不直接报错，留待下方回退到引用区域
                Dim readResult = ExecuteReadRange(New Dictionary(Of String, Object)())
                Dim jsonText As String = "[]"
                If Not readResult.IsError Then
                    jsonText = If(readResult.Content.Count > 0 AndAlso TypeOf readResult.Content(0) Is McpTextContent,
                                  DirectCast(readResult.Content(0), McpTextContent).Text,
                                  "[]")
                Else
                    Dim readErrText = "当前选区读取失败"
                    Try
                        If readResult.Content IsNot Nothing AndAlso readResult.Content.Count > 0 Then
                            readErrText = readResult.Content(0).ToString()
                        End If
                    Catch
                    End Try
                    System.Diagnostics.Debug.WriteLine("[ExecuteRunPythonScript] 当前选区读取失败，尝试回退引用区域: " & readErrText)
                End If
                ' 防御性反序列化：ExecuteReadRange 返回结构化 JSON 对象（含 workbook/worksheet/result 等字段）
                ' 需要提取 .result 数组作为 Python input_data；兼容旧版直接返回数组的格式
                Dim inputData As Object = Nothing
                Try
                    ' 先尝试直接作为数组反序列化（旧版格式）
                    inputData = JsonConvert.DeserializeObject(Of List(Of Object))(jsonText)
                Catch
                    ' 新版格式：JSON 对象，提取 .result 数组
                    Try
                        Dim structured = Newtonsoft.Json.Linq.JObject.Parse(jsonText)
                        Dim resultToken = structured("result")
                        If resultToken IsNot Nothing Then
                            inputData = resultToken.ToObject(Of List(Of Object))()
                            System.Diagnostics.Debug.WriteLine("[ExecuteRunPythonScript] 已从结构化对象中提取 result 数组，行数: " & DirectCast(inputData, List(Of Object)).Count)
                        Else
                            ' 没有 result 字段，把整个对象包装为单元素列表
                            Dim wrapper As New List(Of Object)
                            wrapper.Add(structured)
                            inputData = wrapper
                            System.Diagnostics.Debug.WriteLine("[ExecuteRunPythonScript] JSON 对象无 result 字段，包装为单元素列表")
                        End If
                    Catch objEx As Exception
                        System.Diagnostics.Debug.WriteLine("[ExecuteRunPythonScript] 反序列化失败: " & objEx.Message)
                    End Try
                End Try
                ' 当前选区为空/读取失败时，回退到引用区域数据（用户请求中引用的区域），
                ' 避免因执行时当前选区变化导致 input_data 为空而误报"模板执行错误"
                If (inputData Is Nothing OrElse DirectCast(inputData, List(Of Object)).Count = 0) AndAlso referenceData IsNot Nothing Then
                    For Each refItem In referenceData
                        Dim refObj = TryCast(refItem, Newtonsoft.Json.Linq.JObject)
                        If refObj IsNot Nothing Then
                            Dim rowsToken = refObj("rows")
                            If rowsToken IsNot Nothing AndAlso rowsToken.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                                Dim rows = rowsToken.ToObject(Of List(Of Object))()
                                If rows IsNot Nothing AndAlso rows.Count > 0 Then
                                    inputData = rows
                                    System.Diagnostics.Debug.WriteLine("[ExecuteRunPythonScript] 当前选区为空，已回退到引用区域数据，行数: " & rows.Count)
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
                If inputData Is Nothing OrElse DirectCast(inputData, List(Of Object)).Count = 0 Then
                    result.Content.Add(New McpTextContent With {.Text = "当前未选中任何数据，无法执行 Python 脚本"})
                    result.IsError = True
                    Return result
                End If
                ' 提取表头信息，注入 Python 环境供脚本使用
                Dim headers As New List(Of String)
                Try
                    Dim structuredForHeaders = Newtonsoft.Json.Linq.JObject.Parse(jsonText)
                    Dim headersToken = structuredForHeaders("headers")
                    If headersToken IsNot Nothing Then
                        headers = headersToken.ToObject(Of List(Of String))()
                    ElseIf inputData IsNot Nothing AndAlso DirectCast(inputData, List(Of Object)).Count > 0 Then
                        Dim firstRow = DirectCast(inputData, List(Of Object))(0)
                        If TypeOf firstRow Is IDictionary(Of String, Object) Then
                            For Each key In DirectCast(firstRow, IDictionary(Of String, Object)).Keys
                                headers.Add(key)
                            Next
                        End If
                    End If
                Catch headerEx As Exception
                    System.Diagnostics.Debug.WriteLine("[ExecuteRunPythonScript] 提取表头失败: " & headerEx.Message)
                End Try
                ' 提取当前选区结构化数据，注入 Python 环境作为 selection_data
                Dim selectionData As Dictionary(Of String, Object) = Nothing
                Try
                    Dim structured = Newtonsoft.Json.Linq.JObject.Parse(jsonText)
                    selectionData = New Dictionary(Of String, Object) From {
                        {"workbookName", If(structured("workbook")?.ToString(), "")},
                        {"sheetName", If(structured("worksheet")?.ToString(), "")},
                        {"address", If(structured("range")?.ToString(), "")},
                        {"headers", If(structured("headers")?.ToObject(Of List(Of String))(), New List(Of String)())},
                        {"rows", If(structured("result")?.ToObject(Of List(Of Object))(), New List(Of Object)())}
                    }
                    System.Diagnostics.Debug.WriteLine("[ExecuteRunPythonScript] 注入当前选区: " & selectionData("address").ToString())
                Catch selEx As Exception
                    System.Diagnostics.Debug.WriteLine("[ExecuteRunPythonScript] 读取当前选区失败: " & selEx.Message)
                End Try
                ' 使用 Task.Run 避免 COM 同步上下文中的死锁
                ' 将异步执行放到线程池，独立处理超时
                Dim execResult As PythonResult = Nothing
                Try
                    execResult = Task.Run(Function()
                        Dim svc = PythonExecutionService.Instance
                        Dim t = svc.ExecuteScriptAsync(script, inputData, True, headers, referenceData, selectionData)
                        t.Wait(TimeSpan.FromSeconds(timeout + 60))
                        If t.IsCompleted Then
                            Return t.Result
                        End If
                        Return New PythonResult With {
                            .Success = False,
                            .ErrorMessage = $"Python 脚本执行超时（超过 {timeout + 60} 秒）。首次安装 Python 库可能需要更长时间。",
                            .ExitCode = -1
                        }
                    End Function).Result
                Catch aggEx As AggregateException
                    Dim innerEx = aggEx.Flatten().InnerException
                    Throw New Exception("Python 执行异常: " & innerEx?.Message, innerEx)
                End Try
                If execResult.Success Then
                    ' 检查 Python 输出的 success 字段：模板/脚本返回 success=False 时，
                    ' 视为业务失败并上报具体 message，而不是笼统地显示"执行成功"
                    If execResult.OutputData IsNot Nothing AndAlso
                       execResult.OutputData("success") IsNot Nothing AndAlso
                       execResult.OutputData("success").Type = Newtonsoft.Json.Linq.JTokenType.Boolean AndAlso
                       Not execResult.OutputData("success").ToObject(Of Boolean)() Then
                        Dim failMsg = "脚本执行未成功"
                        If execResult.OutputData("message") IsNot Nothing Then
                            failMsg = execResult.OutputData("message").ToString()
                        End If
                        result.Content.Add(New McpTextContent With {.Text = $"脚本执行失败: {failMsg}"})
                        result.IsError = True
                        Return result
                    End If
                    Dim output As New System.Text.StringBuilder()
                    output.AppendLine("Python 脚本执行成功")
                    ' 解析图表路径（[CHART] 标记）
                    Dim chartPaths As New List(Of String)()
                    If Not String.IsNullOrEmpty(execResult.StandardOutput) Then
                        For Each line In execResult.StandardOutput.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None)
                            If line.StartsWith("[CHART]") Then
                                chartPaths.Add(line.Substring(7))
                            End If
                        Next
                    End If
                    ' 添加图表信息
                    If chartPaths.Count > 0 Then
                        output.AppendLine($"生成了 {chartPaths.Count} 个图表:")
                        For Each chartPath In chartPaths
                            output.AppendLine($"   {chartPath}")
                        Next
                        output.AppendLine()
                    End If
                    ' 添加标准输出（过滤掉 [CHART] 行）
                    If Not String.IsNullOrEmpty(execResult.StandardOutput) Then
                        Dim lines = execResult.StandardOutput.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None)
                        Dim cleanOutput = String.Join(vbCrLf, lines.Where(Function(l) Not l.StartsWith("[CHART]")))
                        If Not String.IsNullOrWhiteSpace(cleanOutput) Then
                            output.AppendLine(cleanOutput)
                        End If
                    End If
                    result.Content.Add(New McpTextContent With {.Text = output.ToString()})
                    ' 【修复-趋势图表】附加结构化 OutputData（含图表配置 JSON，如 forecast_x/forecast 等），
                    ' 供前端 [OUTPUT_JSON] 标记优先解析出 ECharts 图表，避免结构化数据被纯文本提取丢弃
                    If execResult.OutputData IsNot Nothing Then
                        result.Content.Add(New McpTextContent With {
                            .Text = "[OUTPUT_JSON]" & vbCrLf &
                                    execResult.OutputData.ToString(Newtonsoft.Json.Formatting.None)
                        })
                    End If
                Else
                    result.Content.Add(New McpTextContent With {.Text = $"Python 脚本执行失败: {execResult.ErrorMessage}"})
                    result.IsError = True
                End If
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"执行 Python 脚本失败: {ex.Message}"})
            End Try
            Return result
        End Function

        Private Function ExecuteTakeScreenshot(arguments As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Dim excelApp = Globals.ThisAddIn.Application
            If excelApp Is Nothing Then
                result.Content.Add(New McpTextContent With {.Text = "Excel 应用程序未就绪"})
                result.IsError = True
                Return result
            End If
            Dim wb = excelApp.ActiveWorkbook
            If wb Is Nothing Then
                result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                result.IsError = True
                Return result
            End If
            Try
                Dim rangeAddress = GetArg(Of String)(arguments, "range", "")
                Dim outputPath = GetArg(Of String)(arguments, "outputPath", "")
                If String.IsNullOrEmpty(outputPath) Then
                    outputPath = IO.Path.Combine(IO.Path.GetTempPath(), $"excel_screenshot_{DateTime.Now:yyyyMMddHHmmss}.png")
                End If
                ' 截图功能需要 ChatLogger，这里返回提示信息
                result.Content.Add(New McpTextContent With {.Text = $"截图功能需要通过界面操作，或使用 ScreenshotService.StartOcr() 方法。目标路径: {outputPath}"})
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"截图失败: {ex.Message}"})
            Finally
                Utils.ComReleaseHelper.Release(wb)
            End Try
            Return result
        End Function

        Private Function ExecuteDisplayHtmlInBrowser(arguments As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim htmlContent = If(arguments.ContainsKey("htmlContent"), arguments("htmlContent")?.ToString(), "")
                Dim title = If(arguments.ContainsKey("title") AndAlso arguments("title") IsNot Nothing,
                               arguments("title").ToString(),
                               "选中数据报表")
                Dim rangeArg = If(arguments.ContainsKey("range"), arguments("range")?.ToString(), "")
                Dim sheetArg = If(arguments.ContainsKey("sheet"), arguments("sheet")?.ToString(), "")
                ' 如果未传 htmlContent，自动读取选区数据并生成表格 HTML
                If String.IsNullOrWhiteSpace(htmlContent) Then
                    htmlContent = BuildRangeTableHtml(rangeArg, sheetArg, title, result)
                    If result.IsError Then Return result
                End If
                ' 直接复用 BrowserDisplayBridge 的能力
                Dim bridge = Bridges.BrowserDisplayBridge.Instance
                Dim jsonResult = bridge.OpenHtmlInBrowser(htmlContent, title)
                Dim parsed = JsonConvert.DeserializeObject(Of JObject)(jsonResult)
                If parsed IsNot Nothing AndAlso parsed("success")?.Value(Of Boolean)() Then
                    Dim filePath = parsed("filePath")?.ToString()
                    result.Content.Add(New McpTextContent With {
                        .Text = $"HTML 页面已在浏览器中打开：{filePath}"
                    })
                Else
                    result.IsError = True
                    Dim errMsg = If(parsed?("error")?.ToString(), "打开 HTML 页面失败")
                    result.Content.Add(New McpTextContent With {.Text = errMsg})
                End If
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"展示 HTML 失败: {ex.Message}"})
            End Try
            Return result
        End Function

        Private Function BuildRangeTableHtml(rangeAddr As String, sheetName As String, title As String, ByRef result As ToolsCallResult) As String
            Dim excelApp = Globals.ThisAddIn.Application
            If excelApp Is Nothing Then
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "Excel 应用程序未就绪"})
                Return ""
            End If
            If excelApp.ActiveWorkbook Is Nothing Then
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿"})
                Return ""
            End If
            Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim rng As Microsoft.Office.Interop.Excel.Range = Nothing
            Try
                ' 确定工作表
                If Not String.IsNullOrEmpty(sheetName) Then
                    ws = DirectCast(excelApp.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                Else
                    ws = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                End If
                ' 确定选区（与 ExecuteClearRange 相同的选区决策链）
                If Not String.IsNullOrEmpty(rangeAddr) Then
                    rng = ws.Range(rangeAddr)
                Else
                    Dim ctxAddr As String = ""
                    If RangeContext.TryGet(ws.Name, ctxAddr) Then
                        Try
                            rng = ws.Range(ctxAddr)
                        Catch
                            rng = Nothing
                            RangeContext.Clear()
                        End Try
                    End If
                    If rng Is Nothing Then
                        Try
                            Dim sel = excelApp.Selection
                            If TypeOf sel Is Microsoft.Office.Interop.Excel.Range Then
                                Dim selRng = DirectCast(sel, Microsoft.Office.Interop.Excel.Range)
                                Dim selSheetName As String = CStr(selRng.Worksheet.Name)
                                If selSheetName = ws.Name Then
                                    rng = selRng
                                Else
                                    Utils.ComReleaseHelper.Release(selRng)
                                End If
                            End If
                        Catch
                            rng = Nothing
                        End Try
                    End If
                    If rng Is Nothing Then
                        rng = ws.UsedRange
                    End If
                End If
                ' 读取数据
                Dim data As Object(,) = Nothing
                Dim rows As Integer = 0, cols As Integer = 0
                Try
                    data = DirectCast(rng.Value2, Object(,))
                    If data IsNot Nothing Then
                        rows = data.GetLength(0)
                        cols = data.GetLength(1)
                    End If
                Catch
                    data = Nothing
                End Try
                If data Is Nothing OrElse rows = 0 OrElse cols = 0 Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "选区为空，没有可展示的数据"})
                    Return ""
                End If
                ' 生成 HTML 表格
                Dim sb As New System.Text.StringBuilder()
                Dim usedAddr As String = rng.Address(False, False)
                Dim wbName As String = excelApp.ActiveWorkbook.Name
                Dim wsName As String = ws.Name
                sb.AppendLine("<!DOCTYPE html>")
                sb.AppendLine("<html lang=""zh-CN""><head><meta charset=""UTF-8"">")
                sb.AppendLine($"<title>{EscapeHtml(title)}</title>")
                sb.AppendLine("<style>")
                sb.AppendLine("body{font-family:-apple-system,BlinkMacSystemFont,""Segoe UI"",Roboto,Arial,sans-serif;margin:24px;background:#f8fafc;color:#1e293b}")
                sb.AppendLine(".container{max-width:1200px;margin:0 auto;background:#fff;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,.1);padding:24px}")
                sb.AppendLine("h1{font-size:20px;margin:0 0 8px}")
                sb.AppendLine(".subtitle{color:#64748b;font-size:13px;margin-bottom:16px}")
                sb.AppendLine("table{width:100%;border-collapse:collapse;font-size:13px}")
                sb.AppendLine("th,td{border:1px solid #e2e8f0;padding:8px 12px;text-align:left}")
                sb.AppendLine("th{background:#f1f5f9;font-weight:600}")
                sb.AppendLine("tr:nth-child(even){background:#f8fafc}")
                sb.AppendLine(".chart-wrap{width:100%;height:380px;margin-top:20px;padding-top:16px;border-top:1px solid #e2e8f0}")
                sb.AppendLine("#report-chart{width:100%;height:340px}")
                sb.AppendLine(".chart-wrap .chart-title{font-size:15px;font-weight:600;color:#1e293b;margin-bottom:8px}")
                sb.AppendLine(".chart-fallback{color:#94a3b8;font-size:12px;padding:12px;text-align:center;border:1px dashed #cbd5e1;border-radius:6px}")
                sb.AppendLine(".kpis{display:flex;gap:12px;flex-wrap:wrap;margin:12px 0}")
                sb.AppendLine(".kpi{background:#f1f5f9;border-radius:6px;padding:10px 14px;min-width:120px}")
                sb.AppendLine(".kpi .kpi-label{font-size:11px;color:#64748b}")
                sb.AppendLine(".kpi .kpi-value{font-size:18px;font-weight:600;color:#0f172a;margin-top:2px}")
                sb.AppendLine("</style></head><body><div class=""container"">")
                sb.AppendLine($"<h1>{EscapeHtml(title)}</h1>")
                sb.AppendLine($"<div class=""subtitle"">{EscapeHtml(wbName)} / {EscapeHtml(wsName)} / {EscapeHtml(usedAddr)} · 共 {rows} 行</div>")
                sb.AppendLine("<table><thead><tr>")
                ' 表头（第一行）
                For c As Integer = 1 To cols
                    Dim hv As Object = data(1, c)
                    Dim hStr As String = If(hv Is Nothing OrElse hv Is DBNull.Value, $"列{c}", FormatCellValueForDisplay(hv))
                    sb.Append($"<th>{EscapeHtml(hStr)}</th>")
                Next
                sb.AppendLine("</tr></thead><tbody>")
                ' 数据行（从第2行开始）
                For r As Integer = 2 To rows
                    sb.Append("<tr>")
                    For c As Integer = 1 To cols
                        Dim v As Object = data(r, c)
                        Dim vStr As String = If(v Is Nothing OrElse v Is DBNull.Value, "", FormatCellValueForDisplay(v))
                        sb.Append($"<td>{EscapeHtml(vStr)}</td>")
                    Next
                    sb.AppendLine("</tr>")
                Next
                sb.AppendLine("</tbody></table>")

                ' ── 升级为"真报表"：在表格之后自动注入 KPI 汇总 + ECharts 图表 ──
                ' 分类轴取第一列，数值列取其余列中含数值的列；无数值列时不注入图表（保持纯表格）。
                Dim categories As New List(Of Object)
                Dim seriesList As New List(Of Dictionary(Of String, Object))
                Dim kpis As New List(Of Dictionary(Of String, Object))
                For r As Integer = 2 To rows
                    Dim catVal = data(r, 1)
                    Dim catStr = If(catVal Is Nothing OrElse catVal Is DBNull.Value, "", FormatCellValueForDisplay(catVal))
                    ' 防 script 闭合注入（分类文本来自 Excel 单元格）
                    categories.Add(catStr.Replace("</script", String.Concat("<", "\/script")))
                Next
                For c As Integer = 2 To cols
                    Dim isNumCol As Boolean = False
                    Dim colVals As New List(Of Object)  ' 与 categories 对齐，非数值位置置 Nothing
                    Dim colSum As Double = 0
                    Dim colCnt As Integer = 0
                    For r As Integer = 2 To rows
                        Dim v = data(r, c)
                        Dim d As Double = 0
                        If TryParseNumber(v, d) Then
                            isNumCol = True
                            colVals.Add(d)
                            colSum += d
                            colCnt += 1
                        Else
                            colVals.Add(Nothing)
                        End If
                    Next
                    If Not isNumCol Then Continue For
                    Dim colHead = data(1, c)
                    Dim colName = If(colHead Is Nothing OrElse colHead Is DBNull.Value, $"列{c}", FormatCellValueForDisplay(colHead))
                    seriesList.Add(New Dictionary(Of String, Object) From {
                        {"name", colName},
                        {"data", colVals}
                    })
                    If colCnt > 0 Then
                        kpis.Add(New Dictionary(Of String, Object) From {
                            {"label", colName & " 合计"},
                            {"value", colSum.ToString("N2")}
                        })
                    End If
                Next
                If seriesList.Count > 0 AndAlso categories.Count > 0 Then
                    sb.AppendLine(BuildReportChartsHtml(categories, seriesList, kpis))
                End If

                sb.AppendLine("<div class=""footer"">Generated by SmartPilot</div>")
                sb.AppendLine("</div></body></html>")
                Return sb.ToString()
            Finally
                Utils.ComReleaseHelper.Release(rng)
                Utils.ComReleaseHelper.Release(ws)
            End Try
        End Function

        Private Function FormatCellValueForDisplay(val As Object) As String
            If val Is Nothing OrElse val Is DBNull.Value Then Return ""
            If TypeOf val Is Double OrElse TypeOf val Is Single OrElse TypeOf val Is Decimal Then
                Dim numVal As Double = Convert.ToDouble(val)
                If numVal >= 1 AndAlso numVal <= 100000 Then
                    Try
                        Dim dt As DateTime = DateTime.FromOADate(numVal)
                        ' 简单判断是否可能为日期（年在合理范围）
                        If dt.Year >= 1990 AndAlso dt.Year <= 2100 Then
                            Return dt.ToString("yyyy-MM-dd")
                        End If
                    Catch
                    End Try
                End If
                Return val.ToString()
            End If
            Return val.ToString()
        End Function

        Private Function EscapeHtml(text As String) As String
            If text Is Nothing Then Return ""
            Return text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("""", "&quot;").Replace("'", "&#39;")
        End Function

        ''' <summary>解析单元格为数值；成功输出 Double。</summary>
        Private Function TryParseNumber(val As Object, ByRef out As Double) As Boolean
            If val Is Nothing OrElse val Is DBNull.Value Then Return False
            If TypeOf val Is Double OrElse TypeOf val Is Single OrElse TypeOf val Is Decimal OrElse
               TypeOf val Is Integer OrElse TypeOf val Is Long Then
                out = Convert.ToDouble(val)
                Return True
            End If
            Dim s = val.ToString()
            If String.IsNullOrWhiteSpace(s) Then Return False
            Return Double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out)
        End Function

        ' 生成"真报表"增强片段：KPI 汇总卡片 + ECharts 图表。
        ' 数据序列化后注入 script；ECharts 用 CDN 加载，离线时优雅降级为纯表格。
        Private Function BuildReportChartsHtml(categories As List(Of Object),
                                               seriesList As List(Of Dictionary(Of String, Object)),
                                               kpis As List(Of Dictionary(Of String, Object))) As String
            Dim sb As New System.Text.StringBuilder()
            Dim catJson = Newtonsoft.Json.JsonConvert.SerializeObject(categories)
            Dim seriesJson = Newtonsoft.Json.JsonConvert.SerializeObject(seriesList)
            Dim kpiJson = Newtonsoft.Json.JsonConvert.SerializeObject(kpis)

            sb.AppendLine("<div class=""kpis"" id=""report-kpis""></div>")
            sb.AppendLine("<div class=""chart-wrap"">")
            sb.AppendLine("<div class=""chart-title"">数据可视化</div>")
            sb.AppendLine("<div id=""report-chart""></div>")
            sb.AppendLine("</div>")
            sb.AppendLine("<script src=""//cdn.jsdelivr.net/npm/echarts@5/dist/echarts.min.js""></script>")
            sb.AppendLine("<script>")
            sb.AppendLine("(function(){")
            sb.AppendLine("var cat=" & catJson & ";")
            sb.AppendLine("var ser=" & seriesJson & ";")
            sb.AppendLine("var kpi=" & kpiJson & ";")
            sb.AppendLine("var kpiEl=document.getElementById('report-kpis');")
            sb.AppendLine("if(kpi&&kpi.length){kpiEl.innerHTML=kpi.map(function(k){return '<div class=""kpi""><div class=""kpi-label"">'+k.label.replace(/</g,'&lt;')+'</div><div class=""kpi-value"">'+k.value.replace(/</g,'&lt;')+'</div></div>';}).join('');}")
            sb.AppendLine("var el=document.getElementById('report-chart');")
            sb.AppendLine("if(window.echarts&&el){")
            sb.AppendLine("var ch=echarts.init(el);")
            sb.AppendLine("var opt={tooltip:{trigger:'axis'},legend:{top:0},grid:{left:50,right:24,bottom:44,top:36},xAxis:{type:'category',data:cat,axisLabel:{interval:0,rotate:30}},yAxis:{type:'value'},series:ser.map(function(s){return {name:s.name,type:'bar',data:s.data};})};")
            sb.AppendLine("ch.setOption(opt);")
            sb.AppendLine("window.addEventListener('resize',function(){ch.resize();});")
            sb.AppendLine("}else{el.innerHTML='<div class=""chart-fallback"">图表组件加载失败（需联网加载 ECharts）</div>';}")
            sb.AppendLine("})();")
            sb.AppendLine("</script>")
            Return sb.ToString()
        End Function

        Private Function ExecuteClipboardRead(arguments As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim clipContent = My.Computer.Clipboard.GetText()
                If String.IsNullOrEmpty(clipContent) Then
                    result.Content.Add(New McpTextContent With {.Text = "剪贴板为空"})
                Else
                    result.Content.Add(New McpTextContent With {.Text = clipContent})
                End If
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"读取剪贴板失败: {ex.Message}"})
            End Try
            Return result
        End Function

        Private Function ExecuteClipboardWrite(arguments As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim content = GetArg(Of String)(arguments, "content", "")
                If String.IsNullOrEmpty(content) Then
                    result.Content.Add(New McpTextContent With {.Text = "必须指定 content 参数"})
                    result.IsError = True
                    Return result
                End If
                My.Computer.Clipboard.SetText(content)
                result.Content.Add(New McpTextContent With {.Text = "内容已写入剪贴板"})
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"写入剪贴板失败: {ex.Message}"})
            End Try
            Return result
        End Function

    End Class

End Namespace
