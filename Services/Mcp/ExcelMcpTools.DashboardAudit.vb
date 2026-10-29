Option Strict On
Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Runtime.InteropServices
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core
Imports ExcelAIVSTOAddIn.Utils

Namespace Services.Mcp

    Partial Public Class ExcelMcpTools

        ''' <summary>
        ''' excel_capture_dashboard — 截取看板工作表区域为 PNG base64。
        ''' 用 Range.CopyPicture + 临时 Chart 导出，不依赖屏幕截图。
        ''' </summary>
        Private Function ExecuteCaptureDashboard(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result As New ToolsCallResult()
            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "看板")
                Dim rangeAddr As String = GetArg(Of String)(args, "range", "A1:Q32")
                Dim excelApp = DirectCast(Globals.ThisAddIn.Application, Microsoft.Office.Interop.Excel.Application)
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿。"})
                    Return result
                End If
                Dim wb = DirectCast(excelApp.ActiveWorkbook, Microsoft.Office.Interop.Excel.Workbook)
                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Try : ws = DirectCast(wb.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet) : Catch : End Try
                If ws Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"找不到工作表 [{sheetName}]"})
                    Return result
                End If

                ' 解析锚点范围，算看板边界
                Dim rng As Microsoft.Office.Interop.Excel.Range = Nothing
                Try
                    rng = ws.Range(rangeAddr)
                Catch
                    rng = Nothing
                End Try
                If rng Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"无法解析范围 {rangeAddr}"})
                    Return result
                End If

                Dim left = CDbl(rng.Left), top = CDbl(rng.Top)
                Dim width = CDbl(rng.Width), height = CDbl(rng.Height)

                ' 复制区域为图片
                rng.CopyPicture(XlPictureAppearance.xlScreen, XlCopyPictureFormat.xlBitmap)

                ' 在同位置建临时 Chart 作为图片容器
                Dim co As ChartObjects = CType(ws.ChartObjects(), ChartObjects)
                Dim tmpCo As ChartObject = co.Add(left, top, width, height)
                tmpCo.Border.LineStyle = XlLineStyle.xlLineStyleNone
                Dim tmpChart = tmpCo.Chart
                tmpChart.ChartArea.Format.Fill.Visible = MsoTriState.msoFalse
                tmpChart.Paste()

                ' 导出为临时 PNG
                Dim tmpPath = Path.Combine(Path.GetTempPath(), "dashboard_capture_" & Environment.TickCount & ".png")
                tmpChart.Export(tmpPath, "PNG")

                ' 删除临时 Chart（粘贴的图片在 Chart 内，随 Chart 一起删除）
                Try : tmpCo.Delete() : Catch : End Try

                ' 读 PNG 转 base64
                Dim bytes = File.ReadAllBytes(tmpPath)
                Dim b64 = Convert.ToBase64String(bytes)
                Try : File.Delete(tmpPath) : Catch : End Try

                result.Content.Add(New McpTextContent With {
                    .Text = $"{{""ok"":true,""imageBase64"":""{b64}"",""width"":{CInt(width)},""height"":{CInt(height)},""sheet"":""{sheetName}"",""range"":""{rangeAddr}""}}"
                })
                Return result
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"截图失败：{ex.Message}"})
                Return result
            End Try
        End Function

        ''' <summary>
        ''' excel_audit_dashboard — 调视觉模型审查看板截图，返回结构化问题列表 JSON。
        ''' 入参：imageBase64（或先调 capture_dashboard）、sheet、focus（可选审查重点）
        ''' </summary>
        Private Function ExecuteAuditDashboard(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result As New ToolsCallResult()
            Try
                Dim imageB64 As String = GetArg(Of String)(args, "imageBase64", "")
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "看板")
                Dim focus As String = GetArg(Of String)(args, "focus", "")

                ' 如果没传 imageBase64，先自动截
                If String.IsNullOrWhiteSpace(imageB64) Then
                    Dim capArgs As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                    capArgs("sheet") = sheetName
                    Dim capRes = ExecuteCaptureDashboard(capArgs)
                    If capRes.IsError Then
                        result.IsError = True
                        result.Content.AddRange(capRes.Content)
                        Return result
                    End If
                    ' 从 capture 结果里解析 imageBase64
                    Try
                        Dim txt = CType(capRes.Content(0), McpTextContent).Text
                        Dim jo = JObject.Parse(txt)
                        imageB64 = jo.Value(Of String)("imageBase64")
                    Catch
                    End Try
                End If

                If String.IsNullOrWhiteSpace(imageB64) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "没有可用的截图（imageBase64 为空，自动截图也失败）"})
                    Return result
                End If

                ' 构造视觉审查 prompt
                Dim prompt = "你是 Excel 看板 UI 审查专家。请审查这张数据看板截图，找出布局/审美问题。" & vbCrLf &
                    "审查清单：" & vbCrLf &
                    "1. KPI 卡片是否整齐对齐？有无文字溢出/裁切？" & vbCrLf &
                    "2. 图表之间是否重叠？大小是否一致？" & vbCrLf &
                    "3. 切片器是否和图表重叠？位置是否合理？" & vbCrLf &
                    "4. 配色是否协调？是否有刺眼/不专业的颜色？" & vbCrLf &
                    "5. 留白是否合理？有无大片空白或拥挤？" & vbCrLf &
                    "6. 标题/标签是否清晰可读？" & vbCrLf &
                    (If(String.IsNullOrEmpty(focus), "", "特别关注：" & focus & vbCrLf)) &
                    vbCrLf &
                    "只返回 JSON，不要解释：" & vbCrLf &
                    "{""issues"":[{""type"":""kpi_alignment|chart_overlap|slicer_overflow|color|spacing|text_readability"",""severity"":""high|medium|low"",""location"":""描述位置"",""problem"":""问题描述"",""fix"":""具体修正建议（如：行高调到22、图表右移一列、卡片底色改浅）""}],""overall_score"":1-10}"

                ' 调多模态引擎
                Dim engine = New MultimodalLLMEngine()
                Dim llmResp = engine.ProcessImageAndTextAsync(imageB64, prompt, "image/png").GetAwaiter().GetResult()
                If llmResp Is Nothing OrElse Not llmResp.Success Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"视觉模型调用失败：{If(llmResp?.ErrorMessage, "未知错误")}"})
                    Return result
                End If

                Dim answer = If(llmResp.Content, "").ToString()
                ' 尝试提取 JSON
                Dim jsonStart = answer.IndexOf("{"c)
                Dim jsonEnd = answer.LastIndexOf("}"c)
                If jsonStart >= 0 AndAlso jsonEnd > jsonStart Then
                    Dim jsonStr = answer.Substring(jsonStart, jsonEnd - jsonStart + 1)
                    result.Content.Add(New McpTextContent With {.Text = jsonStr})
                Else
                    result.Content.Add(New McpTextContent With {.Text = answer})
                End If
                Return result
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"审查失败：{ex.Message}"})
                Return result
            End Try
        End Function

    End Class

End Namespace
