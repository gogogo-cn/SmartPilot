Option Strict On
Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core
Imports ExcelAIVSTOAddIn.Utils

Namespace Services.Mcp

    Partial Public Class ExcelMcpTools

        ' ═══════════════════════════════════════════════════════════════
        ' 一键专业看板布局引擎（硬编码网格 + 统一配色 + 卡片化 KPI）
        ' 设计目标：LLM 只传数据，不碰坐标/颜色/行高列宽，
        '           所有排版由本工具按专业 BI 网格算好，杜绝"散、乱、挤"。
        '
        ' 网格布局（字符列宽 / 磅行高）：
        '   A=2(边距)  B~P=11(主内容15列)  Q=2(间距)  R=14(切片器栏)  S=2(边距)
        '   Row1=8  Row2=30(标题)  Row3=8
        '   Row4~6=20(KPI卡3行高)  Row7=10
        '   Row8~19=17(图表行1)  Row20=10  Row21~32=17(图表行2)
        '
        '   KPI 4 卡: B4:E6 / F4:I6 / J4:M6 / N4:P6
        '   图表 2×2: B8:H19 / J8:P19 / B21:H32 / J21:P32
        '   切片器栏: R 列起，纵向排列，与 KPI 顶部对齐
        ' ═══════════════════════════════════════════════════════════════

        ' ── 专业 BI 配色板（硬编码，LLM 不可覆盖）──
        Private Const C_PRIMARY As String = "1F4E79"   ' 深蓝-标题/大数字
        Private Const C_SECONDARY As String = "2E75B6" ' 中蓝-图表主系列
        Private Const C_ACCENT As String = "ED7D31"    ' 橙-关键项强调
        Private Const C_SUCCESS As String = "70AD47"   ' 绿-正向
        Private Const C_WARNING As String = "FFC000"   ' 黄-中性提醒
        Private Const C_GRAY As String = "A5A5A5"      ' 灰-辅助
        Private Const C_CARD_FILL As String = "F2F6FB" ' 卡片底-极浅蓝
        Private Const C_CARD_LINE As String = "D6E4F0" ' 卡片边-浅蓝
        Private Const C_TITLE_BAR As String = "1F4E79" ' 标题栏底
        Private Const C_TEXT_MUTED As String = "808080" ' 辅助字灰

        ''' <summary>
        ''' excel_build_dashboard — 一键生成专业数据看板。
        ''' 入参只给数据（标题/KPI/图表），布局/配色/卡片/对齐由引擎硬编码决定。
        ''' </summary>
        Private Function ExecuteBuildDashboard(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result As New ToolsCallResult()
            Dim msgs As New List(Of String)()

            Try
                ' ── 1. 解析入参 ──
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "看板")
                Dim title As String = GetArg(Of String)(args, "title", "数据看板")
                Dim kpisJson As String = GetArg(Of String)(args, "kpis", "[]")
                Dim chartsJson As String = GetArg(Of String)(args, "charts", "[]")
                Dim showGridlines As Boolean = GetArg(Of Boolean)(args, "showGridlines", False)
                Dim freezePane As Boolean = GetArg(Of Boolean)(args, "freezeTitle", True)

                Dim excelApp As Microsoft.Office.Interop.Excel.Application =
                    DirectCast(Globals.ThisAddIn.Application, Microsoft.Office.Interop.Excel.Application)
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿。"})
                    Return result
                End If
                Dim wb As Microsoft.Office.Interop.Excel.Workbook = DirectCast(excelApp.ActiveWorkbook, Microsoft.Office.Interop.Excel.Workbook)

                ' ── 2. 创建或重建看板表（清掉旧形状/图表/切片器）──
                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim existed As Boolean = False
                Try
                    ws = DirectCast(wb.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                    existed = True
                Catch
                End Try
                If existed Then
                    ' 删除旧形状（含切片器、KPI 卡片）
                    Try
                        For i As Integer = CType(ws.Shapes, Microsoft.Office.Interop.Excel.Shapes).Count To 1 Step -1
                            Try : CType(ws.Shapes(i), Microsoft.Office.Interop.Excel.Shape).Delete() : Catch : End Try
                        Next
                    Catch : End Try
                    ' 删除旧图表
                    Try
                        For i As Integer = CType(ws.ChartObjects(), Microsoft.Office.Interop.Excel.ChartObjects).Count To 1 Step -1
                            Try : CType(ws.ChartObjects(i), Microsoft.Office.Interop.Excel.ChartObject).Delete() : Catch : End Try
                        Next
                    Catch : End Try
                    ' 清空单元格内容
                    Try : ws.Cells.Clear() : Catch : End Try
                Else
                    ws = DirectCast(wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)),
                                    Microsoft.Office.Interop.Excel.Worksheet)
                    ws.Name = sheetName
                End If
                msgs.Add($"工作表 [{sheetName}] 就绪")

                ' ── 3. 设定网格：列宽行高 ──
                ApplyDashboardGrid(ws)
                msgs.Add("网格(列宽/行高)已应用")

                ' 隐藏网格线（专业看板观感）
                Try
                    If Not showGridlines Then
                        ws.Application.ActiveWindow.DisplayGridlines = False
                    End If
                Catch : End Try

                ' ── 4. 标题栏（B2:Q2 深蓝长条，白色粗体左对齐）──
                DrawTitleBar(ws, title)
                msgs.Add("标题栏已绘制")

                ' ── 5. KPI 卡片（最多 4 张，等分 B:P 列）──
                Dim kpiRanges As String() = {"B4:E6", "F4:I6", "J4:M6", "N4:P6"}
                Dim kpis As JArray = Nothing
                Try : kpis = JArray.Parse(kpisJson) : Catch : kpis = New JArray()
                End Try
                Dim kpiCount As Integer = Math.Min(kpis.Count, 4)
                For i As Integer = 0 To kpiCount - 1
                    Dim kpi As JObject = TryCast(kpis(i), JObject)
                    If kpi Is Nothing Then Continue For
                    Dim label As String = kpi.Value(Of String)("label")
                    Dim value As String = kpi.Value(Of String)("value")
                    Dim sub_ As String = kpi.Value(Of String)("sub")
                    DrawKpiCard(ws, kpiRanges(i), label, value, sub_)
                Next
                msgs.Add($"KPI 卡片 {kpiCount} 张已绘制")

                ' ── 6. 图表 2×2 网格（复用 ExecuteCreateChart）──
                Dim chartAnchors As String() = {"B8:H19", "J8:P19", "B21:H32", "J21:P32"}
                Dim charts As JArray = Nothing
                Try : charts = JArray.Parse(chartsJson) : Catch : charts = New JArray()
                End Try
                Dim chartCount As Integer = Math.Min(charts.Count, 4)
                For i As Integer = 0 To chartCount - 1
                    Dim ch As JObject = TryCast(charts(i), JObject)
                    If ch Is Nothing Then Continue For
                    Try
                        Dim chartArgs As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                        chartArgs("sheet") = sheetName
                        chartArgs("title") = ch.Value(Of String)("title")
                        If ch("chartType") IsNot Nothing Then chartArgs("chartType") = ch.Value(Of String)("chartType")
                        If ch("dataRange") IsNot Nothing Then chartArgs("dataRange") = ch.Value(Of String)("dataRange")
                        If ch("dataValues") IsNot Nothing Then chartArgs("dataValues") = ch.Value(Of String)("dataValues")
                        chartArgs("position") = "{""anchorRange"":""" & chartAnchors(i) & """}"
                        chartArgs("multiChart") = True
                        ' 默认专业样式：有图例、显数据标签（小数值图）
                        If ch("hasLegend") IsNot Nothing Then chartArgs("hasLegend") = ch.Value(Of Boolean)("hasLegend")
                        If ch("showDataLabels") IsNot Nothing Then
                            chartArgs("showDataLabels") = ch.Value(Of Boolean)("showDataLabels")
                        Else
                            chartArgs("showDataLabels") = True
                        End If
                        Dim cr = ExecuteCreateChart(chartArgs)
                        If cr.IsError Then
                            msgs.Add($"⚠️ 图表[{i + 1}]生成失败: {FirstText(cr)}")
                        Else
                            msgs.Add($"图表[{i + 1}]已绘制 → {chartAnchors(i)}")
                        End If
                    Catch ex As Exception
                        msgs.Add($"⚠️ 图表[{i + 1}]异常: {ex.Message}")
                    End Try
                Next

                ' ── 7. 冻结标题区（前 3 行）──
                If freezePane Then
                    Try
                        ws.Application.ActiveWindow.SplitRow = 3
                        ws.Application.ActiveWindow.FreezePanes = True
                    Catch : End Try
                End If

                ' ── 8. 激活看板表 ──
                Try : ws.Activate() : Catch : End Try

                result.Content.Add(New McpTextContent With {
                    .Text = "✅ 专业看板已生成：" & sheetName & vbCrLf & String.Join(vbCrLf, msgs)})
                Return result

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"❌ 看板生成失败：{ex.Message}"})
                Return result
            End Try
        End Function

        ' ═══════════════════════════════════════════════════════════════
        ' 私有辅助：网格 / 标题 / KPI 卡片绘制
        ' ═══════════════════════════════════════════════════════════════

        ''' <summary>设定看板专用列宽行高（字符列宽 / 磅行高）。</summary>
        Private Sub ApplyDashboardGrid(ws As Microsoft.Office.Interop.Excel.Worksheet)
            ' 列宽（字符单位）
            SetColWidth(ws, "A", 2)
            For Each col As String In {"B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"}
                SetColWidth(ws, col, 11)
            Next
            SetColWidth(ws, "Q", 2)
            SetColWidth(ws, "R", 14)   ' 切片器栏
            SetColWidth(ws, "S", 2)

            ' 行高（磅）
            SetRowHeight(ws, 1, 8)
            SetRowHeight(ws, 2, 32)    ' 标题
            SetRowHeight(ws, 3, 8)
            For r As Integer = 4 To 6 : SetRowHeight(ws, r, 22) : Next   ' KPI 卡
            SetRowHeight(ws, 7, 10)
            For r As Integer = 8 To 19 : SetRowHeight(ws, r, 18) : Next ' 图表行1
            SetRowHeight(ws, 20, 10)
            For r As Integer = 21 To 32 : SetRowHeight(ws, r, 18) : Next ' 图表行2
        End Sub

        Private Shared Sub SetColWidth(ws As Microsoft.Office.Interop.Excel.Worksheet, col As String, chars As Double)
            Try : CType(ws.Columns(col), Microsoft.Office.Interop.Excel.Range).ColumnWidth = chars : Catch : End Try
        End Sub

        Private Shared Sub SetRowHeight(ws As Microsoft.Office.Interop.Excel.Worksheet, row As Integer, pts As Double)
            Try : CType(ws.Rows(row), Microsoft.Office.Interop.Excel.Range).RowHeight = pts : Catch : End Try
        End Sub

        ''' <summary>绘制深蓝标题栏：B2:Q2，白色粗体 18 号，左对齐留内边距。</summary>
        Private Sub DrawTitleBar(ws As Microsoft.Office.Interop.Excel.Worksheet, title As String)
            Dim args As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            args("sheet") = ws.Name
            args("range") = "B2:Q2"
            args("shapeType") = "rectangle"
            args("fill") = C_TITLE_BAR
            args("line") = "none"
            args("noShadow") = True
            args("sendToBack") = True
            ' 标题文字：白色 18 号粗体，左对齐（TextFrame2 自带 3pt 内边距）
            Dim titleLine As New JObject()
            titleLine("text") = title
            titleLine("size") = 18.0
            titleLine("bold") = True
            titleLine("color") = "FFFFFF"
            titleLine("alignH") = "left"
            args("lines") = "[" & titleLine.ToString(Formatting.None) & "]"
            Try
                ExecuteAddShape(args)
            Catch : End Try
        End Sub

        ''' <summary>绘制单张 KPI 卡片：圆角浅蓝底 + 大数字(深蓝粗体16号) + 小标签(灰9号)。</summary>
        Private Sub DrawKpiCard(ws As Microsoft.Office.Interop.Excel.Worksheet, rangeAddr As String,
                                label As String, value As String, sub_ As String)
            ' 1) 卡片底板：圆角矩形，浅蓝底，浅蓝边，无阴影
            Dim cardArgs As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
            cardArgs("sheet") = ws.Name
            cardArgs("range") = rangeAddr
            cardArgs("shapeType") = "roundedRect"
            cardArgs("cornerRadius") = 0.08
            cardArgs("fill") = C_CARD_FILL
            cardArgs("line") = C_CARD_LINE
            cardArgs("lineWeight") = 1.0
            cardArgs("noShadow") = True
            cardArgs("sendToBack") = True

            ' 2) 卡片文字：大数字上行(16号粗体深蓝)，小标签下行(9号灰)
            Dim linesArr As New JArray()
            Dim valueLine As New JObject()
            valueLine("text") = If(value, "")
            valueLine("size") = 16.0
            valueLine("bold") = True
            valueLine("color") = C_PRIMARY
            valueLine("alignH") = "left"
            linesArr.Add(valueLine)

            Dim labelText = If(sub_, label)
            Dim labelLine As New JObject()
            labelLine("text") = If(labelText, "")
            labelLine("size") = 9.0
            labelLine("bold") = False
            labelLine("color") = C_TEXT_MUTED
            labelLine("alignH") = "left"
            linesArr.Add(labelLine)

            cardArgs("lines") = linesArr.ToString(Formatting.None)

            Try
                ExecuteAddShape(cardArgs)
            Catch : End Try
        End Sub

        ''' <summary>从 ToolsCallResult 取首条文本（错误诊断用）。</summary>
        Private Shared Function FirstText(r As ToolsCallResult) As String
            If r Is Nothing OrElse r.Content Is Nothing Then Return ""
            For Each c In r.Content
                Dim t = TryCast(c, McpTextContent)
                If t IsNot Nothing AndAlso Not String.IsNullOrEmpty(t.Text) Then Return t.Text
            Next
            Return ""
        End Function

    End Class

End Namespace
