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

        '#Region "工具实现 - 图表"

        ' =============================================================
        ' 图表工具辅助区（excel_list_charts / set_chart_properties / export_chart_image 共用）
        ' =============================================================
        Private Shared Function ResolveWorksheet(excelApp As Microsoft.Office.Interop.Excel.Application, wb As Microsoft.Office.Interop.Excel.Workbook, sheetName As String) As Microsoft.Office.Interop.Excel.Worksheet
            If Not String.IsNullOrEmpty(sheetName) Then
                Return TryCast(wb.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
            End If
            Return TryCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
        End Function

        ''' <summary>根据 (ws, chartName) 查找 ChartObject；未指定名返回最后一个；找不到返回 Nothing。调用方负责 Release。</summary>
        Private Shared Function FindChartObject(ws As Microsoft.Office.Interop.Excel.Worksheet,
                                                chartName As String,
                                                ByRef chartObjects As Object,
                                                ByRef count As Integer) As Microsoft.Office.Interop.Excel.ChartObject
            Dim cos As Microsoft.Office.Interop.Excel.ChartObjects = DirectCast(ws.ChartObjects(), Microsoft.Office.Interop.Excel.ChartObjects)
            chartObjects = cos
            count = CInt(cos.Count)
            If count = 0 Then Return Nothing
            If String.IsNullOrEmpty(chartName) Then Return DirectCast(cos(count), Microsoft.Office.Interop.Excel.ChartObject)
            For i As Integer = 1 To count
                Dim co As Microsoft.Office.Interop.Excel.ChartObject = DirectCast(cos(i), Microsoft.Office.Interop.Excel.ChartObject)
                If CStr(co.Name) = chartName Then
                    Return co
                End If
                Utils.ComReleaseHelper.Release(co)
            Next
            Return Nothing
        End Function

        Private Shared Function ChartTypeToString(typeCode As Integer) As String
            Select Case typeCode
                Case 51 : Return "xlColumnClustered"
                Case 52 : Return "xlColumnStacked"
                Case 53 : Return "xlColumnStacked100"
                Case 57 : Return "xlBarClustered"
                Case 58 : Return "xlBarStacked"
                Case 59 : Return "xlBarStacked100"
                Case 4  : Return "xlLine"
                Case 65 : Return "xlLineMarkers"
                Case 3  : Return "xlArea"
                Case 5  : Return "xlPie"
                Case -4120 : Return "xlDoughnut"
                Case -4169 : Return "xlXYScatter"
                Case -4151 : Return "xlRadar"
                Case 86 : Return "xlRadarFilled"
                Case 15 : Return "xlBubble"
                Case 117 : Return "xlSunburst"
                Case 115 : Return "xlTreemap"
                Case 118 : Return "xlHistogram"
                Case 121 : Return "xlBoxwhisker"
                Case 122 : Return "xlFunnel"
                Case 120 : Return "xlWaterfall"
                Case Else : Return $"custom({typeCode})"
            End Select
        End Function

        Private Shared Function TryParseChartType(name As String, ByRef out As Microsoft.Office.Interop.Excel.XlChartType) As Boolean
            Select Case If(name, "").ToLowerInvariant()
                Case "xlcolumnclustered"   : out = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered
                Case "xlcolumnstacked"     : out = Microsoft.Office.Interop.Excel.XlChartType.xlColumnStacked
                Case "xlcolumnstacked100"  : out = Microsoft.Office.Interop.Excel.XlChartType.xlColumnStacked100
                Case "xlbarclustered"      : out = Microsoft.Office.Interop.Excel.XlChartType.xlBarClustered
                Case "xlbarstacked"        : out = Microsoft.Office.Interop.Excel.XlChartType.xlBarStacked
                Case "xlbarstacked100"     : out = Microsoft.Office.Interop.Excel.XlChartType.xlBarStacked100
                Case "xlline"              : out = Microsoft.Office.Interop.Excel.XlChartType.xlLine
                Case "xllinemarkers"       : out = Microsoft.Office.Interop.Excel.XlChartType.xlLineMarkers
                Case "xlarea"              : out = Microsoft.Office.Interop.Excel.XlChartType.xlArea
                Case "xlpie"               : out = Microsoft.Office.Interop.Excel.XlChartType.xlPie
                Case "xldoughnut"          : out = Microsoft.Office.Interop.Excel.XlChartType.xlDoughnut
                Case "xlxyscatter"         : out = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatter
                Case "xlradar"             : out = Microsoft.Office.Interop.Excel.XlChartType.xlRadar
                Case "xlradarfilled"       : out = Microsoft.Office.Interop.Excel.XlChartType.xlRadarFilled
                Case "xlbubble"            : out = CType(15, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xlsunburst"          : out = CType(117, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xltreemap"           : out = CType(115, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xlhistogram"         : out = CType(118, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xlboxwhisker"        : out = CType(121, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xlfunnel"            : out = CType(122, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xlwaterfall"         : out = CType(120, Microsoft.Office.Interop.Excel.XlChartType)
                Case Else : Return False
            End Select
            Return True
        End Function

        Private Shared Function TryParseLegendPosition(name As String, ByRef out As Microsoft.Office.Interop.Excel.XlLegendPosition) As Boolean
            Select Case If(name, "").ToLowerInvariant()
                Case "xllegendpositionbottom" : out = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionBottom
                Case "xllegendpositiontop"    : out = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionTop
                Case "xllegendpositionleft"   : out = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionLeft
                Case "xllegendpositioncorner" : out = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionCorner
                Case "xllegendpositionright"  : out = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionRight
                Case Else : Return False
            End Select
            Return True
        End Function

        Private Shared Function TryParseLabelPosition(name As String, ByRef out As Integer) As Boolean
            Select Case If(name, "").ToLowerInvariant()
                Case "xllabelpositioncenter"     : out = -4108
                Case "xllabelpositioninsideend"  : out = 4
                Case "xllabelpositioninsidebase" : out = -4109
                Case "xllabelpositionbestfit"    : out = 5
                Case "xllabelpositionabove"      : out = -4160
                Case "xllabelpositionbelow"      : out = -4107
                Case "xllabelpositionright"      : out = -4152
                Case "xllabelpositionleft"       : out = -4131
                Case "xllabelpositionoutsidend"  : out = 3
                Case Else : Return False
            End Select
            Return True
        End Function

        ''' <summary>
        ''' 列出所有图表（名称/类型/位置/大小/标题）——图表编辑的前置动作
        ''' </summary>
        Private Function ExecuteListCharts(arguments As Dictionary(Of String, Object)) As ToolsCallResult
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
                Dim sheetName = GetArg(Of String)(arguments, "sheet", "")
                Dim ws = ResolveWorksheet(excelApp, wb, sheetName)
                If ws Is Nothing Then
                    result.IsError = True : result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: {sheetName}"})
                    Utils.ComReleaseHelper.Release(wb) : Return result
                End If

                Dim chartObjs As Microsoft.Office.Interop.Excel.ChartObjects = DirectCast(ws.ChartObjects(), Microsoft.Office.Interop.Excel.ChartObjects)
                Dim count As Integer = CInt(chartObjs.Count)
                If count = 0 Then
                    result.Content.Add(New McpTextContent With {.Text = $"工作表『{ws.Name}』上没有图表。"})
                Else
                    Dim list = New List(Of Dictionary(Of String, Object))(count)
                    For i As Integer = 1 To count
                        Dim co As Microsoft.Office.Interop.Excel.ChartObject = DirectCast(chartObjs(i), Microsoft.Office.Interop.Excel.ChartObject)
                        Dim ch As Microsoft.Office.Interop.Excel.Chart = co.Chart
                        Dim rec = New Dictionary(Of String, Object)
                        rec("index") = i
                        rec("chartName") = CStr(co.Name)
                        rec("chartType") = ChartTypeToString(CInt(ch.ChartType))
                        rec("left") = CDbl(co.Left)
                        rec("top") = CDbl(co.Top)
                        rec("width") = CDbl(co.Width)
                        rec("height") = CDbl(co.Height)
                        rec("title") = If(CBool(ch.HasTitle), CStr(ch.ChartTitle.Text), "")
                        list.Add(rec)
                        Utils.ComReleaseHelper.Release(ch)
                        Utils.ComReleaseHelper.Release(co)
                    Next

                    Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(list, Newtonsoft.Json.Formatting.Indented)
                    result.Content.Add(New McpTextContent With {.Text = $"共 {count} 个图表（工作表『{ws.Name}』）：" & vbCrLf & json})
                End If

                Utils.ComReleaseHelper.Release(chartObjs)
                Utils.ComReleaseHelper.Release(ws)
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"列出图表失败: {ex.Message}"})
            Finally
                Utils.ComReleaseHelper.Release(wb)
            End Try
            Return result
        End Function

        ''' <summary>
        ''' 统一设置已存在图表的属性
        ''' </summary>
        Private Function ExecuteSetChartProperties(arguments As Dictionary(Of String, Object)) As ToolsCallResult
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
                Dim sheetName = GetArg(Of String)(arguments, "sheet", "")
                Dim chartName = GetArg(Of String)(arguments, "chartName", "")
                Dim ws = ResolveWorksheet(excelApp, wb, sheetName)
                If ws Is Nothing Then
                    result.IsError = True : result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: {sheetName}"})
                    Utils.ComReleaseHelper.Release(wb) : Return result
                End If

                Dim chartObjs As Object = Nothing, count As Integer = 0
                Dim co = FindChartObject(ws, chartName, chartObjs, count)
                If co Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = If(count = 0,
                        "工作表上没有图表，请先调用 excel_create_chart 创建。",
                        $"找不到图表：{chartName}。请先调用 excel_list_charts 查看准确的 chartName。")})
                    Utils.ComReleaseHelper.Release(chartObjs)
                    Utils.ComReleaseHelper.Release(ws)
                    Utils.ComReleaseHelper.Release(wb)
                    Return result
                End If

                Dim ch = co.Chart
                Dim applied = New List(Of String)()

                ' 1) 类型
                Dim ctName = GetArg(Of String)(arguments, "chartType", "")
                If Not String.IsNullOrEmpty(ctName) Then
                    Dim ct As Microsoft.Office.Interop.Excel.XlChartType
                    If Not TryParseChartType(ctName, ct) Then
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = $"chartType 不支持：{ctName}"})
                        GoTo Cleanup
                    End If
                    ch.ChartType = ct
                    applied.Add("chartType=" & ctName)
                End If

                ' 2) 标题（空串删除标题）
                If arguments.ContainsKey("title") Then
                    Dim t = GetArg(Of String)(arguments, "title", Nothing)
                    If String.IsNullOrEmpty(t) Then
                        ch.HasTitle = False
                        applied.Add("title=清除")
                    Else
                        ch.HasTitle = True
                        ch.ChartTitle.Text = t
                        applied.Add("title=" & t)
                    End If
                End If

                ' 3) 位置/大小 JSON（字段缺省不修改）
                ' 支持两种格式：
                '   a) 磅值格式: {left, top, width, height}
                '   b) 单元格锚定: {anchorRange:"C5:F20"} — 引擎自动换算磅坐标（看板布局推荐）
                Dim posJson = GetArg(Of String)(arguments, "position", "")
                If Not String.IsNullOrWhiteSpace(posJson) Then
                    Try
                        Dim j = Newtonsoft.Json.Linq.JObject.Parse(posJson)
                        ' b) anchorRange 优先（看板语义更清晰）
                        Dim anchorRangeTok = j("anchorRange")
                        If anchorRangeTok IsNot Nothing Then
                            Dim anchorAddr = CStr(anchorRangeTok)
                            Dim anchorRng As Microsoft.Office.Interop.Excel.Range = Nothing
                            Try
                                anchorRng = ws.Range(anchorAddr)
                                If anchorRng Is Nothing Then
                                    result.IsError = True
                                    result.Content.Add(New McpTextContent With {.Text = $"anchorRange 无效: {anchorAddr}"})
                                    GoTo Cleanup
                                End If
                                co.Left = CDbl(anchorRng.Left)
                                co.Top = CDbl(anchorRng.Top)
                                ' anchorRange 的宽高可被显式 width/height 覆盖
                                If j("width") Is Nothing Then co.Width = CDbl(anchorRng.Width)
                                If j("height") Is Nothing Then co.Height = CDbl(anchorRng.Height)
                                applied.Add($"anchorRange={anchorAddr} → position=({co.Left:F0},{co.Top:F0}) {co.Width:F0}x{co.Height:F0}")
                            Catch ex As Exception
                                result.IsError = True
                                result.Content.Add(New McpTextContent With {.Text = $"anchorRange 解析失败: {ex.Message}"})
                                GoTo Cleanup
                            Finally
                                If anchorRng IsNot Nothing Then Utils.ComReleaseHelper.Release(anchorRng)
                            End Try
                        Else
                            ' a) 直接磅值
                            If j("left") IsNot Nothing Then co.Left = CDbl(j("left"))
                            If j("top") IsNot Nothing Then co.Top = CDbl(j("top"))
                            If j("width") IsNot Nothing Then co.Width = Math.Max(150, CDbl(j("width")))
                            If j("height") IsNot Nothing Then co.Height = Math.Max(100, CDbl(j("height")))
                            applied.Add($"position=({co.Left:F0},{co.Top:F0}) {co.Width:F0}x{co.Height:F0}")
                        End If
                    Catch ex As Exception
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = $"position JSON 解析失败：{ex.Message}"})
                        GoTo Cleanup
                    End Try
                End If

                ' 4) 系列方向 / 切换行列
                Dim sb = GetArg(Of String)(arguments, "seriesBy", "")
                If Not String.IsNullOrEmpty(sb) Then
                    ch.PlotBy = If(String.Equals(sb, "rows", StringComparison.OrdinalIgnoreCase),
                                    Microsoft.Office.Interop.Excel.XlRowCol.xlRows,
                                    Microsoft.Office.Interop.Excel.XlRowCol.xlColumns)
                    applied.Add("seriesBy=" & sb)
                End If
                If arguments.ContainsKey("switchRowColumn") AndAlso
                   GetArg(Of Boolean)(arguments, "switchRowColumn", False) Then
                    If CInt(ch.PlotBy) = CInt(Microsoft.Office.Interop.Excel.XlRowCol.xlColumns) Then
                        ch.PlotBy = Microsoft.Office.Interop.Excel.XlRowCol.xlRows
                    Else
                        ch.PlotBy = Microsoft.Office.Interop.Excel.XlRowCol.xlColumns
                    End If
                    applied.Add("switchRowColumn=True")
                End If

                ' 5) 样式
                If arguments.ContainsKey("chartStyle") Then
                    Dim cs = GetArg(Of Integer)(arguments, "chartStyle", -1)
                    If cs > 0 AndAlso cs <= 208 Then
                        ch.ChartStyle = cs
                        applied.Add("chartStyle=" & cs.ToString())
                    End If
                End If

                ' 6) 图例
                If arguments.ContainsKey("hasLegend") Then
                    ch.HasLegend = GetArg(Of Boolean)(arguments, "hasLegend", True)
                    applied.Add("hasLegend=" & ch.HasLegend.ToString())
                End If
                Dim lpStr = GetArg(Of String)(arguments, "legendPosition", "")
                If Not String.IsNullOrEmpty(lpStr) Then
                    ch.HasLegend = True
                    Dim lp As Microsoft.Office.Interop.Excel.XlLegendPosition
                    If TryParseLegendPosition(lpStr, lp) Then
                        ch.Legend.Position = lp
                        applied.Add("legendPosition=" & lpStr)
                    End If
                End If

                ' 7) 坐标轴标题（空串=取消）
                Dim xaStr = GetArg(Of String)(arguments, "xAxisTitle", Nothing)
                Dim yaStr = GetArg(Of String)(arguments, "yAxisTitle", Nothing)
                If xaStr IsNot Nothing OrElse yaStr IsNot Nothing Then
                    Dim axes As Microsoft.Office.Interop.Excel.Axes = DirectCast(ch.Axes(), Microsoft.Office.Interop.Excel.Axes)
                    For Each ax As Microsoft.Office.Interop.Excel.Axis In axes
                        Try
                            Dim tc As Integer = CInt(ax.Type)
                            If tc = 1 AndAlso xaStr IsNot Nothing Then
                                ax.HasTitle = Not String.IsNullOrEmpty(xaStr)
                                If ax.HasTitle Then ax.AxisTitle.Text = xaStr
                            End If
                            If tc = 2 AndAlso yaStr IsNot Nothing Then
                                ax.HasTitle = Not String.IsNullOrEmpty(yaStr)
                                If ax.HasTitle Then ax.AxisTitle.Text = yaStr
                            End If
                        Catch : End Try
                    Next
                    Utils.ComReleaseHelper.Release(axes)
                    If xaStr IsNot Nothing Then
                        applied.Add($"xAxisTitle={If(String.IsNullOrEmpty(xaStr), "清除", "设置")}")
                    End If
                    If yaStr IsNot Nothing Then
                        applied.Add($"yAxisTitle={If(String.IsNullOrEmpty(yaStr), "清除", "设置")}")
                    End If
                End If

                ' 8) 数据标签
                If arguments.ContainsKey("showDataLabels") Then
                    Dim show As Boolean = GetArg(Of Boolean)(arguments, "showDataLabels", False)
                    Dim lblPosStr = GetArg(Of String)(arguments, "dataLabelPosition", "")
                    Dim defaultLbl As Integer =
                        If((ch.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlPie OrElse
                            ch.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlDoughnut),
                           5,  ' xlLabelPositionBestFit
                           If((ch.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered OrElse
                               ch.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnStacked OrElse
                               ch.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnStacked100),
                              3,   ' xlLabelPositionOutsideEnd
                              -4160))  ' xlLabelPositionAbove

                    Dim lp As Integer = defaultLbl
                    If Not String.IsNullOrEmpty(lblPosStr) Then
                        If Not TryParseLabelPosition(lblPosStr, lp) Then lp = defaultLbl
                    End If

                    Dim seriesColl As Microsoft.Office.Interop.Excel.SeriesCollection = DirectCast(ch.SeriesCollection(), Microsoft.Office.Interop.Excel.SeriesCollection)
                    For Each s As Microsoft.Office.Interop.Excel.Series In seriesColl
                        Try
                            If show Then
                                s.ApplyDataLabels(Type:=Microsoft.Office.Interop.Excel.XlDataLabelsType.xlDataLabelsShowValue,
                                                  LegendKey:=False, AutoText:=True, HasLeaderLines:=True)
                                DirectCast(s.DataLabels, Microsoft.Office.Interop.Excel.DataLabels).Position =
                                    CType(lp, Microsoft.Office.Interop.Excel.XlDataLabelPosition)
                            Else
                                s.ApplyDataLabels(Type:=Microsoft.Office.Interop.Excel.XlDataLabelsType.xlDataLabelsShowNone)
                            End If
                        Catch : End Try
                    Next
                    Utils.ComReleaseHelper.Release(seriesColl)
                    applied.Add(If(show, "showDataLabels=True", "showDataLabels=False"))
                End If

                ' 9) 趋势线（为第 1 个系列加线性趋势线 + R²）
                If arguments.ContainsKey("addTrendline") AndAlso
                   GetArg(Of Boolean)(arguments, "addTrendline", False) Then
                    Dim sc As Microsoft.Office.Interop.Excel.SeriesCollection = DirectCast(ch.SeriesCollection(), Microsoft.Office.Interop.Excel.SeriesCollection)
                    If CInt(sc.Count) >= 1 Then
                        Try
                            Dim s1 As Microsoft.Office.Interop.Excel.Series = DirectCast(sc(1), Microsoft.Office.Interop.Excel.Series)
                            Dim tl As Microsoft.Office.Interop.Excel.Trendline = DirectCast(DirectCast(s1.Trendlines, Microsoft.Office.Interop.Excel.Trendlines).Add(Type:=Microsoft.Office.Interop.Excel.XlTrendlineType.xlLinear,
                                                         DisplayEquation:=True, DisplayRSquared:=True), Microsoft.Office.Interop.Excel.Trendline)
                            tl.Forward = 2
                            applied.Add("addTrendline=线性(R²+公式+预测2期)")
                            Utils.ComReleaseHelper.Release(tl)
                            Utils.ComReleaseHelper.Release(s1)
                        Catch : End Try
                    End If
                    Utils.ComReleaseHelper.Release(sc)
                End If

                ' 10) 组合图 - 单独系列类型（seriesChartType + seriesIndex，seriesIndex 缺省=最后一个系列）
                Dim sctName = GetArg(Of String)(arguments, "seriesChartType", "")
                If Not String.IsNullOrEmpty(sctName) Then
                    Dim ct As Microsoft.Office.Interop.Excel.XlChartType
                    If Not TryParseChartType(sctName, ct) Then
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = $"seriesChartType 不支持：{sctName}"})
                        GoTo Cleanup
                    End If
                    Dim sci As Integer = GetArg(Of Integer)(arguments, "seriesIndex", -1)
                    Dim sc2 As Microsoft.Office.Interop.Excel.SeriesCollection = DirectCast(ch.SeriesCollection(), Microsoft.Office.Interop.Excel.SeriesCollection)
                    Dim n As Integer = CInt(sc2.Count)
                    If sci <= 0 Then sci = n   ' 未指定或非法→默认最后一个系列
                    If sci < 1 OrElse sci > n Then
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = $"seriesIndex={sci} 越界，该图表仅 {n} 个系列。"})
                        Utils.ComReleaseHelper.Release(sc2)
                        GoTo Cleanup
                    End If
                    Try
                        Dim targetS As Microsoft.Office.Interop.Excel.Series = DirectCast(sc2(sci), Microsoft.Office.Interop.Excel.Series)
                        targetS.ChartType = ct
                        applied.Add($"系列#{sci} 类型→{sctName}")
                        Utils.ComReleaseHelper.Release(targetS)
                    Catch : End Try
                    Utils.ComReleaseHelper.Release(sc2)
                End If

                ' 11) 组合图 - 次坐标轴（secondaryAxisSeriesIndex：正数=设置次轴；0/负=清除该系列次轴；缺省=最后一个系列设为次轴）
                If arguments.ContainsKey("secondaryAxisSeriesIndex") Then
                    Dim idxRaw As Object = Nothing
                    If arguments.TryGetValue("secondaryAxisSeriesIndex", idxRaw) Then
                        Dim parsedOk As Boolean = True
                        Dim axIdx As Integer = -1
                        Try
                            axIdx = CInt(Math.Truncate(CDbl(idxRaw)))
                        Catch
                            parsedOk = False
                        End Try
                        If parsedOk Then
                            Dim sc3 As Microsoft.Office.Interop.Excel.SeriesCollection = DirectCast(ch.SeriesCollection(), Microsoft.Office.Interop.Excel.SeriesCollection)
                            Dim n3 As Integer = CInt(sc3.Count)
                            If n3 = 0 Then
                                Utils.ComReleaseHelper.Release(sc3)
                            Else
                                If axIdx <= 0 Then
                                    ' 0/负 语义：对最后一个系列执行『取消次轴→放回主轴』（也可传 exact index 扩展，当前保持与正向一致的默认最后一个）
                                    Dim tIdx As Integer = If(axIdx = 0, n3, -axIdx)
                                    If tIdx < 1 OrElse tIdx > n3 Then tIdx = n3
                                    Try
                                        Dim targetS As Microsoft.Office.Interop.Excel.Series = DirectCast(sc3(tIdx), Microsoft.Office.Interop.Excel.Series)
                                        targetS.AxisGroup = Microsoft.Office.Interop.Excel.XlAxisGroup.xlPrimary
                                        applied.Add($"系列#{tIdx} 取消次轴（回主轴）")
                                        Utils.ComReleaseHelper.Release(targetS)
                                    Catch : End Try
                                Else
                                    If axIdx > n3 Then axIdx = n3
                                    Try
                                        Dim targetS As Microsoft.Office.Interop.Excel.Series = DirectCast(sc3(axIdx), Microsoft.Office.Interop.Excel.Series)
                                        targetS.AxisGroup = Microsoft.Office.Interop.Excel.XlAxisGroup.xlSecondary
                                        ch.HasAxis(Microsoft.Office.Interop.Excel.XlAxisType.xlValue,
                                                   Microsoft.Office.Interop.Excel.XlAxisGroup.xlSecondary) = True
                                        applied.Add($"系列#{axIdx} 绑定次坐标轴 AxisGroup=2")
                                        Utils.ComReleaseHelper.Release(targetS)
                                    Catch : End Try
                                End If
                                Utils.ComReleaseHelper.Release(sc3)
                            End If
                        End If
                    End If
                End If

                result.Content.Add(New McpTextContent With {
                    .Text = $"图表『{co.Name}』属性已更新：{If(applied.Count = 0, "(未传任何可识别属性，请参考 schema 传参)", String.Join("；", applied))}"
                })
Cleanup:
                Utils.ComReleaseHelper.Release(ch)
                Utils.ComReleaseHelper.Release(co)
                Utils.ComReleaseHelper.Release(chartObjs)
                Utils.ComReleaseHelper.Release(ws)
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"设置图表属性失败: {ex.Message}"})
            Finally
                Utils.ComReleaseHelper.Release(wb)
            End Try
            Return result
        End Function

        ''' <summary>
        ''' 导出图表为 PNG/JPG/GIF，返回保存路径
        ''' </summary>
        Private Function ExecuteExportChartImage(arguments As Dictionary(Of String, Object)) As ToolsCallResult
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
                Dim sheetName = GetArg(Of String)(arguments, "sheet", "")
                Dim chartName = GetArg(Of String)(arguments, "chartName", "")
                Dim outputPath = GetArg(Of String)(arguments, "outputPath", "")
                Dim imgFmt = GetArg(Of String)(arguments, "imageFormat", "png").ToLowerInvariant()

                Dim ws = ResolveWorksheet(excelApp, wb, sheetName)
                If ws Is Nothing Then
                    result.IsError = True : result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: {sheetName}"})
                    Utils.ComReleaseHelper.Release(wb) : Return result
                End If

                Dim chartObjs As Object = Nothing, count As Integer = 0
                Dim co = FindChartObject(ws, chartName, chartObjs, count)
                If co Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = If(count = 0,
                        "工作表上没有图表，请先创建图表。",
                        $"找不到图表：{chartName}，请先调用 excel_list_charts。")})
                    Utils.ComReleaseHelper.Release(chartObjs)
                    Utils.ComReleaseHelper.Release(ws)
                    Utils.ComReleaseHelper.Release(wb)
                    Return result
                End If

                ' 1. 确定最终扩展名
                Dim ext As String = ""
                If Not String.IsNullOrEmpty(outputPath) Then
                    ext = IO.Path.GetExtension(outputPath).TrimStart("."c).ToLowerInvariant()
                End If
                If String.IsNullOrEmpty(ext) Then ext = imgFmt
                If ext = "jpeg" Then ext = "jpg"
                Dim supported As Boolean = (ext = "png" OrElse ext = "jpg" OrElse ext = "gif")
                If Not supported Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"imageFormat 仅支持 png/jpg/gif，实际：{ext}"})
                    GoTo Cleanup
                End If
                Dim filterName As String = If(ext = "png", "PNG", If(ext = "jpg", "JPEG", "GIF"))

                ' 2. 决定输出路径
                If String.IsNullOrWhiteSpace(outputPath) Then
                    Dim safeName = IO.Path.GetInvalidFileNameChars().Aggregate(co.Name, Function(cur, c) cur.Replace(c, "_"c))
                    Dim tmpDir = IO.Path.GetTempPath()
                    outputPath = IO.Path.Combine(tmpDir, safeName & "_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & "." & ext)
                Else
                    ' 扩展名与实际格式不符 → 强制修正
                    If Not IO.Path.GetExtension(outputPath).TrimStart("."c).ToLowerInvariant().Equals(ext, StringComparison.Ordinal) Then
                        outputPath = IO.Path.Combine(IO.Path.GetDirectoryName(outputPath),
                                                     IO.Path.GetFileNameWithoutExtension(outputPath) & "." & ext)
                    End If
                    ' 【防幻觉路径】净化模型编造的占位符/伪造用户目录，避免 CreateDirectory 造出垃圾目录
                    outputPath = Utils.PathHelper.SanitizeUserOutputPath(outputPath)
                    ' 确保父目录存在
                    Dim dir = IO.Path.GetDirectoryName(outputPath)
                    If Not String.IsNullOrEmpty(dir) AndAlso Not IO.Directory.Exists(dir) Then IO.Directory.CreateDirectory(dir)
                End If

                ' 3. 调用 Chart.Export 导出
                Dim ch = co.Chart
                Dim ok As Boolean = CBool(ch.Export(outputPath, filterName, Interactive:=False))
                Utils.ComReleaseHelper.Release(ch)

                If Not ok Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"Chart.Export 返回 False，导出失败。"})
                Else
                    result.Content.Add(New McpTextContent With {.Text = $"图表『{co.Name}』已导出为 {filterName}：{outputPath}（文件大小 {New IO.FileInfo(outputPath).Length} 字节）"})
                End If

Cleanup:
                Utils.ComReleaseHelper.Release(co)
                Utils.ComReleaseHelper.Release(chartObjs)
                Utils.ComReleaseHelper.Release(ws)
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"导出图表图片失败: {ex.Message}"})
            Finally
                Utils.ComReleaseHelper.Release(wb)
            End Try
            Return result
        End Function

        '#End Region

    End Class

End Namespace
