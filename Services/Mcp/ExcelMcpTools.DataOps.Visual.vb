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

        ''' <summary>
        ''' 删除指定工作表上的图表
        ''' </summary>
        Private Function ExecuteDeleteChart(arguments As Dictionary(Of String, Object)) As ToolsCallResult
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

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                If Not String.IsNullOrEmpty(sheetName) Then
                    ws = TryCast(wb.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                Else
                    ws = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                End If

                If ws Is Nothing Then
                    result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: {sheetName}"})
                    result.IsError = True
                    Return result
                End If

                Dim chartObjects As Microsoft.Office.Interop.Excel.ChartObjects = DirectCast(ws.ChartObjects, Microsoft.Office.Interop.Excel.ChartObjects)
                Dim chartCount = chartObjects.Count

                If chartCount = 0 Then
                    result.Content.Add(New McpTextContent With {.Text = "工作表上没有图表。请先调用 excel_create_chart 在该工作表创建图表；若图表可能建在其他工作表，请先用 excel_list_charts 确认其所在 sheet 与准确 chartName，再执行本操作。"})
                    result.IsError = True
                    Utils.ComReleaseHelper.Release(ws)
                    Return result
                End If

                Dim targetChart As Microsoft.Office.Interop.Excel.ChartObject = Nothing
                If String.IsNullOrEmpty(chartName) Then
                    ' 删除最后一个图表
                    targetChart = DirectCast(chartObjects(chartCount), Microsoft.Office.Interop.Excel.ChartObject)
                Else
                    ' 按名称查找
                    For i As Integer = 1 To chartCount
                        Dim co As Microsoft.Office.Interop.Excel.ChartObject = DirectCast(chartObjects(i), Microsoft.Office.Interop.Excel.ChartObject)
                        If co.Name = chartName Then
                            targetChart = co
                            Exit For
                        End If
                        Utils.ComReleaseHelper.Release(co)
                    Next
                End If

                If targetChart Is Nothing Then
                    result.Content.Add(New McpTextContent With {.Text = $"找不到图表: {chartName}"})
                    result.IsError = True
                Else
                    targetChart.Delete()
                    result.Content.Add(New McpTextContent With {.Text = "图表已删除"})
                End If

                Utils.ComReleaseHelper.Release(chartObjects)
                Utils.ComReleaseHelper.Release(ws)

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"删除图表失败: {ex.Message}"})
            Finally
                Utils.ComReleaseHelper.Release(wb)
            End Try

            Return result
        End Function
        ''' <summary>
        ''' 更新图表的数据源范围
        ''' </summary>
        Private Function ExecuteUpdateChartData(arguments As Dictionary(Of String, Object)) As ToolsCallResult
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
                Dim newDataRange = GetArg(Of String)(arguments, "newDataRange", "")

                If String.IsNullOrEmpty(newDataRange) Then
                    result.Content.Add(New McpTextContent With {.Text = "必须指定 newDataRange 参数"})
                    result.IsError = True
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                If Not String.IsNullOrEmpty(sheetName) Then
                    ws = TryCast(wb.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                Else
                    ws = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                End If

                If ws Is Nothing Then
                    result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: {sheetName}"})
                    result.IsError = True
                    Return result
                End If

                Dim chartObjects As Microsoft.Office.Interop.Excel.ChartObjects = DirectCast(ws.ChartObjects, Microsoft.Office.Interop.Excel.ChartObjects)
                Dim chartCount = chartObjects.Count

                If chartCount = 0 Then
                    result.Content.Add(New McpTextContent With {.Text = "工作表上没有图表。请先调用 excel_create_chart 在该工作表创建图表；若图表可能建在其他工作表，请先用 excel_list_charts 确认其所在 sheet 与准确 chartName，再执行本操作。"})
                    result.IsError = True
                    Utils.ComReleaseHelper.Release(ws)
                    Return result
                End If

                Dim targetChart As Microsoft.Office.Interop.Excel.ChartObject = Nothing
                If String.IsNullOrEmpty(chartName) Then
                    targetChart = DirectCast(chartObjects(chartCount), Microsoft.Office.Interop.Excel.ChartObject)
                Else
                    For i As Integer = 1 To chartCount
                        Dim co As Microsoft.Office.Interop.Excel.ChartObject = DirectCast(chartObjects(i), Microsoft.Office.Interop.Excel.ChartObject)
                        If co.Name = chartName Then
                            targetChart = co
                            Exit For
                        End If
                        Utils.ComReleaseHelper.Release(co)
                    Next
                End If

                If targetChart Is Nothing Then
                    result.Content.Add(New McpTextContent With {.Text = $"找不到图表: {chartName}"})
                    result.IsError = True
                Else
                    Dim chart = targetChart.Chart
                    chart.SetSourceData(ws.Range(newDataRange))
                    result.Content.Add(New McpTextContent With {.Text = $"图表数据源已更新为: {newDataRange}"})
                    Utils.ComReleaseHelper.Release(chart)
                End If

                Utils.ComReleaseHelper.Release(chartObjects)
                Utils.ComReleaseHelper.Release(ws)

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"更新图表数据源失败: {ex.Message}"})
            Finally
                Utils.ComReleaseHelper.Release(wb)
            End Try

            Return result
        End Function
        ''' <summary>
        ''' 设置图表标题
        ''' </summary>
        Private Function ExecuteSetChartTitle(arguments As Dictionary(Of String, Object)) As ToolsCallResult
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
                Dim title = GetArg(Of String)(arguments, "title", "")

                If String.IsNullOrEmpty(title) Then
                    result.Content.Add(New McpTextContent With {.Text = "必须指定 title 参数"})
                    result.IsError = True
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                If Not String.IsNullOrEmpty(sheetName) Then
                    ws = TryCast(wb.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                Else
                    ws = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                End If

                If ws Is Nothing Then
                    result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: {sheetName}"})
                    result.IsError = True
                    Return result
                End If

                Dim chartObjects As Microsoft.Office.Interop.Excel.ChartObjects = DirectCast(ws.ChartObjects, Microsoft.Office.Interop.Excel.ChartObjects)
                Dim chartCount = chartObjects.Count

                If chartCount = 0 Then
                    result.Content.Add(New McpTextContent With {.Text = "工作表上没有图表。请先调用 excel_create_chart 在该工作表创建图表；若图表可能建在其他工作表，请先用 excel_list_charts 确认其所在 sheet 与准确 chartName，再执行本操作。"})
                    result.IsError = True
                    Utils.ComReleaseHelper.Release(ws)
                    Return result
                End If

                Dim targetChart As Microsoft.Office.Interop.Excel.ChartObject = Nothing
                If String.IsNullOrEmpty(chartName) Then
                    targetChart = DirectCast(chartObjects(chartCount), Microsoft.Office.Interop.Excel.ChartObject)
                Else
                    For i As Integer = 1 To chartCount
                        Dim co As Microsoft.Office.Interop.Excel.ChartObject = DirectCast(chartObjects(i), Microsoft.Office.Interop.Excel.ChartObject)
                        If co.Name = chartName Then
                            targetChart = co
                            Exit For
                        End If
                        Utils.ComReleaseHelper.Release(co)
                    Next
                End If

                If targetChart Is Nothing Then
                    result.Content.Add(New McpTextContent With {.Text = $"找不到图表: {chartName}"})
                    result.IsError = True
                Else
                    Dim chart = targetChart.Chart
                    chart.HasTitle = True
                    chart.ChartTitle.Text = title
                    result.Content.Add(New McpTextContent With {.Text = $"图表标题已设置为: {title}"})
                    Utils.ComReleaseHelper.Release(chart.ChartTitle)
                    Utils.ComReleaseHelper.Release(chart)
                End If

                Utils.ComReleaseHelper.Release(chartObjects)
                Utils.ComReleaseHelper.Release(ws)

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"设置图表标题失败: {ex.Message}"})
            Finally
                Utils.ComReleaseHelper.Release(wb)
            End Try

            Return result
        End Function
        ''' <summary>
        ''' 刷新数据透视表
        ''' </summary>
        Private Function ExecuteRefreshPivotTable(arguments As Dictionary(Of String, Object)) As ToolsCallResult
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
                Dim pivotTableName = GetArg(Of String)(arguments, "pivotTableName", "")

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                If Not String.IsNullOrEmpty(sheetName) Then
                    ws = TryCast(wb.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                Else
                    ws = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                End If

                If ws Is Nothing Then
                    result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: {sheetName}"})
                    result.IsError = True
                    Return result
                End If

                Dim pivotTables As Microsoft.Office.Interop.Excel.PivotTables = DirectCast(ws.PivotTables, Microsoft.Office.Interop.Excel.PivotTables)
                Dim pivotCount = pivotTables.Count

                If pivotCount = 0 Then
                    ' 【可自愈错误信息】当前工作表没有透视表 ≠ 工作簿没有。
                    ' 实测：透视表建在自动新建的 sheet 上，refresh 却指定了源数据表（sheet=销售数据）→
                    ' 报一句"工作表上没有数据透视表"后 Agent 无从修复。
                    ' 这里扫描全工作簿，把透视表实际所在 sheet 与名称告诉 Agent，让它改参数即可重试。
                    Dim foundList As New List(Of String)()
                    For Each sh As Microsoft.Office.Interop.Excel.Worksheet In wb.Worksheets
                        Dim pts As Microsoft.Office.Interop.Excel.PivotTables = DirectCast(sh.PivotTables, Microsoft.Office.Interop.Excel.PivotTables)
                        If pts IsNot Nothing AndAlso CInt(pts.Count) > 0 Then
                            For i2 As Integer = 1 To CInt(pts.Count)
                                Dim pt As Microsoft.Office.Interop.Excel.PivotTable = DirectCast(pts(i2), Microsoft.Office.Interop.Excel.PivotTable)
                                foundList.Add(sh.Name & " → " & pt.Name)
                                Utils.ComReleaseHelper.Release(pt)
                            Next
                        End If
                        Utils.ComReleaseHelper.Release(pts)
                        ' ⚠ 严禁 Release(sh)：sh 可能就是 ws 本身，释放后下方 ws.Name 与后续
                        ' 代码都在已分离 RCW 上运行（NRE 根源，与 create_pivot_table 同坑）
                    Next
                    If foundList.Count > 0 Then
                        result.Content.Add(New McpTextContent With {.Text = $"工作表『{ws.Name}』上没有数据透视表，但工作簿中存在以下透视表：{String.Join("、", foundList)}。请把 sheet 参数改为对应工作表（或省略 sheet 用活动表），必要时带上 pivotTableName 重试。"})
                    Else
                        result.Content.Add(New McpTextContent With {.Text = "工作簿中没有任何数据透视表。请先调用 excel_create_pivot_table 创建，再执行刷新。"})
                    End If
                    result.IsError = True
                    Utils.ComReleaseHelper.Release(pivotTables)
                    Utils.ComReleaseHelper.Release(ws)
                    Return result
                End If

                Dim targetPivot As Microsoft.Office.Interop.Excel.PivotTable = Nothing
                If String.IsNullOrEmpty(pivotTableName) Then
                    targetPivot = DirectCast(pivotTables(pivotCount), Microsoft.Office.Interop.Excel.PivotTable)
                Else
                    For i As Integer = 1 To pivotCount
                        Dim pt As Microsoft.Office.Interop.Excel.PivotTable = DirectCast(pivotTables(i), Microsoft.Office.Interop.Excel.PivotTable)
                        If pt.Name = pivotTableName Then
                            targetPivot = pt
                            Exit For
                        End If
                        Utils.ComReleaseHelper.Release(pt)
                    Next
                End If

                If targetPivot Is Nothing Then
                    result.Content.Add(New McpTextContent With {.Text = $"找不到数据透视表: {pivotTableName}"})
                    result.IsError = True
                Else
                    targetPivot.RefreshTable()
                    result.Content.Add(New McpTextContent With {.Text = "数据透视表已刷新"})
                    Utils.ComReleaseHelper.Release(targetPivot)
                End If

                Utils.ComReleaseHelper.Release(pivotTables)
                Utils.ComReleaseHelper.Release(ws)

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"刷新数据透视表失败: {ex.Message}"})
            Finally
                Utils.ComReleaseHelper.Release(wb)
            End Try

            Return result
        End Function
        ''' <summary>
        ''' 应用条件格式
        ''' </summary>
        Private Function ExecuteApplyConditionalFormat(arguments As Dictionary(Of String, Object)) As ToolsCallResult
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
                Dim rangeAddress = GetArg(Of String)(arguments, "range", "")
                Dim formatType = GetArg(Of String)(arguments, "formatType", "")
                Dim optionsJson = GetArg(Of String)(arguments, "options", "{}")

                ' ========== Parameter Fallback：和 excel_apply_format 保持一致 ==========
                Dim excelApp2 = Globals.ThisAddIn.Application
                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim rng As Microsoft.Office.Interop.Excel.Range = Nothing

                ' 回退1：用户是否选中了Range？优先级最高（符合"对当前选区应用"心智）
                If String.IsNullOrWhiteSpace(rangeAddress) Then
                    Dim sel As Object = excelApp2.Selection
                    If TypeOf sel Is Microsoft.Office.Interop.Excel.Range Then
                        rng = DirectCast(sel, Microsoft.Office.Interop.Excel.Range)
                        rangeAddress = rng.Address(False, False)
                        If String.IsNullOrWhiteSpace(sheetName) Then
                            ws = DirectCast(rng.Worksheet, Microsoft.Office.Interop.Excel.Worksheet)
                            sheetName = ws.Name
                        End If
                    End If
                End If

                ' 回退2：仍为空 → 用 ActiveSheet.UsedRange（整张已有数据表）
                If String.IsNullOrWhiteSpace(rangeAddress) Then
                    ws = TryCast(excelApp2.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    If ws IsNot Nothing Then
                        Dim used As Microsoft.Office.Interop.Excel.Range = ws.UsedRange
                        If used IsNot Nothing AndAlso used.Cells.Count > 0 Then
                            rng = used
                            rangeAddress = rng.Address(False, False)
                        End If
                    End If
                End If

                ' 三次兜底仍没有 → 报错+附具体操作建议（不要再发干巴巴的"必须指定range"）
                If String.IsNullOrWhiteSpace(rangeAddress) Then
                    result.Content.Add(New McpTextContent With {.Text = "范围不能为空。请先用鼠标选中要应用条件格式的单元格区域（或确保当前工作表有数据），然后重试。"})
                    result.IsError = True
                    Return result
                End If

                ' formatType Fallback：空则默认 colorScale（三色阶最通用，大小值一眼可见）
                Dim fallbackApplied As String = ""
                If String.IsNullOrWhiteSpace(formatType) Then
                    formatType = "colorScale"
                    fallbackApplied = "（上游未传formatType，已自动回退为默认三色阶 colorScale）"
                End If

                ' 工作表解析：兜底填过ws就复用，否则按 sheetName 解析
                If ws Is Nothing Then
                    If Not String.IsNullOrEmpty(sheetName) Then
                        ws = TryCast(wb.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                    Else
                        ws = DirectCast(excelApp2.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    End If
                End If

                If ws Is Nothing Then
                    result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: {sheetName}"})
                    result.IsError = True
                    Return result
                End If

                ' Range 解析：兜底填过 rng 就复用，否则按 rangeAddress 显式取
                If rng Is Nothing Then
                    rng = ws.Range(rangeAddress)
                End If
                If rng Is Nothing OrElse rng.Cells.Count = 0 Then
                    result.Content.Add(New McpTextContent With {.Text = "范围地址无效: " & rangeAddress & " (工作表: " & ws.Name & ")"})
                    result.IsError = True
                    Return result
                End If

                Dim formatConditions = rng.FormatConditions

                ' 解析选项
                Dim options = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(optionsJson)
                If options Is Nothing Then options = New Dictionary(Of String, Object)()

                ' formatType 大小写不敏感 + 别名兼容（上游可能传驼峰 dataBar 也可能全小写 databar）
                Dim ftNorm As String = If(formatType, "").Trim().ToLowerInvariant().Replace("_", "")

                Select Case ftNorm
                    Case "databar", "databars", "bar"
                        Dim removedBars As Integer = DeleteTypedRulesOnSheet(rng, ws, formatConditions, CInt(Microsoft.Office.Interop.Excel.XlFormatConditionType.xlDatabar))
                        Dim fc As Microsoft.Office.Interop.Excel.Databar = DirectCast(formatConditions.AddDatabar(), Microsoft.Office.Interop.Excel.Databar)
                        If options.ContainsKey("color") Then
                            Dim colorStr = options("color").ToString().Replace("#", "")
                            Dim r = Convert.ToInt32(colorStr.Substring(0, 2), 16)
                            Dim g = Convert.ToInt32(colorStr.Substring(2, 2), 16)
                            Dim b = Convert.ToInt32(colorStr.Substring(4, 2), 16)
                            DirectCast(fc.BarColor, Microsoft.Office.Interop.Excel.FormatColor).Color = r + g * 256 + b * 65536
                        End If
                        Try : fc.SetFirstPriority() : Catch : End Try
                        Utils.ComReleaseHelper.Release(fc)
                        result.Content.Add(New McpTextContent With {.Text = $"已应用数据条条件格式到: {rangeAddress}" & ReplacedRuleNote(removedBars) & fallbackApplied})

                    Case "colorscale", "scale", "3scale", "gradient"
                        Dim removedScales As Integer = DeleteTypedRulesOnSheet(rng, ws, formatConditions, CInt(Microsoft.Office.Interop.Excel.XlFormatConditionType.xlColorScale))
                        Dim fc As Microsoft.Office.Interop.Excel.ColorScale = DirectCast(formatConditions.AddColorScale(3), Microsoft.Office.Interop.Excel.ColorScale)
                        Try : fc.SetFirstPriority() : Catch : End Try
                        result.Content.Add(New McpTextContent With {.Text = $"已应用色阶条件格式到: {rangeAddress}" & ReplacedRuleNote(removedScales) & fallbackApplied})
                        Utils.ComReleaseHelper.Release(fc)

                    Case "iconset", "icons", "icon"
                        Dim removedIcons As Integer = DeleteTypedRulesOnSheet(rng, ws, formatConditions, CInt(Microsoft.Office.Interop.Excel.XlFormatConditionType.xlIconSets))
                        Dim fc As Microsoft.Office.Interop.Excel.IconSetCondition = DirectCast(formatConditions.AddIconSetCondition(), Microsoft.Office.Interop.Excel.IconSetCondition)
                        Try : fc.SetFirstPriority() : Catch : End Try
                        result.Content.Add(New McpTextContent With {.Text = $"已应用图标集条件格式到: {rangeAddress}" & ReplacedRuleNote(removedIcons) & fallbackApplied})
                        Utils.ComReleaseHelper.Release(fc)

                    Case "highlight", "rule", "cellrule", "format", "condition"
                        Dim operatorType = 5 ' xlGreater
                        Dim value As Object = Nothing
                        Dim color As Integer = 255 ' 默认红色

                        If options.ContainsKey("operator") Then
                            Dim op = options("operator").ToString().ToLower()
                            Select Case op
                                Case "greaterthan", "greater", ">", "gt" : operatorType = 5
                                Case "lessthan", "less", "<", "lt" : operatorType = 6
                                Case "between", "bw" : operatorType = 1
                                Case "equalto", "equal", "=", "==", "eq" : operatorType = 3
                                Case "notequal", "notequalto", "!=", "<>", "ne" : operatorType = 4
                                Case "greaterthanorequal", ">=", "gte" : operatorType = 7
                                Case "lessthanorequal", "<=", "lte" : operatorType = 8
                            End Select
                        End If

                        If options.ContainsKey("value") Then
                            value = options("value")
                        End If

                        If options.ContainsKey("color") Then
                            Dim colorStr = options("color").ToString().Replace("#", "")
                            Dim r = Convert.ToInt32(colorStr.Substring(0, 2), 16)
                            Dim g = Convert.ToInt32(colorStr.Substring(2, 2), 16)
                            Dim b = Convert.ToInt32(colorStr.Substring(4, 2), 16)
                            color = r + g * 256 + b * 65536
                        End If

                        ' between 需要 Formula2（区间上界）；只给了单值时降级为 greater，
                        ' 否则 Excel 会以"参数无效"拒绝创建（旧实现完全没有这条路径）。
                        Dim effectiveOp As Integer = operatorType
                        Dim value2 As Object = Nothing
                        If options.ContainsKey("value2") Then value2 = options("value2")
                        If effectiveOp = 1 AndAlso value2 Is Nothing Then
                            effectiveOp = 5
                            fallbackApplied &= "（between 缺少上界 value2，已按『大于 value』处理）"
                        End If

                        If value IsNot Nothing Then
                            ' 同一值已经标过色时，旧规则优先级更高，新规则加上去颜色也不会变。
                            ' 先删掉这个值的旧「单元格值等于」规则，再写新颜色。
                            Dim wantedFormula As String = NormalizeCfFormula(If(TypeOf value Is String, CStr(value), Convert.ToString(value, Globalization.CultureInfo.InvariantCulture)))
                            Dim replacedRules As Integer = DeleteMatchingCellValueRules(formatConditions, effectiveOp, wantedFormula)
                            Dim usedForRules As Microsoft.Office.Interop.Excel.Range = Nothing
                            Try
                                usedForRules = ws.UsedRange
                                If usedForRules IsNot Nothing AndAlso usedForRules.Address(False, False) <> rng.Address(False, False) Then
                                    Dim usedFcs As Microsoft.Office.Interop.Excel.FormatConditions = usedForRules.FormatConditions
                                    replacedRules += DeleteMatchingCellValueRules(usedFcs, effectiveOp, wantedFormula)
                                    Utils.ComReleaseHelper.Release(usedFcs)
                                End If
                            Catch
                            Finally
                                Utils.ComReleaseHelper.Release(usedForRules)
                            End Try
                            ' 【修复·E_NOINTERFACE】见 AddHighlightFormatCondition：Type 必须是 xlCellValue，
                            ' 比较符（xlGreater 等）走 Operator；旧实现两者错位建出了 Top10 条件。
                            Dim fc As Microsoft.Office.Interop.Excel.FormatCondition = AddHighlightFormatCondition(formatConditions, effectiveOp, value, value2)
                            fc.Interior.Color = color
                            Try
                                fc.StopIfTrue = True
                                fc.SetFirstPriority()
                            Catch
                            End Try
                            Utils.ComReleaseHelper.Release(fc)
                            Dim replacedNote As String = If(replacedRules > 0, $"（已去掉 {replacedRules} 条旧标色，改为新颜色）", "")
                            result.Content.Add(New McpTextContent With {.Text = $"已应用突出显示规则到: {rangeAddress}" & replacedNote & fallbackApplied})
                        Else
                            ' highlight value 也兜底：默认高亮大于 0 的单元格（否则上游没传 value 还是会失败）
                            Dim fc As Microsoft.Office.Interop.Excel.FormatCondition = AddHighlightFormatCondition(formatConditions, 5, 0)
                            fc.Interior.Color = color
                            Utils.ComReleaseHelper.Release(fc)
                            fallbackApplied &= "（highlight.value未指定，已默认高亮大于0的单元格）"
                            result.Content.Add(New McpTextContent With {.Text = $"已应用突出显示规则到: {rangeAddress}" & fallbackApplied})
                        End If

                    Case Else
                        result.Content.Add(New McpTextContent With {.Text = $"不支持的条件格式类型: {formatType}。支持的值：dataBar、colorScale、iconSet、highlight。" & fallbackApplied})
                        result.IsError = True
                End Select

                Utils.ComReleaseHelper.Release(formatConditions)
                Utils.ComReleaseHelper.Release(rng)
                Utils.ComReleaseHelper.Release(ws)

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"应用条件格式失败: {ex.Message}"})
            Finally
                Utils.ComReleaseHelper.Release(wb)
            End Try

            Return result
        End Function
        ''' <summary>
        ''' 创建"突出显示单元格规则"（highlight）条件格式。
        ''' 【修复·E_NOINTERFACE 根因】旧实现把 operator 枚举值（xlGreater=5 等）当作 **Type**
        ''' 传给 FormatConditions.Add 的第一个参数——而 XlFormatConditionType 里 5 是 xlTop10，
        ''' Excel 于是建出"前 N 项"条件，返回的是 Top10 接口对象；Top10 并不实现
        ''' FormatCondition 接口，强转时 QI 必然失败（IID {00024425-...} = FormatCondition）。
        ''' 副作用是语义也错了：用户要"高亮大于 100"，实际建的是"前 10 项"。
        ''' 正确用法：Type 固定 xlCellValue（单元格值规则），比较符放 Operator 参数。
        ''' Formula1 用 InvariantCulture 字符串（JSON 反序列化出的 Int64/Double 直传不稳）。
        ''' </summary>
        Private Shared Function NormalizeCfFormula(s As String) As String
            Dim t As String = If(s, "").Trim()
            If t.StartsWith("=", StringComparison.Ordinal) Then t = t.Substring(1).Trim()
            If t.Length >= 2 AndAlso t.StartsWith("""", StringComparison.Ordinal) AndAlso t.EndsWith("""", StringComparison.Ordinal) Then
                t = t.Substring(1, t.Length - 2)
            End If
            Return t
        End Function

        ''' <summary>
        ''' 删掉「单元格值 + 同一比较符 + 同一比较值」的旧规则。从后往前删，避免下标错位。
        ''' </summary>
        Private Shared Function ReplacedRuleNote(removed As Integer) As String
            If removed <= 0 Then Return ""
            Return $"（已去掉 {removed} 条旧规则，改为新的）"
        End Function

        ''' <summary>按条件格式类型删旧规则（数据条/色阶/图标集）。再套一次时旧规则仍压在上面，颜色和样式都不会变。</summary>
        Private Shared Function DeleteRulesOfType(formatConditions As Microsoft.Office.Interop.Excel.FormatConditions, ruleType As Integer) As Integer
            Dim removed As Integer = 0
            If formatConditions Is Nothing Then Return 0
            For i As Integer = formatConditions.Count To 1 Step -1
                Dim raw As Object = Nothing
                Try
                    raw = formatConditions.Item(i)
                    Dim t As Integer = -1
                    Try
                        t = CInt(Microsoft.VisualBasic.CallByName(raw, "Type", Microsoft.VisualBasic.CallType.Get))
                    Catch
                    End Try
                    If t = ruleType Then
                        Microsoft.VisualBasic.CallByName(raw, "Delete", Microsoft.VisualBasic.CallType.Method)
                        removed += 1
                    End If
                Catch
                Finally
                    If raw IsNot Nothing Then Utils.ComReleaseHelper.Release(raw)
                End Try
            Next
            Return removed
        End Function

        Private Shared Function DeleteTypedRulesOnSheet(rng As Microsoft.Office.Interop.Excel.Range,
                                                        ws As Microsoft.Office.Interop.Excel.Worksheet,
                                                        primary As Microsoft.Office.Interop.Excel.FormatConditions,
                                                        ruleType As Integer) As Integer
            Dim removed As Integer = DeleteRulesOfType(primary, ruleType)
            Dim used As Microsoft.Office.Interop.Excel.Range = Nothing
            Try
                If ws IsNot Nothing Then used = ws.UsedRange
                If used IsNot Nothing AndAlso rng IsNot Nothing AndAlso used.Address(False, False) <> rng.Address(False, False) Then
                    Dim usedFcs As Microsoft.Office.Interop.Excel.FormatConditions = used.FormatConditions
                    removed += DeleteRulesOfType(usedFcs, ruleType)
                    Utils.ComReleaseHelper.Release(usedFcs)
                End If
            Catch
            Finally
                Utils.ComReleaseHelper.Release(used)
            End Try
            Return removed
        End Function

        Private Shared Function DeleteMatchingCellValueRules(formatConditions As Microsoft.Office.Interop.Excel.FormatConditions,
                                                              operatorType As Integer, wanted As String) As Integer
            Dim removed As Integer = 0
            If formatConditions Is Nothing OrElse String.IsNullOrEmpty(wanted) Then Return 0
            For i As Integer = formatConditions.Count To 1 Step -1
                Dim raw As Object = Nothing
                Try
                    raw = formatConditions.Item(i)
                    Dim fc = TryCast(raw, Microsoft.Office.Interop.Excel.FormatCondition)
                    If fc Is Nothing Then Continue For
                    If CInt(fc.Type) <> CInt(Microsoft.Office.Interop.Excel.XlFormatConditionType.xlCellValue) Then Continue For
                    If CInt(fc.Operator) <> operatorType Then Continue For
                    If Not String.Equals(NormalizeCfFormula(CStr(fc.Formula1)), wanted, StringComparison.Ordinal) Then Continue For
                    fc.Delete()
                    removed += 1
                Catch
                Finally
                    If raw IsNot Nothing Then Utils.ComReleaseHelper.Release(raw)
                End Try
            Next
            Return removed
        End Function

        Private Shared Function AddHighlightFormatCondition(formatConditions As Microsoft.Office.Interop.Excel.FormatConditions,
                                                            operatorType As Integer, value As Object,
                                                            Optional formula2 As Object = Nothing) As Microsoft.Office.Interop.Excel.FormatCondition
            Dim formula1 As String = If(TypeOf value Is String, CStr(value), Convert.ToString(value, Globalization.CultureInfo.InvariantCulture))
            ' 不用三元 If(Of T)：Missing.Value 与 String 混用会触发 BC36912（无法推断通用类型）
            Dim f2 As Object
            If formula2 Is Nothing Then
                f2 = System.Reflection.Missing.Value
            ElseIf TypeOf formula2 Is String Then
                f2 = CStr(formula2)
            Else
                f2 = Convert.ToString(formula2, Globalization.CultureInfo.InvariantCulture)
            End If
            ' PIA 中 FormatConditions.Add 的返回类型是 Object（随 Type 不同可返回 Top10/UniqueValues/
            ' DataBar/...），所以 Option Strict On 下必须显式 DirectCast —— 旧代码的 DirectCast 本身没错，
            ' 错在它传的 Type=5（xlTop10）导致返回的是 Top10，QI 自然失败。
            ' Type=xlCellValue 时返回值确实是 FormatCondition，此处转换必然成功。
            Return DirectCast(formatConditions.Add(Microsoft.Office.Interop.Excel.XlFormatConditionType.xlCellValue,
                                                   DirectCast(CInt(operatorType), Microsoft.Office.Interop.Excel.XlFormatConditionOperator),
                                                   formula1, f2),
                              Microsoft.Office.Interop.Excel.FormatCondition)
        End Function

        ''' <summary>
        ''' 获取透视表聚合函数枚举值。
        ''' 用 CInt 包裹枚举而非硬编码魔术值（-4157 等），让编译器校验。
        ''' Excel XlConsolidationFunction 枚举自 2007 以来稳定不变。
        ''' </summary>
        Private Function GetPivotFunctionEnum(funcName As String) As Integer
            Select Case funcName.ToLower()
                Case "sum" : Return CInt(Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlSum)
                Case "count" : Return CInt(Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlCount)
                Case "average" : Return CInt(Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlAverage)
                Case "max" : Return CInt(Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlMax)
                Case "min" : Return CInt(Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlMin)
                Case Else : Return CInt(Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlSum)
            End Select
        End Function

    End Class

End Namespace