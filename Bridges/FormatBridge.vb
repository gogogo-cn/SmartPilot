' TECHNICAL DEBT: 本文件大量使用 Object 类型进行 Excel COM Interop（Style/Font/Border/PivotField 等返回值均为 Object）
' Option Strict On 需要对几乎每个 COM 调用加 DirectCast，严重影响可读性；优先在应用层/UI层严格类型安全。
Option Strict Off
Option Explicit On
Option Infer On

Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading.Tasks
Imports System.Diagnostics
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ExcelAIVSTOAddIn.Utils
Imports Microsoft.Office.Interop.Excel

Namespace Bridges

    ''' <summary>
    ''' FormatBridge - 格式化与渲染 API 桥梁类
    ''' 提供 Markdown 渲染、表格数据插入、智能数据插入等功能
    ''' </summary>
    <ComVisible(True)>
    <ClassInterface(ClassInterfaceType.AutoDispatch)>
    <ProgId("ExcelAIVSTOAddIn.FormatBridge")>
    Public Class FormatBridge
        Implements IBridge

        ' 单例实例 - 使用 Lazy(Of T) 确保线程安全
        Private Shared ReadOnly _lazyInstance As New Lazy(Of FormatBridge)(Function() New FormatBridge(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        ' 数据写入频率限制器
        Private _dataWriteLimiter As Utils.RateLimiter = Utils.RateLimiter.ForDataWrite

#Region "单例模式"

        ''' <summary>
        ''' 获取 FormatBridge 单例实例（线程安全 Lazy 初始化）
        ''' </summary>
        Public Shared ReadOnly Property Instance As FormatBridge
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        ''' <summary>
        ''' 私有构造函数（单例模式）
        ''' </summary>
        Private Sub New()
        End Sub

#End Region

#Region "格式化与渲染 API"

        ''' <summary>
        ''' 将Markdown文本渲染到Excel，每行文本独立成行显示
        ''' </summary>
        ''' <param name="markdownText">Markdown格式文本</param>
        ''' <param name="startCell">起始单元格地址，如"A1"</param>
        ''' <returns>操作结果JSON</returns>
        <ComVisible(True)>
        Public Function RenderFormattedText(markdownText As String, Optional startCell As String = "") As String
            Dim renderer As New ExcelMarkdownRenderer(Globals.ThisAddIn.Application)
            Return renderer.RenderFormattedText(markdownText, startCell)
        End Function

        ''' <summary>
        ''' 插入表格数据到Excel（保持原始表格布局）
        ''' </summary>
        ''' <param name="rowsJson">JSON格式的二维数组，包含表头和数据行</param>
        ''' <returns>操作结果JSON</returns>
        <ComVisible(True)>
        Public Function InsertTableData(rowsJson As String, Optional writeNumbers As Boolean = False) As String
            Dim workbook As Object = Nothing
            Dim sheet As Object = Nothing
            Dim activeCell As Object = Nothing

            Try
                ' 频率限制检查
                If Not _dataWriteLimiter.TryAcquire() Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "数据写入频率过高，请稍后再试"})
                End If

                Dim excelApp As Application = Globals.ThisAddIn.Application
                If excelApp Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "Excel应用不可用"})
                End If

                ' 解析JSON数据（二维数组格式）
                Dim rows As List(Of List(Of String)) = JsonConvert.DeserializeObject(Of List(Of List(Of String)))(rowsJson)

                If rows Is Nothing OrElse rows.Count = 0 Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "表格数据为空"})
                End If

                ' 获取当前工作表
                workbook = excelApp.ActiveWorkbook
                If workbook Is Nothing Then
                    workbook = excelApp.Workbooks.Add()
                End If

                sheet = workbook.ActiveSheet
                If sheet Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "无法获取活动工作表"})
                End If

                ' 确定起始位置（有选中区域则从选中区域左上角开始，否则从A1开始）
                Dim startRow As Integer = 1
                Dim startCol As Integer = 1
                activeCell = excelApp.ActiveCell
                If activeCell IsNot Nothing Then
                    Dim activeCellRange As Range = CType(activeCell, Range)
                    startRow = activeCellRange.Row
                    startCol = activeCellRange.Column
                    Utils.ComReleaseHelper.Release(activeCellRange)
                End If

                ' 写入表格数据（保持原始布局，同时防护公式注入）
                For rowIndex = 0 To rows.Count - 1
                    Dim row = rows(rowIndex)
                    For colIndex = 0 To row.Count - 1
                        Dim rangeCell As Object = Nothing
                        Try
                            rangeCell = CType(excelApp.Cells(startRow + rowIndex, startCol + colIndex), Range)
                            Dim rawValue = If(row(colIndex), "")
                            Dim target = CType(rangeCell, Range)
                            If writeNumbers AndAlso WriteNumericCell(target, rawValue) Then
                                ' 纯数字按数值写入，方便后续计算
                            Else
                                ' 先设为文本格式，防止 Excel 自动处理（吞空格、转数字等）
                                target.NumberFormatLocal = "@"
                                ' 用 Value2 而非 Value，避免 COM 类型转换吞掉空格
                                target.Value2 = SanitizeCellValue(rawValue)
                            End If
                        Finally
                            Utils.ComReleaseHelper.Release(rangeCell)
                        End Try
                    Next
                Next

                ' 自动调整列宽
                Dim maxCols = rows.Max(Function(r) r.Count)
                Dim sheetWorksheet As Worksheet = CType(sheet, Worksheet)
                For colIndex = 0 To maxCols - 1
                    Dim colRange As Object = Nothing
                    Try
                        colRange = sheetWorksheet.Columns(startCol + colIndex)
                        CType(colRange, Range).AutoFit()
                    Finally
                        Utils.ComReleaseHelper.Release(colRange)
                    End Try
                Next

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .message = $"表格数据已插入，共{rows.Count}行"
                })

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("InsertTableData 错误: " & ex.Message)
                System.Diagnostics.Debug.WriteLine("堆栈: " & ex.StackTrace)
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            Finally
                ' 【修复】释放COM对象防止泄漏
                Utils.ComReleaseHelper.Release(activeCell)
                Utils.ComReleaseHelper.Release(sheet)
                Utils.ComReleaseHelper.Release(workbook)
            End Try
        End Function

        ''' <summary>
        ''' 智能插入数据（支持多种数据格式）
        ''' </summary>
        ''' <param name="payloadJson">包含records和mode的JSON</param>
        ''' <returns>操作结果JSON</returns>
        <ComVisible(True)>
        Public Function InsertDataSmart(payloadJson As String) As String
            Dim workbook As Object = Nothing
            Dim sheet As Object = Nothing
            Dim activeCell As Object = Nothing
            Dim rangeCell As Object = Nothing

            Try
                ' 频率限制检查（已恢复）
                If Not _dataWriteLimiter.TryAcquire() Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "数据写入频率过高，请稍后再试"})
                End If

                Dim excelApp As Application = Globals.ThisAddIn.Application
                If excelApp Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "Excel应用不可用"})
                End If

                ' 解析JSON数据
                Dim payload As Dictionary(Of String, Object) = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(payloadJson)

                If Not payload.ContainsKey("records") Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "数据格式不正确，缺少records字段"})
                End If

                Dim records = payload("records")
                Dim mode = If(payload.ContainsKey("mode"), payload("mode").ToString(), "smart")

                ' 获取当前工作表
                workbook = excelApp.ActiveWorkbook
                If workbook Is Nothing Then
                    workbook = excelApp.Workbooks.Add()
                End If

                sheet = workbook.ActiveSheet
                If sheet Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "无法获取活动工作表"})
                End If

                ' 确定起始位置（有选中区域则从选中区域左上角开始，否则从A1开始）
                Dim startRow As Integer = 1
                Dim startCol As Integer = 1
                activeCell = excelApp.ActiveCell
                If activeCell IsNot Nothing Then
                    Dim activeCellRange As Range = Nothing
                    Try
                        activeCellRange = CType(activeCell, Range)
                        startRow = activeCellRange.Row
                        startCol = activeCellRange.Column
                    Finally
                        Utils.ComReleaseHelper.Release(activeCellRange)
                    End Try
                End If

                ' 处理不同类型的数据
                If TypeOf records Is Newtonsoft.Json.Linq.JArray Then
                    Dim recordsArray = CType(records, Newtonsoft.Json.Linq.JArray)

                    If recordsArray.Count = 0 Then
                        Return JsonConvert.SerializeObject(New With {.success = False, .error = "数据为空"})
                    End If

                    ' 检查第一条记录类型
                    Dim firstRecord = recordsArray(0)

                    If TypeOf firstRecord Is Newtonsoft.Json.Linq.JArray Then
                        ' 二维数组格式（表格数据）
                        Dim rows As List(Of List(Of String)) = recordsArray.ToObject(Of List(Of List(Of String)))()

                        For rowIndex = 0 To rows.Count - 1
                            Dim row = rows(rowIndex)
                            For colIndex = 0 To row.Count - 1
                                Dim cell As Object = Nothing
                                Try
                                    cell = CType(excelApp.Cells(startRow + rowIndex, startCol + colIndex), Range)
                                    CType(cell, Range).NumberFormatLocal = "@"
                                    CType(cell, Range).Value2 = SanitizeCellValue(row(colIndex))
                                Finally
                                    Utils.ComReleaseHelper.Release(cell)
                                End Try
                            Next
                        Next

                        Return JsonConvert.SerializeObject(New With {
                            .success = True,
                            .message = $"表格数据已插入，共{rows.Count}行"
                        })
                    Else
                        ' 对象数组格式（结构化数据）
                        Dim dataArray As List(Of Dictionary(Of String, Object)) = recordsArray.ToObject(Of List(Of Dictionary(Of String, Object)))()

                        If dataArray.Count > 0 Then
                            ' 获取所有列名
                            Dim columns = dataArray(0).Keys.ToList()

                            ' 写入表头
                            For colIndex = 0 To columns.Count - 1
                                Dim headerCell As Object = Nothing
                                Try
                                    headerCell = CType(excelApp.Cells(startRow, startCol + colIndex), Range)
                                    CType(headerCell, Range).NumberFormatLocal = "@"
                                    CType(headerCell, Range).Value2 = columns(colIndex)
                                Finally
                                    Utils.ComReleaseHelper.Release(headerCell)
                                End Try
                            Next

                            ' 写入数据
                            For rowIndex = 0 To dataArray.Count - 1
                                Dim row = dataArray(rowIndex)
                                For colIndex = 0 To columns.Count - 1
                                    Dim colName = columns(colIndex)
                                    Dim value = ""
                                    If row.ContainsKey(colName) AndAlso row(colName) IsNot Nothing Then
                                        value = row(colName).ToString()
                                    End If
                                    Dim cell As Object = Nothing
                                    Try
                                        cell = CType(excelApp.Cells(startRow + 1 + rowIndex, startCol + colIndex), Range)
                                        CType(cell, Range).NumberFormatLocal = "@"
                                        CType(cell, Range).Value2 = SanitizeCellValue(value)
                                    Finally
                                        Utils.ComReleaseHelper.Release(cell)
                                    End Try
                                Next
                            Next

                            Return JsonConvert.SerializeObject(New With {
                                .success = True,
                                .message = $"数据已插入，共{dataArray.Count}行"
                            })
                        End If
                    End If
                End If

                Return JsonConvert.SerializeObject(New With {.success = False, .error = "不支持的数据格式"})

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("InsertDataSmart 错误: " & ex.Message)
                System.Diagnostics.Debug.WriteLine("堆栈: " & ex.StackTrace)
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            Finally
                ' 【修复】释放COM对象防止泄漏
                If activeCell IsNot Nothing Then Utils.ComReleaseHelper.Release(activeCell)
                If sheet IsNot Nothing Then Utils.ComReleaseHelper.Release(sheet)
                If workbook IsNot Nothing Then Utils.ComReleaseHelper.Release(workbook)
            End Try
        End Function

#End Region

#Region "辅助方法"

        ''' <summary>
        ''' 清理单元格值，防止公式注入攻击，并强制身份证号/长数字串作为文本保留
        ''' </summary>
        ''' <param name="value">原始值</param>
        ''' <returns>转义后的安全值</returns>
        Private Function SanitizeCellValue(value As String) As String
            If String.IsNullOrEmpty(value) Then
                Return value
            End If

            Dim trimmedValue = value.Trim()

            ' 身份证号或长数字串应作为文本写入，避免科学计数法或精度丢失
            If ShouldPreserveAsText(trimmedValue) Then
                Return "'" & value
            End If

            ' 检查是否以危险字符开头
            trimmedValue = value.TrimStart()
            If trimmedValue.Length > 0 Then
                Dim firstChar = trimmedValue(0)
                ' Excel公式危险字符: = + - @
                If firstChar = "="c OrElse firstChar = "+"c OrElse firstChar = "-"c OrElse firstChar = "@"c Then
                    ' 在开头添加单引号,使其成为文本而非公式
                    Return "'" & value
                End If
            End If

            Return value
        End Function

        ''' <summary>
        ''' 把纯数字写成数值。带正号的公差保留正号显示，单元格类型仍是数字。
        ''' </summary>
        Private Function WriteNumericCell(cell As Range, raw As String) As Boolean
            If String.IsNullOrWhiteSpace(raw) Then Return False
            Dim text = raw.Trim()
            If ShouldPreserveAsText(text) Then Return False
            If Not System.Text.RegularExpressions.Regex.IsMatch(text, "^[+-]?(?:0|[1-9]\d*)(?:\.\d+)?$") Then Return False

            Dim number As Double
            If Not Double.TryParse(text, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, number) Then Return False

            If text.StartsWith("+"c) Then
                Dim decimals = 0
                Dim dot = text.IndexOf("."c)
                If dot >= 0 Then decimals = text.Length - dot - 1
                Dim pattern = If(decimals > 0, "0." & New String("0"c, decimals), "0")
                cell.NumberFormat = """+""" & pattern & ";-" & pattern & ";0"
            Else
                cell.NumberFormat = "General"
            End If
            cell.Value2 = number
            Return True
        End Function

        ''' <summary>
        ''' 判断字符串是否应作为文本保留（避免 Excel 科学计数法或精度丢失）
        ''' </summary>
        Private Function ShouldPreserveAsText(value As String) As Boolean
            If String.IsNullOrEmpty(value) Then Return False

            ' 18位身份证号：17位数字 + 1位数字或X/x
            If System.Text.RegularExpressions.Regex.IsMatch(value, "^\d{17}[\dXx]$") Then Return True

            ' 超过15位的纯数字（Excel 精度极限）
            If value.Length > 15 AndAlso System.Text.RegularExpressions.Regex.IsMatch(value, "^\d+$") Then Return True

            ' 有前导零的数字串（如订单号、区号等）
            If value.Length > 1 AndAlso value.StartsWith("0") AndAlso System.Text.RegularExpressions.Regex.IsMatch(value, "^\d+$") Then Return True

            Return False
        End Function

#End Region

#Region "【扩展4】增强输出格式 - 图表/透视表/PDF/条件格式"

        ''' <summary>
        ''' 【扩展4】根据 Skill 输出数据生成 Excel 图表
        ''' 支持柱状图、折线图、饼图、散点图、面积图、雷达图
        ''' </summary>
        ''' <param name="dataJson">图表数据的 JSON（含 labels 和 datasets）</param>
        ''' <param name="chartConfigJson">图表配置 JSON（ChartConfig 序列化）</param>
        ''' <returns>操作结果JSON</returns>
        <ComVisible(True)>
        Public Function RenderChart(dataJson As String, Optional chartConfigJson As String = "") As String
            Try
                Dim app = Globals.ThisAddIn.Application
                Dim ws = app.ActiveSheet
                If ws Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "无活动工作表"})
                End If

                ' 解析配置
                Dim config As Skills.ChartConfig = Nothing
                If Not String.IsNullOrWhiteSpace(chartConfigJson) Then
                    config = JsonConvert.DeserializeObject(Of Skills.ChartConfig)(chartConfigJson)
                End If
                If config Is Nothing Then config = New Skills.ChartConfig()

                ' 解析数据
                Dim dataObj = Newtonsoft.Json.Linq.JObject.Parse(dataJson)
                Dim labels As New List(Of String)()
                Dim seriesData As New Dictionary(Of String, List(Of Object))()

                If dataObj.ContainsKey("labels") Then
                    For Each lbl In dataObj("labels")
                        labels.Add(lbl.ToString())
                    Next
                End If

                If dataObj.ContainsKey("datasets") Then
                    For Each ds In dataObj("datasets")
                        Dim dsName = ds("name")?.ToString() OrElse "Series"
                        Dim values As New List(Of Object)()
                        For Each v In ds("values")
                            values.Add(v)
                        Next
                        seriesData(dsName) = values
                    Next
                ElseIf dataObj.ContainsKey("data") Then
                    seriesData("Value") = New List(Of Object)()
                    For Each prop In dataObj("data")
                        labels.Add(prop("key").ToString())
                        seriesData("Value").Add(prop("value"))
                    Next
                End If

                If labels.Count = 0 OrElse seriesData.Count = 0 Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "无有效图表数据"})
                End If

                ' 将数据写入临时区域（用于创建图表的数据源）
                Dim startRow As Integer = ws.UsedRange.Row + ws.UsedRange.Rows.Count + 2
                Dim startCol As Integer = 1

                ' 写入表头
                ws.Cells(startRow, startCol).Value2 = "类别"
                Dim colIdx As Integer = startCol + 1
                For Each seriesName In seriesData.Keys
                    ws.Cells(startRow, colIdx).Value2 = seriesName
                    colIdx += 1
                Next

                ' 写入数据行
                For rowIdx As Integer = 0 To labels.Count - 1
                    ws.Cells(startRow + 1 + rowIdx, startCol).Value2 = labels(rowIdx)
                    colIdx = startCol + 1
                    For Each seriesName In seriesData.Keys
                        Dim valObj = seriesData(seriesName)(rowIdx)
                        If TypeOf valObj Is Double OrElse TypeOf valObj Is Decimal _
                           OrElse TypeOf valObj Is Integer OrElse TypeOf valObj Is Long Then
                            ws.Cells(startRow + 1 + rowIdx, colIdx).Value2 = CDbl(valObj)
                        Else
                            ws.Cells(startRow + 1 + rowIdx, colIdx).Value2 = valObj?.ToString()
                        End If
                        colIdx += 1
                    Next
                Next

                Dim dataEndRow As Integer = startRow + labels.Count
                Dim dataEndCol As Integer = startCol + seriesData.Count

                ' 创建图表
                Dim dataRange As Object = ws.Range(ws.Cells(startRow, startCol), ws.Cells(dataEndRow, dataEndCol))
                Dim chartObj As Object = ws.ChartObjects.Add(
                    Left:=ws.Cells(startRow, dataEndCol + 2).Left,
                    Top:=ws.Cells(startRow, dataEndCol + 2).Top,
                    Width:=config.Width,
                    Height:=config.Height)

                Dim chart As Object = chartObj.Chart

                Select Case config.ChartType.ToLower()
                    Case "line" : chart.ChartType = 4       ' xlLine
                    Case "pie" : chart.ChartType = 5        ' xlPie
                    Case "scatter" : chart.ChartType = -4169 ' xlXYScatter
                    Case "area" : chart.ChartType = 1        ' xlArea
                    Case "radar" : chart.ChartType = -4151   ' xlRadar
                    Case "bar", "column" : chart.ChartType = 51 ' xlColumnClustered
                    Case Else : chart.ChartType = 51         ' xlColumnClustered default
                End Select

                chart.SetSourceData(Source:=dataRange)
                If Not String.IsNullOrWhiteSpace(config.Title) Then
                    chart.HasTitle = True
                    chart.ChartTitle.Text = config.Title
                End If
                chart.HasLegend = config.ShowLegend

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .message = String.Format("已创建 {0} 图表，数据源范围: {1}", config.ChartType, dataRange.Address),
                    .chartAddress = chartObj.TopLeftCell.Address
                })

            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = "图表渲染失败: " & ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 【扩展4】创建数据透视表（与 MCP SheetOps.ExecuteCreatePivotTable 语义完全对齐）
        ''' 对齐项：魔术值→枚举、字段名宽松匹配、全量校验→统一 missing→回滚、COM RCW 释放、名称冲突避让。
        ''' </summary>
        <ComVisible(True)>
        Public Function RenderPivotTable(pivotConfigJson As String) As String
            Try
                Dim app = Globals.ThisAddIn.Application
                Dim wb = app.ActiveWorkbook
                Dim ws = app.ActiveSheet
                If ws Is Nothing OrElse wb Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "无活动工作簿/工作表"})
                End If

                Dim config = JsonConvert.DeserializeObject(Of Skills.PivotTableConfig)(pivotConfigJson)
                If config Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "无效的透视表配置"})
                End If

                ' ── 解析源数据范围 ──
                Dim sourceWs As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim sourceAddr As String = ""
                If Not String.IsNullOrWhiteSpace(config.SourceRange) Then
                    If config.SourceRange.Contains("!") Then
                        Dim parts = config.SourceRange.Split("!"c)
                        sourceWs = DirectCast(wb.Worksheets(parts(0)), Microsoft.Office.Interop.Excel.Worksheet)
                        sourceAddr = parts(1)
                    Else
                        sourceWs = DirectCast(ws, Microsoft.Office.Interop.Excel.Worksheet)
                        sourceAddr = config.SourceRange
                    End If
                Else
                    sourceWs = DirectCast(ws, Microsoft.Office.Interop.Excel.Worksheet)
                    Try : sourceAddr = app.Selection.Address(False, False) : Catch : sourceAddr = "A1" : End Try
                End If
                If sourceWs Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "无法确定数据源工作表"})
                End If
                Dim sourceRng As Microsoft.Office.Interop.Excel.Range = sourceWs.Range(sourceAddr)

                ' ── 目标位置（默认：数据源下方空 3 行）──
                Dim targetRng As Microsoft.Office.Interop.Excel.Range = Nothing
                If Not String.IsNullOrWhiteSpace(config.TargetRange) Then
                    targetRng = DirectCast(ws.Range(config.TargetRange), Microsoft.Office.Interop.Excel.Range)
                Else
                    Dim srcEndRow As Integer = CInt(sourceRng.Row + sourceRng.Rows.Count)
                    targetRng = DirectCast(ws.Cells(srcEndRow + 3, 1), Microsoft.Office.Interop.Excel.Range)
                End If

                ' ── 避让已有透视表名（工作簿级唯一）──
                Dim existingNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                For Each sh As Microsoft.Office.Interop.Excel.Worksheet In wb.Worksheets
                    Try
                        Dim pts As Microsoft.Office.Interop.Excel.PivotTables = DirectCast(sh.PivotTables, Microsoft.Office.Interop.Excel.PivotTables)
                        For i2 As Integer = 1 To CInt(pts.Count)
                            Dim pt2 As Microsoft.Office.Interop.Excel.PivotTable = DirectCast(pts(i2), Microsoft.Office.Interop.Excel.PivotTable)
                            Try : existingNames.Add(CStr(pt2.Name)) : Catch : End Try
                            Utils.ComReleaseHelper.Release(pt2)
                        Next
                        Utils.ComReleaseHelper.Release(pts)
                    Catch
                    End Try
                Next
                Dim pivotName As String = "Pivot_AI"
                Dim seq As Integer = 1
                While existingNames.Contains(pivotName)
                    seq += 1
                    pivotName = "Pivot_AI_" & seq
                End While

                ' ── 创建透视表（枚举替代魔术值）──
                Dim pivotCache As Microsoft.Office.Interop.Excel.PivotCache = DirectCast(
                    wb.PivotCaches.Create(
                        CInt(Microsoft.Office.Interop.Excel.XlPivotTableSourceType.xlDatabase),
                        sourceRng),
                    Microsoft.Office.Interop.Excel.PivotCache)
                Dim pivotTable As Microsoft.Office.Interop.Excel.PivotTable = DirectCast(
                    pivotCache.CreatePivotTable(targetRng, pivotName),
                    Microsoft.Office.Interop.Excel.PivotTable)

                ' ── 收集可用字段（供报错自纠）──
                Dim pfAll As Microsoft.Office.Interop.Excel.PivotFields = DirectCast(pivotTable.PivotFields, Microsoft.Office.Interop.Excel.PivotFields)
                Dim available As New List(Of String)()
                For i0 As Integer = 1 To CInt(pfAll.Count)
                    Dim f0 As Microsoft.Office.Interop.Excel.PivotField = DirectCast(pfAll(i0), Microsoft.Office.Interop.Excel.PivotField)
                    Try
                        If f0.Name IsNot Nothing Then available.Add(CStr(f0.Name))
                    Catch
                    End Try
                    Utils.ComReleaseHelper.Release(f0)
                Next

                ' ═══════════════════════════════════════════════════════════════
                ' Phase 1：全量校验字段存在性（宽松匹配 + 忽略大小写），不动透视表
                ' ═══════════════════════════════════════════════════════════════
                Dim missing As New List(Of String)()

                If config.RowFields IsNot Nothing Then
                    For Each n In config.RowFields
                        If String.IsNullOrWhiteSpace(n) Then Continue For
                        If FormatBridge.FindPivotFieldLoose(pfAll, n) Is Nothing Then missing.Add(n)
                    Next
                End If
                If config.ColumnFields IsNot Nothing Then
                    For Each n In config.ColumnFields
                        If String.IsNullOrWhiteSpace(n) Then Continue For
                        If FormatBridge.FindPivotFieldLoose(pfAll, n) Is Nothing Then missing.Add(n)
                    Next
                End If
                If config.DataFields IsNot Nothing Then
                    For Each df In config.DataFields
                        If df Is Nothing OrElse String.IsNullOrWhiteSpace(df.FieldName) Then Continue For
                        If FormatBridge.FindPivotFieldLoose(pfAll, df.FieldName) Is Nothing Then missing.Add(df.FieldName)
                    Next
                End If
                If config.FilterFields IsNot Nothing Then
                    For Each n In config.FilterFields
                        If String.IsNullOrWhiteSpace(n) Then Continue For
                        If FormatBridge.FindPivotFieldLoose(pfAll, n) Is Nothing Then missing.Add(n)
                    Next
                End If

                If missing.Count > 0 Then
                    ' 回滚：Delete 半成品透视表，兜底 Clear
                    Try : pivotTable.TableRange2.Delete()
                    Catch
                        Try : pivotTable.TableRange2.Clear() : Catch : End Try
                    End Try
                    Utils.ComReleaseHelper.Release(pfAll)
                    Utils.ComReleaseHelper.Release(pivotTable)
                    Utils.ComReleaseHelper.Release(pivotCache)
                    Utils.ComReleaseHelper.Release(sourceRng)
                    Utils.ComReleaseHelper.Release(sourceWs)
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = String.Format("字段不存在: {0}。可用字段: {1}",
                            String.Join("、", missing), String.Join("、", available))
                    })
                End If

                ' ═══════════════════════════════════════════════════════════════
                ' Phase 2：应用（所有字段校验已通过）
                ' ═══════════════════════════════════════════════════════════════
                Dim rowIdx As Integer = 0
                If config.RowFields IsNot Nothing Then
                    For Each n In config.RowFields
                        If String.IsNullOrWhiteSpace(n) Then Continue For
                        Dim f = FormatBridge.FindPivotFieldLoose(pfAll, n)
                        f.Orientation = CInt(Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField)
                        f.Position = rowIdx + 1
                        rowIdx += 1
                        Utils.ComReleaseHelper.Release(f)
                    Next
                End If

                Dim colIdx As Integer = 0
                If config.ColumnFields IsNot Nothing Then
                    For Each n In config.ColumnFields
                        If String.IsNullOrWhiteSpace(n) Then Continue For
                        Dim f = FormatBridge.FindPivotFieldLoose(pfAll, n)
                        f.Orientation = CInt(Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlColumnField)
                        f.Position = colIdx + 1
                        colIdx += 1
                        Utils.ComReleaseHelper.Release(f)
                    Next
                End If

                If config.FilterFields IsNot Nothing Then
                    For Each n In config.FilterFields
                        If String.IsNullOrWhiteSpace(n) Then Continue For
                        Dim f = FormatBridge.FindPivotFieldLoose(pfAll, n)
                        f.Orientation = CInt(Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlPageField)
                        Utils.ComReleaseHelper.Release(f)
                    Next
                End If

                If config.DataFields IsNot Nothing Then
                    For Each df In config.DataFields
                        If df Is Nothing OrElse String.IsNullOrWhiteSpace(df.FieldName) Then Continue For
                        Dim f = FormatBridge.FindPivotFieldLoose(pfAll, df.FieldName)
                        pivotTable.AddDataField(f, , FormatBridge.PivotFunctionEnum(df.Aggregation))
                        Utils.ComReleaseHelper.Release(f)
                    Next
                End If

                pivotTable.RefreshTable()
                Utils.ComReleaseHelper.Release(pfAll)
                Utils.ComReleaseHelper.Release(pivotTable)
                Utils.ComReleaseHelper.Release(pivotCache)
                Utils.ComReleaseHelper.Release(sourceRng)
                Utils.ComReleaseHelper.Release(sourceWs)

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .message = String.Format("已创建数据透视表 '{0}'", pivotName),
                    .pivotTableAddress = targetRng.Address()
                })

            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = "数据透视表创建失败: " & ex.Message
                })
            End Try
        End Function

        ''' <summary>FormatBridge 层 PivotField 宽松匹配（与 MCP SheetOps.ResolvePivotFieldLoose 对齐）：
        ''' ① 精确名（忽略大小写，同时查 Name 和 SourceName）→ ② 包含匹配 → ③ 列字母 A/B/C 按源列序。
        ''' 增强：同 SheetOps 版，PivotField.Name 可能为空或"字段N"，SourceName 更可靠返回源列名。</summary>
        Private Shared Function FindPivotFieldLoose(pfColl As Microsoft.Office.Interop.Excel.PivotFields, spec As String) As Microsoft.Office.Interop.Excel.PivotField
            If pfColl Is Nothing OrElse String.IsNullOrEmpty(spec) Then Return Nothing
            Dim specTrim = spec.Trim()
            If String.IsNullOrEmpty(specTrim) Then Return Nothing
            Try
                ' ① 精确名——同时查 Name 和 SourceName
                For i As Integer = 1 To CInt(pfColl.Count)
                    Dim f As Microsoft.Office.Interop.Excel.PivotField = DirectCast(pfColl(i), Microsoft.Office.Interop.Excel.PivotField)
                    Dim n As String = ""
                    Dim s As String = ""
                    Try : n = CStr(f.Name) : Catch : End Try
                    Try : s = CStr(f.SourceName) : Catch : End Try
                    If String.Equals(n, specTrim, StringComparison.OrdinalIgnoreCase) OrElse
                       String.Equals(s, specTrim, StringComparison.OrdinalIgnoreCase) Then
                        Return f
                    End If
                    Utils.ComReleaseHelper.Release(f)
                Next
                ' ② 包含匹配——同时查 Name 和 SourceName
                For i As Integer = 1 To CInt(pfColl.Count)
                    Dim f As Microsoft.Office.Interop.Excel.PivotField = DirectCast(pfColl(i), Microsoft.Office.Interop.Excel.PivotField)
                    Dim n As String = ""
                    Dim s As String = ""
                    Try : n = CStr(f.Name) : Catch : End Try
                    Try : s = CStr(f.SourceName) : Catch : End Try
                    Dim namesToCheck() As String = {n, s}
                    For Each cand In namesToCheck
                        If cand <> "" AndAlso (cand.IndexOf(specTrim, StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                              specTrim.IndexOf(cand, StringComparison.OrdinalIgnoreCase) >= 0) Then
                            Utils.ComReleaseHelper.Release(f)
                            Return f
                        End If
                    Next
                    Utils.ComReleaseHelper.Release(f)
                Next
                ' ③ 列字母
                If System.Text.RegularExpressions.Regex.IsMatch(specTrim, "^[A-Za-z]{1,3}$") Then
                    Dim idx As Integer = 0
                    For Each c As Char In specTrim
                        idx = idx * 26 + (AscW(Char.ToUpper(c)) - AscW("A"c) + 1)
                    Next
                    If idx >= 1 AndAlso idx <= CInt(pfColl.Count) Then
                        Return DirectCast(pfColl(idx), Microsoft.Office.Interop.Excel.PivotField)
                    End If
                End If
            Catch
            End Try
            Return Nothing
        End Function

        ''' <summary>FormatBridge 层聚合函数枚举：字符串 → XlConsolidationFunction。
        ''' 与 MCP 层 GetPivotFunctionEnum 完全一致（xlMax=-4136 而非旧版错误的 -4137）。</summary>
        Private Shared Function PivotFunctionEnum(funcName As String) As Microsoft.Office.Interop.Excel.XlConsolidationFunction
            Select Case If(funcName, "").Trim().ToLower()
                Case "average" : Return Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlAverage
                Case "count", "counta" : Return Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlCount
                Case "max" : Return Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlMax
                Case "min" : Return Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlMin
                Case Else : Return Microsoft.Office.Interop.Excel.XlConsolidationFunction.xlSum
            End Select
        End Function

        ''' <summary>
        ''' 【扩展4】应用条件格式规则到指定范围
        ''' </summary>
        <ComVisible(True)>
        Public Function ApplyConditionalFormatting(rangeAddress As String, rulesJson As String) As String
            Try
                Dim app = Globals.ThisAddIn.Application
                Dim ws = app.ActiveSheet
                If ws Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "无活动工作表"})
                End If

                Dim targetRange = ws.Range(rangeAddress)
                Dim rules = Newtonsoft.Json.Linq.JArray.Parse(rulesJson)

                For Each ruleObj In rules
                    Dim ruleType = ruleObj("type")?.ToString()?.ToLower()
                    Dim formatStyle = ruleObj("format")?.ToString()

                    Select Case ruleType
                        Case "cellvalue", "value"
                            Dim operatorStr = ruleObj("operator")?.ToString()
                            Dim formula = ruleObj("formula")?.ToString()
                            Dim cf As Object = Nothing

                            If Not String.IsNullOrWhiteSpace(formula) Then
                                cf = targetRange.FormatConditions.Add(
                                    Type:=2, ' xlExpression
                                    Formula1:=formula)
                            ElseIf Not String.IsNullOrWhiteSpace(operatorStr) AndAlso ruleObj.Contains("value") Then
                                Dim value = ruleObj("value")
                                Dim xlOp = GetXlOperator(operatorStr)
                                cf = targetRange.FormatConditions.Add(
                                    Type:=1, ' xlCellValue
                                    Operator:=xlOp,
                                    Formula1:=value.ToString())
                            End If

                            If cf IsNot Nothing Then ApplyFormatStyle(cf, formatStyle)

                        Case "colorscale"
                            Dim cs = targetRange.FormatConditions.AddColorScale(
                                ColorScaleType:=If(ruleObj("colorScaleType")?.ToObject(Of Integer)(), 3))
                            ApplyColorScaleColors(cs, ruleObj)

                        Case "databar"
                            Dim db = targetRange.FormatConditions.AddDatabar()
                            If ruleObj.Contains("barColor") Then
                                db.BarColor.RGB = ParseColor(ruleObj("barColor").ToString())
                            End If

                        Case "iconset"
                            Dim iconSetType = GetIconSetType(ruleObj("iconSet")?.ToString())
                            Dim iset = targetRange.FormatConditions.AddIconSetCondition()
                            iset.IconSet = iconSetType
                    End Select
                Next

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .message = String.Format("已对 {0} 应用 {1} 条条件格式规则", rangeAddress, rules.Count)
                })

            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = "条件格式应用失败: " & ex.Message
                })
            End Try
        End Function

        Private Function GetXlOperator(opStr As String) As Integer
            ' XlFormatConditionOperator 枚举值: xlGreater=5, xlLess=4, xlEqual=1,
            ' xlBetween=1(特殊), xlNotBetween=2, xlGreaterEqual=7, xlLessEqual=8, xlNotEqual=2
            Select Case opStr.ToLower()
                Case "greaterthan", ">" : Return 5   ' xlGreater
                Case "lessthan", "<" : Return 4      ' xlLess
                Case "equal", "=" : Return 1        ' xlEqual
                Case "between" : Return 1            ' xlBetween
                Case "notbetween" : Return 2         ' xlNotBetween
                Case "greaterorequal", ">=" : Return 7 ' xlGreaterEqual
                Case "lessorequal", "<=" : Return 8   ' xlLessEqual
                Case "notequal", "<>", "!=" : Return 2 ' xlNotEqual
                Case Else : Return 5                ' xlGreater default
            End Select
        End Function

        Private Sub ApplyFormatStyle(cf As Object, style As String)
            If String.IsNullOrWhiteSpace(style) Then Return
            Select Case style.ToLower()
                Case "red", "highlight_red"
                    cf.Interior.Color = RGB(255, 199, 206)
                    cf.Font.Color = RGB(156, 0, 6)
                Case "yellow", "highlight_yellow"
                    cf.Interior.Color = RGB(255, 235, 156)
                    cf.Font.Color = RGB(156, 101, 0)
                Case "orange", "highlight_orange"
                    cf.Interior.Color = RGB(255, 217, 163)
                    cf.Font.Color = RGB(156, 67, 0)
                Case "green", "highlight_green"
                    cf.Interior.Color = RGB(198, 239, 206)
                    cf.Font.Color = RGB(0, 97, 0)
                Case "bold" : cf.Font.Bold = True
                Case "italic" : cf.Font.Italic = True
                Case "strikethrough" : cf.Font.Strikethrough = True
            End Select
        End Sub

        Private Function ParseColor(colorStr As String) As Integer
            If colorStr.StartsWith("#") AndAlso colorStr.Length = 7 Then
                Return RGB(
                    Convert.ToInt32(colorStr.Substring(1, 2), 16),
                    Convert.ToInt32(colorStr.Substring(3, 2), 16),
                    Convert.ToInt32(colorStr.Substring(5, 2), 16))
            End If
            Return RGB(0, 0, 0)
        End Function

        Private Sub ApplyColorScaleColors(cs As Object, ruleObj As Newtonsoft.Json.Linq.JObject)
            If ruleObj("minColor") IsNot Nothing Then
                cs.ColorScaleCriteria(1).FormatColor.Color = ParseColor(ruleObj("minColor").ToString())
            End If
            If ruleObj("midColor") IsNot Nothing Then
                cs.ColorScaleCriteria(2).FormatColor.Color = ParseColor(ruleObj("midColor").ToString())
            End If
            If ruleObj("maxColor") IsNot Nothing Then
                cs.ColorScaleCriteria(3).FormatColor.Color = ParseColor(ruleObj("maxColor").ToString())
            End If
        End Sub

        Private Function GetIconSetType(iconSetName As String) As Integer
            ' XlIconSet 枚举值: xl3Arrows=1, xl3Flags=2, xl3TrafficLights1=3, xl3TrafficLights2=4,
            ' xl3Signs=5, xl4Arrows=6, xl4Rating=7, xl5Arrows=8, xl5Quartiles=9
            Select Case iconSetName?.ToLower()
                Case "3arrows" : Return 1   ' xl3Arrows
                Case "3flags" : Return 2     ' xl3Flags
                Case "3trafficlights1" : Return 3 ' xl3TrafficLights1
                Case "3trafficlights2" : Return 4 ' xl3TrafficLights2
                Case "3signs" : Return 5     ' xl3Signs
                Case "4arrows" : Return 6     ' xl4Arrows
                Case "4rating" : Return 7     ' xl4Rating
                Case "5arrows" : Return 8     ' xl5Arrows
                Case "5quartiles" : Return 9  ' xl5Quartiles
                Case Else : Return 1          ' xl3Arrows default
            End Select
        End Function

#End Region

#Region "IBridge 生命周期"

        Public ReadOnly Property BridgeName As String Implements IBridge.BridgeName
            Get
                Return "FormatBridge"
            End Get
        End Property

        Private _disposed As Boolean = False
        Public Sub Dispose() Implements IBridge.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not _disposed Then
                _disposed = True
            End If
        End Sub

#End Region

    End Class

End Namespace