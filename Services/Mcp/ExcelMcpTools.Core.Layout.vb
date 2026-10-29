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
        ''' excel_get_layout_snapshot — 获取当前工作表的布局快照
        ''' 通用框架层工具（非看板专用），多步骤复合操作通用
        ''' 返回：usedRange, charts(位置/大小/数据源), occupiedRegions, kpiCells, suggestedNextAnchor
        ''' </summary>
        Private Function ExecuteGetLayoutSnapshot(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "Excel 应用不可用"})
                    Return result
                End If

                Dim sheetName = GetArg(Of String)(args, "sheet", "")
                Dim includeHeader = GetArg(Of Boolean)(args, "includeHeader", True)
                Dim includeKpiCells = GetArg(Of Boolean)(args, "includeKpiCells", True)

                Dim wb As Microsoft.Office.Interop.Excel.Workbook = excelApp.ActiveWorkbook
                If wb Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何工作簿"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Try
                    If Not String.IsNullOrEmpty(sheetName) Then
                        ws = TryCast(wb.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                    Else
                        ws = TryCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    End If
                    If ws Is Nothing Then
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: {sheetName}"})
                        Return result
                    End If

                    Dim snapshot = New Dictionary(Of String, Object)()
                    snapshot("sheet") = ws.Name

                    ' ── 1. UsedRange ──
                    Dim usedRange = ws.UsedRange
                    Dim usedAddr As String = ""
                    Dim usedRows As Integer = 0, usedCols As Integer = 0
                    If usedRange IsNot Nothing Then
                        usedAddr = usedRange.Address(False, False)
                        usedRows = usedRange.Rows.Count
                        usedCols = usedRange.Columns.Count
                    End If
                    snapshot("usedRange") = usedAddr
                    snapshot("usedRows") = usedRows
                    snapshot("usedCols") = usedCols

                    ' ── 2. Charts 列表 ──
                    Dim charts = New List(Of Dictionary(Of String, Object))()
                    Dim chartRegions = New List(Of Dictionary(Of String, Object))()
                    Dim cos As Microsoft.Office.Interop.Excel.ChartObjects = Nothing
                    Try
                        cos = DirectCast(ws.ChartObjects(), Microsoft.Office.Interop.Excel.ChartObjects)
                        Dim chartCount = CInt(cos.Count)
                        For i As Integer = 1 To chartCount
                            Dim co As Microsoft.Office.Interop.Excel.ChartObject = Nothing
                            Try
                                co = DirectCast(cos(i), Microsoft.Office.Interop.Excel.ChartObject)
                                Dim coName As String = CStr(co.Name)
                                Dim coLeft As Double = CDbl(co.Left)
                                Dim coTop As Double = CDbl(co.Top)
                                Dim coWidth As Double = CDbl(co.Width)
                                Dim coHeight As Double = CDbl(co.Height)

                                ' 尝试获取图表标题和数据源
                                Dim coTitle As String = ""
                                Dim coType As String = ""
                                Dim coDataSource As String = ""
                                Try
                                    Dim cht = co.Chart
                                    If cht IsNot Nothing Then
                                        ' 标题
                                        If cht.HasTitle Then
                                            coTitle = CStr(cht.ChartTitle.Text)
                                        End If
                                        ' 类型
                                        coType = ChartTypeToString(CInt(cht.ChartType))
                                        ' 数据源（尝试从 SeriesCollection 取第一个 Series 的 Formula）
                                        Try
                                            Dim sc As Microsoft.Office.Interop.Excel.SeriesCollection = Nothing
                                            Try : sc = DirectCast(cht.SeriesCollection, Microsoft.Office.Interop.Excel.SeriesCollection) : Catch : sc = Nothing : End Try
                                            If sc IsNot Nothing AndAlso sc.Count > 0 Then
                                                Dim s1 As Microsoft.Office.Interop.Excel.Series = DirectCast(sc(1), Microsoft.Office.Interop.Excel.Series)
                                                Dim f As Object = s1.Formula
                                                If f IsNot Nothing Then
                                                    coDataSource = CStr(f)
                                                End If
                                                Utils.ComReleaseHelper.Release(s1)
                                            End If
                                            Utils.ComReleaseHelper.Release(sc)
                                        Catch
                                        End Try
                                        Utils.ComReleaseHelper.Release(cht)
                                    End If
                                Catch
                                End Try

                                ' 把位置从磅转换成单元格锚点（C5 格式）
                                Dim anchorCell = RangeFromPoint(ws, coLeft, coTop)
                                Dim bottomRightCell = RangeFromPoint(ws, coLeft + coWidth, coTop + coHeight)
                                Dim occupiedAddr As String = ""
                                If anchorCell IsNot Nothing And bottomRightCell IsNot Nothing Then
                                    occupiedAddr = anchorCell.Address(False, False) & ":" & bottomRightCell.Address(False, False)
                                End If

                                charts.Add(New Dictionary(Of String, Object) From {
                                    {"name", coName},
                                    {"type", coType},
                                    {"title", coTitle},
                                    {"anchor", If(anchorCell IsNot Nothing, anchorCell.Address(False, False), "")},
                                    {"widthPt", Math.Round(coWidth, 1)},
                                    {"heightPt", Math.Round(coHeight, 1)},
                                    {"dataSource", coDataSource}
                                })

                                If Not String.IsNullOrEmpty(occupiedAddr) Then
                                    chartRegions.Add(New Dictionary(Of String, Object) From {
                                        {"range", occupiedAddr},
                                        {"type", "chart"},
                                        {"name", coName}
                                    })
                                End If

                                Utils.ComReleaseHelper.Release(co)
                            Catch ex As Exception
                                System.Diagnostics.Debug.WriteLine($"[GetLayoutSnapshot] 读取第 {i} 个图表失败: {ex.Message}")
                            End Try
                        Next
                    Finally
                        If cos IsNot Nothing Then Utils.ComReleaseHelper.Release(cos)
                    End Try
                    snapshot("charts") = charts
                    snapshot("chartCount") = charts.Count

                    ' ── 3. 已占用区域合并（chart + 数据区）──
                    Dim occupied = New List(Of Dictionary(Of String, Object))()
                    For Each cr In chartRegions
                        occupied.Add(cr)
                    Next
                    If Not String.IsNullOrEmpty(usedAddr) AndAlso usedRows > 0 Then
                        occupied.Add(New Dictionary(Of String, Object) From {
                            {"range", usedAddr},
                            {"type", "data"}
                        })
                    End If
                    snapshot("occupiedRegions") = occupied

                    ' ── 4. 表头样本 ──
                    If includeHeader AndAlso usedRange IsNot Nothing AndAlso usedRows >= 1 Then
                        Dim headerRow = New List(Of String)()
                        Try
                            Dim firstRow = DirectCast(usedRange.Rows(1), Microsoft.Office.Interop.Excel.Range)
                            Dim rowValue = firstRow.Value
                            If rowValue IsNot Nothing Then
                                If TypeOf rowValue Is Object(,) Then
                                    Dim arr = DirectCast(rowValue, Object(,))
                                    For j As Integer = 1 To arr.GetLength(1)
                                        Dim v = arr(1, j)
                                        headerRow.Add(If(v Is Nothing, "", CStr(v)))
                                    Next
                                ElseIf TypeOf rowValue Is Object() Then
                                    For Each v In DirectCast(rowValue, Object())
                                        headerRow.Add(If(v Is Nothing, "", CStr(v)))
                                    Next
                                Else
                                    headerRow.Add(If(rowValue Is Nothing, "", CStr(rowValue)))
                                End If
                            End If
                            Utils.ComReleaseHelper.Release(firstRow)
                        Catch
                        End Try
                        snapshot("headerRow") = headerRow
                    End If

                    ' ── 5. KPI 单元格样本（前 20 个非空单元格）──
                    If includeKpiCells AndAlso usedRange IsNot Nothing AndAlso usedRows > 0 Then
                        Dim kpiCells = New List(Of Dictionary(Of String, Object))()
                        Dim maxScan = Math.Min(usedRows, 50)  ' 最多扫描 50 行
                        Try
                            Dim scanRange = ws.Range(ws.Cells(1, 1), ws.Cells(maxScan, Math.Min(usedCols, 20)))
                            Dim scanValue = scanRange.Value
                            If scanValue IsNot Nothing AndAlso TypeOf scanValue Is Object(,) Then
                                Dim arr = DirectCast(scanValue, Object(,))
                                Dim limit As Integer = 20
                                For i As Integer = 1 To arr.GetLength(0)
                                    For j As Integer = 1 To arr.GetLength(1)
                                        If kpiCells.Count >= limit Then Exit For
                                        Dim v = arr(i, j)
                                        If v IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(CStr(v)) Then
                                            Dim cellRng As Microsoft.Office.Interop.Excel.Range = DirectCast(ws.Cells(i, j), Microsoft.Office.Interop.Excel.Range)
                                            Dim addr As String = CStr(cellRng.Address(False, False))
                                            Utils.ComReleaseHelper.Release(cellRng)
                                            kpiCells.Add(New Dictionary(Of String, Object) From {
                                                {"cell", addr},
                                                {"value", CStr(v)}
                                            })
                                        End If
                                    Next
                                    If kpiCells.Count >= limit Then Exit For
                                Next
                            End If
                            Utils.ComReleaseHelper.Release(scanRange)
                        Catch
                        End Try
                        snapshot("kpiCells") = kpiCells
                    End If

                    ' ── 6. 建议下一个锚点 ──
                    ' 简单算法：找到已占用区域的最大行号 + 3，A 列作为建议锚点
                    Dim suggestedRow As Integer = 1
                    If usedRows > 0 Then suggestedRow = usedRows + 3
                    If charts.Count > 0 Then
                        ' 用最后一个 chart 的 bottom 位置
                        Dim lastChart = charts.Last()
                        Dim lastAnchor = CStr(lastChart("anchor"))
                        If Not String.IsNullOrEmpty(lastAnchor) Then
                            Dim rowMatch = System.Text.RegularExpressions.Regex.Match(lastAnchor, "\d+")
                            If rowMatch.Success Then
                                Dim chartBottom = Integer.Parse(rowMatch.Value) + 15  ' 假设 chart 占 ~15 行
                                suggestedRow = Math.Max(suggestedRow, chartBottom + 3)
                            End If
                        End If
                    End If
                    snapshot("suggestedNextAnchor") = "A" & suggestedRow
                    snapshot("usedRangeBottom") = "A" & Math.Max(usedRows, 1)

                    ' ── 7. 输出 ──
                    Dim json = JsonConvert.SerializeObject(snapshot, Formatting.Indented)
                    result.Content.Add(New McpTextContent With {.Text = json})
                    Return result

                Finally
                    If ws IsNot Nothing Then Utils.ComReleaseHelper.Release(ws)
                End Try

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"GetLayoutSnapshot 失败: {ex.Message}"})
                Return result
            End Try
        End Function

        ''' <summary>根据 (left, top) 磅值反推 Excel 单元格地址</summary>
        Private Shared Function RangeFromPoint(ws As Microsoft.Office.Interop.Excel.Worksheet,
                                                leftPt As Double, topPt As Double) As Microsoft.Office.Interop.Excel.Range
            Try
                Dim excelApp = ws.Application
                Dim cell = DirectCast(excelApp.ActiveWindow.RangeFromPoint(CInt(leftPt), CInt(topPt)),
                                     Microsoft.Office.Interop.Excel.Range)
                Return cell
            Catch
                Return Nothing
            End Try
        End Function

    End Class

End Namespace


