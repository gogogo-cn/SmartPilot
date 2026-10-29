Imports System
Imports System.Collections.Generic
Imports System.Linq

Namespace Services

    ''' <summary>
    ''' OCR 坐标表格解析引擎
    ''' 基于 _fix_send.py 参考实现移植
    ''' 核心算法：Y轴聚类分群 → X轴排序对齐 → 跨行合并
    ''' 支持：多级表头、跨行合并(rowspan)、不规则单元格
    ''' </summary>
    Public Class OcrTableParser

        Private _cells As List(Of OcrTableCell)
        Private _yTolerance As Double
        Private _xTolerance As Double
        Private _rows As List(Of List(Of OcrTableCell))
        Private _columns As List(Of Double)
        Private _table As List(Of List(Of String))
        Private _rowspanMap As Dictionary(Of Tuple(Of Integer, Integer), Integer)

        ''' <summary>
        ''' 构造表格解析器
        ''' </summary>
        ''' <param name="elements">带 BoundingBox 的 OCR 元素列表</param>
        ''' <param name="yTolerance">Y轴行容差（同一行判定阈值）</param>
        ''' <param name="xTolerance">X轴列容差（同一列判定阈值）</param>
        Public Sub New(elements As List(Of OcrElement), Optional yTolerance As Double = 50, Optional xTolerance As Double = 200)
            _cells = elements.Where(Function(e)
                Return e.BoundingBox IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(e.Text)
            End Function).Select(Function(e) New OcrTableCell(e)).ToList()

            _yTolerance = yTolerance
            _xTolerance = xTolerance

            _rows = New List(Of List(Of OcrTableCell))()
            _columns = New List(Of Double)()
            _table = New List(Of List(Of String))()
            _rowspanMap = New Dictionary(Of Tuple(Of Integer, Integer), Integer)()
        End Sub

        ''' <summary>
        ''' 执行完整解析，返回 OcrTableData
        ''' </summary>
        Public Function Parse() As OcrTableData
            Try
                If _cells.Count < 2 Then Return Nothing

                ClusterByY()
                CleanOutlierRows()
                DetectColumns()

                If _columns.Count < 2 Then Return Nothing

                Dim grid = AssignToGrid()
                MergeRowspan(grid)

                Return BuildTableData()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[OcrTableParser] Parse error: {ex.Message}")
                Return Nothing
            End Try
        End Function

#Region "Step 1: Y轴聚类分群"

        Private Sub ClusterByY()
            If _cells.Count = 0 Then Return

            ' 1. 按 Y 坐标排序
            Dim sortedCells = _cells.OrderBy(Function(c) c.CenterY).ToList()
            Dim allY = sortedCells.Select(Function(c) c.CenterY).ToList()

            ' 2. 计算相邻单元格的 Y 间距，取低百分位作为行内典型间距
            Dim gaps As New List(Of Double)()
            For i As Integer = 1 To allY.Count - 1
                gaps.Add(allY(i) - allY(i - 1))
            Next

            If gaps.Count = 0 Then
                _rows = New List(Of List(Of OcrTableCell))() From {
                    New List(Of OcrTableCell)(sortedCells)
                }
                Return
            End If

            Dim sortedGaps = gaps.OrderBy(Function(g) g).ToList()
            Dim intraRowGap = sortedGaps(CInt(sortedGaps.Count * 0.33)) ' 33 百分位，避免被行间距拉高
            Dim avgHeight = _cells.Average(Function(c) c.Height)
            Dim threshold = Math.Min(_yTolerance, Math.Max(12, Math.Min(intraRowGap * 2.0, avgHeight * 0.7)))

            ' 3. 按间距粗分行
            Dim roughRows As New List(Of List(Of OcrTableCell))()
            Dim currentRow As New List(Of OcrTableCell)() From {sortedCells(0)}
            For i As Integer = 1 To sortedCells.Count - 1
                Dim gap = sortedCells(i).CenterY - sortedCells(i - 1).CenterY
                If gap <= threshold Then
                    currentRow.Add(sortedCells(i))
                Else
                    currentRow.Sort(Function(a, b) a.CenterX.CompareTo(b.CenterX))
                    roughRows.Add(currentRow)
                    currentRow = New List(Of OcrTableCell)() From {sortedCells(i)}
                End If
            Next
            If currentRow.Count > 0 Then
                currentRow.Sort(Function(a, b) a.CenterX.CompareTo(b.CenterX))
                roughRows.Add(currentRow)
            End If

            ' 4. 合并 Y 中心过近的碎行（兼容 OCR 按列返回导致的同一行被拆分）
            Dim mergedRows As New List(Of List(Of OcrTableCell))()
            For Each row In roughRows
                Dim rowCenter = row.Average(Function(c) c.CenterY)
                Dim merged = False
                For Each existing In mergedRows
                    Dim existingCenter = existing.Average(Function(c) c.CenterY)
                    If Math.Abs(rowCenter - existingCenter) <= threshold Then
                        existing.AddRange(row)
                        merged = True
                        Exit For
                    End If
                Next
                If Not merged Then
                    mergedRows.Add(row)
                End If
            Next

            ' 5. 最终按 X 排序并移除空行
            _rows = New List(Of List(Of OcrTableCell))()
            For Each row In mergedRows
                If row.Count = 0 Then Continue For
                row.Sort(Function(a, b) a.CenterX.CompareTo(b.CenterX))
                _rows.Add(row)
            Next
        End Sub

        ''' <summary>
        ''' 剔除孤立行（如水印、页脚）：单行且与相邻行间距明显过大
        ''' </summary>
        Private Sub CleanOutlierRows()
            If _rows.Count <= 2 Then Return

            Dim rowCenters = _rows.Select(Function(r) r.Average(Function(c) c.CenterY)).ToList()
            Dim gaps As New List(Of Double)()
            For i As Integer = 1 To rowCenters.Count - 1
                gaps.Add(rowCenters(i) - rowCenters(i - 1))
            Next
            If gaps.Count = 0 Then Return

            Dim avgGap = gaps.Average()
            Dim threshold = avgGap * 2.5

            Dim rowsToKeep As New List(Of List(Of OcrTableCell))()
            For i As Integer = 0 To _rows.Count - 1
                Dim isSparse = _rows(i).Count <= 1
                Dim gapBefore = If(i > 0, rowCenters(i) - rowCenters(i - 1), 0.0)
                Dim gapAfter = If(i < rowCenters.Count - 1, rowCenters(i + 1) - rowCenters(i), 0.0)

                If isSparse AndAlso (gapBefore > threshold OrElse gapAfter > threshold) Then
                    Continue For
                End If
                rowsToKeep.Add(_rows(i))
            Next

            _rows = rowsToKeep
        End Sub

#End Region

#Region "Step 2: X轴列检测"

        ''' <summary>
        ''' 用单元格最多的行来确定列中心，避免缩进/空行列带偏全局聚类
        ''' </summary>
        Private Sub DetectColumns()
            If _rows.Count = 0 Then
                _columns = New List(Of Double)()
                Return
            End If

            Dim bestRow = _rows.OrderByDescending(Function(r) r.Count).FirstOrDefault()
            If bestRow Is Nothing OrElse bestRow.Count < 2 Then
                _columns = New List(Of Double)()
                Return
            End If

            _columns = bestRow.Select(Function(c) c.CenterX).OrderBy(Function(x) x).ToList()
        End Sub

#End Region

#Region "Step 3: 网格分配"

        Private Function AssignToGrid() As List(Of List(Of OcrTableCell))
            Dim grid As New List(Of List(Of OcrTableCell))()
            Dim halfTolerance = _xTolerance / 2.0

            For Each rowCells In _rows
                Dim gridRow As New List(Of OcrTableCell)(New OcrTableCell(_columns.Count - 1) {})
                Dim rowHasContent = False
                Dim pendingFar As New List(Of OcrTableCell)()
                For Each cell In rowCells
                    Dim bestCol = -1
                    Dim bestDist = Double.MaxValue
                    For i As Integer = 0 To _columns.Count - 1
                        Dim dist = Math.Abs(cell.CenterX - _columns(i))
                        If dist < bestDist Then
                            bestDist = dist
                            bestCol = i
                        End If
                    Next

                    If bestCol < 0 Then Continue For

                    ' 孤立且远离列中心的框（水印、边注）先挂起。
                    ' 同一行已有内容时再补上：身份证号这类长文本的中心会偏到列中心右侧。
                    If bestDist > halfTolerance AndAlso Not rowHasContent Then
                        pendingFar.Add(cell)
                        Continue For
                    End If

                    PlaceCell(gridRow, bestCol, cell)
                    rowHasContent = True
                Next

                If rowHasContent Then
                    For Each cell In pendingFar
                        PlaceCell(gridRow, NearestColumn(cell), cell)
                    Next
                    grid.Add(gridRow)
                ElseIf pendingFar.Count > 0 AndAlso grid.Count > 0 Then
                    ' 号码单独成行、紧挨着标签行时，并回上一行，避免整格丢失
                    Dim prev = grid(grid.Count - 1)
                    Dim prevCells = prev.Where(Function(c) c IsNot Nothing).ToList()
                    If prevCells.Count > 0 Then
                        Dim prevY = prevCells.Average(Function(c) c.CenterY)
                        For Each cell In pendingFar
                            If Math.Abs(cell.CenterY - prevY) <= _yTolerance * 2 Then
                                PlaceCell(prev, NearestColumn(cell), cell)
                            End If
                        Next
                    End If
                End If
            Next
            Return grid
        End Function

        Private Function NearestColumn(cell As OcrTableCell) As Integer
            Dim bestCol = 0
            Dim bestDist = Double.MaxValue
            For i As Integer = 0 To _columns.Count - 1
                Dim dist = Math.Abs(cell.CenterX - _columns(i))
                If dist < bestDist Then
                    bestDist = dist
                    bestCol = i
                End If
            Next
            Return bestCol
        End Function

        Private Sub PlaceCell(gridRow As List(Of OcrTableCell), col As Integer, cell As OcrTableCell)
            Dim existing = gridRow(col)
            If existing Is Nothing Then
                gridRow(col) = cell
                Return
            End If
            existing.Text = existing.Text & " " & cell.Text
            If cell.Right > existing.Right Then existing.Right = cell.Right
            If cell.Left < existing.Left Then existing.Left = cell.Left
        End Sub

#End Region

#Region "Step 4: 跨行合并"

        Private Sub MergeRowspan(grid As List(Of List(Of OcrTableCell)))
            Dim numCols = _columns.Count
            Dim pendingMerges As New Dictionary(Of Integer, Tuple(Of Integer, String))()
            Dim result As New List(Of List(Of String))()

            For rowIdx As Integer = 0 To grid.Count - 1
                Dim resultRow As New List(Of String)(New String(numCols - 1) {})
                Dim gridRow = grid(rowIdx)

                For colIdx As Integer = 0 To numCols - 1
                    Dim cell = gridRow(colIdx)

                    If cell IsNot Nothing Then
                        resultRow(colIdx) = cell.Text

                        ' 判断是否是跨行单元格开始：高度明显大于普通行高
                        If cell.Height > _yTolerance * 1.5 Then
                            pendingMerges(colIdx) = Tuple.Create(rowIdx, cell.Text)
                            _rowspanMap(Tuple.Create(rowIdx, colIdx)) = 1
                        Else
                            ' 当前列出现新内容，结束之前的跨行合并
                            If pendingMerges.ContainsKey(colIdx) Then
                                pendingMerges.Remove(colIdx)
                            End If
                        End If
                    ElseIf pendingMerges.ContainsKey(colIdx) Then
                        ' 跨行单元格的后续行保持空白（由 HTML rowspan 显示），不再重复填充
                        resultRow(colIdx) = ""
                        _rowspanMap(Tuple.Create(pendingMerges(colIdx).Item1, colIdx)) += 1
                    Else
                        resultRow(colIdx) = ""
                    End If
                Next

                result.Add(resultRow)
            Next

            _table = result
        End Sub

#End Region

#Region "输出转换"

        Private Function BuildTableData() As OcrTableData
            Dim tableData As New OcrTableData()
            If _table.Count = 0 Then Return tableData

            tableData.Header.AddRange(_table(0))
            For i As Integer = 1 To _table.Count - 1
                tableData.Rows.Add(_table(i).ToList())
            Next

            tableData.ColumnCount = _columns.Count
            tableData.RowCount = _table.Count - 1
            tableData.RowSpanMap = _rowspanMap.ToDictionary(
                Function(k) k.Key,
                Function(k) k.Value)

            Return tableData
        End Function

        ''' <summary>
        ''' 输出 Markdown 表格
        ''' </summary>
        Public Function ToMarkdown() As String
            If _table Is Nothing OrElse _table.Count = 0 Then Return String.Empty

            Dim lines As New List(Of String)()
            lines.Add("| " + String.Join(" | ", _table(0).Select(Function(c) If(c, ""))) + " |")
            lines.Add("| " + String.Join(" | ", Enumerable.Repeat("---", _table(0).Count)) + " |")

            For i As Integer = 1 To _table.Count - 1
                lines.Add("| " + String.Join(" | ", _table(i).Select(Function(c) If(c, ""))) + " |")
            Next

            Return String.Join(vbLf, lines)
        End Function

        ''' <summary>
        ''' 输出 HTML 表格（支持 rowspan）
        ''' </summary>
        Public Function ToHtml() As String
            If _table Is Nothing OrElse _table.Count = 0 Then Return String.Empty

            Dim skipCells As New HashSet(Of Tuple(Of Integer, Integer))()
            For Each kvp In _rowspanMap
                For i As Integer = 1 To kvp.Value - 1
                    skipCells.Add(Tuple.Create(kvp.Key.Item1 + i, kvp.Key.Item2))
                Next
            Next

            Dim html As New List(Of String)()
            html.Add("<table border='1' cellpadding='6' style='border-collapse:collapse;font-family:微软雅黑,sans-serif;font-size:14px;'>")

            ' 表头
            html.Add("  <thead>")
            html.Add("    <tr style='background:#c8a96e;color:#fff;'>")
            For Each col In _table(0)
                html.Add($"      <th style='border:1px solid #a08050;padding:8px;'>{col}</th>")
            Next
            html.Add("    </tr>")
            html.Add("  </thead>")

            ' 数据
            html.Add("  <tbody>")
            For rowIdx As Integer = 1 To _table.Count - 1
                html.Add("    <tr>")
                For colIdx As Integer = 0 To _table(rowIdx).Count - 1
                    If skipCells.Contains(Tuple.Create(rowIdx, colIdx)) Then Continue For

                    Dim attrs = ""
                    Dim style = "border:1px solid #d4c4a0;padding:8px;"
                    If _rowspanMap.ContainsKey(Tuple.Create(rowIdx, colIdx)) Then
                        Dim span = _rowspanMap(Tuple.Create(rowIdx, colIdx))
                        If span > 1 Then
                            attrs = $" rowspan=""{span}"""
                            style += "background:#f5f0e0;vertical-align:middle;"
                        End If
                    End If
                    html.Add($"      <td{attrs} style='{style}'>{_table(rowIdx)(colIdx)}</td>")
                Next
                html.Add("    </tr>")
            Next
            html.Add("  </tbody>")
            html.Add("</table>")

            Return String.Join(vbLf, html)
        End Function

#End Region

    End Class

    ''' <summary>
    ''' 表格解析内部使用的单元格模型
    ''' </summary>
    Public Class OcrTableCell
        Public Text As String
        Public Confidence As Single
        Public Box As OcrBoundingBox
        Public CenterX As Double
        Public CenterY As Double
        Public Top As Double
        Public Bottom As Double
        Public Left As Double
        Public Right As Double
        Public Height As Double

        Public Sub New(element As OcrElement)
            Text = element.Text
            Confidence = CSng(element.Confidence)
            Box = element.BoundingBox
            Left = Box.Left
            Right = Box.Right
            Top = Box.Top
            Bottom = Box.Bottom
            CenterX = (Left + Right) / 2
            CenterY = (Top + Bottom) / 2
            Height = Bottom - Top
        End Sub
    End Class

End Namespace