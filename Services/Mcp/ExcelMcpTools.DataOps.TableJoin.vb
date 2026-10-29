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
        ''' 跨表关联（数据库风格 JOIN：内/左/右/全连接）
        ''' </summary>
        Private Function ExecuteJoinTables(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Dim app = Globals.ThisAddIn.Application
            Dim displayAlertsOriginal = app.DisplayAlerts
            Dim pr As New StepRecorder()
            Try
                Dim leftSheetName As String = GetArg(Of String)(args, "leftSheet", "")
                Dim rightSheetName As String = GetArg(Of String)(args, "rightSheet", "")

                ' 关联键：兼容单数键名 leftKeyColumn/rightKeyColumn 与复数 JSON 数组 leftKeyColumns/rightKeyColumns
                Dim leftKeys As List(Of String) = ParseKeyColumns(args, "leftKeyColumn", "leftKeyColumns")
                Dim rightKeys As List(Of String) = ParseKeyColumns(args, "rightKeyColumn", "rightKeyColumns")
                Dim joinType As String = GetArg(Of String)(args, "joinType", "inner").Trim().ToLower()
                Dim targetSheetName As String = GetArg(Of String)(args, "targetSheet", "")
                Dim includeHeaders As Boolean = GetArg(Of Boolean)(args, "includeHeaders", True)

                If String.IsNullOrEmpty(leftSheetName) OrElse String.IsNullOrEmpty(rightSheetName) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "必须指定左右两个工作表"})
                    Return result
                End If

                Dim wb = app.ActiveWorkbook
                If wb Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何工作簿"})
                    Return result
                End If

                If leftKeys Is Nothing OrElse leftKeys.Count = 0 OrElse rightKeys Is Nothing OrElse rightKeys.Count = 0 OrElse leftKeys.Count <> rightKeys.Count Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "左右关联键列必须是一一对应的非空列表（长度相等）"})
                    Return result
                End If

                ' 读取两个表
                Dim lws = DirectCast(wb.Worksheets(leftSheetName), Microsoft.Office.Interop.Excel.Worksheet)
                Dim leftFormats As String() = Nothing
                Dim leftArr = ReadRangeTo0BasedWithFormats(lws.UsedRange, leftFormats)
                Utils.ComReleaseHelper.Release(lws)
                Dim rws = DirectCast(wb.Worksheets(rightSheetName), Microsoft.Office.Interop.Excel.Worksheet)
                Dim rightFormats As String() = Nothing
                Dim rightArr = ReadRangeTo0BasedWithFormats(rws.UsedRange, rightFormats)
                Utils.ComReleaseHelper.Release(rws)
                If leftArr Is Nothing OrElse leftArr.GetLength(0) < 2 OrElse rightArr Is Nothing OrElse rightArr.GetLength(0) < 2 Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "左右表至少需要一行表头和一行数据"})
                    Return result
                End If
                pr.JoinRead(leftSheetName, leftArr.GetLength(0), rightSheetName, rightArr.GetLength(0))

                ' 解析关联键列索引
                Dim leftCols = ResolveKeyColumns(leftArr, leftKeys)
                Dim rightCols = ResolveKeyColumns(rightArr, rightKeys)
                If leftCols Is Nothing OrElse rightCols Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "无法解析关联键列，请检查左/右表的键列是否存在"})
                    Return result
                End If

                ' 右表哈希索引：key -> 数据行下标
                Dim rightHash As New Dictionary(Of String, List(Of Integer))()
                For r = 1 To rightArr.GetLength(0) - 1
                    Dim k = MakeRowKey(rightArr, r, rightCols)
                    If Not rightHash.ContainsKey(k) Then rightHash(k) = New List(Of Integer)()
                    rightHash(k).Add(r)
                Next
                Dim rightMatched(rightArr.GetLength(0) - 1) As Boolean

                ' 输出列：左表全部 + 右表除关联键外的列
                Dim rightColCount = rightArr.GetLength(1)
                Dim rightOutCols As New List(Of Integer)()
                For c = 0 To rightColCount - 1
                    If Not rightCols.Contains(c) Then rightOutCols.Add(c)
                Next
                If rightOutCols.Count = 0 Then rightOutCols.Add(rightCols(0))
                Dim outColInfo As New List(Of KeyValuePair(Of Integer, Integer))() ' (isLeft=0/1, colIndex 0 基)
                For c = 0 To leftArr.GetLength(1) - 1 : outColInfo.Add(New KeyValuePair(Of Integer, Integer)(0, c)) : Next
                For Each rc In rightOutCols : outColInfo.Add(New KeyValuePair(Of Integer, Integer)(1, rc)) : Next
                Dim outCols = outColInfo.Count

                Dim outRows As New List(Of List(Of Object))()
                If includeHeaders Then
                    Dim hdr = New List(Of Object)()
                    For Each info In outColInfo
                        Dim srcArr = If(info.Key = 0, leftArr, rightArr)
                        hdr.Add(NormalizeCell(srcArr(0, info.Value)))
                    Next
                    outRows.Add(hdr)
                End If

                Select Case joinType
                    Case "inner"
                        For lr = 1 To leftArr.GetLength(0) - 1
                            Dim k = MakeRowKey(leftArr, lr, leftCols)
                            If rightHash.ContainsKey(k) Then
                                For Each rr In rightHash(k)
                                    rightMatched(rr) = True
                                    outRows.Add(ComposeJoinRow(leftArr, lr, rightArr, rr, outColInfo))
                                Next
                            End If
                        Next
                    Case "left"
                        For lr = 1 To leftArr.GetLength(0) - 1
                            Dim k = MakeRowKey(leftArr, lr, leftCols)
                            If rightHash.ContainsKey(k) Then
                                For Each rr In rightHash(k)
                                    rightMatched(rr) = True
                                    outRows.Add(ComposeJoinRow(leftArr, lr, rightArr, rr, outColInfo))
                                Next
                            Else
                                outRows.Add(ComposeJoinRow(leftArr, lr, Nothing, -1, outColInfo))
                            End If
                        Next
                    Case "right"
                        For rr = 1 To rightArr.GetLength(0) - 1
                            Dim k = MakeRowKey(rightArr, rr, rightCols)
                            Dim lr = FindLeftRow(leftArr, leftCols, k)
                            If lr >= 0 Then
                                outRows.Add(ComposeJoinRow(leftArr, lr, rightArr, rr, outColInfo))
                            Else
                                outRows.Add(ComposeJoinRow(Nothing, -1, rightArr, rr, outColInfo))
                            End If
                        Next
                    Case "full"
                        For lr = 1 To leftArr.GetLength(0) - 1
                            Dim k = MakeRowKey(leftArr, lr, leftCols)
                            If rightHash.ContainsKey(k) Then
                                For Each rr In rightHash(k)
                                    rightMatched(rr) = True
                                    outRows.Add(ComposeJoinRow(leftArr, lr, rightArr, rr, outColInfo))
                                Next
                            Else
                                outRows.Add(ComposeJoinRow(leftArr, lr, Nothing, -1, outColInfo))
                            End If
                        Next
                        For rr = 1 To rightArr.GetLength(0) - 1
                            If Not rightMatched(rr) Then
                                outRows.Add(ComposeJoinRow(Nothing, -1, rightArr, rr, outColInfo))
                            End If
                        Next
                    Case Else
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = $"不支持的关联类型: {joinType}（支持 inner/left/right/full）"})
                        Return result
                End Select

                Dim dataCount = outRows.Count - (If(includeHeaders, 1, 0))
                If dataCount <= 0 Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "关联结果为空，没有匹配到任何行"})
                    Return result
                End If
                pr.JoinMatch(joinType, String.Join(", ", leftKeys), String.Join(", ", rightKeys), dataCount)

                ' ---------- 关联结果返回；若调用方传入 targetSheet，则落地到该表 ----------
                Dim landedSheetName = WriteResultToTargetSheet(args, outRows, "关联结果")
                Dim outData As New List(Of Object)()
                For Each orow In outRows
                    outData.Add(orow.ToArray())
                Next
                Dim resultObj = New Dictionary(Of String, Object) From {
                    {"success", True},
                    {"joinType", joinType},
                    {"outputRows", dataCount},
                    {"includeHeaders", includeHeaders},
                    {"data", outData}
                }
                If landedSheetName IsNot Nothing Then resultObj("writtenTo") = landedSheetName
                result.Content.Add(New McpTextContent With {.Text = JsonConvert.SerializeObject(resultObj, Formatting.Indented)})

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"跨表关联失败: {ex.Message}"})
            Finally
                app.DisplayAlerts = displayAlertsOriginal
            End Try
            Return result
        End Function
        ''' <summary>
        ''' 读取 Range 的 Value2 并归一化为 0 基二维数组（单单元格也统一为 1x1）
        ''' </summary>
        Private Function ReadRangeTo0Based(rng As Microsoft.Office.Interop.Excel.Range) As Object(,)
            If rng Is Nothing Then Return Nothing
            Dim rc = rng.Rows.Count
            Dim cc = rng.Columns.Count
            If rc < 1 OrElse cc < 1 Then Return Nothing
            Dim raw As Object = Nothing
            Try
                raw = rng.Value2
            Catch
                Return Nothing
            End Try
            If raw Is Nothing Then Return Nothing
            If TypeOf raw Is Object(,) Then
                Dim src = DirectCast(raw, Object(,))
                Dim l0 = src.GetLowerBound(0) : Dim u0 = src.GetUpperBound(0)
                Dim l1 = src.GetLowerBound(1) : Dim u1 = src.GetUpperBound(1)
                Dim rows = u0 - l0 + 1 : Dim cols = u1 - l1 + 1
                Dim out = New Object(rows - 1, cols - 1) {}
                For i = 0 To rows - 1
                    For j = 0 To cols - 1
                        out(i, j) = src(l0 + i, l1 + j)
                    Next
                Next
                Return out
            Else
                Dim out = New Object(0, 0) {}
                out(0, 0) = raw
                Return out
            End If
        End Function
        ''' <summary>
        ''' 读取 Range 数据并同时捕获每列的数字格式（取首个数据行的 NumberFormat），用于写回后还原数据类型显示。
        ''' NumberFormat 取该列第一个数据单元格的格式（日期/百分比/货币等的格式掩码）。
        ''' </summary>
        Private Function ReadRangeTo0BasedWithFormats(rng As Microsoft.Office.Interop.Excel.Range, ByRef formats As String()) As Object(,)
            Dim arr = ReadRangeTo0Based(rng)
            formats = Nothing
            If arr Is Nothing OrElse rng Is Nothing Then Return arr
            Dim urr = TryCast(rng, Microsoft.Office.Interop.Excel.Range)
            If urr Is Nothing Then Return arr
            Dim ws = TryCast(urr.Worksheet, Microsoft.Office.Interop.Excel.Worksheet)
            If ws Is Nothing Then Return arr
            Dim cols = arr.GetLength(1)
            Dim formatsArr = New String(cols - 1) {}
            Dim firstDataRowOnSheet = urr.Row + 1
            If firstDataRowOnSheet > urr.Row + urr.Rows.Count - 1 Then firstDataRowOnSheet = urr.Row + urr.Rows.Count - 1
            If firstDataRowOnSheet < urr.Row Then firstDataRowOnSheet = urr.Row
            For c = 0 To cols - 1
                Try
                    Dim cellCol = urr.Column + c
                    formatsArr(c) = CStr(DirectCast(ws.Cells(firstDataRowOnSheet, cellCol), Microsoft.Office.Interop.Excel.Range).NumberFormat)
                Catch
                    formatsArr(c) = ""
                End Try
            Next
            formats = formatsArr
            Return arr
        End Function
        ''' <summary>
        ''' 将每列数字格式应用到目标工作表的数据区域（保留日期/百分比/货币等显示格式）。
        ''' 仅覆盖非 General 的列，且从第 startRow 行开始、作用于 dataRows 行。
        ''' </summary>
        Private Sub ApplyColumnFormats(ws As Microsoft.Office.Interop.Excel.Worksheet, formats As String(), startRow As Integer, dataRows As Integer, maxCols As Integer)
            If ws Is Nothing OrElse formats Is Nothing OrElse maxCols <= 0 OrElse dataRows <= 0 Then Return
            For c = 0 To maxCols - 1
                Dim fmt As String = ""
                If c < formats.Length Then fmt = formats(c)
                If String.IsNullOrWhiteSpace(fmt) OrElse fmt.Trim() = "General" Then Continue For
                Try
                    Dim fr = DirectCast(ws.Range(ws.Cells(startRow, c + 1), ws.Cells(startRow + dataRows - 1, c + 1)), Microsoft.Office.Interop.Excel.Range)
                    fr.NumberFormat = fmt
                    Utils.ComReleaseHelper.Release(fr)
                Catch
                End Try
            Next
        End Sub
        ''' <summary>
        ''' 将二维数组写入工作表，并按每列格式掩码还原数据类型显示（用于拆分场景）。
        ''' </summary>
        Private Sub WriteArrayWithFormats(ws As Microsoft.Office.Interop.Excel.Worksheet, outArr As Object(,), hasHeader As Boolean, formats As String())
            Dim r = outArr.GetLength(0)
            Dim c = outArr.GetLength(1)
            If r <= 0 OrElse c <= 0 OrElse ws Is Nothing Then Return
            Dim rng = DirectCast(ws.Range(ws.Cells(1, 1), ws.Cells(r, c)), Microsoft.Office.Interop.Excel.Range)
            rng.Value2 = outArr
            Utils.ComReleaseHelper.Release(rng)
            Dim dataStart = If(hasHeader, 2, 1)
            Dim dataRows = r - (If(hasHeader, 1, 0))
            ApplyColumnFormats(ws, formats, dataStart, dataRows, c)
        End Sub
        ''' <summary>
        ''' 将可能为 COM 错误/其他类型的单元格值归一化为可写回的对象
        ''' </summary>
        Private Function NormalizeCell(value As Object) As Object
            If value Is Nothing Then Return Nothing
            Try
                Dim t = value.GetType()
                If t Is GetType(String) OrElse t Is GetType(Double) OrElse t Is GetType(Single) OrElse
                   t Is GetType(Integer) OrElse t Is GetType(Long) OrElse t Is GetType(Decimal) OrElse
                   t Is GetType(Boolean) OrElse t Is GetType(DateTime) OrElse t Is GetType(Byte) OrElse
                   t Is GetType(Short) Then Return value
                Dim strVal = value.ToString()
                If strVal IsNot Nothing AndAlso strVal.StartsWith("Error ") Then Return Nothing
                Return strVal
            Catch
                Return Nothing
            End Try
        End Function
        ''' <summary>
        ''' 将一维行单元的规范化文本与另一行对比，判断是否为同一"表头"。
        ''' 仅对比非空单元格的前 5 个，全部一致则视为表头行（兼容不同源表表头列数/排列差异）。
        ''' </summary>
        Private Function HeaderRowsMatch(a As Object(), b As Object()) As Boolean
            If a Is Nothing OrElse b Is Nothing Then Return False
            Dim ka As New List(Of String)()
            Dim kb As New List(Of String)()
            For i = 0 To a.Length - 1
                Dim v = a(i)
                If v IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(v.ToString()) Then ka.Add(v.ToString().Trim())
            Next
            For i = 0 To b.Length - 1
                Dim v = b(i)
                If v IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(v.ToString()) Then kb.Add(v.ToString().Trim())
            Next
            If ka.Count = 0 AndAlso kb.Count = 0 Then Return True
            If ka.Count = 0 OrElse kb.Count = 0 Then Return False
            Dim n = Math.Min(Math.Min(ka.Count, kb.Count), 5)
            For i = 0 To n - 1
                If Not String.Equals(ka(i), kb(i), StringComparison.OrdinalIgnoreCase) Then Return False
            Next
            Return True
        End Function
        ''' <summary>
        ''' 判定一行是否全空（用于裁剪 UsedRange 的幽灵空行/空列）。
        ''' </summary>
        Private Function RowIsBlank(rowVals As Object()) As Boolean
            If rowVals Is Nothing Then Return True
            For Each v In rowVals
                If v IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(v.ToString()) Then Return False
            Next
            Return True
        End Function
        ''' <summary>
        ''' 裁剪一个数据矩阵的尾部/首部完全空白的行，以及尾部完全空白的列（去掉 UsedRange 幽灵空区）。
        ''' 至少保留 1 行 1 列。返回可能为新的更紧凑数组。
        ''' </summary>
        Private Function TrimPhantomRanges(arr As Object(,)) As Object(,)
            If arr Is Nothing Then Return arr
            Dim rows = arr.GetLength(0)
            Dim cols = arr.GetLength(1)
            If rows = 0 OrElse cols = 0 Then Return arr

            Dim firstRow As Integer = 0
            While firstRow < rows AndAlso RowIsBlank(RowOf(arr, firstRow))
                firstRow += 1
            End While
            Dim lastRow As Integer = rows - 1
            While lastRow > firstRow AndAlso RowIsBlank(RowOf(arr, lastRow))
                lastRow -= 1
            End While
            If lastRow < firstRow Then Return New Object(0, 0) {}

            Dim lastCol As Integer = cols - 1
            While lastCol > 0
                Dim blank As Boolean = True
                For r = firstRow To lastRow
                    If arr(r, lastCol) IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(arr(r, lastCol).ToString()) Then
                        blank = False
                        Exit For
                    End If
                Next
                If Not blank Then Exit While
                lastCol -= 1
            End While

            Dim newRows = lastRow - firstRow + 1
            Dim newCols = lastCol + 1
            Dim trimmed = New Object(newRows - 1, newCols - 1) {}
            For r = 0 To newRows - 1
                For c = 0 To newCols - 1
                    trimmed(r, c) = arr(firstRow + r, c)
                Next
            Next
            Return trimmed
        End Function
        ''' <summary>
        ''' 取出某一行（作为 Object() 数组）。
        ''' </summary>
        Private Function RowOf(arr As Object(,), r As Integer) As Object()
            Dim cols = arr.GetLength(1)
            Dim rowVals = New Object(cols - 1) {}
            For c = 0 To cols - 1
                rowVals(c) = arr(r, c)
            Next
            Return rowVals
        End Function
        ''' <summary>
        ''' 判定一个矩阵是否整表空白（裁剪后仍无任何非空内容，用于跳过空源表/幽灵表）。
        ''' </summary>
        Private Function ArrayIsAllBlank(arr As Object(,)) As Boolean
            If arr Is Nothing Then Return True
            For r = 0 To arr.GetLength(0) - 1
                For c = 0 To arr.GetLength(1) - 1
                    If arr(r, c) IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(arr(r, c).ToString()) Then Return False
                Next
            Next
            Return True
        End Function
        ''' <summary>
        ''' 将一组行从指定起始行写入工作表（整块一次写，供逐表插入分步使用）。
        ''' </summary>
        Private Sub WriteBlockRows(ws As Microsoft.Office.Interop.Excel.Worksheet, rows As List(Of List(Of Object)), startRow As Integer, cols As Integer)
            If rows Is Nothing OrElse rows.Count = 0 Then Return
            Dim n = rows.Count
            Dim cc = Math.Max(cols, 1)
            Dim blk = New Object(n - 1, cc - 1) {}
            For i = 0 To n - 1
                For c = 0 To cc - 1
                    If rows(i) IsNot Nothing AndAlso c < rows(i).Count Then blk(i, c) = rows(i)(c) Else blk(i, c) = Nothing
                Next
            Next
            Dim rng = DirectCast(ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + n - 1, cc)), Microsoft.Office.Interop.Excel.Range)
            rng.Value2 = blk
            Utils.ComReleaseHelper.Release(rng)
        End Sub
        ''' <summary>
        ''' 将行集合写入指定工作表（自动覆盖到整块区域）
        ''' </summary>
        Private Sub WriteObjectRowsToSheet(ws As Microsoft.Office.Interop.Excel.Worksheet, rows As List(Of List(Of Object)), cols As Integer)
            If rows Is Nothing OrElse rows.Count = 0 Then Return
            Dim out = New Object(rows.Count - 1, Math.Max(cols, 1) - 1) {}
            For i = 0 To rows.Count - 1
                Dim src = rows(i)
                For j = 0 To Math.Max(cols, 1) - 1
                    If src IsNot Nothing AndAlso j < src.Count Then out(i, j) = src(j) Else out(i, j) = Nothing
                Next
            Next
            Dim rng = DirectCast(ws.Range(ws.Cells(1, 1), ws.Cells(rows.Count, Math.Max(cols, 1))), Microsoft.Office.Interop.Excel.Range)
            rng.Value2 = out
            Utils.ComReleaseHelper.Release(rng)
        End Sub
        ''' <summary>
        ''' 【结果落地】计算工具的可选写回：调用方传入非空 targetSheet 时，把 rows 写入该工作表（不存在则新建，含表头时选中表头行样式加粗），并返回落地后的表名；
        ''' 未传 targetSheet 或写入失败时返回 Nothing，调用方仅返回数据。用于落实“算”与“存”解耦后结果的明确落地路径。
        ''' </summary>
        Private Function WriteResultToTargetSheet(args As Dictionary(Of String, Object), rows As List(Of List(Of Object)), Optional defaultName As String = "结果") As String
            Dim ts = GetArg(Of String)(args, "targetSheet", "")
            If String.IsNullOrWhiteSpace(ts) Then Return Nothing
            Dim app = Globals.ThisAddIn.Application
            If app Is Nothing OrElse app.ActiveWorkbook Is Nothing Then Return Nothing
            Dim wb = app.ActiveWorkbook
            Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Try
                Try
                    ws = DirectCast(wb.Worksheets(ts), Microsoft.Office.Interop.Excel.Worksheet)
                Catch
                    ws = DirectCast(wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)), Microsoft.Office.Interop.Excel.Worksheet)
                    ws.Name = ts
                End Try
                If rows IsNot Nothing AndAlso rows.Count > 0 Then
                    Dim cols = 0
                    For Each r In rows
                        If r IsNot Nothing AndAlso r.Count > cols Then cols = r.Count
                    Next
                    ' 【关键修复】写入新数据前先清 UsedRange 残留——
                    ' 防止旧表更多行的数据在新表行数减少时残留，导致 create_chart 误读旧数据混入
                    Try
                        Dim oldUsed = DirectCast(ws.UsedRange, Microsoft.Office.Interop.Excel.Range)
                        oldUsed.ClearContents()
                        Utils.ComReleaseHelper.Release(oldUsed)
                    Catch
                    End Try
                    WriteObjectRowsToSheet(ws, rows, cols)
                    ' 表头行加粗，便于阅读
                    If rows(0) IsNot Nothing AndAlso rows(0).Count > 0 Then
                        Dim hdrRng = DirectCast(ws.Range(ws.Cells(1, 1), ws.Cells(1, rows(0).Count)), Microsoft.Office.Interop.Excel.Range)
                        hdrRng.Font.Bold = True
                        Utils.ComReleaseHelper.Release(hdrRng)
                    End If
                    ws.Activate()
                End If
                Utils.ComReleaseHelper.Release(ws)
            Catch
                If ws IsNot Nothing Then Utils.ComReleaseHelper.Release(ws)
                Return Nothing
            End Try
            Return ts
        End Function
        ''' <summary>
        ''' 解析关联键列参数：优先读取单数值 primaryKey，否则读取复数 JSON 数组；单字符串也包装成单个元素
        ''' </summary>
        Private Function ParseKeyColumns(args As Dictionary(Of String, Object), primaryKey As String, arrayKey As String) As List(Of String)
            Dim out As New List(Of String)()
            If args Is Nothing Then Return out
            If args.ContainsKey(primaryKey) AndAlso args(primaryKey) IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(args(primaryKey).ToString()) Then
                out.Add(args(primaryKey).ToString())
                Return out
            End If
            If args.ContainsKey(arrayKey) AndAlso args(arrayKey) IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(args(arrayKey).ToString()) Then
                Dim raw = args(arrayKey).ToString()
                Try
                    Dim parsed = JsonConvert.DeserializeObject(Of List(Of String))(raw)
                    If parsed IsNot Nothing Then Return parsed
                Catch
                End Try
                ' 不是 JSON 数组则当单个键处理
                out.Add(raw)
                Return out
            End If
            Return out
        End Function
        ''' <summary>
        ''' 解析关联键列（表头名 / 1 基列序号 / 列字母），返回 0 基列下标列表；任一列无法解析则返回 Nothing
        ''' </summary>
        Private Function ResolveKeyColumns(arr As Object(,), keys As List(Of String)) As List(Of Integer)
            Dim colCount = arr.GetLength(1)
            Dim out As New List(Of Integer)()
            For Each k In keys
                Dim idx = -1
                Dim trimmed = k.Trim()
                ' 1) 表头名
                For c = 0 To colCount - 1
                    Dim h = arr(0, c)
                    If h IsNot Nothing AndAlso String.Equals(h.ToString().Trim(), trimmed, StringComparison.OrdinalIgnoreCase) Then
                        idx = c
                        Exit For
                    End If
                Next
                ' 2) 1 基列序号
                If idx < 0 Then
                    Dim n As Integer
                    If Integer.TryParse(trimmed, n) AndAlso n >= 1 AndAlso n <= colCount Then
                        idx = n - 1
                    End If
                End If
                ' 3) 列字母
                If idx < 0 AndAlso trimmed.Length <= 3 Then
                    Dim isLetterKey = True
                    For Each ch In trimmed
                        If Not Char.IsLetter(ch) Then isLetterKey = False : Exit For
                    Next
                    If isLetterKey Then
                        Dim cn = ExcelMcpCoreUtils.ColumnLetterToNumber(trimmed.ToUpper())
                        If cn >= 1 AndAlso cn <= colCount Then idx = cn - 1
                    End If
                End If
                If idx < 0 Then Return Nothing
                out.Add(idx)
            Next
            Return out
        End Function
        ''' <summary>
        ''' 生成行的关联键（多列拼接，逗号/空值安全）
        ''' </summary>
        Private Function MakeRowKey(arr As Object(,), rowIdx As Integer, cols As List(Of Integer)) As String
            Dim parts(cols.Count - 1) As String
            For i = 0 To cols.Count - 1
                Dim v = arr(rowIdx, cols(i))
                If v Is Nothing Then parts(i) = "" Else parts(i) = v.ToString()
            Next
            Return String.Join(ChrW(1), parts)
        End Function
        ''' <summary>
        ''' 拼接左右表的一行数据到输出行
        ''' </summary>
        Private Function ComposeJoinRow(leftArr As Object(,), lr As Integer, rightArr As Object(,), rr As Integer, outColInfo As List(Of KeyValuePair(Of Integer, Integer))) As List(Of Object)
            Dim row = New List(Of Object)()
            For Each info In outColInfo
                If info.Key = 0 Then
                    If leftArr IsNot Nothing AndAlso lr >= 0 Then
                        row.Add(NormalizeCell(leftArr(lr, info.Value)))
                    Else
                        row.Add(Nothing)
                    End If
                Else
                    If rightArr IsNot Nothing AndAlso rr >= 0 Then
                        row.Add(NormalizeCell(rightArr(rr, info.Value)))
                    Else
                        row.Add(Nothing)
                    End If
                End If
            Next
            Return row
        End Function
        ''' <summary>
        ''' 在左表中按关联键查找第一行（0 基数据行，找不到返回 -1）
        ''' </summary>
        Private Function FindLeftRow(leftArr As Object(,), leftCols As List(Of Integer), key As String) As Integer
            For r = 1 To leftArr.GetLength(0) - 1
                If MakeRowKey(leftArr, r, leftCols) = key Then Return r
            Next
            Return -1
        End Function

    End Class

End Namespace