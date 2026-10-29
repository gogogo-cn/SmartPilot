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
        ''' 【算】纯去重：读取范围、按指定列去重，返回唯一数据（二维数组）。
        ''' 不写回 Excel —— 落盘交给 excel_write_range。不再支持就地删除/标记（那属"存/画"职责）。
        ''' </summary>
        Private Function ExecuteRemoveDuplicates(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim rangeAddr As String = GetArg(Of String)(args, "range", "")
                Dim columnsArg As Object = Nothing
                args?.TryGetValue("columns", columnsArg)
                Dim hasHeaderArg = GetArg(Of String)(args, "hasHeaderRow", "")
                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If
                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim rng As Microsoft.Office.Interop.Excel.Range = Nothing
                Try
                    ' ---------- 1) 工作表解析 ----------
                    If Not String.IsNullOrEmpty(sheetName) Then
                        ws = TryCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                    End If
                    If ws Is Nothing Then ws = TryCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    If ws Is Nothing Then
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = "无法定位有效工作表，请先激活包含数据的工作表。"})
                        Return result
                    End If
                    ' ---------- 2) 范围解析：上下文选区 → Selection → UsedRange 回退链 ----------
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
                                    Dim selSheet As String = ""
                                    Try : selSheet = CStr(selRng.Worksheet.Name) : Catch : End Try
                                    If selSheet = ws.Name Then rng = selRng
                                End If
                            Catch : End Try
                        End If
                        If rng Is Nothing Then rng = ws.UsedRange
                    End If
                    If rng Is Nothing OrElse rng.Rows.Count < 2 Then
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = $"范围 ({If(rng Is Nothing, "空", rng.Address(False, False))}) 数据不足 2 行，没有可去重的内容。"})
                        Return result
                    End If
                    Dim firstCol As Integer = rng.Column
                    Dim lastCol As Integer = rng.Column + rng.Columns.Count - 1
                    Dim totalCols As Integer = lastCol - firstCol + 1
                    Dim addrDisplay As String = rng.Address(False, False)
                    ' ---------- 3) hasHeaderRow 自动检测 ----------
                    Dim hasHeader As Microsoft.Office.Interop.Excel.XlYesNoGuess
                    If String.Equals(hasHeaderArg, "true", StringComparison.OrdinalIgnoreCase) OrElse
                       String.Equals(hasHeaderArg, "1", StringComparison.OrdinalIgnoreCase) OrElse
                       String.Equals(hasHeaderArg, "yes", StringComparison.OrdinalIgnoreCase) Then
                        hasHeader = Microsoft.Office.Interop.Excel.XlYesNoGuess.xlYes
                    ElseIf String.Equals(hasHeaderArg, "false", StringComparison.OrdinalIgnoreCase) OrElse
                           String.Equals(hasHeaderArg, "0", StringComparison.OrdinalIgnoreCase) OrElse
                           String.Equals(hasHeaderArg, "no", StringComparison.OrdinalIgnoreCase) Then
                        hasHeader = Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo
                    Else
                        hasHeader = Microsoft.Office.Interop.Excel.XlYesNoGuess.xlGuess
                    End If
                    ' ---------- 4) columns 参数解析 ----------
                    Dim colList As New List(Of Integer)()
                    If columnsArg Is Nothing Then
                        For c = firstCol To lastCol : colList.Add(c) : Next
                    Else
                        Dim headerArr As Object(,) = Nothing
                        Dim headerRow As Microsoft.Office.Interop.Excel.Range = Nothing
                        Try
                            If hasHeader <> Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo Then
                                headerRow = DirectCast(rng.Rows(1), Microsoft.Office.Interop.Excel.Range)
                                If rng.Columns.Count = 1 Then
                                    Dim v = headerRow.Value2
                                    headerArr = New Object(1, 1) {}
                                    headerArr(1, 1) = v
                                Else
                                    headerArr = TryCast(headerRow.Value2, Object(,))
                                End If
                            End If
                        Catch : End Try
                        Dim colInputs As New List(Of String)()
                        Try
                            If TypeOf columnsArg Is String Then
                                Dim s = CStr(columnsArg).Trim()
                                If s.StartsWith("[") AndAlso s.EndsWith("]") Then
                                    Try
                                        Dim parsed = Newtonsoft.Json.JsonConvert.DeserializeObject(Of List(Of Object))(s)
                                        For Each o In parsed
                                            If o Is Nothing Then Continue For
                                            colInputs.Add(o.ToString())
                                        Next
                                    Catch
                                        colInputs.Add(s)
                                    End Try
                                Else
                                    For Each p In s.Split({","c, ";"c, "|"c, " "c, "，"c, "；"c}, StringSplitOptions.RemoveEmptyEntries)
                                        colInputs.Add(p.Trim())
                                    Next
                                End If
                            ElseIf TypeOf columnsArg Is System.Collections.IEnumerable Then
                                For Each o In DirectCast(columnsArg, System.Collections.IEnumerable)
                                    If o Is Nothing Then Continue For
                                    colInputs.Add(o.ToString())
                                Next
                            Else
                                colInputs.Add(columnsArg.ToString())
                            End If
                        Catch
                            colInputs.Clear()
                        End Try
                        For Each raw In colInputs
                            If String.IsNullOrWhiteSpace(raw) Then Continue For
                            Dim token As String = raw.Trim().Trim(""""c, "'"c, "["c, "]"c, "("c, ")"c).Trim()
                            If token = "" Then Continue For
                            Dim resolved As Integer = -1
                            Dim n As Integer
                            If Integer.TryParse(token, n) AndAlso n > 0 Then
                                If n <= totalCols Then
                                    resolved = firstCol + n - 1
                                ElseIf n >= firstCol AndAlso n <= lastCol Then
                                    resolved = n
                                End If
                            End If
                            If resolved = -1 AndAlso System.Text.RegularExpressions.Regex.IsMatch(token, "^[A-Za-z]{1,3}$") Then
                                Try
                                    Dim up As String = token.ToUpperInvariant()
                                    Dim acc As Integer = 0
                                    For Each ch In up
                                        acc = acc * 26 + (AscW(ch) - AscW("A"c) + 1)
                                    Next
                                    If acc > 0 Then
                                        If acc <= totalCols Then
                                            resolved = firstCol + acc - 1
                                        ElseIf acc >= firstCol AndAlso acc <= lastCol Then
                                            resolved = acc
                                        End If
                                    End If
                                Catch : End Try
                            End If
                            If resolved = -1 AndAlso headerArr IsNot Nothing Then
                                For c = 1 To headerArr.GetLength(1)
                                    Dim hv As Object = headerArr(1, c)
                                    If hv IsNot Nothing AndAlso String.Equals(hv.ToString().Trim(), token, StringComparison.OrdinalIgnoreCase) Then
                                        resolved = firstCol + c - 1
                                        Exit For
                                    End If
                                Next
                            End If
                            If resolved >= firstCol AndAlso resolved <= lastCol AndAlso Not colList.Contains(resolved) Then
                                colList.Add(resolved)
                            End If
                        Next
                        If colList.Count = 0 Then
                            For c = firstCol To lastCol : colList.Add(c) : Next
                        End If
                    End If
                    ' ---------- 5) 读取数据并内存去重 ----------
                    Dim values As Object(,) = Nothing
                    If totalCols = 1 AndAlso rng.Rows.Count = 1 Then
                        values = New Object(0, 0) {{rng.Value2}}
                    Else
                        values = TryCast(rng.Value2, Object(,))
                    End If
                    If values Is Nothing Then
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = "无法读取范围数据。"})
                        Return result
                    End If
                    Dim valueCols As New List(Of Integer)()
                    For Each gc In colList
                        Dim vc = gc - firstCol + 1
                        If vc >= 1 AndAlso vc <= values.GetLength(1) Then valueCols.Add(vc)
                    Next
                    If valueCols.Count = 0 Then
                        For vc = 1 To values.GetLength(1) : valueCols.Add(vc) : Next
                    End If
                    Dim startRow As Integer = 1
                    Dim hasHeaderRowFlag As Boolean = True
                    If hasHeader = Microsoft.Office.Interop.Excel.XlYesNoGuess.xlYes Then
                        startRow = 2
                    ElseIf hasHeader = Microsoft.Office.Interop.Excel.XlYesNoGuess.xlGuess Then
                        Dim isBold As Boolean = False
                        Try
                            isBold = (CBool(DirectCast(DirectCast(rng.Rows(1), Microsoft.Office.Interop.Excel.Range).Font, Microsoft.Office.Interop.Excel.Font).Bold) = True)
                        Catch : End Try
                        If isBold Then
                            startRow = 2
                        Else
                            hasHeaderRowFlag = False
                        End If
                    Else
                        hasHeaderRowFlag = False
                    End If
                    ' 表头数组
                    Dim headers() As Object = Nothing
                    If hasHeaderRowFlag AndAlso values.GetLength(0) >= 1 Then
                        Dim hdr(values.GetLength(1) - 1) As Object
                        For cc = 1 To values.GetLength(1)
                            hdr(cc - 1) = If(values(1, cc) Is Nothing, "", values(1, cc))
                        Next
                        headers = hdr
                    End If
                    ' 去重
                    Dim hs As New HashSet(Of String)(StringComparer.Ordinal)
                    Dim uniqueRows As New List(Of Object())(50)
                    Dim duplicateRows As New List(Of Integer)()   ' 重复行的绝对行号（供前端"仅高亮"标黄）
                    Dim duplicateCount As Integer = 0              ' 实际重复数据行数（不含空行）
                    Dim skippedBlankRows As Integer = 0            ' 全空行的数量（空行不参与去重校验）
                    For r = startRow To values.GetLength(0)
                        ' 全空行不参与去重：既不当"唯一数据"保留，也不误判为"重复行"。
                        ' 否则 3 行空行 key 都是 "|||..." 相同串，会被扣成 1 行并上报成"去重重复行"。
                        ' 空行属脏数据，跳过留给"删除空行"/写回清残留物理删除。
                        Dim isBlankRow As Boolean = True
                        For Each vc In valueCols
                            Dim vCk As Object = values(r, vc)
                            If Not (vCk Is Nothing OrElse IsDBNull(vCk) OrElse CStr(vCk).Trim() = "") Then
                                isBlankRow = False
                                Exit For
                            End If
                        Next
                        If isBlankRow Then
                            skippedBlankRows += 1
                            Continue For
                        End If
                        Dim sb As New System.Text.StringBuilder()
                        For Each vc In valueCols
                            Dim vObj As Object = values(r, vc)
                            If vObj Is Nothing Then
                                sb.Append("||")
                            Else
                                sb.Append("|").Append(vObj.ToString())
                            End If
                        Next
                        If hs.Add(sb.ToString()) Then
                            Dim rowArr(values.GetLength(1) - 1) As Object
                            For cc = 1 To values.GetLength(1)
                                rowArr(cc - 1) = values(r, cc)
                            Next
                            uniqueRows.Add(rowArr)
                        Else
                            ' 重复行：记录绝对行号（rng.Row 为顶行行号，r 为相对 1 基偏移）
                            duplicateCount += 1
                            duplicateRows.Add(rng.Row + r - 1)
                        End If
                    Next
                    ' ---------- 6) 构建返回 ----------
                    Dim dataRows As New List(Of Object)()
                    ' 合并输出：header 行（若有）+ 唯一数据行；若无表头则只返回数据行
                    If headers IsNot Nothing Then dataRows.Add(headers)
                    For Each urow In uniqueRows : dataRows.Add(urow) : Next
                    ' 可选落地：调用方传入 targetSheet 时写回结果到该表（存在则覆盖，不存在则新建）
                    Dim landedSheetName As String = Nothing
                    Dim targetTs = GetArg(Of String)(args, "targetSheet", "")
                    If Not String.IsNullOrWhiteSpace(targetTs) Then
                        Dim rowLists As New List(Of List(Of Object))()
                        For Each dr In dataRows
                            Dim oarr = TryCast(dr, Object())
                            If oarr IsNot Nothing Then rowLists.Add(oarr.ToList())
                        Next
                        landedSheetName = WriteResultToTargetSheet(args, rowLists, "去重结果")
                    End If
                    Dim output As New Dictionary(Of String, Object)()
                    output("success") = True
                    output("deduplicated") = True
                    output("workbook") = excelApp.ActiveWorkbook.Name
                    output("worksheet") = ws.Name
                    output("range") = addrDisplay
                    output("uniqueCount") = uniqueRows.Count
                    output("rawCount") = values.GetLength(0) - (startRow - 1)
                    output("removedCount") = duplicateCount
                    output("duplicateCount") = duplicateCount
                    output("skippedBlankRows") = skippedBlankRows
                    output("duplicateRows") = duplicateRows
                    output("hasHeader") = hasHeaderRowFlag
                    output("headers") = headers
                    output("data") = dataRows
                    If landedSheetName IsNot Nothing Then output("writtenTo") = landedSheetName
                    result.Content.Add(New McpTextContent With {.Text = JsonConvert.SerializeObject(output, Formatting.Indented)})
                    result.TaskCompleted = True
                Catch ex2 As Exception
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"去重计算异常：{ex2.Message}"})
                    Return result
                Finally
                    If rng IsNot Nothing Then Utils.ComReleaseHelper.Release(rng)
                    If ws IsNot Nothing Then Utils.ComReleaseHelper.Release(ws)
                End Try
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"工具初始化失败 [{ex.GetType().Name}]: {ex.Message}"})
                Return result
            End Try
            Return result
        End Function

    End Class

End Namespace