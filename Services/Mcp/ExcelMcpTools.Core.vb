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

        '#Region "工具实现 - 核心读写"

        Private Function ExecuteReadRange(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "Excel 应用程序未就绪"})
                    Return result
                End If
                If excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim rangeAddr As String = GetArg(Of String)(args, "range", "")
                Dim hasHeaders As Boolean = GetArg(Of Boolean)(args, "hasHeaders", True)
                Dim maxRows As Integer = GetArg(Of Integer)(args, "maxRows", 10000)
                Dim returnFormat As String = GetArg(Of String)(args, "returnFormat", "array")
                Dim skipEmptyRows As Boolean = GetArg(Of Boolean)(args, "skipEmptyRows", False)
                Dim maxColumns As Integer = GetArg(Of Integer)(args, "maxColumns", 100)
                Dim readFormulas As Boolean = GetArg(Of Boolean)(args, "readFormulas", False)
                ' 自动探测（默认开启）：任何需要读 Excel 数据的任务，未显式传 range 且选区为单格时，
                ' 视为"用户未真正框选数据区"，回退 UsedRange 探测整表数据区再读取。
                ' 影响所有读取路径：skill/py/html-display/text 各 pipeline、CRUD、MultiStepEngine 等。
                ' 若调用方只想读当前单格，可显式传 autoProbe=false 关闭。
                Dim autoProbe As Boolean = GetArg(Of Boolean)(args, "autoProbe", True)

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim wsNeedsRelease As Boolean = False
                If Not String.IsNullOrEmpty(sheetName) Then
                    Try
                        ws = TryCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                    Catch ex As Exception
                        ' 工作表不存在（如 LLM 从用户消息猜测的表名与真实表名不符）→ 回退活动工作表
                        ' 否则 Worksheets(sheetName) 会抛 DISP_E_BADINDEX(0x8002000B) 无效索引
                        System.Diagnostics.Debug.WriteLine($"[ExecuteReadRange] 工作表 '{sheetName}' 不存在: {ex.Message}，回退到活动工作表")
                        ws = Nothing
                    End Try
                    wsNeedsRelease = True
                End If
                If ws Is Nothing Then
                    ws = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    wsNeedsRelease = False
                End If

                Dim rng As Microsoft.Office.Interop.Excel.Range = Nothing
                If String.IsNullOrEmpty(rangeAddr) Then
                    ' 无参数时：优先使用当前选中区域（Selection），与excel_clear_range保持一致
                    ' 用户说"读取选中数据"时期望读的是当前选中的区域，而非整个UsedRange
                    Try
                        Dim sel = excelApp.Selection
                        If TypeOf sel Is Microsoft.Office.Interop.Excel.Range Then
                            Dim selRng = DirectCast(sel, Microsoft.Office.Interop.Excel.Range)
                            ' 确保选区在当前工作表（通过名称比较避免COM对象泄漏）
                            Dim selSheetName As String = ""
                            Try
                                selSheetName = CStr(selRng.Worksheet.Name)
                            Catch
                                selSheetName = ""
                            End Try
                            If selSheetName <> "" AndAlso selSheetName = ws.Name Then
                                rng = selRng
                                ' 自动探测：选区为单格（用户默认是想分析"这张表"）时，优先用 CurrentRegion 扩展为
                                ' 围绕锚点的连续数据区，更贴合"点进表格里就想分析整张表"的意图；仅无连续区时回退
                                ' UsedRange。避免只用 UsedRange 时因 Excel"使用区域"未刷新而只拾取活动单元格，
                                ' 导致只分析到第一行（表现为"只拿到一行数据"）。
                                If autoProbe Then
                                    Try
                                        If rng.Rows.Count = 1 AndAlso rng.Columns.Count = 1 Then
                                            Dim expanded As Microsoft.Office.Interop.Excel.Range = Nothing
                                            Try
                                                expanded = rng.CurrentRegion
                                            Catch
                                                expanded = Nothing
                                            End Try
                                            Utils.ComReleaseHelper.Release(rng)
                                            If expanded IsNot Nothing AndAlso expanded.Rows.Count > 1 Then
                                                rng = expanded
                                                System.Diagnostics.Debug.WriteLine($"[ExecuteReadRange] autoProbe: 单格选区→CurrentRegion {rng.Address}")
                                            Else
                                                rng = ws.UsedRange
                                                If expanded IsNot Nothing Then Utils.ComReleaseHelper.Release(expanded)
                                                System.Diagnostics.Debug.WriteLine($"[ExecuteReadRange] autoProbe: 单格选区→UsedRange {rng.Address}")
                                            End If
                                        End If
                                    Catch
                                    End Try
                                End If
                            Else
                                Utils.ComReleaseHelper.Release(selRng)
                                rng = Nothing
                            End If
                        End If
                    Catch
                        rng = Nothing
                    End Try
                    If rng Is Nothing Then
                        rng = ws.UsedRange
                        System.Diagnostics.Debug.WriteLine($"[ExecuteReadRange] Selection不可用，回退到UsedRange")
                    Else
                        System.Diagnostics.Debug.WriteLine($"[ExecuteReadRange] 使用当前Selection: {rng.Address}")
                    End If
                Else
                    Try
                        rng = ws.Range(rangeAddr)
                        ' autoProbe：单格（用户未真框选数据区）时，优先用 CurrentRegion 扩展为连续数据区
                        '$ CurrentRegion 更贴合"点进表格里就想分析整张表"的意图；仅无连续区时回退 UsedRange。
                        If autoProbe Then
                            Try
                                If rng.Rows.Count = 1 AndAlso rng.Columns.Count = 1 Then
                                    Dim expanded As Microsoft.Office.Interop.Excel.Range = Nothing
                                    Try
                                        expanded = rng.CurrentRegion
                                    Catch
                                        expanded = Nothing
                                    End Try
                                    Utils.ComReleaseHelper.Release(rng)
                                    If expanded IsNot Nothing AndAlso expanded.Rows.Count > 1 Then
                                        rng = expanded
                                        System.Diagnostics.Debug.WriteLine($"[ExecuteReadRange] autoProbe: 注入单格{rangeAddr}→CurrentRegion {rng.Address}")
                                    Else
                                        rng = ws.UsedRange
                                        If expanded IsNot Nothing Then Utils.ComReleaseHelper.Release(expanded)
                                        System.Diagnostics.Debug.WriteLine($"[ExecuteReadRange] autoProbe: 注入单格{rangeAddr}→UsedRange {rng.Address}")
                                    End If
                                End If
                            Catch
                            End Try
                        End If
                        System.Diagnostics.Debug.WriteLine($"[ExecuteReadRange] 使用指定范围: {rangeAddr}")
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine($"[ExecuteReadRange] Range调用失败: {ex.Message}，尝试回退")
                        ' 回退 1：尝试使用当前选区
                        Try
                            rng = DirectCast(excelApp.Selection, Microsoft.Office.Interop.Excel.Range)
                            System.Diagnostics.Debug.WriteLine($"[ExecuteReadRange] 回退到当前选区: {rng.Address}")
                        Catch ex2 As Exception
                            System.Diagnostics.Debug.WriteLine($"[ExecuteReadRange] 当前选区回退失败: {ex2.Message}")
                        End Try
                        ' 回退 2：如果选区也失败，使用 UsedRange
                        If rng Is Nothing Then
                            rng = ws.UsedRange
                            System.Diagnostics.Debug.WriteLine($"[ExecuteReadRange] 回退到 UsedRange")
                        End If
                    End Try
                End If

                If rng Is Nothing Then
                    result.Content.Add(New McpTextContent With {.Text = "[]"})
                    Return result
                End If

                System.Diagnostics.Debug.WriteLine($"[ExecuteReadRange] 范围地址: {rng.Address}, 行数: {rng.Rows.Count}, 列数: {rng.Columns.Count}")

                ' 读取数据为二维数组
                Dim data As Object(,)
                If readFormulas Then
                    data = DirectCast(rng.Formula, Object(,))
                Else
                    data = DirectCast(rng.Value2, Object(,))
                End If
                If data Is Nothing Then
                    result.Content.Add(New McpTextContent With {.Text = "[]"})
                    Return result
                End If

                ' 裁剪到实际数据区域（去除末尾空行/空列）
                ' 解决 UsedRange 返回整表缓存大范围（如 A1:Z1000）但实际数据只有小区域的问题
                Dim _lastDataRow = data.GetUpperBound(0)
                Dim _lastDataCol = data.GetUpperBound(1)
                Dim _foundCell As Boolean

                _foundCell = False
                For _r = data.GetUpperBound(0) To data.GetLowerBound(0) Step -1
                    For _c = data.GetLowerBound(1) To data.GetUpperBound(1)
                        Dim _v = data(_r, _c)
                        If _v IsNot Nothing AndAlso Not Convert.IsDBNull(_v) Then
                            _lastDataRow = _r
                            _foundCell = True
                            Exit For
                        End If
                    Next
                    If _foundCell Then Exit For
                Next

                _foundCell = False
                For _c = data.GetUpperBound(1) To data.GetLowerBound(1) Step -1
                    For _r = data.GetLowerBound(0) To _lastDataRow
                        Dim _v = data(_r, _c)
                        If _v IsNot Nothing AndAlso Not Convert.IsDBNull(_v) Then
                            _lastDataCol = _c
                            _foundCell = True
                            Exit For
                        End If
                    Next
                    If _foundCell Then Exit For
                Next

                If _lastDataRow < data.GetUpperBound(0) OrElse _lastDataCol < data.GetUpperBound(1) Then
                    Dim _newRows = _lastDataRow - data.GetLowerBound(0) + 1
                    Dim _newCols = _lastDataCol - data.GetLowerBound(1) + 1
                    Dim _trimmed(_newRows - 1, _newCols - 1) As Object
                    For _r = 0 To _newRows - 1
                        For _c = 0 To _newCols - 1
                            _trimmed(_r, _c) = data(data.GetLowerBound(0) + _r, data.GetLowerBound(1) + _c)
                        Next
                    Next
                    data = _trimmed
                    ' 更新 rng 引用到实际数据区域
                    Dim _startCell = rng.Cells(1, 1)
                    Dim _endCell = rng.Cells(_newRows, _newCols)
                    Dim _newRng = ws.Range(_startCell, _endCell)
                    Utils.ComReleaseHelper.Release(rng)
                    rng = _newRng
                    System.Diagnostics.Debug.WriteLine($"[ExecuteReadRange] 裁剪到实际数据: {rng.Address}, 行数={_newRows}, 列数={_newCols}")
                End If

                ' maxColumns：限制列数，防止单元格爆炸
                Dim _colTruncated As Boolean = False
                Dim _actualCols = data.GetUpperBound(1) - data.GetLowerBound(1) + 1
                If _actualCols > maxColumns Then
                    Dim _newRows = data.GetUpperBound(0) - data.GetLowerBound(0) + 1
                    Dim _trimmedCols(_newRows - 1, maxColumns - 1) As Object
                    For _r = 0 To _newRows - 1
                        For _c = 0 To maxColumns - 1
                            _trimmedCols(_r, _c) = data(data.GetLowerBound(0) + _r, data.GetLowerBound(1) + _c)
                        Next
                    Next
                    data = _trimmedCols
                    _colTruncated = True
                    System.Diagnostics.Debug.WriteLine($"[ExecuteReadRange] 列数截断: {_actualCols} → {maxColumns}")
                End If

                ' 读取NumberFormat用于日期检测（多种方式尝试，确保鲁棒性）
                Dim formats(,) As Object = Nothing
                Dim singleFormat As String = ""
                Try
                    ' 尝试作为二维数组获取（混合格式选区）
                    formats = DirectCast(rng.NumberFormat, Object(,))
                Catch
                    formats = Nothing
                    Try
                        ' 尝试作为单个字符串获取（统一格式选区）
                        Dim sf = rng.NumberFormat
                        If sf IsNot Nothing Then singleFormat = sf.ToString()
                    Catch
                        singleFormat = ""
                    End Try
                End Try

                ' 日期格式检测辅助函数
                Dim IsDateFormat = Function(nf As String) As Boolean
                                       If String.IsNullOrEmpty(nf) Then Return False
                                       Return nf.IndexOfAny({"y"c, "m"c, "d"c, "Y"c, "M"c, "D"c, "/"c, "-"c, "年"c, "月"c, "日"c}) >= 0
                                   End Function
                Dim IsTimeFormat = Function(nf As String) As Boolean
                                       If String.IsNullOrEmpty(nf) Then Return False
                                       Return nf.IndexOfAny({"h"c, "H"c, "s"c, "S"c}) >= 0
                                   End Function

                ' 辅助函数：判断NumberFormat是否为日期格式，并格式化日期值
                Dim FormatCellValue = Function(val As Object, rowIdx As Integer, colIdx As Integer, colHeader As String) As Object
                    If val Is Nothing OrElse val Is DBNull.Value Then Return Nothing
                    ' Excel 错误值（#DIV/0! / #N/A 等）通过 COM 读取 Value2 会变成内部错误码（负数），
                    ' 需要转成可读的字符串，否则会被当成普通数值导致误判。
                    If Microsoft.VisualBasic.Information.IsError(val) Then
                        Try
                            Dim cell As Microsoft.Office.Interop.Excel.Range = DirectCast(rng.Cells(rowIdx - CInt(rng.Row) + 1, colIdx - CInt(rng.Column) + 1), Microsoft.Office.Interop.Excel.Range)
                            Dim errText As String = CStr(cell.Text)
                            Utils.ComReleaseHelper.Release(cell)
                            Return errText
                        Catch
                            Return "#ERROR"
                        End Try
                    End If
                    ' 检测是否为数字（日期序列号可以是Double、Single、Decimal、Integer等类型）
                    Dim numVal As Double
                    Dim isNumericVal As Boolean = False
                    If TypeOf val Is Double OrElse TypeOf val Is Single OrElse TypeOf val Is Decimal OrElse TypeOf val Is Integer Then
                        numVal = Convert.ToDouble(val)
                        isNumericVal = True
                    Else
                        Dim dblResult As Double
                        If IsNumeric(val) AndAlso Double.TryParse(val.ToString(), dblResult) Then
                            numVal = dblResult
                            isNumericVal = True
                        End If
                    End If

                    If Not isNumericVal Then Return val
                    ' Excel日期范围：1900-01-01 = 1, 2100-01-01 ≈ 73050
                    If numVal < 1 OrElse numVal > 100000 Then Return val

                    ' 获取该单元格的NumberFormat（多种方式）
                    Dim nf As String = ""
                    If formats IsNot Nothing Then
                        Try
                            Dim nfObj As Object = formats(rowIdx, colIdx)
                            nf = If(nfObj IsNot Nothing, nfObj.ToString(), "")
                        Catch
                            nf = ""
                        End Try
                    End If
                    If nf = "" AndAlso singleFormat <> "" Then nf = singleFormat
                    ' 如果批量获取失败，尝试逐单元格获取
                    If nf = "" Then
                        Try
                            Dim cell As Microsoft.Office.Interop.Excel.Range = DirectCast(rng.Cells(rowIdx - CInt(rng.Row) + 1, colIdx - CInt(rng.Column) + 1), Microsoft.Office.Interop.Excel.Range)
                            Dim cn = cell.NumberFormat
                            If cn IsNot Nothing Then nf = cn.ToString()
                        Catch
                            nf = ""
                        End Try
                    End If

                    ' 策略1：NumberFormat明确包含日期符号 → 转换为日期
                    Dim hasDateFormat As Boolean = IsDateFormat(nf)
                    ' 策略2：列标题包含日期/时间关键词，且值在 Excel 日期全范围 → 也转换为日期
                    '   不再用 30000-70000 人为限制（那是 1990-2090 子集），
                    '   统一用 1-100000（覆盖 1900-2173），避免边界日期漏掉
                    Dim hasDateHeader As Boolean = False
                    If Not hasDateFormat AndAlso colHeader IsNot Nothing Then
                        Dim hl = colHeader.ToLower()
                        If hl.Contains("日期") OrElse hl.Contains("时间") OrElse hl.Contains("date") OrElse hl.Contains("time") Then
                            If numVal >= 1 AndAlso numVal <= 100000 Then
                                hasDateHeader = True
                            End If
                        End If
                    End If

                    If hasDateFormat OrElse hasDateHeader Then
                        Try
                            Dim dt As DateTime = DateTime.FromOADate(numVal)
                            If IsTimeFormat(nf) Then
                                Return dt.ToString("yyyy-MM-dd HH:mm:ss")
                            Else
                                Return dt.ToString("yyyy-MM-dd")
                            End If
                        Catch
                            ' 转换失败，返回原值
                        End Try
                    End If
                    Return val
                End Function

                ' 转换为 JSON
                Dim rows = New List(Of Object)()
                Dim startRow As Integer = If(hasHeaders, 1, 0)
                Dim headers As String() = Nothing
                Dim useObjectFormat As Boolean = (returnFormat = "object")

                If hasHeaders OrElse useObjectFormat Then
                    headers = New String(data.GetUpperBound(1) - data.GetLowerBound(1)) {}
                    For col = data.GetLowerBound(1) To data.GetUpperBound(1)
                        Dim val = FormatCellValue(data(data.GetLowerBound(0), col), data.GetLowerBound(0), col, "")
                        headers(col - data.GetLowerBound(1)) = If(val IsNot Nothing, val.ToString(), "")
                    Next
                End If

                Dim rowCount As Integer = 0
                For row = data.GetLowerBound(0) + startRow To data.GetUpperBound(0)
                    ' 限制最大读取行数
                    If rowCount >= maxRows Then Exit For

                    If hasHeaders OrElse useObjectFormat Then
                        Dim rowDict = New Dictionary(Of String, Object)()
                        Dim isEmptyRow As Boolean = skipEmptyRows
                        For col = data.GetLowerBound(1) To data.GetUpperBound(1)
                            Dim key = headers(col - data.GetLowerBound(1))
                            Dim val = FormatCellValue(data(row, col), row, col, key)
                            rowDict(key) = val
                            If skipEmptyRows AndAlso isEmptyRow AndAlso val IsNot Nothing AndAlso Not (TypeOf val Is String AndAlso DirectCast(val, String).Trim() = "") Then
                                isEmptyRow = False
                            End If
                        Next
                        If Not isEmptyRow Then rows.Add(rowDict)
                        If isEmptyRow Then rowCount -= 1
                    Else
                        Dim rowList = New List(Of Object)()
                        Dim isEmptyRow As Boolean = skipEmptyRows
                        For col = data.GetLowerBound(1) To data.GetUpperBound(1)
                            Dim val = FormatCellValue(data(row, col), row, col, "")
                            rowList.Add(val)
                            If skipEmptyRows AndAlso isEmptyRow AndAlso val IsNot Nothing AndAlso Not (TypeOf val Is String AndAlso DirectCast(val, String).Trim() = "") Then
                                isEmptyRow = False
                            End If
                        Next
                        If Not isEmptyRow Then rows.Add(rowList)
                        If isEmptyRow Then rowCount -= 1
                    End If
                    rowCount += 1
                Next

                Dim truncated As Boolean = (data.GetUpperBound(0) - data.GetLowerBound(0) - startRow) > maxRows OrElse _colTruncated

                ' 返回结构化上下文，同时保留原有 result 数组兼容旧调用方
                Dim workbookName As String = ""
                Try
                    If excelApp.ActiveWorkbook IsNot Nothing Then workbookName = excelApp.ActiveWorkbook.Name
                Catch
                    workbookName = ""
                End Try

                Dim worksheetName As String = ""
                Try
                    If ws IsNot Nothing Then worksheetName = ws.Name
                Catch
                    worksheetName = ""
                End Try

                Dim actualHeaders As String() = If(headers IsNot Nothing, headers, New String(data.GetUpperBound(1) - data.GetLowerBound(1)) {})
                Dim previewRows = rows.Take(5).ToList()

                ' 获取range地址（A1格式，用于RangeContext和返回结果）
                Dim rngAddress As String = If(rng IsNot Nothing, rng.Address(False, False), rangeAddr)

                ' 更新选区上下文：记录本次读取的选区，供后续写工具无参数调用时自动复用
                If rng IsNot Nothing AndAlso Not String.IsNullOrEmpty(worksheetName) Then
                    RangeContext.Update(worksheetName, rngAddress)
                End If

                Dim output = New With {
                    .workbook = workbookName,
                    .worksheet = worksheetName,
                    .range = If(rng IsNot Nothing, rng.Address, rangeAddr),
                    .rows = rows.Count,
                    .columns = actualHeaders.Length,
                    .headers = actualHeaders,
                    .previewRows = previewRows,
                    .truncated = truncated,
                    .result = rows
                }

                Dim jsonResult = JsonConvert.SerializeObject(output, Formatting.Indented)
                If truncated Then
                    Dim notes As New List(Of String)()
                    If (data.GetUpperBound(0) - data.GetLowerBound(0) - startRow) > maxRows Then
                        notes.Add($"行数超过 maxRows={maxRows}")
                    End If
                    If _colTruncated Then
                        notes.Add($"列数超过 maxColumns={maxColumns}")
                    End If
                    jsonResult &= $"{Environment.NewLine}/* ⚠️ 数据已截断：{String.Join("；", notes)}。请缩小 range 范围后重新读取 */"
                End If
                result.Content.Add(New McpTextContent With {
                    .Text = jsonResult
                })

                Utils.ComReleaseHelper.Release(rng)
                If wsNeedsRelease AndAlso ws IsNot Nothing Then Utils.ComReleaseHelper.Release(ws)

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"读取范围失败: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteWriteRange(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim rangeAddr As String = GetArg(Of String)(args, "range", "A1")
                ' 归一化：LLM 常把 read_range 的完整区域（如 "$A$1:$I$101"）抄进 range，
                ' 写入语义是"以该区域左上角为起点"，故取冒号前起始格并去掉 '$'。
                If Not String.IsNullOrEmpty(rangeAddr) AndAlso rangeAddr.Contains(":"c) Then
                    rangeAddr = rangeAddr.Split(":"c)(0)
                End If
                rangeAddr = If(rangeAddr, "A1").Replace("$", "").Trim()
                If String.IsNullOrEmpty(rangeAddr) Then rangeAddr = "A1"
                Dim valuesJson As String = GetArg(Of String)(args, "values", "[]")
                Dim clearFirst As Boolean = GetArg(Of Boolean)(args, "clearFirst", False)
                Dim writeMode As String = GetArg(Of String)(args, "writeMode", "overwrite")

                Dim values As List(Of List(Of Object)) = Nothing
                Try
                    values = JsonConvert.DeserializeObject(Of List(Of List(Of Object)))(valuesJson)
                Catch parseEx As JsonException
                    result.IsError = True
                    Dim preview As String = If(valuesJson IsNot Nothing AndAlso valuesJson.Length > 80, valuesJson.Substring(0, 80) & "...", valuesJson)
                    result.Content.Add(New McpTextContent With {.Text = $"values 不是合法的二维数组 JSON（如 [[""名称"",123],[""名称2"",456]]）。收到: {preview}"})
                    Return result
                End Try
                If values Is Nothing OrElse values.Count = 0 Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "values 为空：请传二维数组 JSON（如 [[""名称"",123],[""名称2"",456]]）。看板场景只写 KPI 小数据块（≤10 行），禁止把原始数据整表塞进 values。"})
                    Return result
                End If

                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Try
                    If Not String.IsNullOrEmpty(sheetName) Then
                        ' 注意：Worksheets(sheetName) 对尚不存在的新表名会抛 DISP_E_BADINDEX(0x8002000B)，
                        ' TryCast 不捕获该 COM 异常。写入目标表可能由 LLM 规划为"新表"，故找不到时须自动创建，
                        ' 否则会报"写入数据失败: 无效索引"。
                        Dim found As Boolean = False
                        Try
                            ws = DirectCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                            found = True
                        Catch ex As Exception
                            found = False
                        End Try
                        If Not found Then
                            ws = DirectCast(excelApp.ActiveWorkbook.Worksheets.Add(After:=
                                excelApp.ActiveWorkbook.Worksheets(excelApp.ActiveWorkbook.Worksheets.Count)),
                                Microsoft.Office.Interop.Excel.Worksheet)
                            ws.Name = sheetName
                        End If
                    End If
                    If ws Is Nothing Then
                        ws = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    End If

                    ' 计算目标范围大小
                    Dim rows = values.Count
                    Dim cols = values(0).Count

                    ' 追加模式：计算实际写入位置
                    Dim actualStartAddr = rangeAddr
                    If writeMode = "append" Then
                        Dim usedRange = ws.UsedRange
                        If usedRange IsNot Nothing Then
                            Dim lastRow As Integer = usedRange.Row + usedRange.Rows.Count - 1
                            Dim startRow As Integer = lastRow + 1
                            Dim colPart = New String(rangeAddr.Where(Function(c) Char.IsLetter(c)).ToArray())
                            actualStartAddr = colPart & startRow.ToString()
                            Utils.ComReleaseHelper.Release(usedRange)
                        End If
                    End If

                    Dim endCell = GetEndCell(actualStartAddr, rows, cols)
                    Dim fullRange = $"{actualStartAddr}:{endCell}"

                    ' 写入前清空目标范围：只清内容、保留格式/边框/批注。
                    ' 此前用 Range.Clear() 会把清除范围内的 NumberFormat 一并抹掉，
                    ' 导致每次写回（如去重删除）把用户设好的日期等格式清成 General，日期列变回序列号或乱格式。
                    If clearFirst Then
                        Dim clearRange = ws.Range(fullRange)
                        clearRange.ClearContents()
                        Utils.ComReleaseHelper.Release(clearRange)
                    End If

                    ' 转换为二维数组写入
                    Dim dataArr(rows - 1, cols - 1) As Object
                    For r = 0 To rows - 1
                        For c = 0 To Math.Min(cols - 1, values(r).Count - 1)
                            Dim val = values(r)(c)
                            ' ISO 日期字符串（"2026-06-16" / "2026-06-16T00:00:00"，由清洗 normalizeDates 产生的 DateTime 序列化而来）
                            ' → 转成真实 DateTime 写入，避免被 Excel 当文本"2026-06-16T00:00:00"落盘，保证日期列是真正的日期值。
                            If TypeOf val Is String AndAlso val.ToString().Length >= 10 AndAlso
                               System.Text.RegularExpressions.Regex.IsMatch(CStr(val), "^[0-9]{4}-[0-9]{1,2}-[0-9]{1,2}([ T].*)?$") Then
                                Dim dt As DateTime
                                If DateTime.TryParse(CStr(val), Globalization.CultureInfo.InvariantCulture,
                                                     Globalization.DateTimeStyles.None, dt) Then
                                    dataArr(r, c) = dt
                                    Continue For
                                End If
                            End If
                            dataArr(r, c) = val
                        Next
                    Next

                    Dim targetRange = ws.Range(fullRange)
                    targetRange.Value2 = dataArr

                    ' 写入后将覆盖该区域的背景形状重新置于底层，防止白色背景形状遮挡单元格数据
                    Try
                        Dim shpCol = ws.Shapes
                        Dim shpCount As Integer = shpCol.Count
                        For si As Integer = shpCount To 1 Step -1
                            Dim shp = shpCol.Item(si)
                            Try
                                Dim tlR As Integer = shp.TopLeftCell.Row
                                Dim tlC As Integer = shp.TopLeftCell.Column
                                Dim brR As Integer = shp.BottomRightCell.Row
                                Dim brC As Integer = shp.BottomRightCell.Column
                                Dim tR As Integer = targetRange.Row
                                Dim tC As Integer = targetRange.Column
                                Dim tBR As Integer = tR + targetRange.Rows.Count - 1
                                Dim tBC As Integer = tC + targetRange.Columns.Count - 1
                                If tlR <= tBR AndAlso brR >= tR AndAlso tlC <= tBC AndAlso brC >= tC Then
                                    shp.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack)
                                End If
                            Catch : End Try
                            Utils.ComReleaseHelper.Release(shp)
                        Next
                        Utils.ComReleaseHelper.Release(shpCol)
                    Catch : End Try
                    Utils.ComReleaseHelper.Release(targetRange)

                    result.Content.Add(New McpTextContent With {
                        .Text = $"成功写入 {rows} 行 x {cols} 列数据到工作表 [{ws.Name}] 的 {fullRange}（模式：{writeMode}）"
                    })

                Catch ex As Exception
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"写入数据失败: {ex.Message}"})
                Finally
                    Utils.ComReleaseHelper.Release(ws)
                End Try

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"写入数据失败: {ex.Message}"})
            End Try
            Return result
        End Function
        ''' <summary>
        ''' 
        ''' 避免 VB.NET 因项目级 Imports Microsoft.Office.Tools.Excel 将 Worksheet/Range 解析为 VSTO 宿主接口
        ''' 而导致 COM QueryInterface 失败 (E_NOINTERFACE 0x80004002)。
        ''' 所有返回值和参数类型全部使用 FULL QUALIFIED 的 Interop 命名空间（无一处未限定名）。
        ''' </summary>
        Private Shared Function ResolveRangeSafe(app As Microsoft.Office.Interop.Excel.Application,
                                                 fullAddr As String,
                                                 sheetNameExplicit As String) As Microsoft.Office.Interop.Excel.Range
            If app Is Nothing OrElse String.IsNullOrEmpty(fullAddr) Then Return Nothing

            ' 无效前缀黑名单（前端未定义变量字符串化污染）
            Dim invalidNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
                "undefined", "null", "[object Object]", "none", "nan", ""
            }

            Dim prefix As String = ""
            Dim localAddr As String = fullAddr
            Dim exclIdx = fullAddr.LastIndexOf("!"c)
            If exclIdx > 0 Then
                prefix = fullAddr.Substring(0, exclIdx).Trim()
                localAddr = fullAddr.Substring(exclIdx + 1).Trim()
                If prefix.Length >= 2 AndAlso prefix.StartsWith("'") AndAlso prefix.EndsWith("'") Then
                    prefix = prefix.Substring(1, prefix.Length - 2)
                End If
                If invalidNames.Contains(prefix) Then prefix = ""
            End If

            ' 解析工作表对象（优先级：显式参数 > dataRange内嵌 > ActiveSheet）
            Dim finalSheetName As String = ""
            If Not String.IsNullOrEmpty(sheetNameExplicit) AndAlso Not invalidNames.Contains(sheetNameExplicit) Then
                finalSheetName = sheetNameExplicit
            ElseIf Not String.IsNullOrEmpty(prefix) Then
                finalSheetName = prefix
            End If

            Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
            Dim wsObj As Object = Nothing
            Dim wsInterop As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Try
                wb = app.ActiveWorkbook
            Catch : wb = Nothing : End Try
            If wb IsNot Nothing AndAlso Not String.IsNullOrEmpty(finalSheetName) Then
                Try : wsObj = wb.Worksheets(finalSheetName) : Catch : wsObj = Nothing : End Try
            End If
            If wsObj Is Nothing Then
                Try : wsObj = app.ActiveSheet : Catch : wsObj = Nothing : End Try
            End If
            wsInterop = TryCast(wsObj, Microsoft.Office.Interop.Excel.Worksheet)

            If wsInterop Is Nothing OrElse String.IsNullOrEmpty(localAddr) Then Return Nothing

            ' 【关键】只使用 Worksheets._Default[localAddr] 不调用任何命名空间歧义的属性/方法
            Dim rngObj As Object = Nothing
            Try
                rngObj = wsInterop.Range(localAddr)
            Catch
                rngObj = Nothing
            End Try
            Return TryCast(rngObj, Microsoft.Office.Interop.Excel.Range)
        End Function
        Private Function ExecuteGetWorkbookInfo(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "Excel 应用不可用"})
                    Return result
                End If

                ' 支持按路径打开指定工作簿
                Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
                Dim openedByPath = False
                Dim filePath = GetArg(Of String)(args, "filePath", "")
                
                If Not String.IsNullOrEmpty(filePath) AndAlso IO.File.Exists(filePath) Then
                    Try
                        wb = excelApp.Workbooks.Open(filePath)
                        openedByPath = True
                    Catch ex As Exception
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = $"无法打开文件: {filePath} - {ex.Message}"})
                        Return result
                    End Try
                Else
                    wb = excelApp.ActiveWorkbook
                End If
                
                If wb Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                ' 只统计可见工作表（与 Excel UI 一致）
                Dim visibleCount = 0
                Dim totalSheets = wb.Worksheets.Count
                For i As Integer = 1 To totalSheets
                    Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                    Try
                        ws = DirectCast(wb.Worksheets(i), Microsoft.Office.Interop.Excel.Worksheet)
                        If ws.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible Then visibleCount += 1
                    Finally
                        Utils.ComReleaseHelper.Release(ws)
                    End Try
                Next
                Dim info = New Dictionary(Of String, Object) From {
                    {"name", wb.Name},
                    {"fullName", wb.FullName},
                    {"path", wb.Path},
                    {"sheetCount", visibleCount},
                    {"saved", wb.Saved}
                }
                
                ' 如果是按路径打开的，读取完后关闭（不保存）
                If openedByPath Then
                    wb.Close(SaveChanges:=False)
                    Utils.ComReleaseHelper.Release(wb)
                End If

                result.Content.Add(New McpTextContent With {
                    .Text = JsonConvert.SerializeObject(info, Formatting.Indented)
                })

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"获取工作簿信息失败: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteGetOpenWorkbooks(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                If _bridge Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "Excel 桥接不可用"})
                    Return result
                End If

                Dim json = DirectCast(_bridge, ExcelAIVSTOAddIn.ExcelBridge).GetOpenWorkbooks()
                result.Content.Add(New McpTextContent With {.Text = json})
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"获取打开的工作簿列表失败: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteFillData(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim dataStr As String = GetArg(Of String)(args, "data", "[]")
                Dim dataType As String = GetArg(Of String)(args, "dataType", "general")

                Dim fillResult = DirectCast(_bridge, ExcelAIVSTOAddIn.ExcelBridge).FillDataToExcel(dataStr, dataType)
                result.Content.Add(New McpTextContent With {.Text = fillResult})

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"填入数据失败: {ex.Message}"})
            End Try
            Return result
        End Function

        '#End Region

        ''' <summary>
        ''' excel_check_errors - 扫描指定/活动工作表中的公式错误单元格（#DIV/0!、#N/A、#VALUE! 等）。
        ''' 用 COM SpecialCells(xlCellTypeFormulas, xlErrors) 一次性定位，不逐格遍历，不走 skill 参数装配链。
        ''' </summary>
        Private Function ExecuteCheckErrors(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Dim wsToRelease As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim rangeToRelease As New List(Of Microsoft.Office.Interop.Excel.Range)()
            Try
                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "Excel 应用程序未就绪"})
                    Return result
                End If
                If excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim rangeAddr As String = GetArg(Of String)(args, "range", "")

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                If Not String.IsNullOrEmpty(sheetName) Then
                    Try
                        ws = TryCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                    Catch ex As Exception
                        ' 工作表不存在 → 回退活动工作表（避免 0x800A03EB 无效索引）
                        System.Diagnostics.Debug.WriteLine($"[ExecuteCheckErrors] 工作表 '{sheetName}' 不存在: {ex.Message}，回退活动工作表")
                        ws = Nothing
                    End Try
                    If ws IsNot Nothing Then wsToRelease = ws
                End If
                If ws Is Nothing Then
                    ws = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                End If

                Dim scanRange As Microsoft.Office.Interop.Excel.Range
                If String.IsNullOrEmpty(rangeAddr) Then
                    scanRange = ws.UsedRange
                Else
                    scanRange = ws.Range(rangeAddr)
                End If
                rangeToRelease.Add(scanRange)

                ' SpecialCells 在无错误单元格时抛 COMException(0x800A03EB)，按"无错误"处理
                Dim errRange As Microsoft.Office.Interop.Excel.Range = Nothing
                Try
                    errRange = scanRange.SpecialCells(
                        Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeFormulas,
                        Microsoft.Office.Interop.Excel.XlSpecialCellsValue.xlErrors)
                Catch
                    errRange = Nothing
                End Try

                Dim errorCells As New List(Of Dictionary(Of String, Object))()
                If errRange IsNot Nothing Then
                    rangeToRelease.Add(errRange)
                    For Each cell As Microsoft.Office.Interop.Excel.Range In errRange.Cells
                        Try
                            Dim info As New Dictionary(Of String, Object)()
                            info("cell") = CStr(cell.Address(False, False))
                            info("error") = CStr(cell.Text)
                            Try
                                info("formula") = CStr(cell.Formula)
                            Catch
                            End Try
                            errorCells.Add(info)
                        Catch
                        Finally
                            Utils.ComReleaseHelper.Release(cell)
                        End Try
                    Next
                End If

                Dim scannedCells As Double
                Try
                    scannedCells = CDbl(scanRange.CountLarge)
                Catch
                    scannedCells = CDbl(CLng(scanRange.Cells.Count))
                End Try

                Dim payload = New With {
                    .success = True,
                    .sheet = CStr(ws.Name),
                    .scannedRange = CStr(scanRange.Address(False, False)),
                    .scannedCells = scannedCells,
                    .errorCount = errorCells.Count,
                    .hasErrors = errorCells.Count > 0,
                    .errors = errorCells
                }
                result.Content.Add(New McpTextContent With {.Text = JsonConvert.SerializeObject(payload)})
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"检查错误失败: {ex.Message}"})
            Finally
                For Each r In rangeToRelease
                    Utils.ComReleaseHelper.Release(r)
                Next
                If wsToRelease IsNot Nothing Then Utils.ComReleaseHelper.Release(wsToRelease)
            End Try
            Return result
        End Function

    End Class

End Namespace

