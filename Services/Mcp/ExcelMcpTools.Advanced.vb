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

        '#Region "工具实现 - 高级与清洗"

        ''' <summary>
        ''' BGR 组合淡黄色高亮（RGB 255,242,180）——Interior.Color 用 BGR 整型。
        ''' </summary>
        Private Shared Function ValidRgbColor() As Integer
            Return 180 + 242 * 256 + 255 * 65536   ' RGB(255,242,180) →
        End Function

        Private Shared Function CleanCellIsBlank(v As Object) As Boolean
            If v Is Nothing OrElse IsDBNull(v) Then Return True
            Return CStr(v).Trim() = ""
        End Function

        ''' <summary>
        ''' 去重比较键。日期按 yyyy-MM-dd，金额按数值（¥9,163.00 与 9163 视为相同），文本去首尾空白。
        ''' </summary>
        Private Shared Function CleanCellKey(v As Object) As String
            If v Is Nothing OrElse IsDBNull(v) Then Return ""
            If TypeOf v Is DateTime Then
                Dim dt = DirectCast(v, DateTime)
                If dt.TimeOfDay = TimeSpan.Zero Then Return dt.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture)
                Return dt.ToString("yyyy-MM-dd HH:mm:ss", Globalization.CultureInfo.InvariantCulture)
            End If
            Dim num As Double
            If TypeOf v Is Double OrElse TypeOf v Is Single OrElse TypeOf v Is Decimal OrElse
               TypeOf v Is Integer OrElse TypeOf v Is Long OrElse TypeOf v Is Short Then
                num = Convert.ToDouble(v, Globalization.CultureInfo.InvariantCulture)
                Return num.ToString("0.######", Globalization.CultureInfo.InvariantCulture)
            End If
            Dim s = CStr(v).Trim()
            Dim stripped = s.Replace("¥", "").Replace("￥", "").Replace(","c, "").Replace("，", "").Replace(" ", "")
            If Double.TryParse(stripped, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, num) Then
                Return num.ToString("0.######", Globalization.CultureInfo.InvariantCulture)
            End If
            Return s
        End Function

        ''' <summary>
        ''' excel_clean_data — 智能数据清洗（原生 Excel COM，替代易报错 VBA）
        ''' 支持操作：trimSpaces / fillEmpty / normalizeDates / normalizeNumbers / standardizeText / removeEmptyRows
        ''' 范围回退链：显式 range → 上下文选区 → Selection → UsedRange（与 excel_remove_duplicates 一致）
        ''' </summary>
        Private Function ExecuteCleanData(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim rangeAddr As String = GetArg(Of String)(args, "range", "")
                Dim hasHeaderArg = GetArg(Of String)(args, "hasHeaderRow", "")
                Dim fillMethod As String = GetArg(Of String)(args, "fillMethod", "ffill").Trim().ToLowerInvariant()
                Dim fillValue As String = GetArg(Of String)(args, "fillValue", "待补充")
                Dim operationsArg As Object = Nothing
                args?.TryGetValue("operations", operationsArg)
                ' 去重依据列（原始内容，解析成 1 基列号）
                Dim dedupeColsRaw As Object = Nothing
                args?.TryGetValue("columns", dedupeColsRaw)

                ' ---------- 0) 解析操作列表 ----------
                Dim ops As New List(Of String)()
                If TypeOf operationsArg Is String Then
                    Dim s = CStr(operationsArg).Trim()
                    If s.StartsWith("[") AndAlso s.EndsWith("]") Then
                        Try
                            Dim parsed = Newtonsoft.Json.JsonConvert.DeserializeObject(Of List(Of Object))(s)
                            For Each o In parsed
                                If o IsNot Nothing Then ops.Add(o.ToString().Trim())
                            Next
                        Catch
                            For Each p In s.Split({","c, ";"c, "|"c, " "c, "，"c, "；"c}, StringSplitOptions.RemoveEmptyEntries)
                                ops.Add(p.Trim())
                            Next
                        End Try
                    Else
                        For Each p In s.Split({","c, ";"c, "|"c, " "c, "，"c, "；"c}, StringSplitOptions.RemoveEmptyEntries)
                            ops.Add(p.Trim())
                        Next
                    End If
                ElseIf TypeOf operationsArg Is System.Collections.IEnumerable Then
                    For Each o In DirectCast(operationsArg, System.Collections.IEnumerable)
                        If o IsNot Nothing Then ops.Add(o.ToString().Trim())
                    Next
                ElseIf operationsArg IsNot Nothing Then
                    ops.Add(operationsArg.ToString().Trim())
                End If

                ' 【默认兜底】"清洗这份数据"这类泛化诉求经 dispatcher/LLM 规划后可能不带 operations，
                ' 直接报错会让整轮失败（实测 ✗失败）。改默认执行两个无破坏性的基础清洗：
                ' 去首尾空格 + 删除整行空白；破坏性操作（fillEmpty 填充/normalizeDates 改写）仍需显式指定。
                Dim defaultOpsNote As String = ""
                If ops.Count = 0 Then
                    ops.Add("trimSpaces")
                    ops.Add("removeDuplicates")
                    ops.Add("removeEmptyRows")
                    defaultOpsNote = "（未指定 operations，已默认执行基础清洗：去首尾空格、删除整行重复、删除空行；如需填充空值、日期/数字规范化等请在指令中说明）"
                End If

                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim rng As Microsoft.Office.Interop.Excel.Range = Nothing
                ' 表名不存在而回退活动工作表时的说明（追加到结果里，避免"静默换表"）
                Dim sheetFallbackNote As String = ""

                Try
                    ' ---------- 1) 工作表解析 ----------
                    If Not String.IsNullOrEmpty(sheetName) Then
                        ' 【必须吞异常】Worksheets(不存在的表名) 抛 DISP_E_BADINDEX(0x8002000B)"无效索引"，
                        ' 若不捕获会直接冒到下面 catch，整个清洗失败并只回报一句"无效索引"
                        ' （实测："清洗这份数据"时规划器猜了个不存在的表名）。这里回退活动工作表并记录说明。
                        Try
                            ws = TryCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                        Catch ex As Exception
                            ws = Nothing
                        End Try
                        If ws Is Nothing Then
                            sheetFallbackNote = $"（工作表『{sheetName}』不存在，已改用当前活动工作表）"
                        End If
                    End If
                    If ws Is Nothing Then ws = TryCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    If ws Is Nothing Then
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = "无法定位有效工作表，请先激活包含数据的工作表。"})
                        Return result
                    End If

                    ' ---------- 2) 范围解析：显式 range → 上下文选区 → Selection → UsedRange ----------
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
                    If rng Is Nothing OrElse rng.Rows.Count < 1 Then
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = "未找到可清洗的数据区域。"})
                        Return result
                    End If

                    Dim totalRows As Integer = rng.Rows.Count
                    Dim totalCols As Integer = rng.Columns.Count
                    Dim addrDisplay As String = rng.Address(False, False)

                    ' ---------- 3) hasHeaderRow ----------
                    Dim hasHeader As Boolean = True
                    If String.Equals(hasHeaderArg, "false", StringComparison.OrdinalIgnoreCase) OrElse
                       String.Equals(hasHeaderArg, "0", StringComparison.OrdinalIgnoreCase) OrElse
                       String.Equals(hasHeaderArg, "no", StringComparison.OrdinalIgnoreCase) Then
                        hasHeader = False
                    End If
                    Dim dataStartRow As Integer = If(hasHeader, 2, 1)

                    If dataStartRow > totalRows Then
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = "数据区域只有表头，没有可清洗的数据行。"})
                        Return result
                    End If

                    ' ---------- 4) 读取值到二维数组（1基下标） ----------
                    ' 用 rng.Value 而非 Value2：Value 会把真正的日期单元格转成 DateTime 类型，
                    ' 便于 normalizeNumbers 能识别"纯整数日期序列号"列（Value2 下日期是 Double，无法与整数 ID 区分）。
                    Dim values As Object(,) = Nothing
                    If totalRows = 1 AndAlso totalCols = 1 Then
                        values = New Object(0, 0) {{rng.Value}}
                    Else
                        values = TryCast(rng.Value, Object(,))
                    End If
                    If values Is Nothing Then
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = "读取数据失败。"})
                        Return result
                    End If

                    ' ---------- 5) 内存中应用操作 ----------
                    Dim rowsToDelete As New List(Of Integer)()   ' 1基行号（相对 rng 顶部）
                    Dim dateCols As New HashSet(Of Integer)()    ' 1基列号（相对 rng 左列）
                    Dim numCols As New HashSet(Of Integer)()
                    ' 各操作的"实际改动量"统计（供前端 buildCleanOpSummary 如实展示，杜绝"已统一但无明细/未改"）
                    Dim trimmedCells As Integer = 0
                    Dim normalizedDates As Integer = 0
                    Dim normalizedNumbers As Integer = 0
                    Dim removedEmptyRows As Integer = 0
                    Dim removedDuplicateRows As Integer = 0
                    Dim duplicateAbsRows As New List(Of Integer)()
                    Dim resultObjWrittenBack As Boolean = False

                    For Each op In ops
                        Select Case op.ToLowerInvariant()
                            Case "trimspaces"
                                For r = dataStartRow To values.GetLength(0)
                                    For c = 1 To values.GetLength(1)
                                        Dim v = values(r, c)
                                        If TypeOf v Is String Then
                                            ' 只清默认空白会漏中文全角空格 U+3000、NBSP U+00A0、em-space、零宽等，
                                            ' 中文/网页数据里这些"空格"扫出来是"未发现多余空格"但仍脏。故用扩展空白集：
                                            '  首尾清 半角/制表/换行/回车/全角/NBSP/en-em-space/细空格/零宽/BOM
                                            '  并把中间 2+ 个连续空格(半角/全角/制表)压缩为单个半角空格。
                                            Dim tv = CStr(v).Trim(" "c, ChrW(&H9), ChrW(&HA), ChrW(&HD), ChrW(&H3000), ChrW(&HA0),
                                                                      ChrW(&H2002), ChrW(&H2003), ChrW(&H2009), ChrW(&H200B), ChrW(&HFEFF))
                                            tv = System.Text.RegularExpressions.Regex.Replace(tv, "[ \u3000\u00A0\t]{2,}", " ")
                                            If tv <> CStr(v) Then trimmedCells += 1
                                            values(r, c) = tv
                                        End If
                                    Next
                                Next

                            Case "standardizetext"
                                For r = dataStartRow To values.GetLength(0)
                                    For c = 1 To values.GetLength(1)
                                        Dim v = values(r, c)
                                        If v IsNot Nothing AndAlso Not IsDBNull(v) AndAlso CStr(v).Trim() <> "" Then
                                            values(r, c) = CStr(v).Trim()
                                        End If
                                    Next
                                Next

                            Case "fillempty"
                                If fillMethod = "delete" Then
                                    For r = dataStartRow To values.GetLength(0)
                                        Dim hasEmpty As Boolean = False
                                        For c = 1 To values.GetLength(1)
                                            Dim v = values(r, c)
                                            If v Is Nothing OrElse IsDBNull(v) OrElse CStr(v).Trim() = "" Then
                                                hasEmpty = True
                                                Exit For
                                            End If
                                        Next
                                        If hasEmpty AndAlso Not rowsToDelete.Contains(r) Then rowsToDelete.Add(r)
                                    Next
                                ElseIf fillMethod = "zero" Then
                                    For r = dataStartRow To values.GetLength(0)
                                        For c = 1 To values.GetLength(1)
                                            Dim v = values(r, c)
                                            If v Is Nothing OrElse IsDBNull(v) OrElse CStr(v).Trim() = "" Then values(r, c) = 0
                                        Next
                                    Next
                                ElseIf fillMethod = "placeholder" Then
                                    For r = dataStartRow To values.GetLength(0)
                                        For c = 1 To values.GetLength(1)
                                            Dim v = values(r, c)
                                            If v Is Nothing OrElse IsDBNull(v) OrElse CStr(v).Trim() = "" Then values(r, c) = fillValue
                                        Next
                                    Next
                                Else ' ffill 前值填充
                                    Dim lastVal(totalCols) As Object
                                    For r = dataStartRow To values.GetLength(0)
                                        For c = 1 To values.GetLength(1)
                                            Dim v = values(r, c)
                                            If v Is Nothing OrElse IsDBNull(v) OrElse CStr(v).Trim() = "" Then
                                                If lastVal(c) IsNot Nothing Then
                                                    values(r, c) = lastVal(c)
                                                Else
                                                    values(r, c) = fillValue
                                                End If
                                            Else
                                                lastVal(c) = v
                                            End If
                                        Next
                                    Next
                                End If

                            Case "normalizedates"
                                ' 旧逻辑逐单元格 TryParse，
                                '   任一含 ./- 的数值（"1234.56""1.5""2026.55"）被宽松解析成日期即把整列加入
                                '   dateCols，随后整列 NumberFormat 被设为 yyyy-mm-dd，导致整列数字显示成日期。
                                '   修复：① 列级整体判定——仅整列被确认为日期列才转换+设格式；
                                '        ② 严格日期形态——4位年份开头分段 / 含"年"，且 TryParseExact 显式格式，
                                '           彻底排除纯数字、小数、比例文本。
                                Dim colHasDtm As New HashSet(Of Integer)()
                                Dim colFmtDtm As New HashSet(Of Integer)()
                                Dim colNonEmpty As New Dictionary(Of Integer, Integer)()
                                Dim colDateOk As New Dictionary(Of Integer, Integer)()
                                Dim dateFormats() As String = {"yyyy-MM-dd", "yyyy-M-d", "yyyy/MM/dd", "yyyy/M/d",
                                    "yyyy.MM.dd", "yyyy.M.d", "yyyy年MM月dd日", "yyyy年M月d日", "yyyy年MM月dd日 HH:mm:ss"}
                                Dim IsStrictDateLike = Function(ss As String) As Boolean
                                    Return ss.IndexOf("年"c) >= 0 OrElse
                                        System.Text.RegularExpressions.Regex.IsMatch(ss, "^[0-9]{4}[./-][0-9]{1,2}([./-][0-9]{1,2})?$")
                                End Function
                                For c = 1 To values.GetLength(1)
                                    Dim cfmt As String = ""
                                    Try : cfmt = CStr(DirectCast(rng.Columns(c), Microsoft.Office.Interop.Excel.Range).NumberFormat) : Catch : End Try
                                    If cfmt <> "" AndAlso IsDateTimeFormat(cfmt) Then colFmtDtm.Add(c)
                                Next
                                ' 第一遍：收集每列证据（DateTime 类型 / 严格日期形态字符串）
                                For r = dataStartRow To values.GetLength(0)
                                    For c = 1 To values.GetLength(1)
                                        Dim v = values(r, c)
                                        If v Is Nothing OrElse IsDBNull(v) Then Continue For
                                        If TypeOf v Is DateTime Then
                                            colHasDtm.Add(c)
                                            Continue For
                                        End If
                                        Dim isNumeric = TypeOf v Is Double OrElse TypeOf v Is Single OrElse TypeOf v Is Decimal OrElse
                                            TypeOf v Is SByte OrElse TypeOf v Is Byte OrElse TypeOf v Is Short OrElse TypeOf v Is UShort OrElse
                                            TypeOf v Is Integer OrElse TypeOf v Is UInteger OrElse TypeOf v Is Long OrElse TypeOf v Is ULong
                                        If isNumeric Then Continue For   ' 数值单元格：不参与日期形态统计，绝不转换
                                        Dim s = CStr(v).Trim()
                                        If s = "" Then Continue For
                                        Dim cntNonEmpty As Integer = 0
                                        colNonEmpty.TryGetValue(c, cntNonEmpty)
                                        colNonEmpty(c) = cntNonEmpty + 1
                                        If IsStrictDateLike(s) Then
                                            Dim dt As DateTime
                                            If DateTime.TryParseExact(s, dateFormats, System.Globalization.CultureInfo.InvariantCulture,
                                                                      System.Globalization.DateTimeStyles.None, dt) Then
                                                Dim cntDateOk As Integer = 0
                                                colDateOk.TryGetValue(c, cntDateOk)
                                                colDateOk(c) = cntDateOk + 1
                                            End If
                                        End If
                                    Next
                                Next
                                ' ③ 数值日期序列号列：表头含"日期/date"且数值全部落在 Excel 日期序列号区间(1~2958465)，
                                '    视为日期列并统一应用日期显示格式（修复"统一日期格式对纯数值日期序列号不生效"）。
                                Dim colNumericDate As New HashSet(Of Integer)()
                                If hasHeader Then
                                    For c = 1 To values.GetLength(1)
                                        Dim hv = values(1, c)
                                        If hv Is Nothing OrElse IsDBNull(hv) Then Continue For
                                        Dim hdr = CStr(hv).Trim()
                                        If hdr.IndexOf("日期", StringComparison.OrdinalIgnoreCase) < 0 AndAlso
                                           hdr.IndexOf("date", StringComparison.OrdinalIgnoreCase) < 0 Then Continue For
                                        Dim allInRange As Boolean = True
                                        Dim hasNum As Boolean = False
                                        For r = dataStartRow To values.GetLength(0)
                                            Dim v = values(r, c)
                                            If v Is Nothing OrElse IsDBNull(v) Then Continue For
                                            If TypeOf v Is DateTime Then Continue For
                                            Dim isNumeric = TypeOf v Is Double OrElse TypeOf v Is Single OrElse TypeOf v Is Decimal OrElse
                                                TypeOf v Is SByte OrElse TypeOf v Is Byte OrElse TypeOf v Is Short OrElse TypeOf v Is UShort OrElse
                                                TypeOf v Is Integer OrElse TypeOf v Is UInteger OrElse TypeOf v Is Long OrElse TypeOf v Is ULong
                                            If isNumeric Then
                                                hasNum = True
                                                Dim dv = CDbl(v)
                                                If dv < 1 OrElse dv > 2958465 Then allInRange = False
                                            Else
                                                Dim sv = CStr(v).Trim()
                                                If sv <> "" Then allInRange = False
                                            End If
                                        Next
                                        If hasNum AndAlso allInRange Then colNumericDate.Add(c)
                                    Next
                                End If
                                ' 列级决策：整列确认为日期列
                                Dim confirmedDates As New HashSet(Of Integer)()
                                For c = 1 To values.GetLength(1)
                                    Dim isDateCol As Boolean = False
                                    If colFmtDtm.Contains(c) Then
                                        isDateCol = True
                                    ElseIf colHasDtm.Contains(c) Then
                                        isDateCol = True
                                    ElseIf colNumericDate.Contains(c) Then
                                        isDateCol = True
                                    ElseIf colNonEmpty.ContainsKey(c) AndAlso colNonEmpty(c) > 0 AndAlso
                                           colDateOk.ContainsKey(c) AndAlso colDateOk(c) / colNonEmpty(c) >= 0.5 Then
                                        isDateCol = True
                                    End If
                                    If isDateCol Then confirmedDates.Add(c)
                                Next
                                If confirmedDates.Count > 0 Then
                                    ' 第二遍：仅对确认的日期列，
                                    '   ① 把"严格日期形态"字符串转换为 DateTime
                                    '   ② 把"数值日期序列号"（如 45454，表头含日期/date 且整列∈1~2958465）转成真实 DateTime
                                    '      此前只判定 dateCols 不转换值，返回的 data 仍是序列号，写回后 Excel 仍是序列号 →"伪完成"。
                                    For r = dataStartRow To values.GetLength(0)
                                        For c = 1 To values.GetLength(1)
                                            If Not confirmedDates.Contains(c) Then Continue For
                                            Dim v = values(r, c)
                                            If TypeOf v Is String Then
                                                Dim s = CStr(v).Trim()
                                                If s <> "" AndAlso IsStrictDateLike(s) Then
                                                    Dim dt As DateTime
                                                    If DateTime.TryParseExact(s, dateFormats, System.Globalization.CultureInfo.InvariantCulture,
                                                                              System.Globalization.DateTimeStyles.None, dt) Then
                                                        values(r, c) = dt
                                                        normalizedDates += 1
                                                    End If
                                                End If
                                            ElseIf colNumericDate.Contains(c) Then
                                                Dim isNumN = TypeOf v Is Double OrElse TypeOf v Is Single OrElse TypeOf v Is Decimal OrElse
                                                    TypeOf v Is SByte OrElse TypeOf v Is Byte OrElse TypeOf v Is Short OrElse TypeOf v Is UShort OrElse
                                                    TypeOf v Is Integer OrElse TypeOf v Is UInteger OrElse TypeOf v Is Long OrElse TypeOf v Is ULong
                                                If isNumN Then
                                                    Dim dv = CDbl(v)
                                                    If dv >= 1 AndAlso dv <= 2958465 Then
                                                        Try
                                                            values(r, c) = DateTime.FromOADate(dv)
                                                            normalizedDates += 1
                                                        Catch
                                                        End Try
                                                    End If
                                                End If
                                            End If
                                        Next
                                    Next
                                    For Each dc In confirmedDates
                                        dateCols.Add(dc)
                                    Next
                                End If

                            Case "removeduplicates", "removeduplicate", "dedupe", "dedup"
                                ' excel_clean_data 也支持去重：把重复行标进 rowsToDelete，写回后删除。
                                Dim rdupcCols As New List(Of Integer)()
                                If dedupeColsRaw IsNot Nothing Then
                                    Dim rawItems As New List(Of Object)()
                                    If TypeOf dedupeColsRaw Is String Then
                                        Dim s0 = CStr(dedupeColsRaw).Trim()
                                        If s0.StartsWith("[") AndAlso s0.EndsWith("]") Then
                                            Try
                                                Dim parsed = Newtonsoft.Json.JsonConvert.DeserializeObject(Of List(Of Object))(s0)
                                                For Each o In parsed
                                                    If o IsNot Nothing Then rawItems.Add(o)
                                                Next
                                            Catch : End Try
                                        Else
                                            For Each p In s0.Split({","c, ";"c, "|"c, " "c, "，"c, "；"c}, StringSplitOptions.RemoveEmptyEntries)
                                                rawItems.Add(p)
                                            Next
                                        End If
                                    ElseIf TypeOf dedupeColsRaw Is System.Collections.IEnumerable Then
                                        For Each o In DirectCast(dedupeColsRaw, System.Collections.IEnumerable)
                                            If o Is Nothing Then Continue For
                                            ' Newtonsoft 反序列化 JSON 数组元素是 JValue（非 String/Integer），需解包底层值
                                            If TypeOf o Is Newtonsoft.Json.Linq.JValue Then
                                                rawItems.Add(DirectCast(o, Newtonsoft.Json.Linq.JValue).Value)
                                            Else
                                                rawItems.Add(o)
                                            End If
                                        Next
                                    Else
                                        rawItems.Add(dedupeColsRaw)
                                    End If
                                    For Each it In rawItems
                                        Dim col As Integer = -1
                                        If TypeOf it Is String Then
                                            Dim st = CStr(it).Trim()
                                            If st <> "" AndAlso System.Text.RegularExpressions.Regex.IsMatch(st, "^[A-Za-z]+$") AndAlso st.Length <= 3 Then
                                                col = ExcelMcpCoreUtils.ColumnLetterToNumber(st)
                                            ElseIf st <> "" Then
                                                For cc = 1 To totalCols
                                                    Dim hv = values(1, cc)
                                                    If hv IsNot Nothing AndAlso Not IsDBNull(hv) AndAlso CStr(hv).Trim() = st Then
                                                        col = cc
                                                        Exit For
                                                    End If
                                                Next
                                            End If
                                        Else
                                            Dim n As Integer = 0
                                            If TypeOf it Is Double Then
                                                n = CInt(DirectCast(it, Double))
                                            ElseIf TypeOf it Is Integer Then
                                                n = DirectCast(it, Integer)
                                            ElseIf TypeOf it Is Int64 Then
                                                n = CInt(DirectCast(it, Int64))
                                            Else
                                                Integer.TryParse(CStr(it), n)
                                            End If
                                            If n >= 1 Then col = n
                                        End If
                                        If col >= 1 AndAlso col <= totalCols AndAlso Not rdupcCols.Contains(col) Then rdupcCols.Add(col)
                                    Next
                                End If
                                Dim seenSet As New HashSet(Of String)(StringComparer.Ordinal)
                                For rcol = dataStartRow To values.GetLength(0)
                                    Dim colsToScan As New List(Of Integer)()
                                    If rdupcCols.Count > 0 Then
                                        colsToScan.AddRange(rdupcCols)
                                    Else
                                        For ccol = 1 To values.GetLength(1)
                                            colsToScan.Add(ccol)
                                        Next
                                    End If
                                    Dim blank As Boolean = True
                                    For Each cc In colsToScan
                                        If Not CleanCellIsBlank(values(rcol, cc)) Then
                                            blank = False
                                            Exit For
                                        End If
                                    Next
                                    ' 全空行交给 removeEmptyRows，避免两行空白被算成「重复行」
                                    If blank Then Continue For
                                    Dim sbK As New System.Text.StringBuilder()
                                    For Each cc In colsToScan
                                        sbK.Append(CleanCellKey(values(rcol, cc))).Append(vbTab)
                                    Next
                                    Dim rkey = sbK.ToString()
                                    If seenSet.Contains(rkey) Then
                                        If Not rowsToDelete.Contains(rcol) Then
                                            rowsToDelete.Add(rcol)
                                            removedDuplicateRows += 1
                                            duplicateAbsRows.Add(rng.Row + rcol - 1)
                                        End If
                                    Else
                                        seenSet.Add(rkey)
                                    End If
                                Next

                            Case "normalizenumbers"
                                ' 统一数字绝不能破坏日期/时间/文本列。
                                '   判定策略（三重防线，覆盖"合并后 NumberFormat 被重置为 General"的隐蔽场景）：
                                '   ① NumberFormat 明确是日期/时间格式 → 整列排除；
                                '   ② 列内存在 DateTime 类型的值（即使格式为 General）→ 整列排除；
                                '   ③ 列主要数值落在 Excel 日期序列号区间(1~2958465) 且列内存在"非整数值"（带时间小数部分），
                                '      说明很可能是日期+时间序列号 → 整列排除（避免误伤整数序号/ID）。
                                Dim nonNumericCols As New HashSet(Of Integer)()
                                Dim fmtDetected As New HashSet(Of Integer)()
                                For c = 1 To values.GetLength(1)
                                    Dim fmt As String = ""
                                    Try : fmt = CStr(DirectCast(rng.Columns(c), Microsoft.Office.Interop.Excel.Range).NumberFormat) : Catch : End Try
                                    If fmt <> "" AndAlso IsDateTimeFormat(fmt) Then
                                        nonNumericCols.Add(c)
                                        fmtDetected.Add(c)
                                    End If
                                Next

                                ' ② 列内 DateTime 值检测 & ③ 日期序列号区间检测
                                Dim hasDateVal(totalCols) As Boolean
                                Dim allInDateRange(totalCols) As Boolean
                                Dim hasFractional(totalCols) As Boolean   ' 是否存在带小数（时间部分）的数值
                                For idx = 0 To totalCols - 1 : allInDateRange(idx) = True : Next
                                For r = dataStartRow To values.GetLength(0)
                                    For c = 1 To values.GetLength(1)
                                        If nonNumericCols.Contains(c) Then Continue For
                                        Dim v = values(r, c)
                                        If v Is Nothing OrElse IsDBNull(v) Then Continue For
                                        If TypeOf v Is DateTime Then
                                            hasDateVal(c) = True
                                            nonNumericCols.Add(c)
                                            Continue For
                                        End If
                                        If allInDateRange(c) Then
                                            If TypeOf v Is Double OrElse TypeOf v Is Single OrElse TypeOf v Is Decimal Then
                                                Dim dv = CDbl(v)
                                                If dv < 1 OrElse dv > 2958465 Then
                                                    allInDateRange(c) = False
                                                Else
                                                    If dv <> Math.Floor(dv) Then hasFractional(c) = True
                                                End If
                                            ElseIf TypeOf v Is String Then
                                                ' 文本型非纯数字 → 不可能整列是日期序列号
                                                Dim sv = CStr(v).Trim()
                                                If Not System.Text.RegularExpressions.Regex.IsMatch(sv, "^[+-]?\d+(\.\d+)?$") Then allInDateRange(c) = False
                                            Else
                                                allInDateRange(c) = False
                                            End If
                                        End If
                                    Next
                                Next

                                For c = 1 To values.GetLength(1)
                                    ' General 格式 + 列内存在 DateTime 值 → 极可能是日期列，排除
                                    If Not fmtDetected.Contains(c) AndAlso hasDateVal(c) Then nonNumericCols.Add(c)
                                    ' General 格式 + 全列数值都在日期序列号区间 + 至少一个带时间小数部分的数值 → 可能是日期时间列，排除
                                    If Not fmtDetected.Contains(c) AndAlso allInDateRange(c) AndAlso hasFractional(c) Then nonNumericCols.Add(c)
                                Next

                                For r = dataStartRow To values.GetLength(0)
                                    For c = 1 To values.GetLength(1)
                                        If nonNumericCols.Contains(c) Then Continue For
                                        Dim v = values(r, c)
                                        If v Is Nothing OrElse IsDBNull(v) Then Continue For
                                        ' 日期值本身（DateTime）绝不当作数字统一
                                        If TypeOf v Is DateTime Then
                                            If Not dateCols.Contains(c) Then dateCols.Add(c)
                                            Continue For
                                        End If
                                        If TypeOf v Is Double OrElse TypeOf v Is Single OrElse TypeOf v Is Decimal OrElse
                                           TypeOf v Is Integer OrElse TypeOf v Is Long OrElse TypeOf v Is Short Then
                                            numCols.Add(c)
                                        Else
                                            Dim s = CStr(v).Trim()
                                            ' 仅转换纯数字串（无千分位逗号），避免 "1,234" 歧义；同时排除日期样字符串
                                            If System.Text.RegularExpressions.Regex.IsMatch(s, "^[+-]?(\d+(\.\d+)?|\.\d+)$") Then
                                                Dim d As Double
                                                If Double.TryParse(s, d) Then
                                                    values(r, c) = d
                                                    normalizedNumbers += 1
                                                    numCols.Add(c)
                                                End If
                                            End If
                                        End If
                                    Next
                                Next

                            Case "removeemptyrows"
                                For r = dataStartRow To values.GetLength(0)
                                    Dim allEmpty As Boolean = True
                                    For c = 1 To values.GetLength(1)
                                        Dim v = values(r, c)
                                        If Not (v Is Nothing OrElse IsDBNull(v) OrElse CStr(v).Trim() = "") Then
                                            allEmpty = False
                                            Exit For
                                        End If
                                    Next
                                    If allEmpty AndAlso Not rowsToDelete.Contains(r) Then
                                        rowsToDelete.Add(r)
                                        removedEmptyRows += 1
                                    End If
                                Next
                        End Select
                    Next

                    ' ────────────────────────────────────────────────────────────
                    ' 6) 副作用控制（三档安全，默认零副作用）
                    '
                    '   层级 A — dryRun=True（默认）
                    '     → 只返回 cleanRows / removedRows 数据，Excel 原表完全不动。
                    '     → Agent 拿到数据后可以：
                    '         a) 先展示预览给用户确认，再决定要不要落盘
                    '         b) 调 excel_write_range(targetSheet=...) 把结果写入新表
                    '
                    '   层级 B — writeBack=True
                    '     → 就地写回源表（改值 + 删空行），但**不自动 Save**。
                    '     → Excel 工作簿被标记 Dirty，关闭时弹保存确认。
                    '     → 返回值里会加 "dirtyWorkbook": True 和 "hint": "请手动保存或确认放弃"。
                    '
                    '   层级 C — writeBack=True + autoSave=True
                    '     → 就地写回 + wb.Save()。用于"用户显式同意修改并保存"的场景。
                    '
                    ' 为什么这么设计？
                    '   - Excel 里就地修改是破坏性操作（Ctrl+Z 撤销栈有限），
                    '     Agent 不应该在没有用户确认的情况下直接改源数据。
                    '   - ToolDefinition 文档明确承诺"仅返回数据，不回写"——
                    '     旧代码偷偷改了默认行为，必须跟文档对齐。
                    '   - targetSheet 落地是零副作用的替代：结果写入新表，
                    '     原表纹丝不动，用户觉得满意再手动替换。
                    ' ────────────────────────────────────────────────────────────
                    Dim dryRun As Boolean = GetArg(Of Boolean)(args, "dryRun", True)   ' 改默认 True ← 零副作用
                    Dim writeBack As Boolean = GetArg(Of Boolean)(args, "writeBack", False)
                    Dim autoSave As Boolean = GetArg(Of Boolean)(args, "autoSave", False)
                    Dim shouldWriteBack As Boolean = (Not dryRun) AndAlso writeBack   ' dryRun 互斥
                    Dim didWriteBack As Boolean = False
                    Dim didAutoSave As Boolean = False

                    If shouldWriteBack Then
                        Try
                            ' 6a 写回就地修改（trim/normalize/fill/standardize）到对应源单元格
                            For r = dataStartRow To values.GetLength(0)
                                If rowsToDelete.Contains(r) Then Continue For
                                For c = 1 To values.GetLength(1)
                                    Dim cell As Microsoft.Office.Interop.Excel.Range = Nothing
                                    Try
                                        cell = DirectCast(rng.Cells(r, c), Microsoft.Office.Interop.Excel.Range)
                                        Dim newVal = values(r, c)
                                        Dim curVal = cell.Value
                                        Dim curStr = If(curVal Is Nothing, "", CStr(curVal))
                                        Dim newStr = If(newVal Is Nothing OrElse IsDBNull(newVal), "", CStr(newVal))
                                        If curStr <> newStr Then cell.Value = newVal
                                    Catch
                                    Finally
                                        If cell IsNot Nothing Then Marshal.ReleaseComObject(cell)
                                    End Try
                                Next
                            Next
                            ' 6b 从后往前删除空行（绝对行号），避免前删导致下方行号上移错位
                            If rowsToDelete.Count > 0 Then
                                rowsToDelete.Sort()
                                For i = rowsToDelete.Count - 1 To 0 Step -1
                                    Dim absRow As Integer = rng.Row + (rowsToDelete(i) - 1)
                                    Dim rowRng As Microsoft.Office.Interop.Excel.Range = Nothing
                                    Try
                                        rowRng = DirectCast(ws.Rows(absRow), Microsoft.Office.Interop.Excel.Range)
                                        rowRng.Delete()
                                    Catch
                                    Finally
                                        If rowRng IsNot Nothing Then Marshal.ReleaseComObject(rowRng)
                                    End Try
                                Next
                            End If
                            didWriteBack = True
                            resultObjWrittenBack = True

                            ' 层级 C —— autoSave 显式保存（通常用户同意后才开）
                            If autoSave Then
                                Try
                                    excelApp.ActiveWorkbook.Save()
                                    didAutoSave = True
                                Catch exS As Exception
                                    System.Diagnostics.Debug.WriteLine($"[ExecuteCleanData] Save 失败: {exS.Message}")
                                End Try
                            End If
                        Catch exW As Exception
                            System.Diagnostics.Debug.WriteLine($"[ExecuteCleanData] 写回源表失败: {exW.Message}")
                        End Try
                    End If

                    ' ---------- 6c) 构建清洗后的数据（用于返回预览/落地表） ----------
                    ' 表头行
                    Dim headers() As Object = Nothing
                    If hasHeader AndAlso values.GetLength(0) >= 1 Then
                        Dim hdr(values.GetLength(1) - 1) As Object
                        For cc = 1 To values.GetLength(1)
                            hdr(cc - 1) = values(1, cc)
                        Next
                        headers = hdr
                    End If
                    Dim cleanRows As New List(Of Object())()
                    If headers IsNot Nothing Then cleanRows.Add(headers)
                    For r = dataStartRow To values.GetLength(0)
                        If rowsToDelete.Contains(r) Then Continue For
                        Dim rowArr(values.GetLength(1) - 1) As Object
                        For cc = 1 To values.GetLength(1)
                            rowArr(cc - 1) = values(r, cc)
                        Next
                        cleanRows.Add(rowArr)
                    Next
                    ' 可选落地：调用方传入 targetSheet 时写回清洗结果到该表（存在则覆盖，不存在则新建）
                    Dim landedSheetName As String = Nothing
                    Dim targetTs = GetArg(Of String)(args, "targetSheet", "")
                    If Not String.IsNullOrWhiteSpace(targetTs) Then
                        Dim rowLists As New List(Of List(Of Object))()
                        For Each cr In cleanRows
                            Dim oarr = TryCast(cr, Object())
                            If oarr IsNot Nothing Then rowLists.Add(oarr.ToList())
                        Next
                        landedSheetName = WriteResultToTargetSheet(args, rowLists, "清洗结果")
                    End If
                    Dim resultObj2 As New Dictionary(Of String, Object)()
                    resultObj2("success") = True
                    resultObj2("cleaned") = True
                    ' ── 副作用状态显式暴露给 Agent ──
                    resultObj2("mode") = If(dryRun, "dryRun_preview",
                                        If(didWriteBack,
                                            If(didAutoSave, "writeBack_saved", "writeBack_dirty"),
                                            "landed_new_sheet"))
                    resultObj2("sideEffect") = didWriteBack OrElse (landedSheetName IsNot Nothing)
                    resultObj2("dirtyWorkbook") = didWriteBack AndAlso Not didAutoSave
                    ' ──────────────────────────────
                    resultObj2("workbook") = excelApp.ActiveWorkbook.Name
                    resultObj2("worksheet") = ws.Name
                    resultObj2("range") = addrDisplay
                    resultObj2("hasHeader") = hasHeader
                    resultObj2("headers") = headers
                    resultObj2("cleanedRows") = cleanRows.Count - If(hasHeader, 1, 0)
                    resultObj2("removedRows") = rowsToDelete.Count
                    ' 各操作的实际改动量（前端 buildCleanOpSummary 据此如实展示，避免"已统一但无明细"）
                    resultObj2("trimmedCells") = trimmedCells
                    resultObj2("normalizedDates") = normalizedDates
                    resultObj2("normalizedNumbers") = normalizedNumbers
                    resultObj2("removedEmptyRows") = removedEmptyRows
                    resultObj2("removedDuplicateRows") = removedDuplicateRows
                    resultObj2("duplicateRows") = duplicateAbsRows
                    If normalizedDates > 0 OrElse dateCols.Count > 0 Then resultObj2("dateCols") = dateCols.ToList()
                    If normalizedNumbers > 0 OrElse numCols.Count > 0 Then resultObj2("numberCols") = numCols.ToList()
                    ' 落盘说明（优先级：autoSave 成功 > writeBack Dirty > landed 新表 > dryRun）
                    If didAutoSave Then
                        resultObj2("writtenTo") = ws.Name & "（已修改并保存）"
                    ElseIf didWriteBack Then
                        resultObj2("writtenTo") = ws.Name & "（已直接修改，工作簿待保存）"
                    ElseIf landedSheetName IsNot Nothing Then
                        resultObj2("writtenTo") = landedSheetName & "（新表落地，不影响源数据）"
                    ElseIf dryRun Then
                        resultObj2("writtenTo") = "（dryRun 预览，未修改工作表）"
                    End If
                    ' dirty 状态下给 Agent 一个明确提示让它告诉用户
                    If didWriteBack AndAlso Not didAutoSave Then
                        Dim hint As String = "工作簿已被就地修改但未保存。建议：a) 调 excel_save_workbook 显式保存；" &
                                           "b) 或用 targetSheet 参数把结果落地到新表再让用户确认；" &
                                           "c) 或在 autoSave=True 时再次调用并确认要保存。"
                        resultObj2("hint") = hint
                    End If
                    If sheetFallbackNote <> "" Then resultObj2("note") = sheetFallbackNote
                    If defaultOpsNote <> "" Then
                        ' 【修复】Dictionary 直接索引不存在的键会抛 KeyNotFoundException
                        '（实测"给定关键字不在字典中"），必须先 TryGetValue。
                        Dim existingNote As String = ""
                        If resultObj2.ContainsKey("note") Then existingNote = CStr(resultObj2("note"))
                        resultObj2("note") = existingNote & defaultOpsNote
                    End If
                    resultObj2("data") = cleanRows
                    result.Content.Add(New McpTextContent With {
                        .Text = JsonConvert.SerializeObject(resultObj2, Formatting.Indented)
                    })
                    result.TaskCompleted = True
                    Return result

                Catch ex2 As Exception
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"清洗操作异常：{ex2.Message}" & sheetFallbackNote})
                    Return result
                Finally
                    Try : Marshal.ReleaseComObject(rng) : Catch : End Try
                    Try : Marshal.ReleaseComObject(ws) : Catch : End Try
                End Try
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"工具初始化失败 [{ex.GetType().Name}]: {ex.Message}"})
                Return result
            End Try
        End Function

        ''' <summary>
        ''' JSON 字符串转义辅助函数
        ''' </summary>
        Private Shared Function EscapeJson(s As String) As String
            If s Is Nothing Then Return ""
            Return s.Replace("\", "\\").Replace("""", "\""").Replace(vbCr, "\r").Replace(vbLf, "\n").Replace(vbTab, "\t")
        End Function

        ''' <summary>
        ''' 判断 Excel NumberFormat 是否为日期/时间类格式（含 y/M/d/h/s/年月日 等），
        ''' 用于将日期/时间列从"数字统一"中排除，避免把日期值强套成纯数字格式。
        ''' </summary>
        Private Shared Function IsDateTimeFormat(fmt As String) As Boolean
            If String.IsNullOrWhiteSpace(fmt) Then Return False
            Dim f = fmt.Trim()
            ' 内置日期/时间格式（Excel 常见样式码，按系统区域设置可能不同）
            If f.StartsWith("m/d/") OrElse f.StartsWith("mm/dd/") OrElse f.StartsWith("d/m/") OrElse f.StartsWith("yyyy/") Then Return True
            ' 中文日期格式（年月日）
            If f.IndexOf("年") >= 0 OrElse f.IndexOf("月") >= 0 OrElse f.IndexOf("日") >= 0 Then Return True
            ' 含日期符号(y/d) 且含时间符号(h/:) → 日期时间格式
            Dim hasDateSym = f.IndexOf("y"c) >= 0 OrElse f.IndexOf("Y"c) >= 0 OrElse f.IndexOf("d"c) >= 0 OrElse f.IndexOf("D"c) >= 0
            Dim hasTimeSym = f.IndexOf("h"c) >= 0 OrElse f.IndexOf("H"c) >= 0 OrElse f.IndexOf(":"c) >= 0 OrElse
                             f.IndexOf("上午") >= 0 OrElse f.IndexOf("下午") >= 0
            If hasDateSym AndAlso hasTimeSym Then Return True
            ' 单独时间格式（h:mm / h:mm:ss / 上下午 h）
            If hasTimeSym Then Return True
            ' 纯日期（yyyy-mm-dd / mm/dd/yyyy）——含 y 且含 m/d
            Dim hasMonth = f.IndexOf("m"c) >= 0 OrElse f.IndexOf("M"c) >= 0
            If hasDateSym AndAlso hasMonth Then Return True
            Return False
        End Function

        '#End Region

    End Class

End Namespace
