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

        '#Region "工具实现 - 数据与图表"

        ''' <summary>
        ''' 批量执行多个 Excel 操作
        ''' </summary>
        Private Function ExecuteBatchExecute(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()

            Try
                Dim operationsJson As String = GetArg(Of String)(args, "operations", "[]")
                Dim operations = JsonConvert.DeserializeObject(Of List(Of Dictionary(Of String, Object)))(operationsJson)

                If operations Is Nothing OrElse operations.Count = 0 Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "操作列表不能为空"})
                    Return result
                End If

                Dim results As New List(Of Dictionary(Of String, Object))()
                Dim successCount = 0
                Dim failCount = 0

                For Each op In operations
                    Dim toolName = If(op.ContainsKey("toolName"), op("toolName").ToString(), "")
                    Dim argumentsJson = If(op.ContainsKey("arguments"), op("arguments").ToString(), "{}")

                    Dim opArgs As Dictionary(Of String, Object)
                    Try
                        opArgs = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(argumentsJson)
                    Catch
                        opArgs = New Dictionary(Of String, Object)()
                    End Try

                    Dim opResult = ExecuteTool(toolName, opArgs)

                    Dim opResultObj = New Dictionary(Of String, Object) From {
                        {"toolName", toolName},
                        {"success", Not opResult.IsError},
                        {"message", If(opResult.Content.Count > 0 AndAlso TypeOf opResult.Content(0) Is McpTextContent, DirectCast(opResult.Content(0), McpTextContent).Text, "")}
                    }
                    results.Add(opResultObj)

                    If opResult.IsError Then
                        failCount += 1
                    Else
                        successCount += 1
                    End If
                Next

                Dim batchResult = New Dictionary(Of String, Object) From {
                    {"total", operations.Count},
                    {"success", successCount},
                    {"failed", failCount},
                    {"results", results}
                }

                result.Content.Add(New McpTextContent With {
                    .Text = JsonConvert.SerializeObject(batchResult, Formatting.Indented)
                })

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"批量执行失败: {ex.Message}"})
            End Try

            Return result
        End Function
        ''' <summary>
        ''' 对指定范围排序 —— 完整还原 Excel 原生"排序"对话框能力（2026-10-16 重写）
        ''' 原生对话框四要素：
        '''   ① 列：多级"排序依据 / 次要关键字"（sortBy 数组，按 Priority 依次生效）
        '''   ② 排序依据：单元格值 / 单元格颜色 / 字体颜色 / 条件格式图标（sortBy[].sortOn）
        '''   ③ 次序：升序 / 降序（sortBy[].order；颜色排序时 ascending=指定颜色排最前、descending=排最后）
        '''   ④ 选项（对话框"选项"按钮）：区分大小写 matchCase / 方向 orientation（按列|按行）/ 排序方法 sortMethod（拼音|笔划）
        '''
        ''' ⚠️ 历史坑（本次实测定修复）：旧实现把 SortOn 硬编码为 1。本机实测（Excel 16.0 build 20430，
        '''    与 PIA v15 枚举一致）：xlSortOnValues=0 / xlSortOnCellColor=1 / xlSortOnFontColor=2 / xlSortOnIcon=3。
        '''    硬编码 1 = "按单元格颜色"且未指定颜色 → Excel 直接空操作，"按值排序"从未生效（已用独立
        '''    Excel 实例复现：SortOn=1 排序后数据一字未动，SortOn=0 才真正排序）。
        ''' ⚠️ 颜色排序的正确姿势：Add 之后取 SortField.SortOnValue 再设其 "Color" 成员（反射后期绑定）；
        '''    本机实测该对象既不是 ColorFormat（无 RGB 成员），也不能整体赋值（报"找不到成员"），
        '''    且不带颜色等于空操作。
        ''' ⚠️ Sort 对象的 MatchCase/Orientation/SortMethod 在 Worksheet 上是持久状态，每次必须显式重置，
        '''    否则上一次"按行排序/按笔划"会残留污染本次排序。
        ''' </summary>
        Private Function ExecuteSortRange(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()

            Try
                Dim rangeAddr As String = GetArg(Of String)(args, "range", "")
                Dim sortByJson As String = GetArg(Of String)(args, "sortBy", "[]")
                Dim hasHeader As Boolean = GetArg(Of Boolean)(args, "hasHeader", True)
                ' ── 原生"选项"对话框三要素 ──
                Dim matchCase As Boolean = GetArg(Of Boolean)(args, "matchCase", False)
                Dim orientationArg As String = GetArg(Of String)(args, "orientation", "")
                Dim sortMethodArg As String = GetArg(Of String)(args, "sortMethod", "")
                ' 排序后回传前 N 行预览（默认 5）：Agent 总结时就有真实数据可填表格，
                ' 不再"承诺画表格却只能输出 '$' 占位行"（2026-10-16 实测根因）
                Dim previewRows As Integer = GetArg(Of Integer)(args, "previewRows", 5)
                If previewRows < 0 Then previewRows = 0
                If previewRows > 50 Then previewRows = 50

                If String.IsNullOrEmpty(rangeAddr) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "范围不能为空"})
                    Return result
                End If

                Dim app = Globals.ThisAddIn.Application
                If app Is Nothing OrElse app.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何工作簿"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = TryCast(app.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                Dim rng As Microsoft.Office.Interop.Excel.Range = ws.Range(rangeAddr)

                ' 解析排序列
                Dim sortColumns As List(Of Dictionary(Of String, Object))
                Try
                    sortColumns = JsonConvert.DeserializeObject(Of List(Of Dictionary(Of String, Object)))(sortByJson)
                Catch
                    sortColumns = New List(Of Dictionary(Of String, Object))()
                End Try

                ' 如果没有指定排序列，默认按第一列升序
                If sortColumns Is Nothing OrElse sortColumns.Count = 0 Then
                    sortColumns.Add(New Dictionary(Of String, Object) From {
                        {"column", "A"},
                        {"order", "ascending"}
                    })
                End If

                ' 执行排序
                ' 【修复·_Sort 方法无效】旧实现 DirectCast(rng.Sort, Excel.Sort) —— Range 上只有
                ' Sort **方法**（旧式 30 参数排序），没有 Sort **属性**；把方法当属性访问会以
                ' 空参数调用 _Sort，COM 报"类 Range 的 _Sort 方法无效"（实测 2026-10-06 截图）。
                ' 带 SortFields/SetRange 的 Sort 对象挂在 **Worksheet** 上：ws.Sort。
                Dim sortObj As Microsoft.Office.Interop.Excel.Sort = ws.Sort
                sortObj.SortFields.Clear()
                ' 每次都显式重置三个选项：Sort 对象状态挂在 Worksheet 上、跨次持久，
                ' 上一次的"按行排序/区分大小写/按笔划"会残留污染本次（实测确认）。
                sortObj.MatchCase = matchCase
                Dim wantRows = String.Equals(orientationArg, "rows", StringComparison.OrdinalIgnoreCase) OrElse
                               String.Equals(orientationArg, "按行", StringComparison.OrdinalIgnoreCase) OrElse
                               String.Equals(orientationArg, "leftToRight", StringComparison.OrdinalIgnoreCase)
                sortObj.Orientation = If(wantRows,
                    Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortRows,
                    Microsoft.Office.Interop.Excel.XlSortOrientation.xlSortColumns)
                Dim sortMethod = ResolveSortMethod(sortMethodArg)
                If sortMethod.HasValue Then
                    sortObj.SortMethod = sortMethod.Value
                End If

                Dim keyDescriptions As New List(Of String)()
                Dim keyCount = 0

                For Each sortDef In sortColumns
                    Dim colLetter = _SortDictStr(sortDef, "column", "A").Trim()
                    Dim orderStr = _SortDictStr(sortDef, "order", "ascending").Trim()
                    Dim sortOn = ResolveSortOn(sortDef)
                    Dim isDescending = (orderStr.ToLower() = "descending") OrElse (orderStr = "降序") OrElse (orderStr.ToLower() = "desc")

                    ' 计算排序列的绝对列号，三级解析：
                    '   ① 数字（"2"）= 范围内相对列号；② 字母（"H"）= 绝对列号；
                    '   ③ 其它（中文列名，如"销售额"——弱模型高频传法，实测截图）= 与范围首行表头逐格匹配。
                    Dim absCol As Integer = 0
                    If IsNumeric(colLetter) Then
                        absCol = rng.Column + CInt(CDbl(colLetter)) - 1
                    ElseIf System.Text.RegularExpressions.Regex.IsMatch(colLetter, "^[A-Za-z]{1,3}$") Then
                        absCol = ExcelMcpCoreUtils.ColumnLetterToNumber(colLetter)
                    Else
                        For c As Integer = 1 To CInt(rng.Columns.Count)
                            Dim headerCell As Microsoft.Office.Interop.Excel.Range = DirectCast(rng.Cells(1, c), Microsoft.Office.Interop.Excel.Range)
                            Dim headerText As String = ""
                            Try
                                If headerCell.Value2 IsNot Nothing Then headerText = Trim(CStr(headerCell.Value2))
                            Catch
                            End Try
                            Utils.ComReleaseHelper.Release(headerCell)
                            If String.Equals(headerText, colLetter, StringComparison.OrdinalIgnoreCase) Then
                                absCol = rng.Column + c - 1
                                Exit For
                            End If
                        Next
                        If absCol = 0 Then
                            ' 找不到列名 → 把可用表头全部列出，Agent 看一眼就能改对参数（可自愈）
                            Dim headers As New List(Of String)()
                            For c2 As Integer = 1 To CInt(rng.Columns.Count)
                                Dim hc As Microsoft.Office.Interop.Excel.Range = DirectCast(rng.Cells(1, c2), Microsoft.Office.Interop.Excel.Range)
                                Try
                                    headers.Add(If(hc.Value2 Is Nothing, "(空)", Trim(CStr(hc.Value2))))
                                Catch
                                    headers.Add("(不可读)")
                                End Try
                                Utils.ComReleaseHelper.Release(hc)
                            Next
                            result.IsError = True
                            result.Content.Add(New McpTextContent With {.Text = $"排序列『{colLetter}』不是列字母/数字，且在 {rangeAddr} 的表头中未找到。可用表头：{String.Join("、", headers)}。请改用对应列字母（或精确表头名）重试。"})
                            Utils.ComReleaseHelper.Release(rng)
                            Utils.ComReleaseHelper.Release(ws)
                            Return result
                        End If
                    End If
                    If absCol < rng.Column OrElse absCol >= rng.Column + rng.Columns.Count Then
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = $"排序列超出范围: 列『{colLetter}』不在 {rangeAddr}（共 {CInt(rng.Columns.Count)} 列）内"})
                        Utils.ComReleaseHelper.Release(rng)
                        Utils.ComReleaseHelper.Release(ws)
                        Return result
                    End If
                    Dim sortKey As Microsoft.Office.Interop.Excel.Range = DirectCast(ws.Cells(rng.Row, absCol), Microsoft.Office.Interop.Excel.Range)
                    ' SortOn 必须显式传枚举值：硬编码 1 = xlSortOnCellColor（不带颜色即空操作，历史 bug）
                    Dim sf As Microsoft.Office.Interop.Excel.SortField =
                        sortObj.SortFields.Add(sortKey, CInt(sortOn), If(isDescending, 2, 1)) ' xlDescending=2, xlAscending=1
                    Utils.ComReleaseHelper.Release(sortKey)

                    If sortOn = Microsoft.Office.Interop.Excel.XlSortOn.xlSortOnCellColor OrElse
                       sortOn = Microsoft.Office.Interop.Excel.XlSortOn.xlSortOnFontColor Then
                        Dim colorInt = ResolveSortColor(sortDef)
                        If Not colorInt.HasValue Then
                            ' 未指定颜色 → 直接列出该列实际出现的颜色，Agent 一眼就能改对参数（与表头错误同风格）
                            Dim colorName = If(sortOn = Microsoft.Office.Interop.Excel.XlSortOn.xlSortOnFontColor, "字体颜色", "单元格颜色")
                            result.IsError = True
                            result.Content.Add(New McpTextContent With {
                                .Text = $"按{colorName}排序必须指定 color 参数（如 #FF0000）。列『{colLetter}』中实际出现的颜色：{ListColumnColors(ws, rng, absCol, sortOn)}。请带上 color 重试。"
                            })
                            Utils.ComReleaseHelper.Release(rng)
                            Utils.ComReleaseHelper.Release(ws)
                            Return result
                        End If
                        SetSortFieldColor(sf, colorInt.Value)
                    ElseIf sortOn = Microsoft.Office.Interop.Excel.XlSortOn.xlSortOnIcon Then
                        Try
                            Dim icon = ResolveSortIcon(app, sortDef)
                            sf.SetIcon(icon)
                        Catch exIcon As Exception
                            result.IsError = True
                            result.Content.Add(New McpTextContent With {
                                .Text = $"按条件格式图标排序失败：{exIcon.Message}。请检查 iconSet（图标集序号或名称，如 3Arrows）与 icon（图标序号，从 1 开始）参数。"
                            })
                            Utils.ComReleaseHelper.Release(rng)
                            Utils.ComReleaseHelper.Release(ws)
                            Return result
                        End Try
                    End If

                    Dim sortOnName = SortOnDisplayName(sortOn, sortDef)
                    Dim keyDesc = colLetter & If(sortOnName <> "", $"（{sortOnName}）", "") & If(isDescending, " 降序", " 升序")
                    If keyCount = 0 Then
                        keyDescriptions.Add(keyDesc)
                    Else
                        keyDescriptions.Add("次要关键字 " & keyDesc)
                    End If
                    keyCount += 1
                Next

                sortObj.Header = CType(If(hasHeader, 1, 2), Microsoft.Office.Interop.Excel.XlYesNoGuess) ' xlYes=1, xlNo=2
                ' 【必须 SetRange】不设的话 Apply 只按第一个 Key 的 CurrentRegion 取数据区，
                ' 与用户显式传入的 range 不一定一致（截图场景两者恰好相同才没暴露）。
                sortObj.SetRange(rng)
                sortObj.Apply()

                ' ── 排序后回传预览（表头 + 前 N 行），让 Agent 有真实数据可总结 ──
                Dim rowCount As Integer = CInt(rng.Rows.Count)
                Dim dataRowCount As Integer = If(hasHeader, rowCount - 1, rowCount)
                Dim previewHeaders As New List(Of Object)()
                Dim preview As New List(Of List(Of Object))()
                Dim truncated As Boolean = False
                Try
                    If hasHeader AndAlso rowCount >= 1 Then
                        Dim hRng = ws.Range(ws.Cells(rng.Row, rng.Column), ws.Cells(rng.Row, rng.Column + rng.Columns.Count - 1))
                        previewHeaders.AddRange(ReadRowValues(hRng))
                        Utils.ComReleaseHelper.Release(hRng)
                    Else
                        For c As Integer = 1 To CInt(rng.Columns.Count)
                            previewHeaders.Add("列" & c.ToString())
                        Next
                    End If
                    If previewRows > 0 AndAlso dataRowCount > 0 Then
                        Dim take As Integer = Math.Min(previewRows, dataRowCount)
                        truncated = take < dataRowCount
                        Dim firstRow As Integer = rng.Row + If(hasHeader, 1, 0)
                        Dim pRng = ws.Range(ws.Cells(firstRow, rng.Column), ws.Cells(firstRow + take - 1, rng.Column + rng.Columns.Count - 1))
                        Dim values = TryCast(pRng.Value2, Object(,))
                        For r As Integer = 1 To take
                            Dim rowVals As New List(Of Object)()
                            For c As Integer = 1 To CInt(rng.Columns.Count)
                                ' 单格范围（1行1列）时 Value2 是标量，TryCast 失败 → 逐格兜底
                                ' ⚠️ Cells(...) 在本 PIA 返回 Object，必须先 DirectCast 成 Range（Option Strict On）
                                Dim cellObj As Microsoft.Office.Interop.Excel.Range = DirectCast(pRng.Cells(r, c), Microsoft.Office.Interop.Excel.Range)
                                rowVals.Add(If(values IsNot Nothing, values(r, c), cellObj.Value2))
                                Utils.ComReleaseHelper.Release(cellObj)
                            Next
                            preview.Add(rowVals)
                        Next
                        Utils.ComReleaseHelper.Release(pRng)
                    End If
                Catch
                    ' 预览读取失败不影响排序结果本身
                End Try

                Utils.ComReleaseHelper.Release(rng)
                Utils.ComReleaseHelper.Release(ws)

                Dim payload = New With {
                    .success = True,
                    .message = $"已按 {String.Join("、", keyDescriptions)} 排序（范围 {rangeAddr}）",
                    .range = rangeAddr,
                    .rowCount = dataRowCount,
                    .sortedBy = keyDescriptions,
                    .headers = previewHeaders,
                    .previewRows = preview,
                    .truncated = truncated
                }
                result.Content.Add(New McpTextContent With {
                    .Text = JsonConvert.SerializeObject(payload, Formatting.Indented)
                })

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"排序失败: {ex.Message}"})
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 从 sortBy 项里宽容取字符串（容忍值本身是数字/布尔/null）
        ''' </summary>
        Private Function _SortDictStr(dict As Dictionary(Of String, Object), key As String, defaultValue As String) As String
            If dict Is Nothing OrElse Not dict.ContainsKey(key) OrElse dict(key) Is Nothing Then Return defaultValue
            Return Convert.ToString(dict(key), System.Globalization.CultureInfo.InvariantCulture)
        End Function

        ''' <summary>
        ''' 解析"排序依据"（对齐原生对话框下拉：单元格值 / 单元格颜色 / 字体颜色 / 条件格式图标）。
        ''' 兼容中英文与常见别名；未知值按"单元格值"处理。
        ''' </summary>
        Private Function ResolveSortOn(sortDef As Dictionary(Of String, Object)) As Microsoft.Office.Interop.Excel.XlSortOn
            Dim raw = _SortDictStr(sortDef, "sortOn", "values").Trim()
            Select Case raw.ToLowerInvariant()
                Case "cellcolor", "color", "cell", "background", "背景色", "填充色", "单元格颜色"
                    Return Microsoft.Office.Interop.Excel.XlSortOn.xlSortOnCellColor
                Case "fontcolor", "font", "文字颜色", "字体颜色"
                    Return Microsoft.Office.Interop.Excel.XlSortOn.xlSortOnFontColor
                Case "icon", "iconset", "条件格式图标", "图标"
                    Return Microsoft.Office.Interop.Excel.XlSortOn.xlSortOnIcon
                Case Else
                    Return Microsoft.Office.Interop.Excel.XlSortOn.xlSortOnValues
            End Select
        End Function

        ''' <summary>排序依据的展示名（空串=按值，不展示）</summary>
        Private Function SortOnDisplayName(sortOn As Microsoft.Office.Interop.Excel.XlSortOn, sortDef As Dictionary(Of String, Object)) As String
            Select Case sortOn
                Case Microsoft.Office.Interop.Excel.XlSortOn.xlSortOnCellColor
                    Dim cc = ResolveSortColor(sortDef)
                    Return If(cc.HasValue, $"单元格颜色 {ColorToHex(cc.Value)}", "单元格颜色")
                Case Microsoft.Office.Interop.Excel.XlSortOn.xlSortOnFontColor
                    Dim fc = ResolveSortColor(sortDef)
                    Return If(fc.HasValue, $"字体颜色 {ColorToHex(fc.Value)}", "字体颜色")
                Case Microsoft.Office.Interop.Excel.XlSortOn.xlSortOnIcon
                    Return "条件格式图标"
                Case Else
                    Return ""
            End Select
        End Function

        ''' <summary>
        ''' 解析排序颜色：支持 #RRGGBB / RRGGBB / "r,g,b" / 十进制 RGB 整数 / 常见中文颜色名。
        ''' 返回 Excel 的 RGB 整数（与 Interior.Color/Font.Color 同一空间，如 红色=255）。
        ''' </summary>
        Private Function ResolveSortColor(sortDef As Dictionary(Of String, Object)) As Integer?
            Dim raw = _SortDictStr(sortDef, "color", "").Trim()
            If raw = "" Then Return Nothing
            ' 中文颜色名
            Select Case raw
                Case "红", "红色" : Return 255
                Case "绿", "绿色" : Return 32768
                Case "蓝", "蓝色" : Return 16711680
                Case "黄", "黄色" : Return 65535
                Case "橙", "橙色", "橙黄" : Return 42495
                Case "紫", "紫色" : Return 8388736
                Case "青", "青色" : Return 8421376
                Case "黑", "黑色" : Return 0
                Case "白", "白色" : Return 16777215
                Case "灰", "灰色" : Return 8421504
            End Select
            ' #RRGGBB / RRGGBB
            Dim hex = raw.TrimStart("#"c)
            If System.Text.RegularExpressions.Regex.IsMatch(hex, "^[0-9a-fA-F]{6}$") Then
                Return Convert.ToInt32(hex, 16)
            End If
            ' r,g,b
            Dim parts = raw.Split(","c)
            If parts.Length = 3 Then
                Dim rr, gg, bb As Integer
                If Integer.TryParse(parts(0).Trim(), rr) AndAlso Integer.TryParse(parts(1).Trim(), gg) AndAlso Integer.TryParse(parts(2).Trim(), bb) Then
                    Return (rr And &HFF) << 16 Or (gg And &HFF) << 8 Or (bb And &HFF)
                End If
            End If
            ' 十进制 RGB 整数
            Dim n As Integer
            If Integer.TryParse(raw, n) Then Return n
            Return Nothing
        End Function

        ''' <summary>RGB 整数 → #RRGGBB（用于错误提示/摘要）</summary>
        Private Function ColorToHex(rgb As Integer) As String
            Return "#" & (rgb And &HFFFFFF).ToString("X6")
        End Function

        ''' <summary>
        ''' 给颜色排序字段设置颜色。⚠️ 本机实测（Excel 16.0 b20430）：SortOnValue 运行时既不是
        ''' ColorFormat（没有 RGB 成员，设 .RGB 抛"找不到属性"），也不能直接整体赋值（报"找不到成员"）；
        ''' 唯一可行路径是反射后期绑定设置其 "Color" 成员（IDispatch）。保留 ColorFormat.RGB
        ''' 作为次选路径，兼容 SortOnValue 是真 ColorFormat 的其它 Excel 构建。
        ''' </summary>
        Private Sub SetSortFieldColor(sf As Microsoft.Office.Interop.Excel.SortField, rgb As Integer)
            Dim sov As Object = sf.SortOnValue
            If sov Is Nothing Then Throw New Exception("无法获取排序字段的颜色载体（SortOnValue 为 Nothing）")

            ' 路径1（本机实测可用）：反射后期绑定设置 "Color"
            Try
                sov.GetType().InvokeMember("Color",
                    BindingFlags.SetProperty Or BindingFlags.Public Or BindingFlags.Instance,
                    Nothing, sov, New Object() {CObj(rgb)})
                Return
            Catch
                ' 落到路径2
            End Try

            ' 路径2（其它构建兼容）：SortOnValue 是真 ColorFormat 时用 RGB
            Dim cf = TryCast(sov, Microsoft.Office.Interop.Excel.ColorFormat)
            If cf IsNot Nothing Then
                cf.RGB = rgb
                Return
            End If
            Throw New Exception("无法设置排序颜色：SortOnValue 既不支持 Color 也不支持 RGB")
        End Sub

        ''' <summary>
        ''' 解析"排序方法"（原生对话框"选项"→排序方法）：pinYin=按字母顺序（拼音）/ stroke=按笔划。
        ''' 空/未识别返回 Nothing（保持 Excel 默认，不显式设置）。
        ''' </summary>
        Private Function ResolveSortMethod(raw As String) As Microsoft.Office.Interop.Excel.XlSortMethod?
            If String.IsNullOrEmpty(raw) Then Return Nothing
            Select Case raw.Trim().ToLowerInvariant()
                Case "pinyin", "pinYin", "字母", "拼音", "按字母", "按字母顺序"
                    Return Microsoft.Office.Interop.Excel.XlSortMethod.xlPinYin
                Case "stroke", "笔划", "笔画", "按笔划", "按笔画"
                    Return Microsoft.Office.Interop.Excel.XlSortMethod.xlStroke
                Case Else
                    Return Nothing
            End Select
        End Function

        ''' <summary>
        ''' 解析条件格式图标：iconSet 支持数字（图标集序号，1 基）或名称（如 "3Arrows"/"xl3Arrows"）；
        ''' icon 为图标序号（1 基，通常是 3/4/5 个）。
        ''' 名称→序号：IconSet.ID 是 XlIconSet 枚举，枚举值与集合序号一一对应（实测 xl3Arrows=1=第1个），
        ''' 因此 Enum.Parse 后取 CInt 即得集合序号，无需硬编码名称表。
        ''' </summary>
        Private Function ResolveSortIcon(app As Microsoft.Office.Interop.Excel.Application, sortDef As Dictionary(Of String, Object)) As Microsoft.Office.Interop.Excel.Icon
            Dim setRaw = _SortDictStr(sortDef, "iconSet", "").Trim()
            Dim iconRaw = _SortDictStr(sortDef, "icon", "").Trim()
            If setRaw = "" Then
                Throw New Exception("缺少 iconSet 参数（图标集序号或名称，如 3Arrows / 3TrafficLights1）")
            End If
            If iconRaw = "" Then
                Throw New Exception("缺少 icon 参数（图标序号，从 1 开始）")
            End If
            Dim iconIdx As Integer
            If Not Integer.TryParse(iconRaw, iconIdx) OrElse iconIdx < 1 Then
                Throw New Exception($"icon 参数『{iconRaw}』不是合法序号（应为从 1 开始的整数）")
            End If

            Dim iconSets = app.ActiveWorkbook.IconSets
            Dim setIdx As Integer
            Dim setNum As Integer
            If Integer.TryParse(setRaw, setNum) Then
                setIdx = setNum
            Else
                ' 名称匹配：XlIconSet 枚举（枚举名带 xl 前缀，如 xl3Arrows），两种写法都试
                Dim enumVal As Object = Nothing
                Dim ok = False
                Try
                    enumVal = System.Enum.Parse(GetType(Microsoft.Office.Interop.Excel.XlIconSet), setRaw, True)
                    ok = True
                Catch
                    ok = False
                End Try
                If Not ok Then
                    Try
                        enumVal = System.Enum.Parse(GetType(Microsoft.Office.Interop.Excel.XlIconSet), "xl" & setRaw, True)
                        ok = True
                    Catch
                        ok = False
                    End Try
                End If
                If ok Then
                    setIdx = CInt(enumVal)
                Else
                    ' 列出可用图标集名称，便于自愈
                    Dim names As New List(Of String)()
                    Dim total = iconSets.Count
                    For i As Integer = 1 To Math.Min(total, 20)
                        ' ⚠️ IconSets.Item(...) 在本 PIA 返回 Object，须 DirectCast（Option Strict On）
                        Dim isetName As Microsoft.Office.Interop.Excel.IconSet = DirectCast(iconSets.Item(i), Microsoft.Office.Interop.Excel.IconSet)
                        names.Add([Enum].GetName(GetType(Microsoft.Office.Interop.Excel.XlIconSet), isetName.ID))
                    Next
                    Throw New Exception($"图标集『{setRaw}』不存在。可用图标集（前 {names.Count} 个）：{String.Join("、", names)}")
                End If
            End If

            Dim iconSet As Microsoft.Office.Interop.Excel.IconSet = DirectCast(iconSets.Item(setIdx), Microsoft.Office.Interop.Excel.IconSet)
            If iconIdx > iconSet.Count Then
                Throw New Exception($"图标序号 {iconIdx} 超出图标集『{setRaw}』的范围（共 {iconSet.Count} 个图标）")
            End If
            Return DirectCast(iconSet.Item(iconIdx), Microsoft.Office.Interop.Excel.Icon)
        End Function

        ''' <summary>
        ''' 列出某列实际出现的填充色/字体色（去重、限量），用于"按颜色排序但未指定颜色"的自愈错误提示。
        ''' </summary>
        Private Function ListColumnColors(ws As Microsoft.Office.Interop.Excel.Worksheet,
                                          rng As Microsoft.Office.Interop.Excel.Range,
                                          absCol As Integer,
                                          sortOn As Microsoft.Office.Interop.Excel.XlSortOn) As String
            Dim found As New List(Of Integer)()
            Dim scanRows As Integer = Math.Min(CInt(rng.Rows.Count), 200)
            For r As Integer = rng.Row To rng.Row + scanRows - 1
                Dim cell As Microsoft.Office.Interop.Excel.Range = DirectCast(ws.Cells(r, absCol), Microsoft.Office.Interop.Excel.Range)
                Dim c As Integer
                Try
                    If sortOn = Microsoft.Office.Interop.Excel.XlSortOn.xlSortOnFontColor Then
                        c = CInt(cell.Font.Color)
                    Else
                        c = CInt(cell.Interior.Color)
                    End If
                Catch
                    c = -1
                End Try
                Utils.ComReleaseHelper.Release(cell)
                If c >= 0 AndAlso Not found.Contains(c) Then
                    found.Add(c)
                    If found.Count >= 8 Then Exit For
                End If
            Next
            If found.Count = 0 Then Return "（未检测到颜色）"
            Return String.Join("、", found.Select(Function(x) ColorToHex(x)))
        End Function

        ''' <summary>读取单行区域的值为 List（兼容单列/多列/单格）</summary>
        Private Function ReadRowValues(rowRng As Microsoft.Office.Interop.Excel.Range) As List(Of Object)
            Dim vals As New List(Of Object)()
            Dim v = TryCast(rowRng.Value2, Object(,))
            If v Is Nothing Then
                ' 单格或整行空：逐格取（单格 Value2 是标量，TryCast 到二维数组会失败）
                For c As Integer = 1 To CInt(rowRng.Columns.Count)
                    ' ⚠️ Cells(...) 在本 PIA 返回 Object，必须先 DirectCast 成 Range（Option Strict On）
                    Dim cellObj As Microsoft.Office.Interop.Excel.Range = DirectCast(rowRng.Cells(1, c), Microsoft.Office.Interop.Excel.Range)
                    vals.Add(cellObj.Value2)
                    Utils.ComReleaseHelper.Release(cellObj)
                Next
                Return vals
            End If
            For c As Integer = 1 To CInt(rowRng.Columns.Count)
                vals.Add(v(1, c))
            Next
            Return vals
        End Function
        ''' <summary>
        ''' 对指定范围应用自动筛选
        ''' </summary>
        Private Function ExecuteApplyFilter(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()

            Try
                Dim rangeAddr As String = GetArg(Of String)(args, "range", "")
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim criteriaJson As String = GetArg(Of String)(args, "criteria", "{}")
                Dim autoRange As Boolean = GetArg(Of Boolean)(args, "autoRange", False)

                Dim app = Globals.ThisAddIn.Application
                If app Is Nothing OrElse app.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何工作簿"})
                    Return result
                End If

                ' ========== 范围三级兜底（与 excel_format_range 同语义）==========
                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim rng As Microsoft.Office.Interop.Excel.Range = Nothing

                ' 回退1：显式 autoRange=true → ActiveSheet.UsedRange（"筛选整张表"）。
                ' 顺序很关键：Excel 永远存在选区（哪怕只是一个空单元格），若把"实时选区"放在前面，
                ' autoRange 永远轮不到，单格选区会把筛选退化成"只筛一个格" → 故 autoRange 优先。
                If String.IsNullOrWhiteSpace(rangeAddr) AndAlso autoRange Then
                    ws = TryCast(app.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    If ws IsNot Nothing Then
                        Dim used As Microsoft.Office.Interop.Excel.Range = ws.UsedRange
                        If used IsNot Nothing AndAlso used.Cells.Count > 0 Then
                            rng = used
                            rangeAddr = rng.Address(False, False)
                        End If
                    End If
                End If

                ' 回退2：界面实时选区（用户明确框选了要筛选的区域）
                If String.IsNullOrWhiteSpace(rangeAddr) Then
                    Dim sel As Object = app.Selection
                    If TypeOf sel Is Microsoft.Office.Interop.Excel.Range Then
                        rng = DirectCast(sel, Microsoft.Office.Interop.Excel.Range)
                        rangeAddr = rng.Address(False, False)
                        If String.IsNullOrWhiteSpace(sheetName) Then
                            ws = DirectCast(rng.Worksheet, Microsoft.Office.Interop.Excel.Worksheet)
                            sheetName = ws.Name
                        End If
                    End If
                End If

                If String.IsNullOrWhiteSpace(rangeAddr) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "范围不能为空。请先用鼠标选中需要筛选的表格区域（或确保当前工作表有数据），然后重试筛选。"})
                    Return result
                End If

                ' 工作表解析
                If ws Is Nothing Then
                    If Not String.IsNullOrWhiteSpace(sheetName) Then
                        Try
                            ws = DirectCast(app.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                        Catch
                            ws = Nothing
                        End Try
                        If ws Is Nothing Then
                            result.IsError = True
                            result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: [{sheetName}]，请先创建该工作表"})
                            Return result
                        End If
                    Else
                        ws = TryCast(app.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    End If
                End If
                If ws Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "找不到工作表：" & If(sheetName, "(空，请确认选中或指定正确的工作表名称)")})
                    Return result
                End If

                If rng Is Nothing Then rng = ws.Range(rangeAddr)

                ' 解析筛选条件
                Dim criteria As Dictionary(Of String, Object)
                Try
                    criteria = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(criteriaJson)
                Catch
                    criteria = New Dictionary(Of String, Object)()
                End Try

                ' 卫生处理：先清掉工作表上可能已存在的旧自动筛选。
                ' 若不清，对"与既有筛选范围不一致"的区域再调 Range.AutoFilter 会抛
                ' 0x800A03EC「类 Range 的 _AutoFilter 方法无效」。
                Try
                    If ws.AutoFilterMode Then ws.AutoFilterMode = False
                Catch
                End Try

                Dim firstCol As Integer = rng.Column
                Dim lastCol As Integer = firstCol + rng.Columns.Count - 1
                Dim totalRows As Integer = rng.Rows.Count

                ' ===== 表头行定位 =====
                ' 实际表格常把"合并大标题"放第 1 行、真正表头在第 2 行；而 AutoFilter 要求
                ' 传入区域的首行必须是表头行，首行是合并标题时直接抛 0x800A03EC。
                ' 故在区域前若干行内选"命中条件键最多"的一行作为表头行（命中 0 时回落到区域首行）。
                Dim headerOffset As Integer = 0
                Dim headerScore As Integer = -1
                Dim scanRows As Integer = Math.Min(10, totalRows)
                For r As Integer = 0 To scanRows - 1
                    Dim score As Integer = 0
                    For Each kvpScore In criteria
                        Dim kScore As String = If(kvpScore.Key, "").Trim()
                        If kScore = "" Then Continue For
                        If ResolveHeaderColumn(rng, kScore, firstCol, lastCol, r) >= 0 Then score += 1
                    Next
                    If score > headerScore Then
                        headerScore = score
                        headerOffset = r
                    End If
                    If score >= criteria.Count Then Exit For
                Next

                ' 用"表头行 → 数据末行"重建真正用于筛选的区域（剔除表头上方的合并标题行）。
                Dim fRange As Microsoft.Office.Interop.Excel.Range = rng
                If headerOffset > 0 Then
                    Dim hr As Integer = rng.Row + headerOffset
                    Dim br As Integer = rng.Row + totalRows - 1
                    Try
                        fRange = DirectCast(ws.Range(ws.Cells(hr, firstCol), ws.Cells(br, lastCol)), Microsoft.Office.Interop.Excel.Range)
                    Catch
                        fRange = rng
                    End Try
                End If
                Dim fCols As Integer = fRange.Columns.Count

                ' ===== 条件键解析：既支持列字母(A/F)，也支持表头名(订单金额/区域) =====
                ' 修复：此前只认列字母，且 fieldIndex 直接用"绝对列号"，范围不从 A 列起时会错列。
                Dim applied As New List(Of String)()
                Dim missing As New List(Of String)()

                For Each kvp In criteria
                    Dim key As String = If(kvp.Key, "").Trim()
                    If key = "" Then Continue For
                    Dim filterValue As String = If(kvp.Value Is Nothing, "", Convert.ToString(kvp.Value))

                    Dim colNum As Integer = -1
                    ' ① 纯列字母键（A / F / AA）→ 按列字母解析（向后兼容）
                    If System.Text.RegularExpressions.Regex.IsMatch(key, "^[A-Za-z]{1,3}$") Then
                        colNum = ExcelMcpCoreUtils.ColumnLetterToNumber(key)
                    End If
                    ' ② 非列字母键，或列字母越界 → 按表头名解析（精确 > 包含 > 被包含），与定位到的表头行对齐
                    If colNum < firstCol OrElse colNum > lastCol Then
                        colNum = ResolveHeaderColumn(rng, key, firstCol, lastCol, headerOffset)
                    End If
                    If colNum < 0 Then
                        missing.Add(key)
                        Continue For
                    End If

                    ' 关键修正：AutoFilter 的 field 是"相对筛选区域起点"的序号
                    Dim fieldIndex As Integer = colNum - firstCol + 1
                    If fieldIndex < 1 OrElse fieldIndex > fCols Then
                        missing.Add(key)
                        Continue For
                    End If

                    Try
                        If String.IsNullOrWhiteSpace(filterValue) Then
                            fRange.AutoFilter(fieldIndex)
                        Else
                            fRange.AutoFilter(fieldIndex, filterValue)
                        End If
                    Catch afEx As Exception
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text =
                            $"应用筛选失败：列「{key}」条件「{filterValue}」在区域 {fRange.Address(False, False)} 上设置失败 -> {afEx.Message}" &
                            $"（诊断：表头行={rng.Row + headerOffset}，工作表受保护={CBool(ws.ProtectContents)}；" &
                            "若表头行含合并单元格，请先取消合并，或改用不含标题行的数据区域）"})
                        Utils.ComReleaseHelper.Release(rng)
                        Utils.ComReleaseHelper.Release(ws)
                        Return result
                    End Try

                    Dim headerText As String = ""
                    Try
                        headerText = Convert.ToString(DirectCast(rng.Cells(headerOffset + 1, fieldIndex), Microsoft.Office.Interop.Excel.Range).Text)
                    Catch
                        headerText = ""
                    End Try
                    applied.Add($"{headerText}({ExcelMcpCoreUtils.NumberToColumnLetter(colNum)})={filterValue}")
                Next

                ' 无任何条件命中时，至少把筛选下拉箭头挂到区域上
                If applied.Count = 0 Then
                    Try
                        fRange.AutoFilter()
                    Catch ex0 As Exception
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text =
                            $"应用筛选失败 -> {ex0.Message}（区域 {fRange.Address(False, False)}，工作表受保护={CBool(ws.ProtectContents)}）"})
                        Utils.ComReleaseHelper.Release(rng)
                        Utils.ComReleaseHelper.Release(ws)
                        Return result
                    End Try
                End If

                If Not Object.ReferenceEquals(fRange, rng) Then Utils.ComReleaseHelper.Release(fRange)
                Utils.ComReleaseHelper.Release(rng)
                Utils.ComReleaseHelper.Release(ws)

                Dim msg As String = $"成功应用筛选到范围: {rangeAddr}，共 {applied.Count} 个条件"
                If applied.Count > 0 Then msg &= "：" & String.Join("、", applied)
                If missing.Count > 0 Then msg &= $"；未找到列（已跳过）: {String.Join("、", missing)}"
                result.Content.Add(New McpTextContent With {.Text = msg})

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"应用筛选失败: {ex.Message}"})
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 把"表头名"解析成列号（用于 excel_apply_filter 支持按列名筛选，如 "订单金额"/"区域"）。
        ''' 匹配优先级：精确 > 表头包含键（"销售区域" ⊃ "区域"）> 键包含表头；全不中返回 -1。
        ''' </summary>
        Private Function ResolveHeaderColumn(rng As Microsoft.Office.Interop.Excel.Range, key As String, firstCol As Integer, lastCol As Integer, Optional rowOffset As Integer = 0) As Integer
            Dim k As String = If(key, "").Trim()
            If k = "" Then Return -1
            Dim kUpper As String = k.ToUpperInvariant()
            Dim containsCol As Integer = -1
            Dim containedCol As Integer = -1
            Dim offset As Integer = 1
            For c As Integer = firstCol To lastCol
                Dim txt As String = ""
                Try
                    txt = Convert.ToString(DirectCast(rng.Cells(rowOffset + 1, offset), Microsoft.Office.Interop.Excel.Range).Text)
                Catch
                    txt = ""
                End Try
                offset += 1
                txt = txt.Trim()
                If txt = "" Then Continue For
                Dim tUpper As String = txt.ToUpperInvariant()
                If tUpper = kUpper Then Return c
                If containsCol < 0 AndAlso tUpper.Contains(kUpper) Then containsCol = c
                If containedCol < 0 AndAlso kUpper.Contains(tUpper) Then containedCol = c
            Next
            If containsCol >= 0 Then Return containsCol
            If containedCol >= 0 Then Return containedCol
            Return -1
        End Function
        ''' <summary>
        ''' 清除指定工作表的筛选
        ''' </summary>
        Private Function ExecuteClearFilter(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()

            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")

                Dim app = Globals.ThisAddIn.Application
                If app Is Nothing OrElse app.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何工作簿"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet
                If Not String.IsNullOrEmpty(sheetName) Then
                    ws = TryCast(app.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                Else
                    ws = TryCast(app.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                End If

                If ws Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: {sheetName}"})
                    Return result
                End If

                ' 清除筛选
                If ws.AutoFilterMode Then
                    ws.AutoFilterMode = False
                End If

                ' 如果有高级筛选，也清除
                Try
                    ws.ShowAllData()
                Catch
                    ' 没有筛选时可能抛出异常，忽略
                End Try

                Utils.ComReleaseHelper.Release(ws)

                result.Content.Add(New McpTextContent With {
                    .Text = $"成功清除工作表筛选: {ws.Name}"
                })

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"清除筛选失败: {ex.Message}"})
            End Try

            Return result
        End Function
        ''' <summary>
        ''' 安全获取参数值
        ''' </summary>
        Private Function GetArg(Of T)(args As Dictionary(Of String, Object), key As String, defaultValue As T) As T
            If args Is Nothing OrElse Not args.ContainsKey(key) Then
                Return defaultValue
            End If
            Try
                Dim value = args(key)
                If value Is Nothing Then Return defaultValue
                If GetType(T) = GetType(String) Then Return CType(CObj(value.ToString()), T)
                If GetType(T) = GetType(Boolean) Then
                    If TypeOf value Is Boolean Then Return CType(value, T)
                    Dim strVal = value.ToString().Trim()
                    If strVal.Equals("true", StringComparison.OrdinalIgnoreCase) Then Return CType(CObj(True), T)
                    If strVal.Equals("false", StringComparison.OrdinalIgnoreCase) Then Return CType(CObj(False), T)
                    Return defaultValue
                End If
                Return CType(value, T)
            Catch
                Return defaultValue
            End Try
        End Function
        ''' <summary>
        ''' 计算范围结束单元格
        ''' </summary>
        Private Function GetEndCell(startCell As String, rows As Integer, cols As Integer) As String
            If rows <= 0 OrElse cols <= 0 Then Return startCell
            ' 归一化：去掉 '$'；若传入的是 "A1:B10"/"$A$1:$I$101" 这类完整区域，
            ' 取冒号前的左上角起始单元格。此前把 "$A$1:$I$101" 的字母拼成 "AI"、数字拼成 "1101"，
            ' 算出 "AK1102" 这种错误结束格，导致写入区域膨胀成 A1:AK1102（整表被旧公式残留覆盖）。
            Dim anchor = If(startCell, "").Trim().Replace("$", "")
            Dim colonIdx As Integer = anchor.IndexOf(":"c)
            If colonIdx >= 0 Then anchor = anchor.Substring(0, colonIdx)
            Dim colPart = New String(anchor.Where(Function(c) Char.IsLetter(c)).ToArray())
            Dim rowPart = New String(anchor.Where(Function(c) Char.IsDigit(c)).ToArray())

            Dim colNum = ExcelMcpCoreUtils.ColumnLetterToNumber(colPart)
            Dim rowNum = If(String.IsNullOrEmpty(rowPart), 1, Integer.Parse(rowPart))

            Dim endCol = ExcelMcpCoreUtils.NumberToColumnLetter(colNum + cols - 1)
            Dim endRow = (rowNum + rows - 1).ToString()

            Return endCol & endRow
        End Function

        ' ========== <Layer>3 - 系统与扩展</Layer> ==========

        '#End Region

    End Class

End Namespace
