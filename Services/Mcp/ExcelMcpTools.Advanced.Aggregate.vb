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
        ''' 【算】纯聚合：读取数据、按维度列分组聚合，返回结果 JSON。
        ''' 不建表、不画图 —— 落盘交给 excel_write_range，画图交给 excel_create_chart。
        ''' </summary>
        Private Function ExecuteAggregate(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            ' COM 对象跟踪，用于最终释放
            Dim comObjects As New List(Of Object)()
            Try
                ' 纯计算参数：聚合后只返回结果
                Dim dimColumn As String = GetArg(Of String)(args, "dimColumn", "")
                Dim requestedDimColumn As String = dimColumn
                Dim intentDeviated As Boolean = False
                Dim valueColumn As String = GetArg(Of String)(args, "valueColumn", "")
                Dim aggFunc As String = GetArg(Of String)(args, "aggFunc", "sum").ToLower()
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim rangeAddr As String = GetArg(Of String)(args, "range", "")
                Dim skipEmptyRows As Boolean = GetArg(Of Boolean)(args, "skipEmptyRows", False)
                ' dimColumn 可选：留空时 fallback 自动选择典型维度列（优先级：部门>地区>产品>...）
                If String.IsNullOrEmpty(dimColumn) Then
                    ' 不报错，让 fallback 逻辑（L247+）自动选最佳维度列
                End If
                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿"})
                    Return result
                End If
                Dim wb = excelApp.ActiveWorkbook
                comObjects.Add(wb)
                ' 获取源工作表
                Dim srcWs As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                If Not String.IsNullOrEmpty(sheetName) Then
                    srcWs = TryCast(wb.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                End If
                If srcWs Is Nothing Then
                    srcWs = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                End If
                comObjects.Add(srcWs)
                ' 获取源数据范围（行为与 excel_create_chart 一致：RangeContext → Selection → UsedRange）
                Dim srcRange As Microsoft.Office.Interop.Excel.Range = Nothing
                If String.IsNullOrEmpty(rangeAddr) Then
                    Dim ctxAddr As String = ""
                    If RangeContext.TryGet(srcWs.Name, ctxAddr) Then
                        Try
                            srcRange = srcWs.Range(ctxAddr)
                            If srcRange IsNot Nothing AndAlso srcRange.Rows.Count < 2 Then
                                srcRange = Nothing
                                RangeContext.Clear()
                            End If
                        Catch ex As Exception
                            srcRange = Nothing
                            RangeContext.Clear()
                        End Try
                    End If
                    If srcRange Is Nothing Then
                        Try
                            srcRange = DirectCast(excelApp.Selection, Microsoft.Office.Interop.Excel.Range)
                        Catch ex As Exception
                            srcRange = Nothing
                        End Try
                    End If
                    If srcRange IsNot Nothing AndAlso srcRange.Columns.Count < 2 Then
                        Try
                            Dim usedRange = srcWs.UsedRange
                            If usedRange IsNot Nothing AndAlso usedRange.Rows.Count > 1 AndAlso usedRange.Columns.Count >= 2 Then
                                srcRange = usedRange
                            End If
                            If usedRange IsNot Nothing Then comObjects.Add(usedRange)
                        Catch ex As Exception
                        End Try
                    End If
                Else
                    Try
                        srcRange = srcWs.Range(rangeAddr)
                    Catch ex As Exception
                        ' 回退 1：当前选区
                        Try
                            srcRange = DirectCast(excelApp.Selection, Microsoft.Office.Interop.Excel.Range)
                        Catch ex2 As Exception
                        End Try
                        ' 回退 2：UsedRange
                        If srcRange Is Nothing Then
                            Try
                                srcRange = srcWs.UsedRange
                            Catch ex3 As Exception
                            End Try
                        End If
                    End Try
                End If
                If srcRange IsNot Nothing Then comObjects.Add(srcRange)
                ' 最终安全兜底：仍无数据则强试 UsedRange
                If srcRange Is Nothing OrElse srcRange.Rows.Count < 2 Then
                    Try
                        Dim usedRange = srcWs.UsedRange
                        If usedRange IsNot Nothing AndAlso usedRange.Rows.Count > 1 AndAlso usedRange.Columns.Count >= 2 Then
                            srcRange = usedRange
                            If srcRange IsNot Nothing Then comObjects.Add(srcRange)
                        End If
                    Catch ex As Exception
                    End Try
                End If
                If srcRange Is Nothing OrElse srcRange.Rows.Count < 2 Then
                    result.IsError = True
                    Dim activeName As String = ""
                    Try
                        activeName = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet).Name
                    Catch : End Try
                    result.Content.Add(New McpTextContent With {.Text = $"未选中有效数据区域（至少需要表头+1行数据）。当前活动工作表是""{activeName}""——" &
                        If(activeName <> "", "若它是刚创建的空看板表，就是没传 sheet 参数导致在空表上聚合；", "") &
                        "请传 sheet=源数据表名（数据所在的表）后重试。"})
                    Return result
                End If
                ' 读取数据到二维数组
                Dim rowCount = srcRange.Rows.Count
                Dim colCount = srcRange.Columns.Count
                Dim rawData As Object(,) = TryCast(srcRange.Value2, Object(,))
                If rawData Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "读取数据失败"})
                    Return result
                End If
                ' 第一行是表头，找到维度列和数值列的索引
                Dim dimColIndex As Integer = -1
                Dim valColIndex As Integer = -1
                Dim headers(colCount - 1) As String
                For c = 1 To colCount
                    Dim headerVal = rawData(1, c)
                    Dim headerStr = If(headerVal IsNot Nothing, headerVal.ToString().Trim(), "")
                    headers(c - 1) = headerStr
                    If String.Equals(headerStr, dimColumn, StringComparison.OrdinalIgnoreCase) Then
                        dimColIndex = c
                    End If
                    If String.Equals(headerStr, valueColumn, StringComparison.OrdinalIgnoreCase) Then
                        valColIndex = c
                    End If
                Next
                If dimColIndex = -1 AndAlso Not String.IsNullOrEmpty(dimColumn) Then
                    For c = 1 To colCount
                        Dim headerVal = rawData(1, c)
                        Dim headerStr = If(headerVal IsNot Nothing, headerVal.ToString().Trim(), "")
                        If headerStr.IndexOf(dimColumn, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            dimColIndex = c
                            Exit For
                        End If
                    Next
                End If
                If dimColIndex = -1 AndAlso Not String.IsNullOrEmpty(dimColumn) Then
                    For c = 1 To colCount
                        Dim headerVal = rawData(1, c)
                        Dim headerStr = If(headerVal IsNot Nothing, headerVal.ToString().Trim(), "")
                        If SemanticMatch(headerStr, dimColumn) Then
                            dimColIndex = c
                            dimColumn = headerStr
                            Exit For
                        End If
                    Next
                End If
                ' 【列字母支持】维度列同数值列：字母（"B"）按范围内相对列号解析为表头名。
                ' 放在语义匹配之后：真实表头名恰为单字母时仍走名称链路。
                If dimColIndex = -1 AndAlso Not String.IsNullOrEmpty(dimColumn) AndAlso
                   System.Text.RegularExpressions.Regex.IsMatch(dimColumn, "^[A-Za-z]{1,3}$") Then
                    Dim dimLetterIdx As Integer = ExcelMcpCoreUtils.ColumnLetterToNumber(dimColumn)
                    If dimLetterIdx >= 1 AndAlso dimLetterIdx <= colCount Then
                        dimColIndex = dimLetterIdx
                        dimColumn = If(rawData(1, dimLetterIdx) IsNot Nothing, rawData(1, dimLetterIdx).ToString().Trim(), dimColumn)
                    End If
                End If
                If valColIndex = -1 AndAlso Not String.IsNullOrEmpty(valueColumn) Then
                    For c = 1 To colCount
                        Dim headerVal = rawData(1, c)
                        Dim headerStr = If(headerVal IsNot Nothing, headerVal.ToString().Trim(), "")
                        If headerStr.IndexOf(valueColumn, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            valColIndex = c
                            Exit For
                        End If
                    Next
                End If
                If valColIndex = -1 AndAlso Not String.IsNullOrEmpty(valueColumn) Then
                    For c = 1 To colCount
                        Dim headerVal = rawData(1, c)
                        Dim headerStr = If(headerVal IsNot Nothing, headerVal.ToString().Trim(), "")
                        If SemanticMatch(headerStr, valueColumn) Then
                            valColIndex = c
                            valueColumn = headerStr
                            Exit For
                        End If
                    Next
                End If
                ' 【列字母支持】弱模型高频传列字母（"D"）而非表头名（实测 2026-10-06：
                ' valueColumn=D、dimColumn=B → "未找到数值列 'D'"）。
                ' 字母按"范围内相对列号"解析成该列表头名，之后与名称链路等价。
                ' 注意放在名称三级匹配之后：真实表头名恰好是单字母时仍走名称匹配，不受影响。
                If valColIndex = -1 AndAlso Not String.IsNullOrEmpty(valueColumn) AndAlso
                   System.Text.RegularExpressions.Regex.IsMatch(valueColumn, "^[A-Za-z]{1,3}$") Then
                    Dim letterIdx As Integer = ExcelMcpCoreUtils.ColumnLetterToNumber(valueColumn)
                    If letterIdx >= 1 AndAlso letterIdx <= colCount Then
                        valColIndex = letterIdx
                        valueColumn = If(rawData(1, letterIdx) IsNot Nothing, rawData(1, letterIdx).ToString().Trim(), valueColumn)
                    End If
                End If
                ' 数值列未指定时，自动推断合适的数值列（金额/指标类优先）
                If aggFunc <> "count" AndAlso valColIndex = -1 AndAlso String.IsNullOrEmpty(valueColumn) Then
                    Dim moneyHints() As String = {"销售额", "销售", "利润", "成本", "收入", "金额", "价格", "单价", "费用", "毛利", "净利", "合计", "总额", "总计", "产值", "业绩", "revenue", "sales", "profit", "cost", "amount", "price", "income"}
                    Dim numericCols As New List(Of Integer)()
                    For c = 1 To colCount
                        If c = dimColIndex Then Continue For
                        Dim headerVal = rawData(1, c)
                        Dim headerStr = If(headerVal IsNot Nothing, headerVal.ToString().Trim(), "")
                        If headerStr.IndexOf("编号", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           headerStr.IndexOf("姓名", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           headerStr.IndexOf("名称", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           headerStr.IndexOf("日期", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           headerStr.IndexOf("时间", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           headerStr.IndexOf("备注", StringComparison.OrdinalIgnoreCase) >= 0 Then
                            Continue For
                        End If
                        Dim isNumeric As Boolean = False
                        For r = 2 To rowCount
                            Dim cellVal = rawData(r, c)
                            If cellVal IsNot Nothing Then
                                Dim d As Double = 0
                                If Double.TryParse(cellVal.ToString(), d) Then
                                    isNumeric = True
                                    Exit For
                                End If
                            End If
                        Next
                        If isNumeric Then numericCols.Add(c)
                    Next
                    Dim matchedCol As Integer = -1
                    Dim matchedName As String = ""
                    For Each c In numericCols
                        Dim headerStr = If(rawData(1, c) IsNot Nothing, rawData(1, c).ToString().Trim(), "")
                        Dim hit As Boolean = False
                        For Each hint In moneyHints
                            If headerStr.IndexOf(hint, StringComparison.OrdinalIgnoreCase) >= 0 Then
                                matchedCol = c
                                matchedName = headerStr
                                hit = True
                                Exit For
                            End If
                        Next
                        If hit Then Exit For
                    Next
                    If matchedCol = -1 AndAlso numericCols.Count > 0 Then
                        matchedCol = numericCols(0)
                        matchedName = If(rawData(1, matchedCol) IsNot Nothing, rawData(1, matchedCol).ToString().Trim(), "")
                    End If
                    if matchedCol <> -1 Then
                        valColIndex = matchedCol
                        valueColumn = matchedName
                    End If
                End If
                If dimColIndex = -1 Then
                    ' 兜底：指定维度列不存在时，按"典型性"优先级自动选最佳分类型列
                    ' 旧逻辑的 Bug：遍历 headers 遇到"部门/地区"任意关键词就 Exit For，
                    ' 如果 headers 顺序是 [日期, 地区, 部门] 会先选"地区"而错失更典型的"部门"
                    ' 新逻辑：收集所有候选后按关键词优先级打分，选分数最高的
                    Dim fallbackDimCol As Integer = -1
                    Dim fallbackDimName As String = ""
                    Dim fallbackDimScore As Integer = -1
                    Dim firstNonNumericCol As Integer = -1
                    Dim firstNonNumericName As String = ""
                    For c = 1 To colCount
                        If c = valColIndex Then Continue For
                        Dim headerVal = rawData(1, c)
                        Dim headerStr = If(headerVal IsNot Nothing, headerVal.ToString().Trim(), "")
                        If headerStr = "" Then Continue For
                        If headerStr.IndexOf("日期", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           headerStr.IndexOf("时间", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           headerStr.IndexOf("编号", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           headerStr.IndexOf("销售额", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           headerStr.IndexOf("成本", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           headerStr.IndexOf("利润", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           headerStr.IndexOf("金额", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           headerStr.IndexOf("数量", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           headerStr.IndexOf("客户数", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           headerStr.IndexOf("单价", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           headerStr.IndexOf("价格", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           headerStr.IndexOf("收入", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                           headerStr.IndexOf("备注", StringComparison.OrdinalIgnoreCase) >= 0 Then
                            Continue For
                        End If
                        If firstNonNumericCol = -1 Then
                            firstNonNumericCol = c
                            firstNonNumericName = headerStr
                        End If
                        ' 关键词优先级打分（分越高越典型，数值列不参与）。
                        ' 打分表统一收敛到 KeywordDimScore —— 与「维度择优」共用同一套口径，避免两处维护漂移。
                        Dim _score As Integer = KeywordDimScore(headerStr)
                        If _score > fallbackDimScore Then
                            fallbackDimScore = _score
                            fallbackDimCol = c
                            fallbackDimName = headerStr
                        End If
                    Next
                    If fallbackDimCol = -1 Then
                        fallbackDimCol = firstNonNumericCol
                        fallbackDimName = firstNonNumericName
                    End If
                    If fallbackDimCol <> -1 Then
                        dimColIndex = fallbackDimCol
                        ' 用户显式指定了维度列但数据中不存在，回退成其他列属"意图降级"
                        If Not String.IsNullOrEmpty(requestedDimColumn) AndAlso
                           Not String.Equals(fallbackDimName, requestedDimColumn, StringComparison.OrdinalIgnoreCase) Then
                            intentDeviated = True
                        End If
                        dimColumn = fallbackDimName
                    End If
                End If
                If dimColIndex = -1 OrElse (aggFunc <> "count" AndAlso valColIndex = -1) Then
                    ' 【可自愈报错】列出全部可用表头，Agent 看一眼即可改对参数重试
                    Dim availCols As New List(Of String)()
                    For c = 1 To colCount
                        Dim hv = rawData(1, c)
                        availCols.Add(If(hv IsNot Nothing AndAlso hv.ToString().Trim() <> "", hv.ToString().Trim(), "(第" & c & "列空)"))
                    Next
                    result.IsError = True
                    If dimColIndex = -1 Then
                        result.Content.Add(New McpTextContent With {.Text = $"未找到维度列 '{dimColumn}'。可用列（按 A={availCols(0)} 起顺序）：{String.Join("、", availCols)}。可改用精确表头名或列字母（A/B/C…）重试。"})
                    Else
                        result.Content.Add(New McpTextContent With {.Text = $"未找到数值列 '{valueColumn}'。可用列（按 A={availCols(0)} 起顺序）：{String.Join("、", availCols)}。可改用精确表头名或列字母（A/B/C…）重试。"})
                    End If
                    Return result
                End If
                ' 【维度择优·第 1 步】扫描所有分类列的取值基数（distinct 数）。
                ' 动机（实测 2026-10-28）：源表「地区」列 20 行全是"华南"，模型按"地区"聚合后
                ' 只有 1 组 → 图表只有一根柱/一个扇区，毫无信息量，而模型还会照常收尾说"已完成"。
                ' 旧逻辑只检查"列名能不能匹配上"，从不看数据分布，所以永远发现不了这种退化。
                ' 现在：① 把候选维度（含基数）回传给模型继续择优；② 结果退化为 1 组时自动换维度重算。
                Dim dataRowCount As Integer = rowCount - 1
                Dim dimCandidates As New List(Of Dictionary(Of String, Object))()
                Dim candColIdx As New List(Of Integer)()
                Dim candDistinct As New List(Of Integer)()
                Dim candScore As New List(Of Integer)()
                If dataRowCount > 0 Then
                    For c = 1 To colCount
                        Dim hdr As String = If(headers(c - 1), "")
                        If hdr = "" Then Continue For
                        If IsNonDimHeader(hdr) Then Continue For
                        If IsNumericColumn(rawData, rowCount, c) Then Continue For
                        Dim dcount As Integer = CountDistinctValues(rawData, rowCount, c)
                        If dcount <= 0 Then Continue For
                        Dim sc As Integer = ScoreDimCandidate(dcount, dataRowCount, hdr)
                        dimCandidates.Add(New Dictionary(Of String, Object) From {
                            {"column", hdr}, {"distinct", dcount}, {"score", sc}})
                        candColIdx.Add(c)
                        candDistinct.Add(dcount)
                        candScore.Add(sc)
                    Next
                End If
                ' 分组聚合（VB.NET 原生 Dictionary）
                Dim aggResults As List(Of Tuple(Of String, Double)) =
                    AggregateByDimColumn(rawData, rowCount, colCount, dimColIndex, valColIndex, aggFunc, skipEmptyRows)
                ' 【维度择优·第 2 步】退化检测：只有 1 组说明该维度列在本次数据里没有区分度，
                ' 自动改用评分最高（基数合适 + 表头典型）的另一个分类列重算，并在返回里写明换了什么。
                Dim dimSwitched As Boolean = False
                Dim dimSwitchNote As String = ""
                If aggResults.Count <= 1 Then
                    Dim bestIdx As Integer = -1
                    Dim bestScore As Integer = 0
                    For i = 0 To candColIdx.Count - 1
                        Dim ci As Integer = candColIdx(i)
                        If ci = dimColIndex OrElse ci = valColIndex Then Continue For
                        If candDistinct(i) < 2 Then Continue For
                        If candScore(i) > bestScore Then
                            bestScore = candScore(i)
                            bestIdx = ci
                        End If
                    Next
                    If bestIdx <> -1 Then
                        Dim newResults As List(Of Tuple(Of String, Double)) =
                            AggregateByDimColumn(rawData, rowCount, colCount, bestIdx, valColIndex, aggFunc, skipEmptyRows)
                        If newResults.Count > 1 Then
                            Dim oldDim As String = dimColumn
                            Dim newDim As String = If(headers(bestIdx - 1), "")
                            dimSwitchNote = $"维度列 '{oldDim}' 在本次数据中只有 {aggResults.Count} 个取值，" &
                                $"聚合结果没有区分度（图表只会有一根柱/一个扇区），已自动改用 '{newDim}'（{newResults.Count} 个取值）。" &
                                $"若确实要按 '{oldDim}' 汇总，请先让用户在源数据中补充该列的不同取值。"
                            dimColIndex = bestIdx
                            dimColumn = newDim
                            aggResults = newResults
                            dimSwitched = True
                        End If
                    End If
                End If
                ' 为前端四阶段展示提供上下文元数据（工作簿/工作表/选区/表头/预览行）
                Dim previewRows As New List(Of Dictionary(Of String, Object))()
                Dim previewRowCount As Integer = Math.Min(5, rowCount - 1)
                For r = 2 To previewRowCount + 1
                    Dim rowDict As New Dictionary(Of String, Object)()
                    For c = 1 To colCount
                        rowDict(headers(c - 1)) = rawData(r, c)
                    Next
                    previewRows.Add(rowDict)
                Next
                Dim displayValueColumn = If(String.IsNullOrEmpty(valueColumn), "数量", valueColumn)
                Dim dataList As New List(Of Dictionary(Of String, Object))()
                For i = 0 To aggResults.Count - 1
                    Dim dict As New Dictionary(Of String, Object)()
                    dict(dimColumn) = aggResults(i).Item1
                    dict(displayValueColumn) = aggResults(i).Item2
                    dataList.Add(dict)
                Next
                Dim metricsList As New List(Of Object)()
                metricsList.Add(New With {
                    .name = displayValueColumn,
                    .valueColumn = valueColumn,
                    .aggFunc = aggFunc
                })
                ' 可选落地：调用方传入 targetSheet 时写回聚合结果到该表（存在则覆盖，不存在则新建）
                Dim landedSheetName As String = Nothing
                Dim targetTs = GetArg(Of String)(args, "targetSheet", "")
                If Not String.IsNullOrWhiteSpace(targetTs) Then
                    Dim rowLists As New List(Of List(Of Object))()
                    Dim hdrList As New List(Of Object) From {dimColumn, displayValueColumn}
                    rowLists.Add(hdrList)
                    For Each d In dataList
                        Dim rl As New List(Of Object)()
                        If d.ContainsKey(dimColumn) Then rl.Add(d(dimColumn)) Else rl.Add("")
                        If d.ContainsKey(displayValueColumn) Then rl.Add(d(displayValueColumn)) Else rl.Add("")
                        rowLists.Add(rl)
                    Next
                    landedSheetName = WriteResultToTargetSheet(args, rowLists, "聚合结果")
                End If
                Dim output As New Dictionary(Of String, Object)()
                output("success") = True
                output("aggregated") = True
                output("workbook") = wb.Name
                output("worksheet") = srcWs.Name
                output("range") = srcRange.Address
                output("rows") = rowCount
                output("columns") = colCount
                output("headers") = headers
                output("previewRows") = previewRows
                output("dimColumn") = dimColumn
                output("valueColumn") = displayValueColumn
                output("aggFunc") = aggFunc
                output("groupCount") = aggResults.Count
                If intentDeviated Then
                    ' 用户意图与数据不匹配（指定维度列在数据中不存在，已降级为其他列），显式告知
                    output("intentDeviated") = True
                    output("requestedDimColumn") = requestedDimColumn
                    output("actualDimColumn") = dimColumn
                    output("intentNote") = $"数据中不存在维度列 '{requestedDimColumn}'，已按 '{dimColumn}' 进行聚合展示。如需按 '{requestedDimColumn}' 展示，请补充该列数据或说明对应关系。"
                End If
                If dimSwitched Then
                    ' 维度列存在但没有区分度（如"地区"全表只有 1 个取值）→ 已自动换成有区分度的列
                    output("dimSwitched") = True
                    output("requestedDimColumn") = requestedDimColumn
                    output("actualDimColumn") = dimColumn
                    output("dimSwitchNote") = dimSwitchNote
                End If
                ' 候选维度（分类列名 + 取值基数 + 评分）回传：模型据此继续择优（换维度/多图对比），
                ' 不必再盲猜列名。
                If dimCandidates.Count > 0 Then output("dimCandidates") = dimCandidates
                If landedSheetName IsNot Nothing Then output("writtenTo") = landedSheetName
                output("data") = dataList
                output("metrics") = metricsList
                Dim jsonResult = JsonConvert.SerializeObject(output, Formatting.Indented)
                result.Content.Add(New McpTextContent With {
                    .Text = jsonResult
                })
                ' ── 结构化产物摘要（通用框架层 Artifact）──
                ' 补 labels/values 供前端图表渲染直接使用，不再解析 JSON
                Dim labels As New List(Of Object)()
                Dim values As New List(Of Object)()
                For Each d In dataList
                    If d.ContainsKey(dimColumn) Then labels.Add(d(dimColumn)) Else labels.Add("")
                    If d.ContainsKey(valueColumn) Then values.Add(d(valueColumn)) Else values.Add(0)
                Next
                Dim aggArtifact = New Dictionary(Of String, Object) From {
                    {"type", "aggregate"},
                    {"dimColumn", dimColumn},
                    {"valueColumn", valueColumn},
                    {"aggFunc", aggFunc},
                    {"groupCount", aggResults.Count},
                    {"targetSheet", If(landedSheetName, "")},
                    {"labels", labels},
                    {"values", values},
                    {"dimSwitched", dimSwitched},
                    {"requestedDimColumn", If(dimSwitched, requestedDimColumn, "")},
                    {"dataRange", If(landedSheetName, "") & "!A1:B" & (aggResults.Count + 1)}
                }
                If dimCandidates.Count > 0 Then aggArtifact("dimCandidates") = dimCandidates
                result.Artifact = aggArtifact
                System.Diagnostics.Debug.WriteLine($"[ExecuteAggregate] 聚合完成: {aggResults.Count} 组, 维度={dimColumn}, 数值={valueColumn}, 函数={aggFunc}")
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"聚合失败: {ex.Message}"})
                System.Diagnostics.Debug.WriteLine($"[ExecuteAggregate] 异常: {ex.ToString()}")
            Finally
                ' 释放所有 COM 对象
                For Each obj In comObjects
                    Try
                        Utils.ComReleaseHelper.Release(obj)
                    Catch
                    End Try
                Next
            End Try
            Return result
        End Function

        ''' <summary>
        ''' 表头"典型维度"关键词打分（分越高越像分组维度；0 = 不像维度，如"销售员""备注"）。
        ''' 该打分表被「维度列兜底选择」与「维度择优评分」共用，避免两处维护漂移。
        ''' </summary>
        Private Function KeywordDimScore(header As String) As Integer
            If String.IsNullOrEmpty(header) Then Return 0
            If header.IndexOf("部门", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 100
            If header.IndexOf("地区", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 90
            If header.IndexOf("区域", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 88
            If header.IndexOf("大区", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 87
            If header.IndexOf("产品", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 85
            If header.IndexOf("商品", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 84
            If header.IndexOf("品类", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 83
            If header.IndexOf("渠道", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 80
            If header.IndexOf("来源", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 79
            If header.IndexOf("分类", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 75
            If header.IndexOf("类别", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 74
            If header.IndexOf("类型", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 73
            If header.IndexOf("城市", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 70
            If header.IndexOf("省份", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 68
            If header.IndexOf("国家", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 65
            If header.IndexOf("客户", StringComparison.OrdinalIgnoreCase) >= 0 Then Return 60
            Return 0
        End Function

        ''' <summary>明确不适合做分组维度的列：度量列、日期时间、编号类、自由文本。</summary>
        Private Function IsNonDimHeader(header As String) As Boolean
            If String.IsNullOrEmpty(header) Then Return True
            Dim keys As String() = {"日期", "时间", "年份", "年度", "月份", "编号", "序号", "编码", "id",
                                    "销售额", "成本", "利润", "金额", "数量", "客户数", "单价", "价格", "收入", "业绩",
                                    "备注", "说明", "地址", "电话", "邮箱"}
            For Each k As String In keys
                If header.IndexOf(k, StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
            Next
            Return False
        End Function

        ''' <summary>该列是否数值列（Excel 序列日期本质是 Double，会被这里一并排除出维度候选）。</summary>
        Private Function IsNumericColumn(rawData As Object(,), rowCount As Integer, colIdx As Integer) As Boolean
            Dim seen As Integer = 0
            Dim numeric As Integer = 0
            For r As Integer = 2 To rowCount
                Dim v = rawData(r, colIdx)
                If v Is Nothing Then Continue For
                Dim s As String = v.ToString().Trim()
                If s = "" Then Continue For
                seen += 1
                Dim d As Double = 0
                If Double.TryParse(s, d) Then numeric += 1
            Next
            Return seen > 0 AndAlso numeric >= CInt(Math.Ceiling(seen * 0.8))
        End Function

        ''' <summary>统计某列的非空取值个数（distinct），达到 cap 提前返回以免大表无谓开销。</summary>
        Private Function CountDistinctValues(rawData As Object(,), rowCount As Integer, colIdx As Integer, Optional cap As Integer = 500) As Integer
            Dim seen As New HashSet(Of String)()
            For r As Integer = 2 To rowCount
                Dim v = rawData(r, colIdx)
                If v Is Nothing Then Continue For
                Dim s As String = v.ToString().Trim()
                If s = "" Then Continue For
                seen.Add(s)
                If seen.Count >= cap Then Return seen.Count
            Next
            Return seen.Count
        End Function

        ''' <summary>
        ''' 维度候选综合评分 = 基数分（权重 2）+ 表头典型性分（权重 1）。
        ''' 只看表头关键词会选出"地区"这种全表只有 1 个取值的列（实测 2026-10-28：20 行全是"华南"，
        ''' 聚合退化成 1 组、图表只有一根柱），所以基数优先级高于关键词。
        ''' -1 表示不可用作维度（distinct &lt;= 1）。
        ''' </summary>
        Private Function ScoreDimCandidate(distinctCount As Integer, dataRowCount As Integer, header As String) As Integer
            If dataRowCount <= 0 OrElse distinctCount <= 1 Then Return -1
            Dim baseScore As Integer = 0
            If distinctCount >= 2 AndAlso distinctCount <= 20 Then
                baseScore = 90
            ElseIf distinctCount <= 50 Then
                baseScore = 60
            Else
                baseScore = 30
            End If
            ' 每行一个唯一值 = 明细列（姓名/流水号），做维度会产出几十上百根柱
            If distinctCount >= dataRowCount Then baseScore -= 20
            If baseScore < 0 Then baseScore = 0
            Return baseScore * 2 + KeywordDimScore(header)
        End Function

        ''' <summary>
        ''' 按维度列分组聚合。抽成方法是因为「维度择优」在结果退化时需要换列重算一次。
        ''' </summary>
        Private Function AggregateByDimColumn(rawData As Object(,), rowCount As Integer, colCount As Integer,
                                              dimColIdx As Integer, valColIdx As Integer,
                                              aggFunc As String, skipEmptyRows As Boolean) As List(Of Tuple(Of String, Double))
            Dim groupOrder As New List(Of String)()
            Dim groupData As New Dictionary(Of String, List(Of Double))()
            Dim isCountNoValue As Boolean = (aggFunc = "count" AndAlso valColIdx < 1)
            For r As Integer = 2 To rowCount
                If skipEmptyRows Then
                    Dim allEmpty As Boolean = True
                    For c As Integer = 1 To colCount
                        Dim cv = rawData(r, c)
                        If cv IsNot Nothing AndAlso Not (TypeOf cv Is String AndAlso DirectCast(cv, String).Trim() = "") Then
                            allEmpty = False
                            Exit For
                        End If
                    Next
                    If allEmpty Then Continue For
                End If
                Dim dimVal = rawData(r, dimColIdx)
                Dim dimStr = If(dimVal IsNot Nothing, dimVal.ToString().Trim(), "(空白)")
                If Not groupData.ContainsKey(dimStr) Then
                    groupData(dimStr) = New List(Of Double)()
                    groupOrder.Add(dimStr)
                End If
                If isCountNoValue Then
                    groupData(dimStr).Add(1.0)
                Else
                    Dim valNum As Double = 0
                    If valColIdx >= 1 AndAlso rawData(r, valColIdx) IsNot Nothing Then
                        Double.TryParse(rawData(r, valColIdx).ToString(), valNum)
                    End If
                    groupData(dimStr).Add(valNum)
                End If
            Next
            Dim out As New List(Of Tuple(Of String, Double))()
            For Each key In groupOrder
                Dim vals = groupData(key)
                Dim aggVal As Double = 0
                Select Case aggFunc
                    Case "sum" : For Each v In vals : aggVal += v : Next
                    Case "avg"
                        If vals.Count > 0 Then
                            Dim s As Double = 0
                            For Each v In vals : s += v : Next
                            aggVal = s / vals.Count
                        End If
                    Case "count" : aggVal = vals.Count
                    Case "max"
                        If vals.Count > 0 Then
                            aggVal = vals(0)
                            For Each v In vals
                                If v > aggVal Then aggVal = v
                            Next
                        End If
                    Case "min"
                        If vals.Count > 0 Then
                            aggVal = vals(0)
                            For Each v In vals
                                If v < aggVal Then aggVal = v
                            Next
                        End If
                    Case Else : For Each v In vals : aggVal += v : Next
                End Select
                out.Add(Tuple.Create(key, Math.Round(aggVal, 2)))
            Next
            Return out
        End Function

        Private Function CleanColumnName(name As String) As String
            If String.IsNullOrEmpty(name) Then Return ""
            Dim result = name.Trim()
            result = System.Text.RegularExpressions.Regex.Replace(result, "\s*\(.*?\)\s*", "")
            result = System.Text.RegularExpressions.Regex.Replace(result, "\s*\[.*?\]\s*", "")
            result = System.Text.RegularExpressions.Regex.Replace(result, "\s*【.*?】\s*", "")
            result = System.Text.RegularExpressions.Regex.Replace(result, "\s*元\s*$", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "\s*%?\s*$", "")
            result = System.Text.RegularExpressions.Regex.Replace(result, "\s+", "")
            Return result.ToLower()
        End Function

        Private Function SemanticMatch(name1 As String, name2 As String) As Boolean
            If String.IsNullOrEmpty(name1) OrElse String.IsNullOrEmpty(name2) Then Return False
            Dim clean1 = CleanColumnName(name1)
            Dim clean2 = CleanColumnName(name2)
            If clean1 = clean2 Then Return True
            If clean1.IndexOf(clean2) >= 0 Then Return True
            If clean2.IndexOf(clean1) >= 0 Then Return True
            Dim synonyms As New Dictionary(Of String, List(Of String)) From {
                {"销售额", New List(Of String) From {"月度业绩", "业绩", "销售收入", "营业额", "收入", "销售额"}},
                {"业绩", New List(Of String) From {"月度业绩", "业绩", "销售收入", "营业额", "收入", "销售额"}},
                {"提成", New List(Of String) From {"提成", "基础提成", "最终实发提成", "实发提成"}},
                {"部门", New List(Of String) From {"部门", "科室", "组", "小组", "团队", "所属部门", "部门名称"}},
                {"员工", New List(Of String) From {"员工", "员工姓名", "姓名", "人员", "职工", "业务员"}},
                {"数量", New List(Of String) From {"数量", "人数", "个数"}},
                {"金额", New List(Of String) From {"金额", "数额", "数值"}},
                {"地区", New List(Of String) From {"地区", "区域", "大区", "地域", "省市", "省份", "城市"}},
                {"日期", New List(Of String) From {"日期", "时间", "月份", "季度", "年份", "年度", "年月", "月", "年", "周"}},
                {"产品", New List(Of String) From {"产品", "商品", "sku", "品类", "类别", "分类", "产品分类", "产品类别", "商品分类", "商品类别"}},
                {"渠道", New List(Of String) From {"渠道", "方式", "途径", "来源"}}
            }
            For Each kvp In synonyms
                Dim list = kvp.Value
                If list.Any(Function(s) String.Equals(s.ToLower(), clean1, StringComparison.OrdinalIgnoreCase)) AndAlso
                   list.Any(Function(s) String.Equals(s.ToLower(), clean2, StringComparison.OrdinalIgnoreCase)) Then
                    Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' 【算】日期分组：将日期列按 年/季度/月/周/日 粒度分组计算，返回结果 JSON。
        ''' 默认仅返回分组数据，不写回 Excel；传入 targetSheet 时把结果落地到该表（不存在则新建）。
        ''' aggColumn 提供时按组对数值列求和并输出汇总表；不填则给每一行追加一个分组列。
        ''' </summary>
        Private Function ExecuteDateGroup(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Dim excelApp = Globals.ThisAddIn.Application
            Try
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿"})
                    Return result
                End If
                Dim wb = excelApp.ActiveWorkbook

                Dim dateColumn As String = GetArg(Of String)(args, "dateColumn", "")
                Dim groupBy As String = GetArg(Of String)(args, "groupBy", "month").ToLower()
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim rangeAddr As String = GetArg(Of String)(args, "range", "")
                Dim weekStart As String = GetArg(Of String)(args, "weekStart", "monday").ToLower()
                Dim fiscalYearStart As Integer = GetArg(Of Integer)(args, "fiscalYearStart", 1)
                Dim fiscalQuarter As Boolean = GetArg(Of Boolean)(args, "fiscalQuarter", False)
                Dim aggColumn As String = GetArg(Of String)(args, "aggColumn", "")
                Dim outputColumn As String = GetArg(Of String)(args, "outputColumn", "")
                Dim targetSheet As String = GetArg(Of String)(args, "targetSheet", "")

                If String.IsNullOrEmpty(dateColumn) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "必须指定 dateColumn 参数"})
                    Return result
                End If
                If Not (groupBy = "year" OrElse groupBy = "quarter" OrElse groupBy = "month" OrElse groupBy = "week" OrElse groupBy = "day") Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"不支持的 groupBy 粒度: {groupBy}（支持 year/quarter/month/week/day）"})
                    Return result
                End If

                ' 解析源工作表与范围（显式 range → RangeContext → Selection → UsedRange）
                Dim srcWs As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                If Not String.IsNullOrEmpty(sheetName) Then srcWs = TryCast(wb.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                If srcWs Is Nothing Then srcWs = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)

                Dim srcRange As Microsoft.Office.Interop.Excel.Range = Nothing
                If String.IsNullOrEmpty(rangeAddr) Then
                    Dim ctxAddr As String = ""
                    If RangeContext.TryGet(srcWs.Name, ctxAddr) Then
                        Try : srcRange = srcWs.Range(ctxAddr) : Catch : srcRange = Nothing : End Try
                    End If
                    If srcRange Is Nothing Then
                        Try : srcRange = DirectCast(excelApp.Selection, Microsoft.Office.Interop.Excel.Range) : Catch : srcRange = Nothing : End Try
                    End If
                    If srcRange Is Nothing OrElse srcRange.Rows.Count < 2 Then
                        Try : srcRange = srcWs.UsedRange : Catch : srcRange = Nothing : End Try
                    End If
                Else
                    Try : srcRange = srcWs.Range(rangeAddr) : Catch : srcRange = Nothing : End Try
                End If
                If srcRange Is Nothing OrElse srcRange.Rows.Count < 2 Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "未选中有效数据区域（至少需要表头+1行数据）"})
                    Return result
                End If

                Dim fmts As String() = Nothing
                Dim arr = ReadRangeTo0BasedWithFormats(srcRange, fmts)
                ' 裁剪 UsedRange 幽灵空行/空列，避免 104 万行的幽灵区域被当成真实数据
                arr = TrimPhantomRanges(arr)
                If arr Is Nothing OrElse arr.GetLength(0) < 2 Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "读取数据失败"})
                    Return result
                End If
                Dim rows = arr.GetLength(0)
                Dim cols = arr.GetLength(1)

                ' 解析表头，定位日期列与聚合列（精确 → 模糊包含）
                Dim dateColIdx As Integer = -1
                Dim aggColIdx As Integer = -1
                Dim headers As String() = New String(cols - 1) {}
                For c = 0 To cols - 1
                    Dim h = arr(0, c)
                    headers(c) = If(h IsNot Nothing, h.ToString().Trim(), "")
                    If String.Equals(headers(c), dateColumn, StringComparison.OrdinalIgnoreCase) Then dateColIdx = c
                    If Not String.IsNullOrEmpty(aggColumn) AndAlso String.Equals(headers(c), aggColumn, StringComparison.OrdinalIgnoreCase) Then aggColIdx = c
                Next
                If dateColIdx = -1 Then
                    For c = 0 To cols - 1
                        If headers(c).IndexOf(dateColumn, StringComparison.OrdinalIgnoreCase) >= 0 Then dateColIdx = c : Exit For
                    Next
                End If
                If aggColIdx = -1 AndAlso Not String.IsNullOrEmpty(aggColumn) Then
                    For c = 0 To cols - 1
                        If headers(c).IndexOf(aggColumn, StringComparison.OrdinalIgnoreCase) >= 0 Then aggColIdx = c : Exit For
                    Next
                End If
                If dateColIdx = -1 Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"在表头中找不到日期列: {dateColumn}"})
                    Return result
                End If
                If Not String.IsNullOrEmpty(aggColumn) AndAlso aggColIdx = -1 Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"在表头中找不到聚合列: {aggColumn}"})
                    Return result
                End If

                ' 默认分组列名
                If String.IsNullOrEmpty(outputColumn) Then
                    Select Case groupBy
                        Case "year" : outputColumn = "年份"
                        Case "quarter" : outputColumn = "季度"
                        Case "month" : outputColumn = "月份"
                        Case "week" : outputColumn = "周"
                        Case Else : outputColumn = "日期"
                    End Select
                End If

                Dim weekMonday As Boolean = (weekStart = "monday")
                Dim sums As New Dictionary(Of String, Double)()
                Dim counts As New Dictionary(Of String, Integer)()
                Dim groupedRows As New List(Of List(Of Object))()
                Dim orderOfGroups As New List(Of String)()

                For r = 1 To rows - 1
                    Dim dateVal = arr(r, dateColIdx)
                    Dim dt As DateTime
                    If Not TryConvertToDate(dateVal, dt) Then Continue For
                    Dim grpLabel = BuildDateGroupLabel(dt, groupBy, weekMonday, fiscalYearStart, fiscalQuarter)

                    If aggColIdx >= 0 Then
                        Dim v = arr(r, aggColIdx)
                        Dim num As Double = 0
                        If v IsNot Nothing AndAlso Double.TryParse(v.ToString(), num) Then
                            If sums.ContainsKey(grpLabel) Then sums(grpLabel) += num Else sums(grpLabel) = num
                            If counts.ContainsKey(grpLabel) Then counts(grpLabel) += 1 Else counts(grpLabel) = 1
                        End If
                        If Not orderOfGroups.Contains(grpLabel) Then orderOfGroups.Add(grpLabel)
                    Else
                        Dim rowOut As New List(Of Object)()
                        For c = 0 To cols - 1
                            rowOut.Add(NormalizeCell(arr(r, c)))
                        Next
                        rowOut.Add(grpLabel)
                        groupedRows.Add(rowOut)
                        If Not orderOfGroups.Contains(grpLabel) Then orderOfGroups.Add(grpLabel)
                    End If
                Next

                If orderOfGroups.Count = 0 Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "没有可分组的有效日期数据"})
                    Return result
                End If

                ' 组装输出（第一行是表头）
                Dim outRows As New List(Of List(Of Object))()
                If aggColIdx >= 0 Then
                    outRows.Add(New List(Of Object) From {outputColumn, "数量", "合计"})
                    For Each g In orderOfGroups
                        outRows.Add(New List(Of Object) From {g, counts(g), sums(g)})
                    Next
                Else
                    Dim hdrRow = New List(Of Object)()
                    For c = 0 To cols - 1 : hdrRow.Add(headers(c)) : Next
                    hdrRow.Add(outputColumn)
                    outRows.Add(hdrRow)
                    outRows.AddRange(groupedRows)
                End If

                ' 可选落地：调用方传入 targetSheet 时写回分组结果到该表
                Dim landedSheetName As String = Nothing
                If Not String.IsNullOrWhiteSpace(targetSheet) Then
                    landedSheetName = WriteResultToTargetSheet(args, outRows, "日期分组结果")
                End If

                Dim outData As New List(Of Object)()
                For Each r In outRows : outData.Add(r.ToArray()) : Next

                Dim output As New Dictionary(Of String, Object)()
                output("success") = True
                output("dateColumn") = dateColumn
                output("groupBy") = groupBy
                output("outputColumn") = outputColumn
                output("groupCount") = orderOfGroups.Count
                output("aggregated") = (aggColIdx >= 0)
                If landedSheetName IsNot Nothing Then output("writtenTo") = landedSheetName
                output("data") = outData
                result.Content.Add(New McpTextContent With {.Text = JsonConvert.SerializeObject(output, Formatting.Indented)})

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"日期分组失败: {ex.Message}"})
            End Try
            Return result
        End Function

        ''' <summary>
        ''' 尝试把单元格值解析为日期：支持 DateTime、Excel 序列号、常见日期字符串格式。
        ''' </summary>
        Private Function TryConvertToDate(v As Object, ByRef dt As DateTime) As Boolean
            If v Is Nothing Then Return False
            If TypeOf v Is DateTime Then
                dt = DirectCast(v, DateTime)
                Return True
            End If
            Dim s = v.ToString()
            If String.IsNullOrWhiteSpace(s) Then Return False
            Dim dbl As Double
            If Double.TryParse(s, dbl) AndAlso dbl > 25569 AndAlso dbl < 50000 Then
                Try
                    dt = DateTime.FromOADate(dbl)
                    Return True
                Catch
                End Try
            End If
            Dim fmts = New String() {"yyyy-MM-dd", "yyyy/M/d", "yyyy-M-d", "yyyy年M月d日", "yyyy-MM-dd HH:mm:ss", "yyyy/M/d H:mm", "yyyyMMdd", "MM-dd", "M月d日"}
            For Each f In fmts
                If DateTime.TryParseExact(s, f, Nothing, Globalization.DateTimeStyles.None, dt) Then Return True
            Next
            If DateTime.TryParse(s, dt) Then Return True
            Return False
        End Function

        ''' <summary>
        ''' 根据粒度生成分组标签；支持财年（yearStart 起始月份）与周起始日。
        ''' </summary>
        Private Function BuildDateGroupLabel(dt As DateTime, groupBy As String, weekMonday As Boolean, fiscalYearStart As Integer, fiscalQuarter As Boolean) As String
            Select Case groupBy
                Case "year"
                    Return GetFiscalYear(dt, fiscalYearStart).ToString()
                Case "quarter"
                    Dim fy = GetFiscalYear(dt, fiscalYearStart)
                    Dim q = GetQuarter(dt, fiscalYearStart)
                    Return If(fiscalQuarter, $"{fy} Q{q}", $"{dt.Year} 第{GetQuarter(dt, 1)}季度")
                Case "month"
                    Return If(fiscalQuarter, GetFiscalYear(dt, fiscalYearStart).ToString() & "-" & dt.Month.ToString("00"), dt.Year.ToString() & "-" & dt.Month.ToString("00"))
                Case "week"
                    Return FirstDayOfWeek(dt, weekMonday).ToString("yyyy-MM-dd") & " 所在周"
                Case Else
                    Return dt.ToString("yyyy-MM-dd")
            End Select
        End Function

        Private Function GetFiscalYear(dt As DateTime, fiscalStart As Integer) As Integer
            If fiscalStart <= 1 Then Return dt.Year
            If dt.Month < fiscalStart Then Return dt.Year - 1
            Return dt.Year
        End Function

        Private Function GetQuarter(dt As DateTime, fiscalStart As Integer) As Integer
            Dim offset = fiscalStart - 1
            Dim adjustedMonth = (dt.Month - offset + 11) Mod 12 + 1
            Return (adjustedMonth - 1) \ 3 + 1
        End Function

        Private Function FirstDayOfWeek(dt As DateTime, monday As Boolean) As DateTime
            Dim dow = CInt(dt.DayOfWeek)   ' Sunday=0 ... Saturday=6
            Dim delta As Integer
            If monday Then
                delta = (dow - 1 + 7) Mod 7
            Else
                delta = dow
            End If
            Return dt.AddDays(-delta).Date
        End Function

    End Class

End Namespace