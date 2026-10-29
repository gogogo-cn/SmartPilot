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

        ' ═══════════════════════════════════════════════════════════════
        ' 模块级共享：自动图表网格分配器
        ' 当用户未显式传 position 时，按 2 列 N 行固定网格分配统一尺寸槽位，
        ' 避免 LLM 不传 position 导致每张图尺寸不一、位置错开。
        ' 每次 Excel 会话内累加；遇到显式传 position 的调用跳过（不占槽）。
        ' ═══════════════════════════════════════════════════════════════
        Private Shared _autoChartSlotIndex As Integer = 0

        ''' <summary>
        ''' 2 列网格 → 槽号 → (left, top, width, height) 磅坐标。
        ''' 槽尺寸统一：chartSlotW/chartSlotH；列间距/行间距 gap。
        ''' 基准：gridLeft=20, gridTop=200（数据区下方起点）。
        ''' </summary>
        Private Shared Function _GridCoords(slot As Integer,
                                             Optional cellW As Double = 80.0,
                                             Optional cellH As Double = 20.0) As (left As Double, top As Double, width As Double, height As Double)
            ' 统一槽位物理尺寸（磅）：2列均分 + 固定高，所有自动图对齐
            Const _colSpacing As Double = 12.0    ' 两列之间磅间距
            Const _rowSpacing As Double = 20.0    ' 两行之间磅间距
            Const _gridLeft As Double = 20.0      ' 网格起点左（磅）
            Const _gridTop As Double = 200.0      ' 网格起点上（磅，避开数据区/标题区）
            Const _slotHeight As Double = 300.0    ' 槽高（磅，统一后所有图同高）

            Dim slotRow = Math.Floor(slot / 2)
            Dim slotCol = slot Mod 2

            ' 槽宽根据默认列宽估算（80 磅 ≈ 标准 Excel 列宽），2 列 + 间距
            ' 用 cellW × 估算列数，保持总宽在 760 磅 ≈ 10" 左右
            Const _colsPerSlot As Double = 8.0   ' 每槽占 8 列（看板模式 chartSlotCols=8）
            Dim slotW = Math.Max(250.0, cellW * _colsPerSlot)
            slotW = Math.Min(slotW, 500.0)  ' 上限 500 磅，避免超宽

            Dim left = _gridLeft + slotCol * (slotW + _colSpacing)
            Dim top = _gridTop + slotRow * (_slotHeight + _rowSpacing)
            Return (left, top, slotW, _slotHeight)
        End Function

        Private Function ExecuteCreateChart(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                ' 1) 解析 17 个参数（含 T3 新增：useSecondaryAxisLastSeries）
                Dim chartType As String = GetArg(Of String)(args, "chartType", "xlColumnClustered")
                Dim dataRange As String = GetArg(Of String)(args, "dataRange", "")
                ' dataValues（可选）：二维数组 JSON，如 [["部门","销售额"],["华东",1200],["华北",980]]
                ' 提供时作为 dataRange 的内存替代——用数组直接绑定 Series（绕开 SetSourceData 的 Range 依赖），
                ' 避免前端聚合结果被迫落地成"图表数据源/聚合结果"工作表。首行为表头，其余行为明细。
                Dim dataValuesJson As String = GetArg(Of String)(args, "dataValues", "")
                Dim useMemoryData As Boolean = Not String.IsNullOrWhiteSpace(dataValuesJson)
                Dim memRows As New List(Of Object())()
                If useMemoryData Then
                    Try
                        Dim ja = Newtonsoft.Json.Linq.JArray.Parse(dataValuesJson)
                        For Each row As Newtonsoft.Json.Linq.JArray In ja
                            Dim arr = New Object(Math.Max(0, row.Count - 1)) {}
                            For i As Integer = 0 To row.Count - 1
                                arr(i) = DirectCast(row(i), Newtonsoft.Json.Linq.JToken).ToObject(Of Object)()
                            Next
                            memRows.Add(arr)
                        Next
                        If memRows.Count < 2 Then
                            result.IsError = True
                            result.Content.Add(New McpTextContent With {.Text = "❌ dataValues 至少需要 2 行（1 行表头 + 1 行数据）"})
                            Return result
                        End If
                    Catch ex As Exception
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = $"❌ dataValues 数组格式错误：{ex.Message}"})
                        Return result
                    End Try
                End If
                Dim title As String = GetArg(Of String)(args, "title", "")
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim positionJson As String = GetArg(Of String)(args, "position", "")
                Dim seriesBy As String = GetArg(Of String)(args, "seriesBy", "columns")
                Dim switchRowColumn As Boolean = GetArg(Of Boolean)(args, "switchRowColumn", False)
                Dim hasLegend As Boolean = GetArg(Of Boolean)(args, "hasLegend", True)
                Dim legendPosition As String = GetArg(Of String)(args, "legendPosition", "xlLegendPositionRight")
                Dim xAxisTitle As String = GetArg(Of String)(args, "xAxisTitle", "")
                Dim yAxisTitle As String = GetArg(Of String)(args, "yAxisTitle", "")
                Dim showDataLabels As Boolean = GetArg(Of Boolean)(args, "showDataLabels", False)
                Dim dataLabelPosition As String = GetArg(Of String)(args, "dataLabelPosition", "")
                Dim chartStyle As Integer = GetArg(Of Integer)(args, "chartStyle", 0)
                Dim useSecondaryAxisLastSeries As Boolean = GetArg(Of Boolean)(args, "useSecondaryAxisLastSeries", False)
                Dim trendLineSeriesIndex As Integer = GetArg(Of Integer)(args, "trendLineSeriesIndex", 0)
                Dim trendLineType As String = GetArg(Of String)(args, "trendLineType", "")
                ' 看板/多图模式下（前端注入 multiChart=true）保留多个图表，不做清理；
                ' 其余默认单图请求，建图前自清理历史 AI 图，避免重复出图。
                Dim multiChart As Boolean = GetArg(Of Boolean)(args, "multiChart", False)
                ' 收敛说明（追加到成功文案里，让用户知道系统替他做了哪些判断）
                Dim seriesConvergedNote As String = ""
                Dim labelConvergedNote As String = ""

                ' 透视图（PivotChart）：useSlicer 模式下前端只传 pivotSheet/pivotTable，
                ' 不传 dataRange/dataValues，图表由后续 usePivot 分支绑定透视表创建。
                ' 必须在此处提前解析并豁免下面的 dataRange 判空，否则会在绑定透视表之前就被拒绝。
                Dim pivotSheetP As String = GetArg(Of String)(args, "pivotSheet", "")
                Dim pivotTableP As String = GetArg(Of String)(args, "pivotTable", "")
                Dim usePivotRequested As Boolean = (Not String.IsNullOrWhiteSpace(pivotSheetP)) AndAlso (Not String.IsNullOrWhiteSpace(pivotTableP))

                If Not useMemoryData AndAlso String.IsNullOrEmpty(dataRange) AndAlso Not usePivotRequested Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "❌ dataRange（数据范围）不能为空，或改传 dataValues（内存数组）"})
                    Return result
                End If

                Dim excelApp As Microsoft.Office.Interop.Excel.Application = DirectCast(Globals.ThisAddIn.Application, Microsoft.Office.Interop.Excel.Application)
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                ' 所有 Worksheet/Range 对象使用无歧义的 Interop 完全限定名
                ' 过滤无效 sheetName（undefined/null/...），即使传入也当作空
                Dim invalidSheets As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
                    "undefined", "null", "[object Object]", "none", "nan", ""
                }
                If invalidSheets.Contains(sheetName) Then sheetName = ""

                Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
                Try : wb = CType(excelApp.ActiveWorkbook, Microsoft.Office.Interop.Excel.Workbook) : Catch : wb = Nothing : End Try
                If wb Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "❌ 无法获取活动工作簿，请重新打开或激活 Excel 窗口重试。"})
                    Return result
                End If

                ' 【根本修复】不用 Dim ws As ...（任何 Worksheet 声明），
                ' 改为直接使用 ResolveRangeSafe（内部用 Interop.Worksheet，但这是 Shared 函数，命名空间解析不会被当前方法上下文污染）
                ' 但其他地方仍然需要 ws 做 ChartObjects.Add()，这里使用 TryCast + Object 中转：
                Dim wsObj As Object = Nothing
                If Not String.IsNullOrEmpty(sheetName) Then
                    Try : wsObj = wb.Worksheets(sheetName) : Catch : wsObj = Nothing : End Try
                End If
                If wsObj Is Nothing Then
                    Try : wsObj = excelApp.ActiveSheet : Catch : wsObj = Nothing : End Try
                End If
                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = TryCast(wsObj, Microsoft.Office.Interop.Excel.Worksheet)
                If ws Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {
                        .Text = $"❌ 指定的工作表、单元格范围或名称不存在。请确认工作簿中存在对应的工作表（sheetName=""{sheetName}""），或先切换到正确的工作表再重试。"
                    })
                    Return result
                End If
                Try : System.Diagnostics.Debug.WriteLine("[CreateChart] sheetName=" & sheetName & " ws.Name=" & ws.Name) : Catch : End Try
                
                ' 2) 类型解析：20 种；未知返回错误（不做静默降级）。
                ' auto/自动/未指定 → 启发式自动选型（基于数据形状选合理默认），
                ' 避免"生成图表"类泛需求因 LLM 输出 auto 而失败（原子工具应有默认决策路径）。
                Dim rawChartType As String = (If(chartType, "")).Trim()
                If rawChartType = "" OrElse _autoChartTypeTokens.Contains(rawChartType) Then
                    chartType = ResolveAutoChartType(memRows, dataRange, sheetName, excelApp)
                    Try : System.Diagnostics.Debug.WriteLine("[CreateChart] auto 选型 → " & chartType) : Catch : End Try
                End If
                Dim targetType As Microsoft.Office.Interop.Excel.XlChartType
                If Not TryParseChartTypeV2(chartType, targetType) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {
                        .Text = "❌ 不支持的图表类型：" & chartType & vbCrLf &
                                "支持的 20 种：xlColumnClustered / xlColumnStacked / xlColumnStacked100 / " &
                                "xlBarClustered / xlBarStacked / xlBarStacked100 / xlLine / xlLineMarkers / xlArea / " &
                                "xlPie / xlDoughnut / xlXYScatter / xlRadar / xlRadarFilled / xlBubble / " &
                                "xlSunburst / xlTreemap / xlHistogram / xlBoxwhisker / xlFunnel / xlWaterfall"
                    })
                    Utils.ComReleaseHelper.Release(ws)
                    Return result
                End If

                ' 3) 位置解析：JSON 优先（anchorRange 按单元格区域锚定，优先级最高）；否则自动在数据下方并错位，避免重叠
                If Not String.IsNullOrWhiteSpace(positionJson) Then
                    Debug.WriteLine($"[ExecuteCreateChart] position 原文(len={Len(positionJson)}): {positionJson.Substring(0, Math.Min(positionJson.Length, 160))}")
                End If
                Dim posLeft As Double = 0, posTop As Double = 0, posWidth As Double = 500, posHeight As Double = 350
                If Not String.IsNullOrWhiteSpace(positionJson) Then
                    Try
                        Dim posJ = Newtonsoft.Json.Linq.JObject.Parse(positionJson)
                        If posJ("anchorRange") IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(posJ("anchorRange").ToString()) Then
                            ' 看板布局：按单元格区域锚定（如 "B5:G18" 或 "看板!E9:L17"），
                            ' 直接用已验证的 ws 对象解析（L118 已拿到 wb.Worksheets(sheetName)），
                            ' 不走 ResolveRangeSafe —— 避免它对带 ! 前缀地址的二次解析干扰。
                            Dim anchorRaw As String = posJ("anchorRange").ToString().Trim()
                            Dim localAddr As String = anchorRaw
                            Dim exclIdx = anchorRaw.LastIndexOf("!"c)
                            If exclIdx > 0 Then localAddr = anchorRaw.Substring(exclIdx + 1).Trim()
                            Try : System.Diagnostics.Debug.WriteLine($"[CreateChart] anchorRange 原文=""{anchorRaw}"" localAddr=""{localAddr}"" ws.Name=""{ws.Name}""") : Catch : End Try

                            Dim anchorRng As Microsoft.Office.Interop.Excel.Range = Nothing
                            Try
                                anchorRng = ws.Range(localAddr)
                            Catch ex As Exception
                                Try : System.Diagnostics.Debug.WriteLine($"[CreateChart] ws.Range(""{localAddr}"") 失败: {ex.Message}") : Catch : End Try
                            End Try

                            If anchorRng IsNot Nothing Then
                                posLeft = CDbl(anchorRng.Left)
                                posTop = CDbl(anchorRng.Top)
                                posWidth = Math.Max(150, CDbl(anchorRng.Width))
                                posHeight = Math.Max(220, CDbl(anchorRng.Height))
                                Try : System.Diagnostics.Debug.WriteLine($"[CreateChart] anchor 解析成功 → Left={posLeft},Top={posTop},W={posWidth},H={posHeight}") : Catch : End Try
                                Utils.ComReleaseHelper.Release(anchorRng)
                            Else
                                Try : System.Diagnostics.Debug.WriteLine($"[CreateChart] ⚠️ anchor 解析失败，坐标归零（堆叠风险！）") : Catch : End Try
                            End If
                        Else
                            If posJ("left") IsNot Nothing Then posLeft = CDbl(posJ("left"))
                            If posJ("top") IsNot Nothing Then posTop = CDbl(posJ("top"))
                            If posJ("width") IsNot Nothing Then posWidth = Math.Max(150, CDbl(posJ("width")))
                            If posJ("height") IsNot Nothing Then posHeight = Math.Max(220, CDbl(posJ("height")))
                        End If
                    Catch ex As Exception
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = $"❌ position JSON 错误：{ex.Message}（正确示例：{{""anchorRange"":""B5:G18""}} 或 {{""left"":100,""top"":80,""width"":600,""height"":400}}）"})
                        Utils.ComReleaseHelper.Release(ws)
                        Return result
                    End Try
                Else
                    ' ── 自动定位：2 列 N 行固定网格（统一槽位，所有自动图强制对齐）──
                    ' 1) 先看当前 ws 里已有的 ChartObjects 数量 → 从下一个槽号开始，
                    '    避免和已有图重叠（Excel 会话内可能有残留图）
                    Dim chartObjsPre As Microsoft.Office.Interop.Excel.ChartObjects = Nothing
                    Try : chartObjsPre = DirectCast(ws.ChartObjects, Microsoft.Office.Interop.Excel.ChartObjects) : Catch : End Try
                    Dim existingCount As Integer = 0
                    If chartObjsPre IsNot Nothing Then
                        Try : existingCount = CInt(chartObjsPre.Count) : Catch : existingCount = 0 : End Try
                        Utils.ComReleaseHelper.Release(chartObjsPre)
                    End If

                    ' 2) 槽号 = max(会话内计数器, 当前工作表已有图数) — 取较大的那个，
                    '    保证既不在同一 ws 内撞槽，也不会因为多 ws 切换而重置计数器
                    Dim slot As Integer = Math.Max(_autoChartSlotIndex, existingCount)

                    ' 3) 从 dataRange 取参考列宽 + 数据区底部（用于单图紧贴数据下方）
                    Dim refCellW As Double = 80.0
                    Dim dataBottom As Double = 0.0
                    Dim dataRng As Microsoft.Office.Interop.Excel.Range = Nothing
                    If Not (useMemoryData AndAlso String.IsNullOrEmpty(dataRange)) Then
                        ' dataRange 带表前缀时前缀优先（数据源表 ≠ 图表宿主表，如看板模式）
                        Dim dataSheetArg As String = sheetName
                        If Not String.IsNullOrEmpty(dataRange) AndAlso dataRange.Contains("!"c) Then dataSheetArg = ""
                        Try : dataRng = ResolveRangeSafe(excelApp, dataRange, dataSheetArg) : Catch : End Try
                    End If
                    If dataRng IsNot Nothing Then
                        Dim firstCol As Microsoft.Office.Interop.Excel.Range = Nothing
                        Try
                            firstCol = DirectCast(dataRng.Columns(1), Microsoft.Office.Interop.Excel.Range)
                            If firstCol IsNot Nothing Then refCellW = CDbl(firstCol.Width)
                            dataBottom = CDbl(dataRng.Top) + CDbl(dataRng.Height)
                        Catch : End Try
                        If firstCol IsNot Nothing Then Utils.ComReleaseHelper.Release(firstCol)
                        Utils.ComReleaseHelper.Release(dataRng)
                    End If

                    ' 4) 单图请求（已自清理历史图）：强制回退 0 号槽并贴紧数据下方，
                    '    避免会话内 _autoChartSlotIndex 累加导致图表越排越偏/越靠下。
                    If Not multiChart Then
                        slot = 0
                        _autoChartSlotIndex = 0
                    End If

                    ' 5) 计算统一网格坐标
                    Dim gc = _GridCoords(slot, refCellW)
                    posLeft = gc.left
                    posTop = gc.top
                    posWidth = gc.width
                    posHeight = gc.height

                    ' 单图：顶部紧贴数据区下方（无数据时回落到网格基准 200）
                    If Not multiChart AndAlso dataBottom > 0 Then
                        posTop = dataBottom + 16.0
                    End If

                    ' 6) 会话内槽号递增 — 仅多图/看板需要往后一格
                    If multiChart Then _autoChartSlotIndex = slot + 1

                    Try : System.Diagnostics.Debug.WriteLine($"[CreateChart] auto-grid slot={slot} multiChart={multiChart} → ({posLeft:F0},{posTop:F0}) {posWidth:F0}x{posHeight:F0}") : Catch : End Try
                End If

                ' 4) 单图请求自清理：建图前只清理【AI 自己生成】的图表（名称以 AIChart_ 开头），
                ' 覆盖"LLM 重复调用 / 跨表残留 / 历史测试遗留"导致的重复出图。看板(multiChart=true)跳过。
                '
                ' ⚠️【误删修复·重要】旧实现在这里无条件删除【当前工作表上的全部 ChartObjects】——
                ' 不看图表名，用户手工创建或此前已存在的图表（"图表 1" 等）会被一并静默删掉，
                ' 这正是"让 AI 出张图，我原来的图没了"的直接来源（破坏性且不可撤销）。
                ' 本工具新建的图表统一命名为 AIChart_<ticks>（见下方 chartObj.Name 赋值），
                ' 因此只按 AIChart_ 前缀清理即可：防重复出图的效果完全不变，同时绝不碰用户自己的图表。
                ' 与下方跨工作簿清理保持同一口径（两边都只认 AIChart_ 前缀）。
                If Not multiChart Then
                    Try
                        Dim cos As Microsoft.Office.Interop.Excel.ChartObjects = DirectCast(ws.ChartObjects, Microsoft.Office.Interop.Excel.ChartObjects)
                        If cos IsNot Nothing Then
                            Dim del1 As New List(Of Microsoft.Office.Interop.Excel.ChartObject)()
                            For Each co As Microsoft.Office.Interop.Excel.ChartObject In cos
                                Try
                                    If co.Name IsNot Nothing AndAlso co.Name.StartsWith("AIChart_") Then del1.Add(co)
                                Catch : End Try
                            Next
                            For Each co In del1
                                Try : co.Delete() : Catch : End Try
                            Next
                            Utils.ComReleaseHelper.Release(cos)
                        End If
                    Catch : End Try
                    Try
                        For Each sh As Microsoft.Office.Interop.Excel.Worksheet In wb.Worksheets
                            Try
                                Dim cos2 As Microsoft.Office.Interop.Excel.ChartObjects = DirectCast(sh.ChartObjects, Microsoft.Office.Interop.Excel.ChartObjects)
                                If cos2 IsNot Nothing Then
                                    Dim del2 As New List(Of Microsoft.Office.Interop.Excel.ChartObject)()
                                    For Each co As Microsoft.Office.Interop.Excel.ChartObject In cos2
                                        Try
                                            If co.Name IsNot Nothing AndAlso co.Name.StartsWith("AIChart_") Then del2.Add(co)
                                        Catch : End Try
                                    Next
                                    For Each co In del2
                                        Try : co.Delete() : Catch : End Try
                                    Next
                                    Utils.ComReleaseHelper.Release(cos2)
                                End If
                            Catch : End Try
                        Next
                    Catch : End Try
                End If

                ' 4) 创建图表 + 数据源绑定
                Dim chartObjsForAdd As Microsoft.Office.Interop.Excel.ChartObjects = DirectCast(ws.ChartObjects, Microsoft.Office.Interop.Excel.ChartObjects)
                If chartObjsForAdd Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "❌ 无法获取工作表 ChartObjects 集合，可能是工作表受保护。请先取消工作表保护后重试。"})
                    Utils.ComReleaseHelper.Release(ws)
                    Return result
                End If
                Dim chartObj As Microsoft.Office.Interop.Excel.ChartObject = Nothing
                Try
                    chartObj = chartObjsForAdd.Add(posLeft, posTop, posWidth, posHeight)
                Catch ex As Exception
                    Utils.ComReleaseHelper.Release(chartObjsForAdd)
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"❌ 无法在工作表上创建 ChartObject（位置 Left={posLeft},Top={posTop},W={posWidth},H={posHeight}）：{ex.Message}"})
                    Utils.ComReleaseHelper.Release(ws)
                    Return result
                End Try
                Utils.ComReleaseHelper.Release(chartObjsForAdd)
                If chartObj Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "❌ 创建 ChartObject 返回空，创建图表失败。"})
                    Utils.ComReleaseHelper.Release(ws)
                    Return result
                End If
                ' 统一命名为 AIChart_<ticks>，便于后续单图请求跨工作簿识别并清理历史 AI 图。
                Try : chartObj.Name = "AIChart_" & DateTime.Now.Ticks.ToString() : Catch : End Try

                Dim chart As Microsoft.Office.Interop.Excel.Chart = Nothing
                Try : chart = chartObj.Chart : Catch : End Try
                If chart Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "❌ 新建 ChartObject 中未包含 Chart 对象，创建图表失败。"})
                    Utils.ComReleaseHelper.Release(chartObj)
                    Utils.ComReleaseHelper.Release(ws)
                    Return result
                End If

                '# PERSPECTIVE CHART: 图表直接由透视表创建（PivotChart）。
                ' 透视表/PivotCache 由此前 excel_create_slicer 生成，切片器已挂到全部透视表；
                ' 透视图绑定透视表后天然随切片器筛选联动，避免"独立图 + 写死地址指空位"。
                ' pivotSheetP/pivotTableP 已在参数解析段提前解析（供 dataRange 判空豁免透视图调用）
                Dim usePivot As Boolean = usePivotRequested
                Dim pivotBound As Boolean = False
                If usePivot Then
                    Dim ptw As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                    Try : ptw = TryCast(wb.Worksheets(pivotSheetP), Microsoft.Office.Interop.Excel.Worksheet) : Catch : ptw = Nothing : End Try
                    Dim pt As Microsoft.Office.Interop.Excel.PivotTable = Nothing
                    If ptw IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(pivotTableP) Then
                        ' PIA PivotTables 索引器虽然有 String 重载，但不同 Excel 版本对大小写敏感行为不一致。
                        ' 与 SheetOps 透视表查找统一：遍历 + OrdinalIgnoreCase 匹配。
                        Try
                            Dim _pts As Microsoft.Office.Interop.Excel.PivotTables = DirectCast(ptw.PivotTables, Microsoft.Office.Interop.Excel.PivotTables)
                            For _pi As Integer = 1 To CInt(_pts.Count)
                                Dim _cand As Microsoft.Office.Interop.Excel.PivotTable = DirectCast(_pts(_pi), Microsoft.Office.Interop.Excel.PivotTable)
                                Try
                                    If _cand IsNot Nothing AndAlso _cand.Name IsNot Nothing AndAlso
                                       String.Equals(CStr(_cand.Name), pivotTableP, StringComparison.OrdinalIgnoreCase) Then
                                        pt = _cand
                                        Exit For
                                    End If
                                Finally
                                    If _cand IsNot Nothing AndAlso _cand IsNot pt Then Utils.ComReleaseHelper.Release(_cand)
                                End Try
                            Next
                            Utils.ComReleaseHelper.Release(_pts)
                        Catch
                        End Try
                    End If
                    If pt IsNot Nothing Then
                        Dim ptRng As Microsoft.Office.Interop.Excel.Range = Nothing
                        Try : ptRng = pt.TableRange1 : Catch : ptRng = Nothing : End Try
                        If ptRng IsNot Nothing Then
                            ' 无论透视图是否成功，dataRange 都回填为透视表真实输出区，杜绝图表指向空位
                            Try
                                dataRange = "'" & ptw.Name & "'!" & ptRng.Address(True, True)
                            Catch
                            End Try
                            Dim plotByP As Microsoft.Office.Interop.Excel.XlRowCol = If(String.Equals(seriesBy, "rows", StringComparison.OrdinalIgnoreCase),
                                                                                      Microsoft.Office.Interop.Excel.XlRowCol.xlRows,
                                                                                      Microsoft.Office.Interop.Excel.XlRowCol.xlColumns)
                            Try
                                chartObj.Delete()
                                Utils.ComReleaseHelper.Release(chartObj)
                                chartObj = Nothing
                                Utils.ComReleaseHelper.Release(chart)
                                chart = Nothing
                                Dim tempChartSheet As Microsoft.Office.Interop.Excel.Chart = DirectCast(excelApp.Charts.Add(), Microsoft.Office.Interop.Excel.Chart)
                                tempChartSheet.ChartType = targetType
                                tempChartSheet.SetSourceData(Source:=ptRng, PlotBy:=plotByP)
                                Dim _pl As Object = Nothing
                                Try : _pl = tempChartSheet.PivotLayout : Catch : _pl = Nothing : End Try
                                If _pl Is Nothing Then
                                    Throw New InvalidOperationException("Charts.Add + SetSourceData 未生成 PivotLayout，回退到普通图")
                                End If
                                Try : tempChartSheet.ShowReportFilterFieldButtons = False : Catch : End Try
                                Try : tempChartSheet.ShowLegendFieldButtons = False : Catch : End Try
                                Try : tempChartSheet.ShowAxisFieldButtons = False : Catch : End Try
                                Try : tempChartSheet.ShowValueFieldButtons = False : Catch : End Try
                                tempChartSheet.Location(Microsoft.Office.Interop.Excel.XlChartLocation.xlLocationAsObject, ws.Name)
                                Dim cos As Microsoft.Office.Interop.Excel.ChartObjects = DirectCast(ws.ChartObjects, Microsoft.Office.Interop.Excel.ChartObjects)
                                chartObj = DirectCast(cos(cos.Count), Microsoft.Office.Interop.Excel.ChartObject)
                                chart = chartObj.Chart
                                ' Location 后 Excel 会重置位置/大小为默认值，需显式恢复到指定坐标
                                Try : chartObj.Left = posLeft : Catch : End Try
                                Try : chartObj.Top = posTop : Catch : End Try
                                Try : chartObj.Width = posWidth : Catch : End Try
                                Try : chartObj.Height = posHeight : Catch : End Try
                                ' Location 后字段按钮可见性可能重置，需重新隐藏
                                Try : chart.ShowReportFilterFieldButtons = False : Catch : End Try
                                Try : chart.ShowLegendFieldButtons = False : Catch : End Try
                                Try : chart.ShowAxisFieldButtons = False : Catch : End Try
                                Try : chart.ShowValueFieldButtons = False : Catch : End Try
                                Try : chart.ShowAllFieldButtons = False : Catch : End Try

                                ' ── 彻底清理 PivotChart 默认 UI 垃圾 ──
                                ' Option Strict On 不允许 COM 后期绑定，这里用 CType 强制转为具体类型，
                                ' COM 集合的 Item/Count 需要通过 _obj 变量绕过 Strict 检查。
                                Try
                                    Dim _plFinal As Microsoft.Office.Interop.Excel.PivotLayout = Nothing
                                    Try : _plFinal = CType(chart.PivotLayout, Microsoft.Office.Interop.Excel.PivotLayout) : Catch : _plFinal = Nothing : End Try
                                    If _plFinal IsNot Nothing Then
                                        ' (a) 清理 DataFields Caption——去掉 "求和项:" 等前缀
                                        Try
                                            Dim _dfObj As Object = _plFinal.DataFields
                                            If _dfObj IsNot Nothing Then
                                                Dim _dfCnt As Integer = CInt(_dfObj.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, Nothing, _dfObj, Nothing))
                                                For i As Integer = 1 To _dfCnt
                                                    Try
                                                        Dim _df As Object = _dfObj.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, Nothing, _dfObj, New Object() {i})
                                                        Dim _cap As String = String.Empty
                                                        Try : _cap = CStr(_df.GetType().InvokeMember("Caption", System.Reflection.BindingFlags.GetProperty, Nothing, _df, Nothing)) : Catch : End Try
                                                        If Not String.IsNullOrEmpty(_cap) Then
                                                            Dim _cleaned = System.Text.RegularExpressions.Regex.Replace(_cap, "^(求和项[:：]\s*|平均项[:：]\s*|计数项[:：]\s*|最大值项[:：]\s*|最小值项[:：]\s*|数值项[:：]\s*|Sum of\s+|Avg of\s+|Count of\s+|Max of\s+|Min of\s+)", "", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                                                            If _cleaned <> _cap Then
                                                                Try : _df.GetType().InvokeMember("Caption", System.Reflection.BindingFlags.SetProperty, Nothing, _df, New Object() {_cleaned}) : Catch : End Try
                                                            End If
                                                        End If
                                                    Catch : End Try
                                                Next
                                            End If
                                        Catch : End Try

                                        ' (b) 隐藏副标题
                                        Try : chart.GetType().InvokeMember("HasSubtitle", System.Reflection.BindingFlags.SetProperty, Nothing, chart, New Object() {False}) : Catch : End Try

                                        ' (c) Location 后重新隐藏所有字段按钮
                                        Try : chart.GetType().InvokeMember("ShowAllFieldButtons", System.Reflection.BindingFlags.SetProperty, Nothing, chart, New Object() {False}) : Catch : End Try
                                        Try : _plFinal.GetType().InvokeMember("ShowReportFilterFieldButtons", System.Reflection.BindingFlags.SetProperty, Nothing, _plFinal, New Object() {False}) : Catch : End Try
                                        Try : _plFinal.GetType().InvokeMember("ShowLegendFieldButtons", System.Reflection.BindingFlags.SetProperty, Nothing, _plFinal, New Object() {False}) : Catch : End Try
                                        Try : _plFinal.GetType().InvokeMember("ShowAxisFieldButtons", System.Reflection.BindingFlags.SetProperty, Nothing, _plFinal, New Object() {False}) : Catch : End Try
                                        Try : _plFinal.GetType().InvokeMember("ShowValueFieldButtons", System.Reflection.BindingFlags.SetProperty, Nothing, _plFinal, New Object() {False}) : Catch : End Try

                                        ' (d) 隐藏 PivotField 的小计/汇总行
                                        Try
                                            Dim _pfObj As Object = _plFinal.PivotFields
                                            If _pfObj IsNot Nothing Then
                                                Dim _pfCnt As Integer = CInt(_pfObj.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, Nothing, _pfObj, Nothing))
                                                For i As Integer = 1 To _pfCnt
                                                    Try
                                                        Dim _pf As Object = _pfObj.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, Nothing, _pfObj, New Object() {i})
                                                        Try : _pf.GetType().InvokeMember("Subtotals", System.Reflection.BindingFlags.SetProperty, Nothing, _pf, New Object() {New Object() {0}}) : Catch : End Try
                                                    Catch : End Try
                                                Next
                                            End If
                                        Catch : End Try
                                    End If
                                Catch : End Try
                                pivotBound = True
                            Catch exPivot As Exception
                                pivotBound = False
                            End Try
                        End If
                    End If
                End If

                Dim srcRng As Microsoft.Office.Interop.Excel.Range = Nothing
                ' 内存数组模式：跳过 range 解析与 SetSourceData，改由内存数组直接绑定 Series
                If Not useMemoryData Then
                    Try
                        ' 【看板模式关键修复】数据源与图表宿主分离：
                        ' dataRange 带表前缀（如 "聚合结果!A1:B6"）时前缀优先，
                        ' 图表宿主由 sheet 参数（"看板"）决定。
                        ' 此前 sheet 参数覆盖前缀 → "sheet=看板 + dataRange=聚合结果!A1:B6"
                        ' 被解析成读看板表 A1:B6（空）→ "没有可绘制的数值列" → 图表创建失败。
                        Dim dataSheetArg As String = sheetName
                        If Not String.IsNullOrEmpty(dataRange) AndAlso dataRange.Contains("!"c) Then dataSheetArg = ""
                        srcRng = ResolveRangeSafe(excelApp, dataRange, dataSheetArg)
                    Catch ex As Exception
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = $"❌ 无法解析数据范围 dataRange=""{dataRange}""：{ex.Message}"})
                        Utils.ComReleaseHelper.Release(chart)
                        Utils.ComReleaseHelper.Release(chartObj)
                        Utils.ComReleaseHelper.Release(ws)
                        Return result
                    End Try
                    If srcRng Is Nothing Then
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {
                            .Text = $"❌ 指定的工作表、单元格范围或名称不存在。请确认单元格地址格式正确（如 A1:B10），工作簿中存在对应的工作表，或先切换到正确的工作表再重试。（dataRange=""{dataRange}""，当前工作表=""{If(ws IsNot Nothing, ws.Name, "空")}""）"
                        })
                        Utils.ComReleaseHelper.Release(chart)
                        Utils.ComReleaseHelper.Release(chartObj)
                        Utils.ComReleaseHelper.Release(ws)
                        Return result
                    End If
                End If

                chart.ChartType = targetType
                Dim plotBy As Microsoft.Office.Interop.Excel.XlRowCol = If(String.Equals(seriesBy, "rows", StringComparison.OrdinalIgnoreCase),
                                          Microsoft.Office.Interop.Excel.XlRowCol.xlRows,
                                          Microsoft.Office.Interop.Excel.XlRowCol.xlColumns)

                If useMemoryData Then
                    ' 【内存数组绑定】数据在内存（memRows 二维数组），不落地工作表。
                    ' 首行=表头，其余=明细。第1列为分类（XValues），其余数值列各为一个系列（Values）。
                    ' 用 Series.XValues/Values 直接赋内存数组，绕开 SetSourceData 的 Range 依赖。
                    chart.ChartType = targetType
                    Try
                        Dim header = memRows(0)
                        Dim dataRowCount = memRows.Count - 1
                        Dim nCols = header.Length
                        If nCols < 2 Then
                            Throw New InvalidOperationException("dataValues 至少需要 2 列（1 列分类 + 1 列数值）")
                        End If
                        ' 分类数组 = 第 1 列（跳过表头）
                        Dim cats As Object() = New Object(dataRowCount - 1) {}
                        For r As Integer = 0 To dataRowCount - 1
                            cats(r) = memRows(r + 1)(0)
                        Next
                        Dim seriesCols As New List(Of Integer)()
                        For c As Integer = 1 To nCols - 1
                            ' 只有数值列才绑为系列（聚合结果常为 分类|数值 两列）
                            Dim isNum As Boolean = True
                            Dim anyNum As Boolean = False
                            For r As Integer = 0 To dataRowCount - 1
                                Dim v = memRows(r + 1)(c)
                                If v Is Nothing Then Continue For
                                Dim d As Double = 0
                                If Double.TryParse(v.ToString(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, d) Then
                                    anyNum = True
                                Else
                                    isNum = False
                                End If
                            Next
                            If isNum AndAlso anyNum Then seriesCols.Add(c)
                        Next
                        If seriesCols.Count = 0 Then
                            ' 全无数值列（纯文本）→ 退化为第 2 列（若存在）参与绘制，交由诚实校验兜底
                            If nCols >= 2 Then seriesCols.Add(1)
                        End If
                        ' 先清空默认空系列，再逐个 NewSeries
                        Dim sCols As Microsoft.Office.Interop.Excel.SeriesCollection = Nothing
                        Try : sCols = DirectCast(chart.SeriesCollection, Microsoft.Office.Interop.Excel.SeriesCollection) : Catch : sCols = Nothing : End Try
                        If sCols IsNot Nothing Then
                            While CInt(sCols.Count) > 0
                                Try
                                    Dim sTmp As Microsoft.Office.Interop.Excel.Series = DirectCast(sCols(1), Microsoft.Office.Interop.Excel.Series)
                                    sTmp.Delete()
                                    Utils.ComReleaseHelper.Release(sTmp)
                                Catch
                                    Exit While
                                End Try
                            End While
                            For Each c As Integer In seriesCols
                                Dim sNew As Microsoft.Office.Interop.Excel.Series = Nothing
                                Try : sNew = DirectCast(sCols.NewSeries(), Microsoft.Office.Interop.Excel.Series) : Catch : sNew = Nothing : End Try
                                If sNew IsNot Nothing Then
                                    sNew.ChartType = targetType
                                    Try
                                        sNew.XValues = cats
                                    Catch
                                    End Try
                                    Dim vals As Object() = New Object(dataRowCount - 1) {}
                                    For r As Integer = 0 To dataRowCount - 1
                                        vals(r) = memRows(r + 1)(c)
                                    Next
                                    Try
                                        sNew.Values = vals
                                    Catch
                                    End Try
                                    ' 系列名取该列表头
                                    Try
                                        If header(c) IsNot Nothing Then sNew.Name = header(c).ToString()
                                    Catch
                                    End Try
                                    Utils.ComReleaseHelper.Release(sNew)
                                End If
                            Next
                            Utils.ComReleaseHelper.Release(sCols)
                        End If
                    Catch ex As Exception
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = $"❌ 内存数组绑定图表失败：{ex.Message}"})
                        Utils.ComReleaseHelper.Release(chart)
                        Utils.ComReleaseHelper.Release(chartObj)
                        Utils.ComReleaseHelper.Release(ws)
                        Return result
                    End Try
                Else
                    If Not pivotBound Then
                        Try
                            chart.SetSourceData(Source:=srcRng, PlotBy:=plotBy)
                        Catch ex As Exception
                            result.IsError = True
                            result.Content.Add(New McpTextContent With {.Text = $"❌ 绑定数据源到图表失败（范围={dataRange}）：{ex.Message}"})
                            Utils.ComReleaseHelper.Release(srcRng)
                            Utils.ComReleaseHelper.Release(chart)
                            Utils.ComReleaseHelper.Release(chartObj)
                            Utils.ComReleaseHelper.Release(ws)
                            Return result
                        End Try
                    End If
                    If switchRowColumn Then
                        If CInt(chart.PlotBy) = CInt(Microsoft.Office.Interop.Excel.XlRowCol.xlColumns) Then
                            chart.PlotBy = Microsoft.Office.Interop.Excel.XlRowCol.xlRows
                        Else
                            chart.PlotBy = Microsoft.Office.Interop.Excel.XlRowCol.xlColumns
                        End If
                    End If
                End If

                ' 【多列混合表修复】range 模式下 SetSourceData 会把源区域每一列都绑成一个系列：
                ' 文本列（地区/部门/产品/销售员/客户类型…）被当成数值系列，画出来是"文本当数值"的
                ' 乱码数据标签（实测：对整表 A1:I101 建柱状图，X 轴=日期、系列=全部 9 列、标签糊成一团）。
                ' 修复：逐系列检查取值，全是文本（无一个数值）的系列直接删除，只保留数值系列；
                ' 分类列（首列）不受影响——它由 Excel 自动作为各系列的 XValues。
                If srcRng IsNot Nothing AndAlso pivotBound = False AndAlso useMemoryData = False AndAlso
                   CInt(srcRng.Columns.Count) > 2 Then
                    Try
                        Dim scFix As Microsoft.Office.Interop.Excel.SeriesCollection = Nothing
                        Try : scFix = DirectCast(chart.SeriesCollection, Microsoft.Office.Interop.Excel.SeriesCollection) : Catch : scFix = Nothing : End Try
                        If scFix IsNot Nothing Then
                            Dim removedSeries As Integer = 0
                            For si As Integer = CInt(scFix.Count) To 1 Step -1
                                Dim sFix As Microsoft.Office.Interop.Excel.Series = Nothing
                                Try : sFix = DirectCast(scFix.Item(si), Microsoft.Office.Interop.Excel.Series) : Catch : sFix = Nothing : End Try
                                If sFix Is Nothing Then Continue For
                                Dim numericCount As Integer = 0
                                Dim textCount As Integer = 0
                                Try
                                    Dim vv As Object = sFix.Values
                                    If TypeOf vv Is Array Then
                                        For Each cell In DirectCast(vv, Array)
                                            If cell Is Nothing Then Continue For
                                            If TypeOf cell Is Double OrElse TypeOf cell Is Integer OrElse
                                               TypeOf cell Is Long OrElse TypeOf cell Is Decimal OrElse TypeOf cell Is Single Then
                                                numericCount += 1
                                            ElseIf TypeOf cell Is String AndAlso Not String.IsNullOrWhiteSpace(CStr(cell)) Then
                                                textCount += 1
                                            End If
                                        Next
                                    End If
                                Catch
                                End Try
                                If numericCount = 0 AndAlso textCount > 0 Then
                                    Try
                                        sFix.Delete()
                                        removedSeries += 1
                                    Catch
                                    End Try
                                End If
                                Utils.ComReleaseHelper.Release(sFix)
                            Next
                            Utils.ComReleaseHelper.Release(scFix)
                            If removedSeries > 0 Then
                                seriesConvergedNote = $"（已忽略 {removedSeries} 个非数值列：地区/部门/产品等文本列不参与绘图）"
                            End If
                        End If
                    Catch
                    End Try
                End If

                ' 【单系列图表双列修复】饼图/环形图数据源为"分类|数值"两列时，
                ' SetSourceData(PlotBy:=xlColumns) 在默认柱形图上生成"每列一个系列"，
                ' 切到饼图后只显示系列1（文本分类列）→ 空壳图 → 下方校验误报"没有可绘制的数值列"。
                ' （实测：聚合结果表 部门|销售额 两列即触发；频次统计表 类别|数量 同理。）
                ' 修复：第1列数据为文本（分类语义）且第2列含数值时，显式指派
                ' 第1列为分类(XValues)、第2列为系列值(Values)，系列名取第2列表头，并清理多余系列。
                If srcRng IsNot Nothing AndAlso pivotBound = False AndAlso CInt(srcRng.Columns.Count) >= 2 AndAlso
                   (CInt(targetType) = CInt(Microsoft.Office.Interop.Excel.XlChartType.xlPie) OrElse
                    CInt(targetType) = CInt(Microsoft.Office.Interop.Excel.XlChartType.xlDoughnut)) Then
                    Try
                        Dim col1 As Microsoft.Office.Interop.Excel.Range = Nothing
                        Dim col2 As Microsoft.Office.Interop.Excel.Range = Nothing
                        Try : col1 = DirectCast(srcRng.Columns(1), Microsoft.Office.Interop.Excel.Range) : Catch : col1 = Nothing : End Try
                        Try : col2 = DirectCast(srcRng.Columns(2), Microsoft.Office.Interop.Excel.Range) : Catch : col2 = Nothing : End Try
                        If col1 IsNot Nothing AndAlso col2 IsNot Nothing Then
                            Dim col1DataNumeric As Boolean = False
                            Dim col2DataNumeric As Boolean = False
                            Try
                                Dim c1vals As Object = col1.Value2
                                Dim arr1 = TryCast(c1vals, Array)
                                If arr1 IsNot Nothing Then
                                    For Each v As Object In arr1
                                        If v Is Nothing OrElse Convert.IsDBNull(v) Then Continue For
                                        Dim d As Double = 0
                                        If Double.TryParse(v.ToString(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, d) Then
                                            col1DataNumeric = True
                                            Exit For
                                        End If
                                    Next
                                End If
                            Catch : End Try
                            Try
                                Dim c2vals As Object = col2.Value2
                                Dim arr2 = TryCast(c2vals, Array)
                                If arr2 IsNot Nothing Then
                                    For Each v As Object In arr2
                                        If v Is Nothing OrElse Convert.IsDBNull(v) Then Continue For
                                        Dim d As Double = 0
                                        If Double.TryParse(v.ToString(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, d) Then
                                            col2DataNumeric = True
                                            Exit For
                                        End If
                                    Next
                                End If
                            Catch : End Try
                            ' 仅当第1列数据为文本（分类语义）且第2列含数值时显式指派，
                            ' 两列均数值（如 销售额|成本）时保持 SetSourceData 默认行为。
                            If Not col1DataNumeric AndAlso col2DataNumeric Then
                                Dim sc As Microsoft.Office.Interop.Excel.SeriesCollection = Nothing
                                Try : sc = DirectCast(chart.SeriesCollection, Microsoft.Office.Interop.Excel.SeriesCollection) : Catch : sc = Nothing : End Try
                                If sc IsNot Nothing Then
                                    While CInt(sc.Count) > 1
                                        Try
                                            Dim lastS As Microsoft.Office.Interop.Excel.Series = DirectCast(sc(CInt(sc.Count)), Microsoft.Office.Interop.Excel.Series)
                                            lastS.Delete()
                                            Utils.ComReleaseHelper.Release(lastS)
                                        Catch
                                            Exit While
                                        End Try
                                    End While
                                    Dim s1 As Microsoft.Office.Interop.Excel.Series = Nothing
                                    Try : s1 = DirectCast(sc(1), Microsoft.Office.Interop.Excel.Series) : Catch : s1 = Nothing : End Try
                                    If s1 IsNot Nothing Then
                                        Dim nRows = CInt(srcRng.Rows.Count)
                                        If nRows >= 2 Then
                                            Dim wsData As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                                            Try : wsData = DirectCast(srcRng.Worksheet, Microsoft.Office.Interop.Excel.Worksheet) : Catch : wsData = Nothing : End Try
                                            If wsData IsNot Nothing Then
                                                ' No-PIA 下 COM 索引属性（Range.Cells(行,列)）访问须经 DirectCast 中转，
                                                ' 直接链式访问会触发 Option Strict 后期绑定错误。
                                                Dim catCell1 As Microsoft.Office.Interop.Excel.Range = Nothing
                                                Dim catCell2 As Microsoft.Office.Interop.Excel.Range = Nothing
                                                Dim valCell1 As Microsoft.Office.Interop.Excel.Range = Nothing
                                                Dim valCell2 As Microsoft.Office.Interop.Excel.Range = Nothing
                                                Try : catCell1 = DirectCast(col1.Cells(2, 1), Microsoft.Office.Interop.Excel.Range) : Catch : End Try
                                                Try : catCell2 = DirectCast(col1.Cells(nRows, 1), Microsoft.Office.Interop.Excel.Range) : Catch : End Try
                                                Try : valCell1 = DirectCast(col2.Cells(2, 1), Microsoft.Office.Interop.Excel.Range) : Catch : End Try
                                                Try : valCell2 = DirectCast(col2.Cells(nRows, 1), Microsoft.Office.Interop.Excel.Range) : Catch : End Try
                                                If catCell1 IsNot Nothing AndAlso catCell2 IsNot Nothing Then
                                                    Dim catRng As Microsoft.Office.Interop.Excel.Range = Nothing
                                                    Try : catRng = DirectCast(wsData.Range(catCell1, catCell2), Microsoft.Office.Interop.Excel.Range) : Catch : End Try
                                                    If catRng IsNot Nothing Then
                                                        s1.XValues = catRng
                                                        Utils.ComReleaseHelper.Release(catRng)
                                                    End If
                                                End If
                                                If valCell1 IsNot Nothing AndAlso valCell2 IsNot Nothing Then
                                                    Dim valRng As Microsoft.Office.Interop.Excel.Range = Nothing
                                                    Try : valRng = DirectCast(wsData.Range(valCell1, valCell2), Microsoft.Office.Interop.Excel.Range) : Catch : End Try
                                                    If valRng IsNot Nothing Then
                                                        s1.Values = valRng
                                                        Utils.ComReleaseHelper.Release(valRng)
                                                    End If
                                                End If
                                                Utils.ComReleaseHelper.Release(catCell1)
                                                Utils.ComReleaseHelper.Release(catCell2)
                                                Utils.ComReleaseHelper.Release(valCell1)
                                                Utils.ComReleaseHelper.Release(valCell2)
                                                Utils.ComReleaseHelper.Release(wsData)
                                            End If
                                        End If
                                        Dim hdrCell As Microsoft.Office.Interop.Excel.Range = Nothing
                                        Try : hdrCell = DirectCast(col2.Cells(1, 1), Microsoft.Office.Interop.Excel.Range) : Catch : hdrCell = Nothing : End Try
                                        If hdrCell IsNot Nothing Then
                                            Dim hdr2 As Object = Nothing
                                            Try : hdr2 = hdrCell.Value2 : Catch : hdr2 = Nothing : End Try
                                            If hdr2 IsNot Nothing Then s1.Name = hdr2.ToString()
                                            Utils.ComReleaseHelper.Release(hdrCell)
                                        End If
                                        Utils.ComReleaseHelper.Release(s1)
                                    End If
                                    Utils.ComReleaseHelper.Release(sc)
                                End If
                            End If
                        End If
                        Utils.ComReleaseHelper.Release(col1)
                        Utils.ComReleaseHelper.Release(col2)
                    Catch ex As Exception
                        ' 指派失败不阻断：交给下方诚实校验兜底
                    End Try
                End If

                ' 【诚实绘图校验】纯文本列等无可绘数值的数据源，SetSourceData 不抛异常但图表是空壳
                ' （实测：选区只有"客户类型"这类分类列时，饼图只有标题没有扇区，工具却返回 ✅）。
                ' 此处校验系列确有可绘数值；空壳图一律删除并向调用方明确报错，不做假完成。
                ' 豁免 xlSunburst/xlTreemap（可按类别层级自聚合渲染）。
                If CInt(targetType) <> 117 AndAlso CInt(targetType) <> 115 Then
                    Dim emptySource As Boolean = False
                    Dim diagParts As New List(Of String)()
                    Dim scCheck As Microsoft.Office.Interop.Excel.SeriesCollection = Nothing
                    Try : scCheck = DirectCast(chart.SeriesCollection, Microsoft.Office.Interop.Excel.SeriesCollection) : Catch : scCheck = Nothing : End Try
                    If scCheck Is Nothing OrElse CInt(scCheck.Count) = 0 Then
                        emptySource = True
                        diagParts.Add("系列集合为空或不可用（SetSourceData 未生成任何系列）")
                    Else
                        Dim anyPlottable As Boolean = False
                        Dim sIdx As Integer = 0
                        For Each sChk As Microsoft.Office.Interop.Excel.Series In scCheck
                            sIdx += 1
                            If sChk Is Nothing Then Continue For
                            Try
                                Dim valsObj As Object = sChk.Values
                                Dim arr = TryCast(valsObj, Array)
                                If arr IsNot Nothing Then
                                    ' 兼容 1 维（饼/环单系列）与 2 维（多系列）数组、任意元素类型
                                    ' （Double(,)/Object(,)/String(,)/Double() 等），逐个值 TryParse 判定。
                                    ' 旧实现只认 Double(,) 与 Object(,)，饼图 Series.Values 常返回
                                    ' 1 维数组 → 数值被漏判 → 误报"没有可绘制的数值列"（聚合结果表实测复现）。
                                    Dim samples As New List(Of String)()
                                    For Each v As Object In arr
                                        If v Is Nothing OrElse Convert.IsDBNull(v) Then Continue For
                                        Dim dTmp As Double
                                        If Double.TryParse(v.ToString(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, dTmp) Then
                                            anyPlottable = True
                                            Exit For
                                        End If
                                        If samples.Count < 8 Then
                                            Dim vs = v.ToString()
                                            If vs.Length > 16 Then vs = vs.Substring(0, 16) & "…"
                                            samples.Add(vs)
                                        End If
                                    Next
                                    If Not anyPlottable Then
                                        diagParts.Add($"系列{sIdx}:{arr.GetType().Name}(秩{arr.Rank}) 样本=[{String.Join(",", samples)}]")
                                    End If
                                ElseIf valsObj IsNot Nothing Then
                                    diagParts.Add($"系列{sIdx}:{valsObj.GetType().Name} 值={valsObj.ToString()}")
                                End If
                            Catch : End Try
                            Utils.ComReleaseHelper.Release(sChk)
                            If anyPlottable Then Exit For
                        Next
                        If Not anyPlottable Then emptySource = True
                    End If
                    If scCheck IsNot Nothing Then Utils.ComReleaseHelper.Release(scCheck)
                    If emptySource Then
                        Try : chartObj.Delete() : Catch : End Try
                        Utils.ComReleaseHelper.Release(chart)
                        Utils.ComReleaseHelper.Release(chartObj)
                        Utils.ComReleaseHelper.Release(ws)
                        result.IsError = True
                        Dim srcDesc = If(useMemoryData, "dataValues（内存数组）", "数据范围 """ & dataRange & """")
                        result.Content.Add(New McpTextContent With {
                            .Text = "❌ 未创建图表：" & srcDesc & " 中没有可绘制的数值列（数据为纯文本分类列或空值），已自动清理空图表。" & vbCrLf &
                                    "系列诊断：" & String.Join(" | ", diagParts) & vbCrLf &
                                    "请改为以下任一方式：" & vbCrLf &
                                    "① 选择包含数值列的区域（如 销售额/成本 列）再生成图表；" & vbCrLf &
                                    "② 对分类列做占比分析时说""按 XX 列统计数量并绘图""，系统会先统计频次再绘图。"
                        })
                        Return result
                    End If
                End If
                Utils.ComReleaseHelper.Release(srcRng)

                ' 5) 标题
                If Not String.IsNullOrEmpty(title) Then
                    Try
                        chart.HasTitle = True
                        chart.ChartTitle.Text = title
                    Catch : End Try
                End If

                ' 6) 样式（1~208）
                If chartStyle > 0 AndAlso chartStyle <= 208 Then
                    Try : chart.ChartStyle = chartStyle : Catch : End Try
                End If

                ' 7) 图例 + 位置
                Try
                    chart.HasLegend = hasLegend
                Catch : End Try
                If hasLegend Then
                    Dim lp As Microsoft.Office.Interop.Excel.XlLegendPosition
                    Select Case If(legendPosition, "").ToLowerInvariant()
                        Case "xllegendpositionbottom" : lp = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionBottom
                        Case "xllegendpositiontop"    : lp = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionTop
                        Case "xllegendpositionleft"   : lp = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionLeft
                        Case "xllegendpositioncorner" : lp = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionCorner
                        Case Else                     : lp = Microsoft.Office.Interop.Excel.XlLegendPosition.xlLegendPositionRight
                    End Select
                    Try
                        chart.Legend.Position = lp
                    Catch : End Try
                End If

                ' 8) 坐标轴标题（非笛卡尔图如饼/圆环/雷达/漏斗/瀑布/树状/旭日/箱线无轴 → 先判空再迭代，避免 NullReferenceException）
                If Not String.IsNullOrEmpty(xAxisTitle) OrElse Not String.IsNullOrEmpty(yAxisTitle) Then
                    Dim ct As Microsoft.Office.Interop.Excel.XlChartType = chart.ChartType
                    Dim isNonCartesianChart = (
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlPie OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlPieExploded OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xl3DPie OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xl3DPieExploded OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlPieOfPie OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlBarOfPie OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlDoughnut OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlDoughnutExploded OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlRadar OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlRadarFilled OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlRadarMarkers OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlFunnel OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlWaterfall OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlTreemap OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlSunburst OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlBoxwhisker OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlHistogram OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlPareto OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlRegionMap OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlBubble OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlBubble3DEffect OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatter OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatterLines OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatterLinesNoMarkers OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatterSmooth OrElse
                        ct = Microsoft.Office.Interop.Excel.XlChartType.xlXYScatterSmoothNoMarkers
                    )
                    If Not isNonCartesianChart Then
                        Dim axes As Microsoft.Office.Interop.Excel.Axes = Nothing
                        Try : axes = DirectCast(chart.Axes, Microsoft.Office.Interop.Excel.Axes) : Catch : End Try
                        If axes IsNot Nothing Then
                            For Each ax As Microsoft.Office.Interop.Excel.Axis In axes
                                If ax Is Nothing Then Continue For
                                Try
                                    Dim tc As Integer = CInt(ax.Type)
                                    If tc = 1 AndAlso Not String.IsNullOrEmpty(xAxisTitle) Then
                                        ax.HasTitle = True
                                        ax.AxisTitle.Text = xAxisTitle
                                    End If
                                    If tc = 2 AndAlso Not String.IsNullOrEmpty(yAxisTitle) Then
                                        ax.HasTitle = True
                                        ax.AxisTitle.Text = yAxisTitle
                                    End If
                                Catch : End Try
                            Next
                            Utils.ComReleaseHelper.Release(axes)
                        End If
                    End If
                End If

                ' 【数据标签收敛】分类点数过多（>30）时自动关闭数据标签：
                ' 上百个数据点全开标签必然互相重叠、糊成一片（实测整表建图时 100 个标签完全不可读）。
                If showDataLabels Then
                    Dim labelPoints As Integer = 0
                    Try
                        Dim scCnt As Microsoft.Office.Interop.Excel.SeriesCollection = Nothing
                        Try : scCnt = DirectCast(chart.SeriesCollection, Microsoft.Office.Interop.Excel.SeriesCollection) : Catch : scCnt = Nothing : End Try
                        If scCnt IsNot Nothing AndAlso CInt(scCnt.Count) > 0 Then
                            Dim s1 As Microsoft.Office.Interop.Excel.Series = Nothing
                            Try : s1 = DirectCast(scCnt.Item(1), Microsoft.Office.Interop.Excel.Series) : Catch : s1 = Nothing : End Try
                            If s1 IsNot Nothing Then
                                ' 分类点数取"首个系列的取值数组长度"：Points 集合在 Option Strict On 下是后期绑定，不可用
                                Dim pv As Object = Nothing
                                Try : pv = s1.Values : Catch : pv = Nothing : End Try
                                If TypeOf pv Is Array Then labelPoints = CInt(DirectCast(pv, Array).Length)
                                Utils.ComReleaseHelper.Release(s1)
                            End If
                        End If
                        Utils.ComReleaseHelper.Release(scCnt)
                    Catch
                    End Try
                    If labelPoints > 30 Then
                        showDataLabels = False
                        labelConvergedNote = $"（数据点 {labelPoints} 个，已自动关闭数据标签以免重叠）"
                    End If
                End If

                ' 9) 数据标签（饼/圆环默认 BestFit(=5)，柱图默认 OutsideEnd(=3)，其余 Above(=-4160)）
                If showDataLabels Then
                    Dim seriesColl As Microsoft.Office.Interop.Excel.SeriesCollection = Nothing
                    Try : seriesColl = DirectCast(chart.SeriesCollection, Microsoft.Office.Interop.Excel.SeriesCollection) : Catch : End Try
                    If seriesColl IsNot Nothing Then
                        Dim defaultLblPos As Integer
                        Dim ct As Microsoft.Office.Interop.Excel.XlChartType = chart.ChartType
                        If ct = Microsoft.Office.Interop.Excel.XlChartType.xlPie OrElse
                           ct = Microsoft.Office.Interop.Excel.XlChartType.xlDoughnut Then
                            defaultLblPos = 5  ' xlLabelPositionBestFit
                        ElseIf ct = Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered OrElse
                               ct = Microsoft.Office.Interop.Excel.XlChartType.xlColumnStacked OrElse
                               ct = Microsoft.Office.Interop.Excel.XlChartType.xlColumnStacked100 Then
                            defaultLblPos = 3  ' xlLabelPositionOutsideEnd
                        Else
                            defaultLblPos = -4160  ' xlLabelPositionAbove
                        End If

                        Dim lblPos As Integer = defaultLblPos
                        Select Case If(dataLabelPosition, "").ToLowerInvariant()
                            Case "xllabelpositioncenter"      : lblPos = -4108
                            Case "xllabelpositioninsideend"   : lblPos = 4
                            Case "xllabelpositioninsidebase"  : lblPos = -4109
                            Case "xllabelpositionbestfit"     : lblPos = 5
                            Case "xllabelpositionabove"       : lblPos = -4160
                            Case "xllabelpositionbelow"       : lblPos = -4107
                            Case "xllabelpositionright"       : lblPos = -4152
                            Case "xllabelpositionleft"        : lblPos = -4131
                            Case "xllabelpositionoutsidend"   : lblPos = 3
                        End Select

                        For Each s As Microsoft.Office.Interop.Excel.Series In seriesColl
                            If s Is Nothing Then Continue For
                            Try
                                s.ApplyDataLabels(Type:=Microsoft.Office.Interop.Excel.XlDataLabelsType.xlDataLabelsShowValue,
                                                  LegendKey:=False, AutoText:=True, HasLeaderLines:=True)
                                Try
                                    DirectCast(s.DataLabels(), Microsoft.Office.Interop.Excel.DataLabels).Position =
                                        CType(lblPos, Microsoft.Office.Interop.Excel.XlDataLabelPosition)
                                Catch : End Try
                            Catch : End Try
                        Next
                        Utils.ComReleaseHelper.Release(seriesColl)
                    End If
                End If

                ' 10) 组合次轴：把最后一个系列放次坐标轴 + 开次纵轴
                Dim comboInfo As String = ""
                If useSecondaryAxisLastSeries Then
                    Dim sc As Microsoft.Office.Interop.Excel.SeriesCollection = Nothing
                    Try : sc = DirectCast(chart.SeriesCollection, Microsoft.Office.Interop.Excel.SeriesCollection) : Catch : End Try
                    If sc IsNot Nothing Then
                        Try
                            If CInt(sc.Count) >= 2 Then
                                Dim lastSeries As Microsoft.Office.Interop.Excel.Series = Nothing
                                Try
                                    lastSeries = DirectCast(sc(CInt(sc.Count)), Microsoft.Office.Interop.Excel.Series)
                                    lastSeries.AxisGroup = Microsoft.Office.Interop.Excel.XlAxisGroup.xlSecondary
                                    chart.HasAxis(Microsoft.Office.Interop.Excel.XlAxisType.xlValue,
                                                  Microsoft.Office.Interop.Excel.XlAxisGroup.xlSecondary) = True
                                    comboInfo = $"🎯 组合图：第 {sc.Count} 个系列 → 次坐标轴"
                                Catch : End Try
                                Utils.ComReleaseHelper.Release(lastSeries)
                            End If
                        Catch : End Try
                        Utils.ComReleaseHelper.Release(sc)
                    End If
                End If

                ' 11) 趋势线（trendLineSeriesIndex = 0 表示不开；>=1 指定系列号；trendLineType 支持 linear/polynomial/power/exponential/movingAvg/logarithmic）
                Dim trendInfo As String = ""
                If trendLineSeriesIndex >= 1 AndAlso Not String.IsNullOrEmpty(trendLineType) Then
                    Dim sc As Microsoft.Office.Interop.Excel.SeriesCollection = Nothing
                    Try : sc = DirectCast(chart.SeriesCollection, Microsoft.Office.Interop.Excel.SeriesCollection) : Catch : End Try
                    If sc IsNot Nothing Then
                        Try
                            If CInt(sc.Count) >= trendLineSeriesIndex Then
                                Dim s As Microsoft.Office.Interop.Excel.Series = Nothing
                                Try
                                    s = DirectCast(sc(trendLineSeriesIndex), Microsoft.Office.Interop.Excel.Series)
                                    Dim regType As Microsoft.Office.Interop.Excel.XlTrendlineType = Microsoft.Office.Interop.Excel.XlTrendlineType.xlLinear
                                    Select Case trendLineType.Trim().ToLowerInvariant()
                                        Case "polynomial" : regType = Microsoft.Office.Interop.Excel.XlTrendlineType.xlPolynomial
                                        Case "power"      : regType = Microsoft.Office.Interop.Excel.XlTrendlineType.xlPower
                                        Case "exponential", "exp" : regType = Microsoft.Office.Interop.Excel.XlTrendlineType.xlExponential
                                        Case "movingavg", "movingaverage", "ma" : regType = Microsoft.Office.Interop.Excel.XlTrendlineType.xlMovingAvg
                                        Case "logarithmic", "log" : regType = Microsoft.Office.Interop.Excel.XlTrendlineType.xlLogarithmic
                                    End Select
                                    Dim tls As Microsoft.Office.Interop.Excel.Trendlines = DirectCast(s.Trendlines, Microsoft.Office.Interop.Excel.Trendlines)
                                    tls.Add(Type:=regType)
                                    Utils.ComReleaseHelper.Release(tls)
                                    trendInfo = $"｜📈 趋势线：系列 {trendLineSeriesIndex}（{trendLineType}）"
                                Catch : End Try
                                Utils.ComReleaseHelper.Release(s)
                            End If
                        Catch : End Try
                        Utils.ComReleaseHelper.Release(sc)
                    End If
                End If

                ' 12) COM 清理 + 返回（12 项参数全量回显 → UI「返回结果」一目了然）
                Dim createdName As String = ""
                Try
                    If chartObj IsNot Nothing Then createdName = chartObj.Name
                Catch : End Try
                If String.IsNullOrEmpty(createdName) Then createdName = "图表"
                Utils.ComReleaseHelper.Release(chart)
                Utils.ComReleaseHelper.Release(chartObj)
                Utils.ComReleaseHelper.Release(ws)

                Dim parts As New List(Of String)
                parts.Add($"类型={chartType}")
                If Not String.IsNullOrEmpty(sheetName) Then parts.Add($"sheet={sheetName}")
                parts.Add($"dataRange={dataRange}")
                parts.Add($"seriesBy={seriesBy}")
                If switchRowColumn Then parts.Add("switchRowColumn=True")
                If Not String.IsNullOrEmpty(title) Then parts.Add($"title={title}")
                If chartStyle <> 0 Then parts.Add($"chartStyle={chartStyle}")
                parts.Add($"图例={(If(hasLegend, "显示(" & legendPosition & ")", "隐藏"))}")
                If Not String.IsNullOrEmpty(xAxisTitle) Then parts.Add($"xAxisTitle={xAxisTitle}")
                If Not String.IsNullOrEmpty(yAxisTitle) Then parts.Add($"yAxisTitle={yAxisTitle}")
                If showDataLabels Then
                    If Not String.IsNullOrEmpty(dataLabelPosition) Then
                        parts.Add($"数据标签=显示({dataLabelPosition})")
                    Else
                        parts.Add("数据标签=显示")
                    End If
                End If
                parts.Add($"位置=({posLeft:N0},{posTop:N0}) 大小={posWidth:N0}x{posHeight:N0}")

                result.Content.Add(New McpTextContent With {
                    .Text = $"✅ 成功创建图表『{createdName}』" & vbCrLf &
                            "📌 调用参数：" & String.Join("｜", parts) &
                            seriesConvergedNote & labelConvergedNote &
                            If(String.IsNullOrEmpty(comboInfo), "", vbCrLf & comboInfo) &
                            trendInfo
                })

                ' ── 结构化产物摘要（通用框架层 Artifact）──
                Dim anchorAddr As String = ""
                Try
                    Dim anchorRng = RangeFromPoint(ws, posLeft, posTop)
                    If anchorRng IsNot Nothing Then
                        anchorAddr = anchorRng.Address(False, False)
                        Utils.ComReleaseHelper.Release(anchorRng)
                    End If
                Catch
                End Try
                Dim chartArtifact = New Dictionary(Of String, Object) From {
                    {"type", "chart"},
                    {"chartType", chartType},
                    {"chartName", createdName},
                    {"title", title},
                    {"dataRange", dataRange},
                    {"sheetName", sheetName},
                    {"anchor", anchorAddr},
                    {"positionPt", New Dictionary(Of String, Double) From {
                        {"left", Math.Round(posLeft, 1)},
                        {"top", Math.Round(posTop, 1)},
                        {"width", Math.Round(posWidth, 1)},
                        {"height", Math.Round(posHeight, 1)}
                    }}
                }
                result.Artifact = chartArtifact

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"❌ 创建图表失败: {ex.Message}"})
            End Try
            Return result
        End Function
        ' ===============================================================
        ' 【T3 最终版：主文件直接内嵌，彻底消除 Partial 合并/遮蔽/晚期绑定三连环问题】
        ' ===============================================================
        ' Helper：20 种图表类型 enum → 字符串
        Private Shared Function ChartTypeToStringV2(ct As Microsoft.Office.Interop.Excel.XlChartType) As String
            Select Case CInt(ct)
                Case 51 : Return "xlColumnClustered"
                Case 52 : Return "xlColumnStacked"
                Case 53 : Return "xlColumnStacked100"
                Case 57 : Return "xlBarClustered"
                Case 58 : Return "xlBarStacked"
                Case 59 : Return "xlBarStacked100"
                Case 4 : Return "xlLine"
                Case 65 : Return "xlLineMarkers"
                Case 1 : Return "xlArea"
                Case -4102 : Return "xlPie"
                Case -4120 : Return "xlDoughnut"
                Case -4169 : Return "xlXYScatter"
                Case -4151 : Return "xlRadar"
                Case 86 : Return "xlRadarFilled"
                Case 15 : Return "xlBubble"
                Case 117 : Return "xlSunburst"
                Case 115 : Return "xlTreemap"
                Case 118 : Return "xlHistogram"
                Case 121 : Return "xlBoxwhisker"
                Case 122 : Return "xlFunnel"
                Case 120 : Return "xlWaterfall"
                Case Else : Return "xlChartType_" & CInt(ct)
            End Select
        End Function
        ' Helper：字符串 → 20 种图表类型 enum（未知=报错，不做静默降级）
        Private Shared Function TryParseChartTypeV2(name As String,
                                                    <Runtime.InteropServices.Out()> ByRef ct As Microsoft.Office.Interop.Excel.XlChartType) As Boolean
            Select Case (If(name, "")).Trim().ToLowerInvariant()
                ' 标准 Excel 枚举名（优先匹配）
                Case "xlcolumnclustered", "column", "clusteredcolumn", "柱状图", "柱形图", "分组柱", "对比图"
                    ct = CType(51, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xlcolumnstacked", "stackedcolumn", "堆叠柱", "累积柱"
                    ct = CType(52, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xlcolumnstacked100", "stackedcolumn100", "百分比柱", "100堆叠柱"
                    ct = CType(53, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xlbarclustered", "bar", "clusteredbar", "条形图", "横向柱"
                    ct = CType(57, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xlbarstacked", "stackedbar", "堆叠条形图"
                    ct = CType(58, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xlbarstacked100", "stackedbar100", "100堆叠条形图"
                    ct = CType(59, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xlline", "line", "折线图"
                    ct = CType(4, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xllinemarkers", "linemarkers", "折线带点", "趋势图"
                    ct = CType(65, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xlarea", "area", "面积图"
                    ct = CType(1, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xlpie", "pie", "饼图", "扇形图", "占比图"
                    ct = CType(-4102, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xldoughnut", "doughnut", "donut", "环形图", "圆环图"
                    ct = CType(-4120, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xlxyscatter", "scatter", "xyscatter", "散点图", "xy图"
                    ct = CType(-4169, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xlradar", "radar", "雷达图"
                    ct = CType(-4151, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xlradarfilled", "radarfilled", "填充雷达图"
                    ct = CType(86, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xlbubble", "bubble", "气泡图"
                    ct = CType(15, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xlsunburst", "sunburst", "旭日图"
                    ct = CType(117, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xltreemap", "treemap", "树状图"
                    ct = CType(115, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xlhistogram", "histogram", "直方图"
                    ct = CType(118, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xlboxwhisker", "boxwhisker", "box", "箱线图"
                    ct = CType(121, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xlfunnel", "funnel", "漏斗图"
                    ct = CType(122, Microsoft.Office.Interop.Excel.XlChartType)
                Case "xlwaterfall", "waterfall", "瀑布图"
                    ct = CType(120, Microsoft.Office.Interop.Excel.XlChartType)
                Case Else
                    ct = CType(51, Microsoft.Office.Interop.Excel.XlChartType)
                    Return False
            End Select
            Return True
        End Function
        ''' <summary>auto/未指定图表类型时的启发式自动选型：按数据形状选合理默认类型。
        ''' 规则：首列为日期 → 折线带点（时间趋势）；仅 1 个数值列且分类 ≤6 → 饼图（占比）；否则 → 柱状图。</summary>
        Private Shared ReadOnly _autoChartTypeTokens As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            "auto", "自动", "自动选择", "自动选型", "自动推荐", "默认", "default", "any"
        }

        Private Shared Function ResolveAutoChartType(memRows As List(Of Object()),
                                                     dataRange As String,
                                                     sheetName As String,
                                                     app As Microsoft.Office.Interop.Excel.Application) As String
            Dim rows As List(Of Object()) = memRows
            If rows Is Nothing OrElse rows.Count = 0 Then
                If Not String.IsNullOrWhiteSpace(dataRange) Then
                    Try
                        Dim rng As Microsoft.Office.Interop.Excel.Range = ResolveRangeSafe(app, dataRange, sheetName)
                        If rng IsNot Nothing Then
                            Dim v2 As Object = Nothing
                            Try : v2 = rng.Value2 : Catch : v2 = Nothing : End Try
                            Utils.ComReleaseHelper.Release(rng)
                            If v2 IsNot Nothing Then
                                rows = New List(Of Object())()
                                Dim arr As Object(,) = TryCast(v2, Object(,))
                                If arr IsNot Nothing Then
                                    For r As Integer = 1 To arr.GetLength(0)
                                        Dim rowVals(arr.GetLength(1) - 1) As Object
                                        For c As Integer = 1 To arr.GetLength(1)
                                            rowVals(c - 1) = arr(r, c)
                                        Next
                                        rows.Add(rowVals)
                                    Next
                                Else
                                    rows.Add(New Object() {v2})
                                End If
                            End If
                        End If
                    Catch
                    End Try
                End If
            End If
            If rows Is Nothing OrElse rows.Count < 2 Then Return "xlColumnClustered"

            Dim firstColDate As Boolean = False
            Dim numericCols As Integer = 0
            Dim dataRowCount As Integer = rows.Count - 1
            Dim probeRows As Integer = Math.Min(dataRowCount, 20)

            For r As Integer = 1 To probeRows
                Dim rowV = rows(r)
                If rowV Is Nothing OrElse rowV.Length = 0 Then Continue For
                If TypeOf rowV(0) Is DateTime Then firstColDate = True
            Next

            Dim header = rows(0)
            Dim colCount As Integer = If(header IsNot Nothing, header.Length, 0)
            For c As Integer = 1 To colCount - 1
                For r As Integer = 1 To probeRows
                    Dim rowV = rows(r)
                    If rowV Is Nothing OrElse c >= rowV.Length Then Exit For
                    Dim v = rowV(c)
                    If TypeOf v Is Double OrElse TypeOf v Is Single OrElse TypeOf v Is Integer OrElse
                       TypeOf v Is Long OrElse TypeOf v Is Decimal OrElse TypeOf v Is Short Then
                        numericCols += 1
                        Exit For
                    End If
                Next
            Next

            If firstColDate Then Return "xlLineMarkers"
            If numericCols = 1 AndAlso dataRowCount <= 6 Then Return "xlPie"
            Return "xlColumnClustered"
        End Function

    End Class

End Namespace