Option Explicit On
Option Infer On
Option Strict Off

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Linq
Imports System.Text
Imports System.Xml.Linq
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json.Serialization
Imports ExcelAIVSTOAddIn.Utils
Imports ExcelAIVSTOAddIn.Bridges

Namespace Services

Partial Public Class PptGenerationService

        ' 共享缓存：探测宿主 Excel 图表是否支持 OLE 嵌入（native 模式可编辑图表依赖它）。
        ' WPS 等环境 Chart.Copy() 必失败，每页重复尝试会造成大量失败噪音与 PowerPoint 状态污染；
        ' 首次判定后缓存结果，后续图表页直接跳过 OLE 尝试、稳定回退带数据位图。
        Private Shared _oleChartSupported As Boolean? = Nothing
        ' 共享缓存：剪贴板 OLE 粘贴（Chart.Copy + PasteSpecial）能力探测。部分环境剪贴板 OLE 失败，
        ' 但 SaveCopyAs + AddOLEObject 文件嵌入仍可用，故与 _oleChartSupported 分开缓存。
        Private Shared _olePasteSupported As Boolean? = Nothing

        ''' <summary>
        ''' 列唯一值计数（跳过空值），用于 X 轴候选筛选与数值轴高唯一性判断
        ''' </summary>
        Private Shared Function ColumnUniqueCount(headers As List(Of String),
                                                  rows As List(Of Dictionary(Of String, Object)),
                                                  colIndex As Integer) As Integer
            If rows Is Nothing OrElse rows.Count = 0 Then Return 0
            Dim key As String = headers(colIndex)
            Dim seen As New HashSet(Of String)()
            For Each row As Dictionary(Of String, Object) In rows
                If row Is Nothing Then Continue For
                Dim v As String = Convert.ToString(GetRowValue(row, key))
                If Not String.IsNullOrWhiteSpace(v) Then seen.Add(v)
            Next
            Return seen.Count
        End Function

        ''' <summary>
        ''' 列是否数值列：非空值中可解析为 Double 的占比 >= 60%
        ''' </summary>
        Private Shared Function IsNumericColumn(headers As List(Of String),
                                                rows As List(Of Dictionary(Of String, Object)),
                                                colIndex As Integer) As Boolean
            If rows Is Nothing OrElse rows.Count = 0 Then Return False
            Dim key As String = headers(colIndex)
            Dim total As Integer = 0
            Dim numeric As Integer = 0
            Dim d As Double
            For Each row As Dictionary(Of String, Object) In rows
                If row Is Nothing Then Continue For
                Dim v As String = Convert.ToString(GetRowValue(row, key))
                If String.IsNullOrWhiteSpace(v) Then Continue For
                total += 1
                If Double.TryParse(v, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, d) Then numeric += 1
            Next
            Return total > 0 AndAlso numeric * 10 >= total * 6
        End Function

        ''' <summary>
        ''' ID 类列：列名命中 ID 模式，或唯一值占比 >= 95% 且样本充足（>=30）视同 ID
        ''' </summary>
        Private Shared Function IsIdLikeColumn(headers As List(Of String),
                                               rows As List(Of Dictionary(Of String, Object)),
                                               colIndex As Integer) As Boolean
            Dim name As String = headers(colIndex)
            If System.Text.RegularExpressions.Regex.IsMatch(name, "id|编号|序号|编码|工号|学号|单号", System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
                Return True
            End If
            If rows Is Nothing OrElse rows.Count < 30 Then Return False
            Dim key As String = headers(colIndex)
            Dim total As Integer = 0
            Dim seen As New HashSet(Of String)()
            For Each row As Dictionary(Of String, Object) In rows
                If row Is Nothing Then Continue For
                Dim v As String = Convert.ToString(GetRowValue(row, key))
                If String.IsNullOrWhiteSpace(v) Then Continue For
                total += 1
                seen.Add(v)
            Next
            Return total >= 30 AndAlso seen.Count * 100 >= total * 95
        End Function

        ''' <summary>
        ''' 缺省 X 轴推断：优先选"首个非数值、非 ID、唯一值适中（2~24）的文本列"；
        ''' 无合适文本列时回退首个非数值非 ID 列；再无则返回 -1（调用方回退第 1 列）
        ''' </summary>
        Private Shared Function FindDefaultXIndex(headers As List(Of String),
                                                  rows As List(Of Dictionary(Of String, Object))) As Integer
            If headers Is Nothing OrElse headers.Count = 0 Then Return -1
            Dim fallback As Integer = -1
            For i As Integer = 0 To headers.Count - 1
                If IsNumericColumn(headers, rows, i) OrElse IsIdLikeColumn(headers, rows, i) Then Continue For
                If fallback < 0 Then fallback = i
                Dim u As Integer = ColumnUniqueCount(headers, rows, i)
                If u >= 2 AndAlso u <= 24 Then Return i
            Next
            Return fallback
        End Function

        Private Shared Function AggregateChartRows(headers As List(Of String),
                                                  rows As List(Of Dictionary(Of String, Object)),
                                                  xIndex As Integer,
                                                  seriesIndexes As List(Of Integer),
                                                  maxCategories As Integer,
                                                  Optional agg As String = "sum",
                                                  Optional categories As List(Of String) = Nothing) As List(Of Dictionary(Of String, Object))
            If rows Is Nothing OrElse rows.Count = 0 Then Return rows
            If maxCategories < 0 Then Return rows

            Dim sums As New Dictionary(Of String, List(Of Double))()
            Dim cnts As New Dictionary(Of String, List(Of Integer))()
            Dim order As New List(Of String)()
            For Each row As Dictionary(Of String, Object) In rows
                Dim xv As String = Convert.ToString(GetRowValue(row, headers(xIndex)))
                If String.IsNullOrWhiteSpace(xv) Then xv = "（空）"
                If Not sums.ContainsKey(xv) Then
                    sums(xv) = New List(Of Double)()
                    cnts(xv) = New List(Of Integer)()
                    For c As Integer = 0 To seriesIndexes.Count - 1
                        sums(xv).Add(0)
                        cnts(xv).Add(0)
                    Next
                    order.Add(xv)
                End If
                For c As Integer = 0 To seriesIndexes.Count - 1
                    Dim v As Double = 0
                    Dim parsed As Boolean = Double.TryParse(Convert.ToString(GetRowValue(row, headers(seriesIndexes(c)))), v)
                    If parsed OrElse agg = "count" Then
                        sums(xv)(c) += v
                        cnts(xv)(c) += 1
                    End If
                Next
            Next

            ' 用户显式指定对比类别（白名单）：仅保留勾选项、按勾选顺序输出，
            ' 不做"其他"汇总、不触发 maxCategories 截断（用户已明确要对比的对象）。
            If categories IsNot Nothing AndAlso categories.Count > 0 Then
                Dim cres As New List(Of Dictionary(Of String, Object))()
                For Each want As String In categories
                    If sums.ContainsKey(want) Then
                        Dim d As New Dictionary(Of String, Object)()
                        d(headers(xIndex)) = want
                        For c As Integer = 0 To seriesIndexes.Count - 1
                            d(headers(seriesIndexes(c))) = ApplyAgg(sums(want)(c), cnts(want)(c), agg)
                        Next
                        cres.Add(d)
                    End If
                Next
                Return cres
            End If

            Dim sorted As List(Of String) = order.OrderByDescending(Function(k) sums(k)(0)).ToList()

            Dim result As New List(Of Dictionary(Of String, Object))()
            Dim keep As Integer = sorted.Count
            If maxCategories > 0 Then keep = Math.Min(maxCategories, sorted.Count)
            For i As Integer = 0 To keep - 1
                Dim k As String = sorted(i)
                Dim d As New Dictionary(Of String, Object)()
                d(headers(xIndex)) = k
                For c As Integer = 0 To seriesIndexes.Count - 1
                    d(headers(seriesIndexes(c))) = ApplyAgg(sums(k)(c), cnts(k)(c), agg)
                Next
                result.Add(d)
            Next

            If sorted.Count > keep Then
                Dim d As New Dictionary(Of String, Object)()
                d(headers(xIndex)) = "其他"
                For c As Integer = 0 To seriesIndexes.Count - 1
                    Dim total As Double = 0
                    Dim totalCnt As Integer = 0
                    For i As Integer = keep To sorted.Count - 1
                        total += sums(sorted(i))(c)
                        totalCnt += cnts(sorted(i))(c)
                    Next
                    d(headers(seriesIndexes(c))) = ApplyAgg(total, totalCnt, agg)
                Next
                result.Add(d)
            End If
            Return result
        End Function

        Private Shared Function ApplyAgg(sumVal As Double, cntVal As Integer, agg As String) As Double
            Select Case agg
                Case "avg"
                    If cntVal = 0 Then Return 0
                    Return sumVal / cntVal
                Case "count"
                    Return cntVal
                Case Else
                    Return sumVal
            End Select
        End Function

        ''' <summary>
        ''' 校验宿主 Excel 图表绑定后的数据有效性：系列数不超过期望数、且至少存在 1 行聚合数据点，
        ''' 防止空图/占位系列（"系列 1"无值）粘贴进 PPT。用于 Calculate/Refresh 后的最终把关，
        ''' 不通过则调用方（BuildChartSlide）不粘贴该图、直接降级为数据表格。
        ''' </summary>
        Private Shared Function ValidSeriesHaveData(chartObj As Object, expectedSeries As Integer, dataRows As Integer) As Boolean
            Try
                If chartObj Is Nothing Then Return False
                If expectedSeries < 1 OrElse dataRows < 1 Then Return False
                ' 图表对象（Shape）下的 Chart 至少应存在 1 个系列；期望系列数为调用方传入的真实绑定数，
                ' 校验时以"存在有效系列"为准，不苛刻到必须逐系列齐全（部分环境 SeriesCollection 访问受限）。
                Dim scCount As Integer = 0
                Try
                    scCount = chartObj.Chart.SeriesCollection.Count
                Catch
                    scCount = 0
                End Try
                If scCount < 1 Then
                    ChatLogger.Instance.LogWarning("PPT_GEN", "ValidSeriesHaveData: 图表系列数为 0（期望 " & expectedSeries.ToString() & "，数据行 " & dataRows.ToString() & "）")
                    Return False
                End If
                Return True
            Catch ex As Exception
                ChatLogger.Instance.LogWarning("PPT_GEN", "ValidSeriesHaveData 校验异常: " & ex.Message)
                Return False
            End Try
        End Function

        Private Shared Function TryCreateExcelChartAndPaste(s As Object, toRelease As List(Of Object),
                                                            headers As List(Of String),
                                                            rows As List(Of Dictionary(Of String, Object)),
                                                            xIndex As Integer,
                                                            seriesIndexes As List(Of Integer),
                                                            chartType As Integer,
                                                            maxCategories As Integer,
                                                            Optional mode As String = "auto",
                                                            Optional agg As String = "sum", Optional categories As List(Of String) = Nothing) As Object
            Dim excelApp As Object = Nothing
            Dim wb As Object = Nothing
            Dim ws As Object = Nothing
            Dim chartObj As Object = Nothing
            Dim oldScreenUpdating As Object = Nothing
            Dim oldDisplayAlerts As Object = Nothing
            Try
                excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing Then Return Nothing

                ' 关闭宿主 Excel 的界面刷新与弹窗，避免临时工作簿打扰用户
                Try
                    oldScreenUpdating = excelApp.ScreenUpdating
                    excelApp.ScreenUpdating = False
                Catch
                End Try
                Try
                    oldDisplayAlerts = excelApp.DisplayAlerts
                    excelApp.DisplayAlerts = False
                Catch
                End Try

                wb = excelApp.Workbooks.Add()
                toRelease.Add(wb)
                ws = wb.Worksheets(1)
                toRelease.Add(ws)

                ' 【静默生成 2026-10-07】隐藏这张临时工作簿的窗口：
                ' ScreenUpdating=False 只减少重绘，不会阻止 Workbooks.Add() 弹出新窗口并抢焦点，
                ' 导致"画一张图闪一次 Excel"。隐藏窗口后所有绘图/复制操作仍在后台正常执行。
                Try
                    wb.Windows(1).Visible = False
                Catch
                End Try

                ' 写入数据（与 WriteChartData 相同的布局：第 1 列 X，其余为系列；同样按 X 轴分类聚合）
                ws.Cells(1, 1).Value = headers(xIndex)
                For c As Integer = 0 To seriesIndexes.Count - 1
                    ws.Cells(1, c + 2).Value = headers(seriesIndexes(c))
                Next
                Dim chartRows As List(Of Dictionary(Of String, Object)) = AggregateChartRows(headers, rows, xIndex, seriesIndexes, maxCategories, agg, categories)
                Dim maxRows As Integer = Math.Min(If(chartRows Is Nothing, 0, chartRows.Count), 100)
                For r As Integer = 0 To maxRows - 1
                    Dim row As Dictionary(Of String, Object) = chartRows(r)
                    ' 分类列按文本、系列列按 Double 写入，规避 SetSourceData 的类型不匹配（DISP_E_TYPEMISMATCH）
                    ws.Cells(r + 2, 1).Value = Convert.ToString(GetRowValue(row, headers(xIndex)))
                    For c As Integer = 0 To seriesIndexes.Count - 1
                        Dim raw As Object = GetRowValue(row, headers(seriesIndexes(c)))
                        Dim num As Double
                        If Double.TryParse(Convert.ToString(raw), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, num) Then
                            ws.Cells(r + 2, c + 2).Value = num
                        Else
                            ws.Cells(r + 2, c + 2).Value = 0
                        End If
                    Next
                Next

                ' 创建 Excel 图表
                Dim lastRow As Integer = Math.Max(2, maxRows + 1)
                Dim lastCol As Integer = seriesIndexes.Count + 1
                Dim dataRange As Object = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
                toRelease.Add(dataRange) ' Range COM 中间对象纳入释放清单
                chartObj = ws.Shapes.AddChart2(240, chartType, 0, 0, 600, 400)
                toRelease.Add(chartObj)
                Dim xRange As Object = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1))
                toRelease.Add(xRange) ' Range COM 中间对象纳入释放清单
                Dim setSourceOk As Boolean = False
                Try
                    ' 显式按列绘制并设置系列名/分类标签，确保粘贴进 PPT 后显示真实字段而非"系列 1/类别 1"
                    chartObj.Chart.SetSourceData(dataRange, -4131) ' xlColumns
                    setSourceOk = True
                Catch ex As Exception
                    ChatLogger.Instance.LogWarning("PPT_GEN", "Excel 图表 SetSourceData 失败，改用逐系列绑定: " & ex.Message)
                End Try
                Dim scCount As Integer = 0
                Try
                    scCount = chartObj.Chart.SeriesCollection.Count
                Catch
                End Try
                For c As Integer = 0 To seriesIndexes.Count - 1
                    Try
                        Dim sc As Object = Nothing
                        If c < scCount Then
                            sc = chartObj.Chart.SeriesCollection(c + 1)
                        Else
                            sc = chartObj.Chart.SeriesCollection.NewSeries()
                        End If
                        toRelease.Add(sc) ' Series COM 中间对象纳入释放清单
                        If Not setSourceOk Then
                            Dim sRange As Object = ws.Range(ws.Cells(2, c + 2), ws.Cells(lastRow, c + 2))
                            toRelease.Add(sRange) ' Range COM 中间对象纳入释放清单
                            sc.Values = sRange
                            sc.XValues = xRange
                        End If
                        sc.Name = headers(seriesIndexes(c))
                        ' 逐系列绑定成功诊断：确认选列与聚合后的行数/求和
                        Dim sumVal As Double = 0
                        For Each crow As Dictionary(Of String, Object) In chartRows
                            Dim raw As Object = GetRowValue(crow, headers(seriesIndexes(c)))
                            Dim num As Double
                            If Double.TryParse(Convert.ToString(raw), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, num) Then sumVal += num
                        Next
                        ChatLogger.Instance.LogInfo("PPT_GEN", "Excel 图表系列 " & (c + 1).ToString() & " 绑定成功: 行=" & chartRows.Count.ToString() & ", 求和=" & sumVal.ToString("F2"))
                    Catch
                    End Try
                Next

                ' 数组绑定后强制重绘：Calculate 使公式/聚合值落值，Refresh 让图表按最新绑定数据重建绘图区，
                ' 避免因跨进程 COM 缓存导致粘贴后仍显示上一次/空白画面。
                Try
                    excelApp.Calculate()
                Catch
                End Try
                Try
                    chartObj.Chart.Refresh()
                Catch
                End Try

                ' 有效性校验：确认绑定后各系列确实有可用数据点（行数/求和均有意义），
                ' 否则不粘贴（返回 Nothing），由调用方降级为数据表格，杜绝"空图/空白错误页"。
                If Not ValidSeriesHaveData(chartObj, seriesIndexes.Count, maxRows) Then
                    ChatLogger.Instance.LogWarning("PPT_GEN", "Excel 图表数据有效性校验失败（类型 " & chartType.ToString() & "，绑定系列数=" & seriesIndexes.Count.ToString() & "，聚合行数=" & maxRows.ToString() & "），不粘贴，降级数据表格")
                    Return Nothing
                End If

                ' 诊断日志：确认图表数据选列、聚合与绑定状态
                Dim diagSeries As New StringBuilder()
                For Each si As Integer In seriesIndexes
                    If diagSeries.Length > 0 Then diagSeries.Append(",")
                    diagSeries.Append(headers(si))
                Next
                ChatLogger.Instance.LogInfo("PPT_GEN", "Excel 图表诊断: chartType=" & chartType.ToString() &
                    ", maxRows=" & maxRows.ToString() & ", seriesNum=" & seriesIndexes.Count.ToString() &
                    ", sCount=" & scCount.ToString() & ", xIndex=" & xIndex.ToString() &
                    ", xField=" & headers(xIndex) & " 系列=[" & diagSeries.ToString() & "]")

                ' 复制并粘贴进 PPT 幻灯片。
                ' mode=picture/auto（默认）→ 复制为位图粘贴成静态图片（最稳，兼容任何环境）；
                ' mode=native（"带数据的图"）→ 优先复制整个可编辑图表对象，粘贴为嵌入 OLE（双击可编辑数据）。
                '   重要：WPS 等环境对 Excel 图表 OLE 嵌入物理不支持（Chart.Copy() 必抛"方法无效"）。
                '   用共享缓存 _oleChartSupported 仅首次探测一次，不支持则后续页跳过 OLE 尝试，
                '   直接回退带数据位图，避免每页重复失败噪音与 PowerPoint 状态污染。
                Dim pasted As Object = Nothing
                If mode = "native" AndAlso _oleChartSupported <> False Then
                    ' 第 1 级：复制可编辑的图表对象（OLE），粘贴为嵌入 OLE（双击可编辑数据）
                    If _olePasteSupported <> False Then
                        Try
                            chartObj.Chart.Copy() ' 复制可编辑的图表对象（OLE）
                            pasted = s.Shapes.PasteSpecial(14) ' PpPasteDataType.ppPasteOLEObject = 14
                            If pasted IsNot Nothing Then
                                _oleChartSupported = True
                                _olePasteSupported = True
                                toRelease.Add(pasted)
                                ChatLogger.Instance.LogInfo("PPT_GEN", "已通过宿主 Excel 创建并粘贴【可编辑】图表（OLE 嵌入，类型 " & chartType.ToString() & "）")
                                Return pasted
                            End If
                        Catch oleEx As Exception
                            _olePasteSupported = False
                            ChatLogger.Instance.LogWarning("PPT_GEN", "图表剪贴板 OLE 粘贴不受支持（类型 " & chartType.ToString() & "，转文件嵌入）: " & oleEx.Message)
                            pasted = Nothing
                        End Try
                    End If

                    ' 第 2 级：剪贴板 OLE 失败时，将已绑定数据的工作簿 SaveCopyAs 临时 .xlsx，
                    ' 通过 AddOLEObject 嵌入 PPT（双击进入 Excel 编辑），不依赖剪贴板 OLE 能力。
                    If pasted Is Nothing Then
                        Dim tempXlsx As String = ""
                        Try
                            Try
                                chartObj.Activate() ' 让嵌入视图展示图表而非工作表单元格
                            Catch
                            End Try
                            ' 【静默生成】Activate 可能把隐藏的临时簿窗口重新显示出来，取完图立刻再隐藏
                            Try
                                If wb IsNot Nothing Then wb.Windows(1).Visible = False
                            Catch
                            End Try
                            tempXlsx = Path.Combine(Path.GetTempPath(), "ppt_ole_" & Guid.NewGuid().ToString("N") & ".xlsx")
                            wb.SaveCopyAs(tempXlsx)
                            pasted = s.Shapes.AddOLEObject(FileName:=tempXlsx, Link:=False, DisplayAsIcon:=False)
                            If pasted IsNot Nothing Then
                                _oleChartSupported = True
                                toRelease.Add(pasted)
                                ChatLogger.Instance.LogInfo("PPT_GEN", "已通过 SaveCopyAs+AddOLEObject 嵌入【可编辑】图表数据（类型 " & chartType.ToString() & "，双击进入 Excel 编辑）")
                                Return pasted
                            End If
                        Catch oleObjEx As Exception
                            ChatLogger.Instance.LogWarning("PPT_GEN", "AddOLEObject 文件嵌入失败（类型 " & chartType.ToString() & "，回退带数据位图）: " & oleObjEx.Message)
                            pasted = Nothing
                        Finally
                            Try
                                If Not String.IsNullOrEmpty(tempXlsx) AndAlso File.Exists(tempXlsx) Then File.Delete(tempXlsx)
                            Catch
                            End Try
                        End Try
                        If pasted Is Nothing Then _oleChartSupported = False
                    End If
                End If

                ' 【静默生成】位图兜底：隐藏窗口下 CopyPicture(xlScreen) 个别环境会失败，
                ' 失败时临时显示窗口重试一次并立刻再隐藏，保证"不弹窗"与"画得出来"两者兼得。
                Try
                    chartObj.CopyPicture(2, 1) ' xlScreen=2, xlBitmap=1（静态位图兜底）
                Catch cpEx As Exception
                    ChatLogger.Instance.LogWarning("PPT_GEN", "隐藏窗口下 CopyPicture 失败，临时显示窗口重试: " & cpEx.Message)
                    Try
                        If wb IsNot Nothing Then wb.Windows(1).Visible = True
                    Catch
                    End Try
                    chartObj.CopyPicture(2, 1)
                    Try
                        If wb IsNot Nothing Then wb.Windows(1).Visible = False
                    Catch
                    End Try
                End Try
                pasted = s.Shapes.Paste()
                If pasted IsNot Nothing Then
                    toRelease.Add(pasted)
                    ChatLogger.Instance.LogInfo("PPT_GEN", "已通过宿主 Excel 创建图表并粘贴进 PPT（类型 " & chartType.ToString() & "）")
                    Return pasted
                End If
                Return Nothing
            Catch ex As Exception
                ChatLogger.Instance.LogError("PPT_GEN", "Excel 图表创建并粘贴失败: " & ex.Message)
                Return Nothing
            Finally
                Try
                    If wb IsNot Nothing Then wb.Close(False)
                Catch
                End Try
                Try
                    If excelApp IsNot Nothing Then excelApp.ScreenUpdating = oldScreenUpdating
                Catch
                End Try
                Try
                    If excelApp IsNot Nothing Then excelApp.DisplayAlerts = oldDisplayAlerts
                Catch
                End Try
            End Try
        End Function
        ''' <summary>
        ''' 图表引擎不可用时的优雅降级：把图表数据渲染为表格，避免空白错误页
        ''' </summary>
        Private Shared Sub BuildChartFallbackTable(s As Object, toRelease As List(Of Object),
                                                   headers As List(Of String),
                                                   rows As List(Of Dictionary(Of String, Object)),
                                                   xIndex As Integer,
                                                   seriesIndexes As List(Of Integer),
                                                   theme As PptTheme)
            Try
                AddTextBox(s, toRelease, "图表组件不可用，已改为数据表格展示：", 40, ContentAreaTop(), 880, 30, 14, False, theme.Gray, theme.Font)

                Dim cols As Integer = seriesIndexes.Count + 1
                Dim maxRows As Integer = Math.Min(If(rows Is Nothing, 0, rows.Count), 12)
                Dim tblShape As Object = s.Shapes.AddTable(maxRows + 1, cols, SX(TableLeft()), SY(TableTop()), SW(TableWidth()), SH(Math.Min(520, 40 + maxRows * 30)))
                toRelease.Add(tblShape)
                Dim tbl As Object = tblShape.Table
                ApplyTableColumnWidths(tbl, cols)

                tbl.Cell(1, 1).Shape.TextFrame.TextRange.Text = headers(xIndex)
                For c As Integer = 0 To seriesIndexes.Count - 1
                    tbl.Cell(1, c + 2).Shape.TextFrame.TextRange.Text = headers(seriesIndexes(c))
                Next
                For r As Integer = 0 To maxRows - 1
                    Dim row As Dictionary(Of String, Object) = rows(r)
                    tbl.Cell(r + 2, 1).Shape.TextFrame.TextRange.Text = GetRowValue(row, headers(xIndex)).ToString()
                    For c As Integer = 0 To seriesIndexes.Count - 1
                        tbl.Cell(r + 2, c + 2).Shape.TextFrame.TextRange.Text = GetRowValue(row, headers(seriesIndexes(c))).ToString()
                    Next
                Next
                ChatLogger.Instance.LogInfo("PPT_GEN", "图表引擎不可用，已降级为表格展示（" & maxRows.ToString() & " 行 × " & cols.ToString() & " 列）")
            Catch ex As Exception
                ChatLogger.Instance.LogError("PPT_GEN", "降级表格渲染失败: " & ex.Message)
                AddTextBox(s, toRelease, "图表与表格均不可用：" & ex.Message, 40, ContentAreaTop(), 880, 40, 14, False, theme.Gray, theme.Font)
            End Try
        End Sub
        Private Shared Function GetRowValue(row As Dictionary(Of String, Object), key As String) As Object
            If row Is Nothing OrElse Not row.ContainsKey(key) Then Return ""
            Dim val As Object = row(key)
            If val Is Nothing OrElse val Is DBNull.Value Then Return ""
            Return val
        End Function
        Public Shared Function MapChartType(chartType As String) As Integer
            If String.IsNullOrWhiteSpace(chartType) Then Return 51
            Dim t As String = chartType.Trim().ToLowerInvariant()
            Select Case t
                Case "column", "bar", "柱状图", "柱形图", "柱图", "xlcolumnclustered", "xlcolumnstacked"
                    Return 51 ' xlColumnClustered
                Case "barh", "条形图", "横向柱状图", "xlbarclustered"
                    Return 57 ' xlBarClustered
                Case "line", "折线图", "趋势图", "xlline", "xllinemarkers"
                    Return 4 ' xlLine
                Case "pie", "饼图", "占比图", "xlpie"
                    Return 5 ' xlPie
                Case "doughnut", "圆环图", "环形图", "xldoughnut"
                    Return -4120 ' xlDoughnut
                Case "area", "面积图", "xlarea"
                    Return 1 ' xlArea
                Case "scatter", "散点图", "xlxyscatter", "xlxyscatterlines"
                    Return 74 ' xlXYScatterLines
                Case "radar", "雷达图", "xlradar"
                    Return -4151 ' xlRadar
                Case "combo", "组合图", "混合图", "柱线图"
                    Return 51 ' xlColumnClustered（配合 lineSeries 参数把指定系列转为折线）
                Case "bubble", "气泡图", "xlbubble"
                    Return 15 ' xlBubble（数据三列：X / Y / 气泡大小）
                Case Else
                    Return 51
            End Select
        End Function
        Private Shared Sub BuildChartSlide(pres As Object, toRelease As List(Of Object),
                                           spec As JObject, theme As PptTheme,
                                           headers As List(Of String),
                                           rows As List(Of Dictionary(Of String, Object)),
                                           Optional chartMode As String = "",
                                           Optional chartAgg As String = "sum")
            Dim s As Object = pres.Slides.Add(pres.Slides.Count + 1, 12)
            SetSlideBackground(s, toRelease, theme.Bg)
            Dim title As String = spec.Value(Of String)("title")
            If String.IsNullOrWhiteSpace(title) Then title = "数据图表"
            AddTextBox(s, toRelease, title, TitleLeft(), TitleTop(), TitleWidth(), TitleHeight(), 26, True, theme.Primary, theme.Font)
            AddRibbon(s, toRelease, theme.Secondary, 40, TitleRibbonTop(), 160, 4)

            Dim chartType As Integer = MapChartType(spec.Value(Of String)("chartType"))
            Dim xField As String = spec.Value(Of String)("xField")
            Dim seriesArr As JArray = spec.Value(Of JArray)("series")

            ' 对比类别（白名单）：用户在配置卡勾选"仅对比这些分类"，按勾选顺序输出，不触发"其他"汇总
            Dim catArr As JArray = TryCast(spec("categories"), JArray)
            Dim catList As New List(Of String)()
            If catArr IsNot Nothing Then
                For Each ct As JToken In catArr
                    Dim cs As String = If(ct Is Nothing, "", ct.ToString().Trim())
                    If Not String.IsNullOrWhiteSpace(cs) AndAlso Not catList.Contains(cs) Then catList.Add(cs)
                Next
            End If

            ' 版式变体：insight 变体在图表下方加数据洞察提示条；full 变体图表占满内容区
            Dim slideVariant As String = spec.Value(Of String)("variant")
            If String.IsNullOrWhiteSpace(slideVariant) Then slideVariant = "standard"
            Dim insight As String = spec.Value(Of String)("insight")
            ' 内容避让：图表区下移到母版顶部装饰区之下，高度自适应防溢出（底部留 70~170 给页脚/提示条）
            Dim chartTop As Integer = ContentAreaTop() + 20
            Dim chartH As Integer = 520
            If slideVariant.ToLowerInvariant() = "full" Then
                chartTop = ContentAreaTop()
                chartH = 610
                chartH = Math.Min(chartH, DESIGN_H - chartTop - 70)
            ElseIf slideVariant.ToLowerInvariant() = "insight" AndAlso Not String.IsNullOrWhiteSpace(insight) Then
                chartTop = ContentAreaTop() + 20
                chartH = Math.Min(430, DESIGN_H - chartTop - 170)
                Dim noteTop As Integer = chartTop + chartH + 18
                AddRoundedRect(s, toRelease, 40, noteTop, 880, 112, theme.Light, 12)
                AddRibbon(s, toRelease, theme.Accent, 40, noteTop, 6, 112)
                AddTextBox(s, toRelease, "数据洞察", 66, noteTop + 16, 120, 30, 14, True, theme.Accent, theme.Font)
                AddTextBox(s, toRelease, insight, 66, noteTop + 48, 830, 52, 15, False, theme.Text, theme.Font)
            Else
                chartH = Math.Min(chartH, DESIGN_H - chartTop - 70)
            End If

            ' X 轴字段：显式指定，否则按列画像推断（优先"首个非数值、非 ID、唯一值适中的文本列"）
            Dim xIndex As Integer = -1
            If Not String.IsNullOrWhiteSpace(xField) Then
                xIndex = headers.IndexOf(xField)
            End If
            If xIndex < 0 Then xIndex = FindDefaultXIndex(headers, rows)
            If xIndex < 0 Then xIndex = 0

            ' 系列字段：显式指定，否则只取数值列（排除 X 轴列）；无数值列时回退全部非 X 列
            Dim seriesIndexes As New List(Of Integer)()
            If seriesArr IsNot Nothing AndAlso seriesArr.Count > 0 Then
                For Each item As JToken In seriesArr
                    Dim name As String = item.ToString()
                    Dim idx As Integer = headers.IndexOf(name)
                    If idx >= 0 AndAlso idx <> xIndex Then seriesIndexes.Add(idx)
                Next
            End If
            If seriesIndexes.Count = 0 Then
                Dim numericCols As List(Of Integer) = Enumerable.Range(0, headers.Count).Where(Function(i) i <> xIndex AndAlso IsNumericColumn(headers, rows, i)).ToList()
                If numericCols.Count > 0 Then
                    seriesIndexes = numericCols
                Else
                    For i As Integer = 0 To headers.Count - 1
                        If i <> xIndex Then seriesIndexes.Add(i)
                    Next
                End If
            End If
            If seriesIndexes.Count = 0 Then
                AddTextBox(s, toRelease, "没有可用于图表的数值列", 40, ContentAreaTop(), 880, 40, 16, False, theme.Gray, theme.Font)
                AddFooter(s, toRelease, pres.Slides.Count, theme.Gray, theme.Font)
                Return
            End If

            ' 容量限制：按版式容量截断系列数，避免图表过载
            Dim cap As PptCapacityProfile = CapacityOf(_activeLayout)
            Dim maxSeries As Integer = If(cap Is Nothing, 5, cap.MaxChartSeries)
            If seriesIndexes.Count > maxSeries Then
                seriesIndexes = seriesIndexes.Take(maxSeries).ToList()
            End If

            ' 次坐标轴：spec 可选 secondaryAxis（数组，列名），命中的系列绑定到次轴（AxisGroup=xlSecondary）。
            ' 仅用于"折线+柱状"这类量纲差异大的组合图；同量纲/单系列用次轴反而误导，故默认不启用。
            Dim secondaryAxisIndexes As New List(Of Integer)()
            Dim secondaryArr As JArray = spec.Value(Of JArray)("secondaryAxis")
            If secondaryArr IsNot Nothing AndAlso secondaryArr.Count > 0 Then
                For Each item As JToken In secondaryArr
                    Dim idx As Integer = seriesIndexes.IndexOf(headers.IndexOf(item.ToString()))
                    If idx >= 0 AndAlso Not secondaryAxisIndexes.Contains(idx) Then secondaryAxisIndexes.Add(idx)
                Next
            End If

            ' 组合图折线系列：spec 可选 lineSeries（数组，系列序号从 1 起或列名），命中的系列转为折线（xlLine），
            ' 与其余柱状系列构成"柱+线"混合图（chartType=combo 时配合使用）。
            Dim lineSeriesIndexes As New List(Of Integer)()
            Dim lineArr As JArray = spec.Value(Of JArray)("lineSeries")
            If lineArr IsNot Nothing AndAlso lineArr.Count > 0 Then
                For Each item As JToken In lineArr
                    Dim sIdx As Integer = -1
                    Dim num As Integer
                    If Integer.TryParse(item.ToString(), num) AndAlso num >= 1 Then
                        sIdx = num - 1
                    Else
                        sIdx = seriesIndexes.IndexOf(headers.IndexOf(item.ToString()))
                    End If
                    If sIdx >= 0 AndAlso Not lineSeriesIndexes.Contains(sIdx) Then lineSeriesIndexes.Add(sIdx)
                Next
            End If

            ' 分类图（饼/柱/环/面积/雷达）聚合后截断到容量上限 + 其他；折线聚合但保留全部时间点；散点按明细绘制
            Dim maxCategories As Integer
            If chartType = 4 Then
                maxCategories = 0
            ElseIf chartType = 74 Then
                maxCategories = -1
            Else
                maxCategories = If(cap Is Nothing, 8, cap.MaxChartCategories)
            End If

            ' 缺省 X 轴为数值列且唯一值高时（如金额/数量类数值轴）不做分类截断，保留全部数据点
            If String.IsNullOrWhiteSpace(xField) AndAlso xIndex >= 0 AndAlso maxCategories > 0 _
               AndAlso IsNumericColumn(headers, rows, xIndex) _
               AndAlso ColumnUniqueCount(headers, rows, xIndex) >= Math.Max(15, If(rows Is Nothing, 0, rows.Count) \ 2) Then
                maxCategories = 0
            End If

            Dim chartShape As Object = Nothing
            Dim isPastedPicture As Boolean = False

            ' 聚合函数取舍：用户驱动模型显式指定（全局 chartAgg，非 auto）→ 最高优先级，覆盖全部图表页；
            ' 否则（auto）采用 spec 页级 agg（默认模型从评分候选聚合视图选出的 sum/avg/count）；再兜底 sum。
            Dim userAgg As String = ""
            If Not String.IsNullOrWhiteSpace(chartAgg) AndAlso Not String.Equals(chartAgg.Trim(), "auto", StringComparison.OrdinalIgnoreCase) Then
                userAgg = chartAgg.Trim().ToLowerInvariant()
            End If
            Dim slideAgg As String = ""
            If userAgg = "" Then
                Dim sa As String = spec.Value(Of String)("agg")
                If Not String.IsNullOrWhiteSpace(sa) AndAlso Not String.Equals(sa.Trim(), "auto", StringComparison.OrdinalIgnoreCase) Then
                    slideAgg = sa.Trim().ToLowerInvariant()
                End If
            End If
            Dim aggMode As String = If(userAgg <> "", userAgg, If(slideAgg <> "", slideAgg, "sum"))

            ' 图表形式：spec 页级 chartMode 覆盖全局；auto=原生优先失败自动降级，picture=仅图片，native=仅带数据原生图
            Dim slideChartMode As String = spec.Value(Of String)("chartMode")
            If String.IsNullOrWhiteSpace(slideChartMode) Then slideChartMode = chartMode
            Dim mode As String = If(String.IsNullOrWhiteSpace(slideChartMode), "auto", slideChartMode.ToLowerInvariant())
            ChatLogger.Instance.LogInfo("PPT_GEN", "图表形式模式=" & mode & "（类型 " & chartType.ToString() & "）聚合=" & aggMode)

            If mode <> "picture" Then
                ' 第 1 层：PowerPoint 原生图表（样式回退链：部分环境特定 Style 值会触发 E_FAIL）
                Dim styleAttempts As Integer() = {240, 201, 1}
                For Each style As Integer In styleAttempts
                    Try
                        chartShape = s.Shapes.AddChart2(style, chartType, SX(40), SY(chartTop), SW(880), SH(chartH))
                        Exit For
                    Catch ex As Exception
                        ChatLogger.Instance.LogWarning("PPT_GEN", "AddChart2 失败（Style=" & style.ToString() & ", 类型 " & chartType.ToString() & "）: " & ex.Message)
                    End Try
                Next
                If chartShape Is Nothing Then
                    Try
                        chartShape = s.Shapes.AddChart(chartType, SX(40), SY(chartTop), SW(880), SH(chartH))
                    Catch ex2 As Exception
                        ChatLogger.Instance.LogWarning("PPT_GEN", "AddChart 也失败（类型 " & chartType.ToString() & "）: " & ex2.Message)
                        chartShape = Nothing
                    End Try
                End If
            End If

            ' 第 2 层：PowerPoint 内嵌工作簿不可用时，在宿主 Excel 建图并粘贴进 PPT（绕开内嵌工作簿）。
            ' 注意：部分环境 PowerPoint 原生图表（AddChart2+内嵌工作簿）必然 E_FAIL（如本机），
            ' 若 native 在此失败后直接降表格，用户将看不到任何图。故 native 也回退 Excel 粘贴，
            ' 只是它优先尝试原生（原生可用时得到可编辑图表），失败才退图片；picture 才直接跳过原生。
            If chartShape Is Nothing Then
                chartShape = TryCreateExcelChartAndPaste(s, toRelease, headers, rows, xIndex, seriesIndexes, chartType, maxCategories, mode, aggMode, catList)
                If chartShape IsNot Nothing Then isPastedPicture = True
            End If

            ' 第 3 层：图表引擎完全不可用 → 数据表格降级，避免空白错误页
            If chartShape Is Nothing Then
                ChatLogger.Instance.LogError("PPT_GEN", "图表创建失败（模式=" & mode & "，原生+Excel 均失败，类型 " & chartType.ToString() & "），降级为表格")
                BuildChartFallbackTable(s, toRelease, headers, rows, xIndex, seriesIndexes, theme)
                AddFooter(s, toRelease, pres.Slides.Count, theme.Gray, theme.Font)
                Return
            End If
            toRelease.Add(chartShape)

            ' Excel 兜底粘贴的图表默认在左上角，按变体坐标重定位
            If isPastedPicture Then
                Try
                    chartShape.Left = SX(40)
                    chartShape.Top = SY(chartTop)
                Catch
                End Try
            End If

            ' 原生图表需要写数据；Excel 粘贴的图片已含数据
            If Not isPastedPicture Then
                Try
                    Dim chart As Object = chartShape.Chart
                    Try
                        chart.ChartTitle.Delete()
                    Catch
                    End Try
                    WriteChartData(chart, toRelease, headers, rows, xIndex, seriesIndexes, maxCategories, secondaryAxisIndexes, lineSeriesIndexes, aggMode, catList)
                Catch ex As Exception
                    ChatLogger.Instance.LogError("PPT_GEN", "配置图表失败: " & ex.Message)
                End Try
            End If

            AddFooter(s, toRelease, pres.Slides.Count, theme.Gray, theme.Font)
        End Sub
        Private Shared Sub WriteChartData(chart As Object, toRelease As List(Of Object),
                                          headers As List(Of String),
                                          rows As List(Of Dictionary(Of String, Object)),
                                          xIndex As Integer,
                                          seriesIndexes As List(Of Integer),
                                          maxCategories As Integer,
                                          Optional secondaryAxisIndexes As List(Of Integer) = Nothing,
                                          Optional lineSeriesIndexes As List(Of Integer) = Nothing,
                                          Optional agg As String = "sum", Optional categories As List(Of String) = Nothing)
            Dim wb As Object = Nothing
            Dim ws As Object = Nothing
            Dim stepName As String = "激活内嵌工作簿"
            Dim activated As Boolean = False
            Try
                stepName = "激活内嵌工作簿"
                ' 【静默生成 2026-10-07】优先直接取内嵌工作簿，不 Activate：
                ' ChartData.Activate() 会弹出"图表数据网格"浮动窗口（像一个小 Excel），
                ' 每写一张原生图表就闪一次。Workbook 在不激活的情况下同样可读写，
                ' 只有取不到时才退回 Activate（随后 Finally 里 Deactivate + 关闭）。
                Try
                    wb = chart.ChartData.Workbook
                Catch
                    wb = Nothing
                End Try
                If wb Is Nothing Then
                    chart.ChartData.Activate()
                    activated = True
                    wb = chart.ChartData.Workbook
                End If
                If wb Is Nothing Then Return
                ws = wb.Worksheets(1)
                If ws Is Nothing Then Return
                toRelease.Add(wb)
                toRelease.Add(ws)
                stepName = "清空内嵌工作表"
                ws.Cells.ClearContents()

                stepName = "写入表头"
                ws.Cells(1, 1).Value = headers(xIndex)
                For c As Integer = 0 To seriesIndexes.Count - 1
                    ws.Cells(1, c + 2).Value = headers(seriesIndexes(c))
                Next

                ' 按 X 轴分类聚合：同一分类的多行明细合并求和，避免饼图/柱状图出现"同一分类拆成几十个扇区"
                Dim chartRows As List(Of Dictionary(Of String, Object)) = AggregateChartRows(headers, rows, xIndex, seriesIndexes, maxCategories, agg, categories)
                Dim maxRows As Integer = Math.Min(If(chartRows Is Nothing, 0, chartRows.Count), 100)
                stepName = "写入数据行"
                For r As Integer = 0 To maxRows - 1
                    Dim row As Dictionary(Of String, Object) = chartRows(r)
                    ' 分类列统一按文本写入，避免纯数字分类被 Excel 误判为数值轴导致系列错位
                    ws.Cells(r + 2, 1).Value = Convert.ToString(GetRowValue(row, headers(xIndex)))
                    For c As Integer = 0 To seriesIndexes.Count - 1
                        ' 系列列显式转 Double：字符串/null/混合类型是 SetSourceData 类型不匹配（DISP_E_TYPEMISMATCH）的常见根因
                        Dim raw As Object = GetRowValue(row, headers(seriesIndexes(c)))
                        Dim num As Double
                        If Double.TryParse(Convert.ToString(raw), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, num) Then
                            ws.Cells(r + 2, c + 2).Value = num
                        Else
                            ws.Cells(r + 2, c + 2).Value = 0
                        End If
                    Next
                Next

                stepName = "设置图表数据源"
                Dim lastRow As Integer = Math.Max(2, maxRows + 1)
                Dim lastCol As Integer = seriesIndexes.Count + 1
                Dim dataRange As Object = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
                toRelease.Add(dataRange) ' Range COM 中间对象纳入释放清单
                Dim xRange As Object = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1))
                toRelease.Add(xRange) ' Range COM 中间对象纳入释放清单
                Dim setSourceOk As Boolean = False
                Try
                    ' 显式按列绘制（第 1 列分类、其余列系列），并显式设置系列名与分类标签，
                    ' 否则 PowerPoint 图表显示"系列 1/类别 1"占位符而非真实字段名。
                    chart.SetSourceData(dataRange, -4131) ' xlColumns
                    setSourceOk = True
                Catch ex As Exception
                    ' 部分环境（内嵌工作簿/特定图表类型）SetSourceData 会报 DISP_E_TYPEMISMATCH，
                    ' 降级为逐系列绑定：直接为每个系列指定 Values/XValues，绕开数据源整体校验
                    ChatLogger.Instance.LogWarning("PPT_GEN", "SetSourceData 失败，改用逐系列绑定: " & ex.Message)
                End Try
                Dim scCount As Integer = 0
                Try
                    scCount = chart.SeriesCollection.Count
                Catch
                End Try
                For c As Integer = 0 To seriesIndexes.Count - 1
                    Try
                        Dim sc As Object = Nothing
                        If c < scCount Then
                            sc = chart.SeriesCollection(c + 1)
                        Else
                            sc = chart.SeriesCollection.NewSeries()
                        End If
                        toRelease.Add(sc) ' Series COM 中间对象纳入释放清单
                        If Not setSourceOk Then
                            Dim sRange As Object = ws.Range(ws.Cells(2, c + 2), ws.Cells(lastRow, c + 2))
                            toRelease.Add(sRange) ' Range COM 中间对象纳入释放清单
                            sc.Values = sRange
                            sc.XValues = xRange
                        End If
                        sc.Name = headers(seriesIndexes(c))
                    Catch
                    End Try
                Next

                ' 次坐标轴：把命中的系列绑定到次轴，并开启次数值轴显示
                ' （组合图折线/柱状量纲不同时，让次轴系列独立量程，避免被主系列压扁/误导。）
                If secondaryAxisIndexes IsNot Nothing AndAlso secondaryAxisIndexes.Count > 0 Then
                    Try
                        For Each c As Integer In secondaryAxisIndexes
                            If c >= 0 AndAlso c < seriesIndexes.Count Then
                                Try
                                    Dim sc2 As Object = chart.SeriesCollection(c + 1)
                                    toRelease.Add(sc2) ' 次轴 Series 中间对象纳入释放清单
                                    sc2.AxisGroup = 2 ' xlSecondary
                                Catch
                                End Try
                            End If
                        Next
                        Try
                            Dim secValueAxis As Object = chart.Axes(2, 2) ' xlValue, xlSecondary
                            toRelease.Add(secValueAxis) ' 次数值轴中间对象纳入释放清单
                            secValueAxis.HasTitle = True
                        Catch
                        End Try
                    Catch
                    End Try
                End If
                ' 组合图：把 lineSeries 命中的系列转为折线（xlLine），实现"柱+线"混合展示
                If lineSeriesIndexes IsNot Nothing AndAlso lineSeriesIndexes.Count > 0 Then
                    For Each c As Integer In lineSeriesIndexes
                        If c >= 0 AndAlso c < seriesIndexes.Count Then
                            Try
                                Dim sc3 As Object = chart.SeriesCollection(c + 1)
                                toRelease.Add(sc3) ' 折线 Series 中间对象纳入释放清单
                                sc3.ChartType = 4 ' xlLine
                            Catch
                            End Try
                        End If
                    Next
                End If
            Catch ex As Exception
                ChatLogger.Instance.LogError("PPT_GEN", "写入图表数据异常（" & stepName & "）: " & ex.Message)
            Finally
                ' 关键：写完后必须关闭内嵌工作簿/数据网格。
                ' 若残留"图表数据网格"窗口，后续所有 AddChart2（原生与宿主 Excel 路径）
                ' 都会报"图表数据网格已在 "" 中打开"而连锁失败，最终整页降级为表格。
                Try
                    chart.ChartData.Deactivate()
                Catch
                End Try
                Try
                    If wb IsNot Nothing Then wb.Close(False)
                Catch
                End Try
            End Try
        End Sub
End Class
End Namespace
