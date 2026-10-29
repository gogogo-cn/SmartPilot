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

        Private Function ExecuteFormatRange(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result As New ToolsCallResult()
            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim rangeAddr As String = GetArg(Of String)(args, "range", "")
                ' format/alignment 经 JSON 反序列化后常是 JObject，不能 GetArg(Of String)（Dictionary.ToString 会毁掉 JSON）
                Dim formatJson As String = GetObjectArgAsJson(args, "format", "{}")
                Dim conditionalFormatJson As String = GetObjectArgAsJson(args, "conditionalFormat", "")
                Dim alignmentJson As String = GetObjectArgAsJson(args, "alignment", "")
                Dim autoRange As Boolean = GetArg(Of Boolean)(args, "autoRange", False)
                Dim sheetFallbackNote As String = ""

                Dim excelApp As Microsoft.Office.Interop.Excel.Application = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                ' ========== 参数兜底：优先上游传的range/sheet，空则智能回退 ==========
                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim rng As Microsoft.Office.Interop.Excel.Range = Nothing

                ' 回退1：用户是否在界面上选中了Range？优先级最高（符合"一键格式化"心智）
                ' 【选区范围判断·2026-10-22】指定了 sheetName 时，Selection 在别的表上不能拿来用：
                ' 旧逻辑把"表 B 的选区"赋给 rng，而 ws 后来按 sheetName 解析成表 A；rng 非 Nothing 时
                ' 下方不会用 ws.Range(rangeAddr) 重建 → 最终格式化的是表 B 的选区，汇报却带表 A 的
                ' 名字（silently 做错表）。与 Advanced.Execution.vb 的选区决策链对齐：
                ' Selection 所在表与目标表不一致（忽略大小写）→ 弃用，交给后续 autoRange / 明确报错。
                Dim selSheetMismatchNote As String = ""
                ' 显式 autoRange 表示用该表已用区域，不用当前选区（选区经常只是一个格子）。
                If String.IsNullOrWhiteSpace(rangeAddr) AndAlso Not autoRange Then
                    Dim sel As Object = excelApp.Selection
                    If TypeOf sel Is Microsoft.Office.Interop.Excel.Range Then
                        Dim selRng = DirectCast(sel, Microsoft.Office.Interop.Excel.Range)
                        Dim selSheetName As String = ""
                        Try
                            selSheetName = CStr(selRng.Worksheet.Name)
                        Catch
                            selSheetName = ""
                        End Try
                        If String.IsNullOrWhiteSpace(sheetName) OrElse
                           String.Equals(selSheetName, sheetName, StringComparison.OrdinalIgnoreCase) Then
                            rng = selRng
                            rangeAddr = rng.Address(False, False)
                            ' 若同时 sheetName 也为空，则用 Selection 所属的工作表
                            If String.IsNullOrWhiteSpace(sheetName) Then
                                ws = DirectCast(rng.Worksheet, Microsoft.Office.Interop.Excel.Worksheet)
                                sheetName = ws.Name
                            End If
                        Else
                            selSheetMismatchNote = $"（检测到当前选区在工作表 [{selSheetName}] 上，与指定的 [{sheetName}] 不一致，未采用该选区）"
                            Utils.ComReleaseHelper.Release(selRng)
                        End If
                    End If
                End If

                ' 回退2：仍为空、且没指定工作表时，仅当显式 autoRange=true 才用 ActiveSheet.UsedRange。
                ' 指定了 sheet 时不要在这里抢 ActiveSheet——否则「市场部」的调用会格式化到当前表。
                ' 指定表的 UsedRange 在工作表解析之后处理。
                ' 未传 autoRange 时不再静默格式化整张表（易误伤，见 2026-10 单格选区 999 格问题）。
                If String.IsNullOrWhiteSpace(rangeAddr) AndAlso autoRange AndAlso String.IsNullOrWhiteSpace(sheetName) Then
                    ws = TryCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    If ws IsNot Nothing Then
                        Dim used As Microsoft.Office.Interop.Excel.Range = ws.UsedRange
                        If used IsNot Nothing AndAlso used.Cells.Count > 0 Then
                            rng = used
                            rangeAddr = rng.Address(False, False)
                        End If
                    End If
                End If

                ' 范围仍空、且不是「指定了表名 + autoRange」时才在这里报错。
                ' 后者要等下面解析到那张表，再用它的 UsedRange，不能先按空范围失败。
                If String.IsNullOrWhiteSpace(rangeAddr) AndAlso Not (autoRange AndAlso Not String.IsNullOrWhiteSpace(sheetName)) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "范围不能为空。请先用鼠标选中需要格式化的单元格区域（或确保当前工作表有数据），然后重试一键格式化。" & selSheetMismatchNote})
                    Return result
                End If

                ' 工作表解析：兜底逻辑填过ws就用ws，否则按上游 sheetName 找
                If ws Is Nothing Then
                    If Not String.IsNullOrWhiteSpace(sheetName) Then
                        Try
                            ws = DirectCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                        Catch
                            ws = Nothing
                        End Try
                        ' 【自愈·2026-10-01】指定表名不存在时不再硬失败，改为回退当前活动工作表并明确告知。
                        ' 根因：CSV 打开后工作表名 = 文件名（如 全流程综合测试表_20261001_080927），
                        ' 弱模型（qwen2.5）却常猜成"销售数据"这类数据描述名 → 旧逻辑直接报错返回，
                        ' 导致整个任务半途而废、还弹出"续做确认"（见 09:50 日志：
                        ' "找不到工作表: [销售数据]" → excel_autofit 也没执行 → 用户空等）。
                        ' excel_autofit 早已是"找不到就回退活动表"，这里对齐同一心智，保持工具间一致。
                        If ws Is Nothing Then
                            Dim activeWs = TryCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                            If activeWs IsNot Nothing Then
                                sheetFallbackNote = $"⚠️ 指定的工作表 [{sheetName}] 不存在，已自动改为当前活动工作表 [{activeWs.Name}] 执行。"
                                ws = activeWs
                            End If
                        End If
                    Else
                        ws = TryCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    End If
                End If
                If ws Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "找不到工作表：" & If(sheetName, "(空，请确认选中或指定正确的工作表名称)")})
                    Return result
                End If

                ' 显式 autoRange=true 表示"自动推断有数据区域"。
                ' 注意：不再把 range="A1" 当作魔法值——此前 range="A1" 会被误读为"格式化整张已用区域"，
                ' 导致用户/LLM 随手写的 "A1" 误伤整表（见 2026-10 单格选区格式化 999 格问题）。
                ' 需要整表自动推断时由前端显式传 autoRange=true（StepExecutor 的格式化步骤）。
                If rng Is Nothing AndAlso autoRange Then
                    Try
                        Dim used As Microsoft.Office.Interop.Excel.Range = ws.UsedRange
                        If used IsNot Nothing AndAlso used.Cells.Count > 0 Then
                            rng = used
                            rangeAddr = rng.Address(False, False)
                        End If
                    Catch
                    End Try
                End If

                If String.IsNullOrWhiteSpace(rangeAddr) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "范围不能为空。请先用鼠标选中需要格式化的单元格区域（或确保当前工作表有数据），然后重试一键格式化。" & selSheetMismatchNote})
                    Return result
                End If

                ' Range解析：兜底填过rng就用rng，否则按上游地址取
                If rng Is Nothing Then
                    rng = TryCast(ws.Range(rangeAddr), Microsoft.Office.Interop.Excel.Range)
                End If
                If rng Is Nothing OrElse rng.Cells.Count = 0 Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "范围地址无效: " & rangeAddr & " (工作表: " & ws.Name & ")"})
                    Return result
                End If

                ' 裁剪到实际数据区域：避免把复制/选区带来的大量空白单元格一起格式化
                rng = TrimRangeToData(rng)
                rangeAddr = rng.Address(False, False)

                ' 撤销记录：整表备份（必须在格式化之前，撤销时整张表还原）
                Try
                    Bridges.BridgeRegistry.Instance.Resolve(Of Bridges.IUndoBridge)().AddSheetBackup(ws.Name)
                Catch undoEx As Exception
                    Debug.WriteLine($"[ExecuteFormatRange] 撤销记录失败: {undoEx.Message}")
                End Try

                Dim format As Dictionary(Of String, Object) = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(formatJson)
                If format Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "格式参数无效"})
                    Return result
                End If

                ' ============================================================
                ' 【关键兜底】：空 format（上游没传具体样式规范）时，
                '   应用"商务蓝"专业默认美化预设（SmartTableBeautifier/Formatter 的标准承诺）
                '   否则用户看"格式已应用"但视觉完全没变化 → 以为没生效
                ' ============================================================
                Dim appliedDefaultPreset As Boolean = False
                If format.Count = 0 AndAlso String.IsNullOrWhiteSpace(alignmentJson) AndAlso String.IsNullOrWhiteSpace(conditionalFormatJson) Then
                    appliedDefaultPreset = True
                    result.Content.Add(New McpTextContent With {.Text = "上游未指定具体样式规则，已自动应用【商务蓝专业美化预设】。"})
                End If

                ' 回显实际作用的区域（消除"我选了A但它格式化了什么？"疑惑）。
                ' 注意：本行随最终 result 一起返回，用户看到时格式已应用完毕，故用完成时措辞，避免"正在.../✓完成"自相矛盾。
                If sheetFallbackNote <> "" Then
                    result.Content.Add(New McpTextContent With {.Text = sheetFallbackNote})
                End If
                result.Content.Add(New McpTextContent With {.Text = $"已格式化工作表【{ws.Name}】的区域 {rangeAddr}（共 {rng.Cells.Count} 个单元格）"})

                ' 把Range拆成三部分方便预设：①表头（首行）②数据区（除首行外）③整个区域通用属性
                Dim totalRows As Integer = rng.Rows.Count
                Dim totalCols As Integer = rng.Columns.Count
                Dim headerRow As Microsoft.Office.Interop.Excel.Range = Nothing
                Dim dataBody As Microsoft.Office.Interop.Excel.Range = Nothing
                If totalRows >= 2 Then
                    headerRow = DirectCast(rng.Rows(1), Microsoft.Office.Interop.Excel.Range)
                    dataBody = DirectCast(rng.Range(rng.Cells(2, 1), rng.Cells(totalRows, totalCols)), Microsoft.Office.Interop.Excel.Range)
                End If

                ' ========== 格式应用：上游显式指定优先（如 bold/fontSize），否则默认预设补全 ==========
                If appliedDefaultPreset Then
                    Dim presetWarnings As New List(Of String)()
                    Dim presetSuccess As New List(Of String)()

                    ' ---- 默认预设①：全局统一字体 + 字号 ----
                    Try
                        rng.Font.Name = "微软雅黑"
                        rng.Font.Size = 11
                        rng.Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbBlack
                        presetSuccess.Add("微软雅黑11号 全局字体")
                    Catch ex As Exception
                        presetWarnings.Add("字体统一失败（已跳过）: " & ex.Message)
                    End Try

                    ' ---- 默认预设②：表头（首行）蓝底白字加粗 + 居中 ----
                    If headerRow IsNot Nothing Then
                        Try
                            headerRow.Font.Bold = True
                            headerRow.Font.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbWhite
                            headerRow.Interior.Color = CInt(Microsoft.Office.Interop.Excel.XlRgbColor.rgbRoyalBlue)
                            headerRow.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                            headerRow.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                            headerRow.RowHeight = 24
                            presetSuccess.Add("首行蓝底白字加粗居中 行高24px")
                        Catch ex As Exception
                            presetWarnings.Add("表头美化失败（已跳过）: " & ex.Message)
                        End Try
                    End If

                    ' 表身底色不动。隔行/斑马纹只在用户点名时由显式参数处理，
                    ' 默认美化不再给数据区上色。

                    ' ---- 默认预设④：细边框（全部边框）----
                    Try
                        rng.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                        rng.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                        rng.Borders.Color = CInt(Microsoft.Office.Interop.Excel.XlRgbColor.rgbLightGray)
                        presetSuccess.Add("浅灰细边框 全部边框")
                    Catch ex As Exception
                        presetWarnings.Add("细边框应用失败（已跳过）: " & ex.Message)
                    End Try

                    ' ---- 默认预设⑤：自动列宽（EntireColumn：按整列数据撑开，避免只选表头时数据行仍 ######）----
                    Try
                        rng.EntireColumn.AutoFit()
                    Catch ex As Exception
                        Try
                            rng.Columns.AutoFit()
                        Catch ex2 As Exception
                            presetWarnings.Add("AutoFit列宽失败（已跳过）: " & ex2.Message)
                        End Try
                    End Try
                    presetSuccess.Add("AutoFit列宽")

                    ' ---- 默认预设⑥：冻结首行（≥5行数据表）----
                    Dim freezeResultText As String = "冻结首行 跳过(行数<5)"
                    If totalRows >= 5 AndAlso headerRow IsNot Nothing Then
                        Try
                            ws.Activate()
                            DirectCast(rng.Cells(2, 1), Microsoft.Office.Interop.Excel.Range).Select()
                            excelApp.ActiveWindow.FreezePanes = True
                            freezeResultText = "冻结首行 已启用"
                            presetSuccess.Add(freezeResultText)
                        Catch ex As Exception
                            freezeResultText = "冻结首行 失败（已跳过）: " & ex.Message
                            presetWarnings.Add(freezeResultText)
                        End Try
                    End If

                    If presetSuccess.Count > 0 Then
                        result.Content.Add(New McpTextContent With {
                            .Text = "商务蓝预设已应用：成功 [" & presetSuccess.Count & "] 项 → " & String.Join("; ", presetSuccess) & ";" & If(presetWarnings.Count > 0, " 警告 [" & presetWarnings.Count & "] 项: " & String.Join("; ", presetWarnings), "")
                        })
                    End If
                    If presetWarnings.Count > 0 AndAlso presetSuccess.Count = 0 Then
                        result.Content.Add(New McpTextContent With {
                            .Text = "商务蓝预设全部6项均未成功应用：" & String.Join("; ", presetWarnings)
                        })
                    End If
                End If

                ' 应用格式（上游显式参数优先级最高，覆盖预设）——每项独立Try/Catch，不因单属性失败整体报废
                ' 但关键视觉项（底色/边框/字体色）失败必须记入 appliedFail，禁止静默吞掉后仍报"成功美化"
                Dim appliedOk As New List(Of String)()
                Dim appliedFail As New List(Of String)()
                ' 【加粗语义修正】bold=True 意图只作用于表头，不污染 body；bold=False（取消加粗）才应用全范围
                If format.ContainsKey("bold") Then
                    Try
                        Dim _bv As Boolean = CBool(format("bold"))
                        If _bv AndAlso headerRow IsNot Nothing Then
                            headerRow.Font.Bold = True
                            ' 兜底：如果预设或其他操作已经把 body 意外加粗，这里恢复
                            If dataBody IsNot Nothing Then dataBody.Font.Bold = False
                        Else
                            rng.Font.Bold = _bv
                        End If
                        appliedOk.Add("bold")
                    Catch ex As Exception
                        appliedFail.Add("bold:" & ex.Message)
                    End Try
                End If
                If format.ContainsKey("italic") Then
                    Try
                        rng.Font.Italic = CBool(format("italic"))
                        appliedOk.Add("italic")
                    Catch ex As Exception
                        appliedFail.Add("italic:" & ex.Message)
                    End Try
                End If
                If format.ContainsKey("fontSize") Then
                    Try
                        rng.Font.Size = CDbl(format("fontSize"))
                        appliedOk.Add("fontSize")
                    Catch ex As Exception
                        appliedFail.Add("fontSize:" & ex.Message)
                    End Try
                End If
                If format.ContainsKey("fontName") Then
                    Try
                        rng.Font.Name = format("fontName").ToString()
                        appliedOk.Add("fontName")
                    Catch ex As Exception
                        appliedFail.Add("fontName:" & ex.Message)
                    End Try
                End If
                If format.ContainsKey("fontColor") Then
                    Try
                        rng.Font.Color = OLE(format("fontColor").ToString())
                        appliedOk.Add("fontColor")
                    Catch ex As Exception
                        appliedFail.Add("fontColor:" & ex.Message)
                    End Try
                End If
                ' headerFontColor 只改表头字色，并把表身字色收回黑色。
                ' fontColor 作用整段：表头要白字时如果走 fontColor，表身会变成白字，数据看起来像被清空。
                If format.ContainsKey("headerFontColor") Then
                    Try
                        Dim hdrFont As Microsoft.Office.Interop.Excel.Range = rng
                        If headerRow IsNot Nothing Then hdrFont = headerRow
                        hdrFont.Font.Color = OLE(format("headerFontColor").ToString())
                        If headerRow IsNot Nothing AndAlso dataBody IsNot Nothing Then
                            dataBody.Font.Color = OLE("000000")
                        End If
                        appliedOk.Add("headerFontColor")
                    Catch ex As Exception
                        appliedFail.Add("headerFontColor:" & ex.Message)
                    End Try
                End If
                If format.ContainsKey("bgColor") Then
                    Try
                        ' bgScope=header：底色只上表头，表身填充保持原样。
                        ' 未声明时仍涂整个区域（看板色块、整表底色）。
                        Dim fillTarget As Microsoft.Office.Interop.Excel.Range = rng
                        Dim fillScope As String = ""
                        If format.ContainsKey("bgScope") AndAlso format("bgScope") IsNot Nothing Then
                            fillScope = format("bgScope").ToString().Trim().ToLowerInvariant()
                        End If
                        If fillScope = "header" AndAlso headerRow IsNot Nothing Then
                            fillTarget = headerRow
                        ElseIf fillScope = "body" AndAlso dataBody IsNot Nothing Then
                            fillTarget = dataBody
                        End If
                        fillTarget.Interior.Pattern = Microsoft.Office.Interop.Excel.XlPattern.xlPatternSolid
                        fillTarget.Interior.Color = OLE(format("bgColor").ToString())
                        appliedOk.Add("bgColor=" & format("bgColor").ToString() & If(fillScope = "header", "@header", If(fillScope = "body", "@body", "")))
                    Catch ex As Exception
                        appliedFail.Add("bgColor:" & ex.Message)
                    End Try
                End If
                ' 行高/列宽（看板布局：KPI 卡大字号标题需固定行高撑开、卡片列宽需固定）
                If format.ContainsKey("rowHeight") Then
                    Try
                        rng.RowHeight = CDbl(format("rowHeight"))
                        appliedOk.Add("rowHeight")
                    Catch ex As Exception
                        appliedFail.Add("rowHeight:" & ex.Message)
                    End Try
                End If
                If format.ContainsKey("columnWidth") Then
                    Try
                        rng.ColumnWidth = CDbl(format("columnWidth"))
                        appliedOk.Add("columnWidth")
                    Catch ex As Exception
                        appliedFail.Add("columnWidth:" & ex.Message)
                    End Try
                End If
                If format.ContainsKey("numberFormat") Then
                    Try
                        rng.NumberFormat = format("numberFormat").ToString()
                        appliedOk.Add("numberFormat")
                    Catch
                        Try
                            For Each c As Microsoft.Office.Interop.Excel.Range In rng.Cells
                                Dim v As Object = c.Value2
                                If v IsNot Nothing Then
                                    Select Case v.GetType().Name
                                        Case "Double", "Decimal", "Int32", "Int64", "Int16", "Single"
                                            c.NumberFormat = format("numberFormat").ToString()
                                    End Select
                                End If
                            Next
                            appliedOk.Add("numberFormat")
                        Catch ex As Exception
                            appliedFail.Add("numberFormat:" & ex.Message)
                        End Try
                    End Try
                End If
                If format.ContainsKey("borderStyle") OrElse format.ContainsKey("borderColor") Then
                    Try
                        ApplyBorderFormat(rng, format)
                        appliedOk.Add("border")
                    Catch ex As Exception
                        appliedFail.Add("border:" & ex.Message)
                    End Try
                End If

                ' 应用对齐方式
                If Not String.IsNullOrEmpty(alignmentJson) Then
                    Try
                        Dim alignment As Dictionary(Of String, Object) = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(alignmentJson)
                        If alignment IsNot Nothing Then
                            If alignment.ContainsKey("horizontal") Then
                                Try
                                    Dim hAlign As String = alignment("horizontal").ToString()
                                    Select Case hAlign.ToLower()
                                        Case "center" : rng.HorizontalAlignment = CInt(Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter)
                                        Case "centeracross" : rng.HorizontalAlignment = CInt(Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenterAcrossSelection)
                                        Case "left" : rng.HorizontalAlignment = CInt(Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft)
                                        Case "right" : rng.HorizontalAlignment = CInt(Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight)
                                    End Select
                                Catch : End Try
                            End If
                            If alignment.ContainsKey("vertical") Then
                                Try
                                    Dim vAlign As String = alignment("vertical").ToString()
                                    Select Case vAlign.ToLower()
                                        Case "middle" : rng.VerticalAlignment = CInt(Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter)
                                        Case "top" : rng.VerticalAlignment = CInt(Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop)
                                        Case "bottom" : rng.VerticalAlignment = CInt(Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignBottom)
                                    End Select
                                Catch : End Try
                            End If
                            If alignment.ContainsKey("wrapText") Then Try : rng.WrapText = CBool(alignment("wrapText")) : Catch : End Try
                            If alignment.ContainsKey("shrinkToFit") Then Try : rng.ShrinkToFit = CBool(alignment("shrinkToFit")) : Catch : End Try
                        End If
                    Catch
                    End Try
                End If

                ' 应用条件格式
                If Not String.IsNullOrEmpty(conditionalFormatJson) Then
                    Try
                        Dim condFmt As Dictionary(Of String, Object) = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(conditionalFormatJson)
                        If condFmt IsNot Nothing AndAlso condFmt.ContainsKey("type") Then
                            Dim condType As String = condFmt("type").ToString()
                            Try
                                Dim condRange = rng.FormatConditions
                                Select Case condType.ToLower()
                                    Case "databar"
                                        DeleteTypedRulesOnSheet(rng, ws, condRange, CInt(Microsoft.Office.Interop.Excel.XlFormatConditionType.xlDatabar))
                                        Try : condRange.AddDatabar() : Catch : End Try
                                    Case "colorscale"
                                        DeleteTypedRulesOnSheet(rng, ws, condRange, CInt(Microsoft.Office.Interop.Excel.XlFormatConditionType.xlColorScale))
                                        Try : condRange.AddColorScale(3) : Catch : End Try
                                    Case "iconset"
                                        DeleteTypedRulesOnSheet(rng, ws, condRange, CInt(Microsoft.Office.Interop.Excel.XlFormatConditionType.xlIconSets))
                                        Try : condRange.AddIconSetCondition() : Catch : End Try
                                End Select
                            Catch
                            End Try
                        End If
                    Catch
                    End Try
                End If

                ' ============================================================
                ' 智能列类型识别（日期/货币/整数/常规数字）——无论是否使用预设都执行
                '   仅当用户未显式指定 numberFormat 时启用，避免覆盖用户明确意图
                ' ============================================================
                If Not format.ContainsKey("numberFormat") AndAlso dataBody IsNot Nothing Then
                    Dim colTypeSummary As New List(Of String)()
                    Try
                        For j As Integer = 1 To totalCols
                            Dim dataColBody As Microsoft.Office.Interop.Excel.Range = DirectCast(dataBody.Columns(j), Microsoft.Office.Interop.Excel.Range)

                            ' 取表头文本（用于关键词识别列类型）
                            Dim headerText As String = ""
                            Try
                                Dim hdrVal As Object = DirectCast(rng.Cells(1, j), Microsoft.Office.Interop.Excel.Range).Value2
                                If hdrVal IsNot Nothing Then headerText = Convert.ToString(hdrVal)
                            Catch
                            End Try

                            ' 收集列数据样本（最多 50 行）用于类型判断
                            Dim sampleRows As Integer = Math.Min(50, dataBody.Rows.Count)
                            Dim allNumeric As Boolean = True
                            Dim allInteger As Boolean = True
                            Dim allInDateRange As Boolean = True
                            Dim hasNumeric As Boolean = False
                            For s As Integer = 1 To sampleRows
                                Dim cellVal As Object = DirectCast(dataBody.Cells(s, j), Microsoft.Office.Interop.Excel.Range).Value2
                                If TypeOf cellVal Is Double OrElse TypeOf cellVal Is Decimal OrElse TypeOf cellVal Is Integer OrElse TypeOf cellVal Is Long Then
                                    Dim d As Double = Convert.ToDouble(cellVal)
                                    hasNumeric = True
                                    If d <> Math.Floor(d) Then allInteger = False
                                    If d < 20000 OrElse d > 60000 Then allInDateRange = False
                                Else
                                    allNumeric = False
                                    allInteger = False
                                    allInDateRange = False
                                End If
                            Next

                            If allNumeric Then
                                ' 列类型识别优先级：日期 > 货币 > 整数 > 常规数字
                                Dim isDateCol As Boolean = False
                                Dim isMoneyCol As Boolean = False
                                If headerText.Length > 0 Then
                                    isDateCol = headerText.IndexOf("日期", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                        headerText.IndexOf("时间", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                        headerText.IndexOf("date", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                        headerText.IndexOf("time", StringComparison.OrdinalIgnoreCase) >= 0
                                    isMoneyCol = headerText.IndexOf("金额", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                        headerText.IndexOf("销售额", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                        headerText.IndexOf("成本", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                        headerText.IndexOf("利润", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                        headerText.IndexOf("价格", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                        headerText.IndexOf("收入", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                        headerText.IndexOf("单价", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                        headerText.IndexOf("sales", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                        headerText.IndexOf("revenue", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                        headerText.IndexOf("cost", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                        headerText.IndexOf("profit", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                        headerText.IndexOf("price", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                        headerText.IndexOf("amount", StringComparison.OrdinalIgnoreCase) >= 0
                                End If
                                ' 表头无日期关键词但值全在日期序列号范围（1954-2064年）且非货币列 → 也视为日期列
                                If Not isDateCol AndAlso hasNumeric AndAlso allInDateRange AndAlso Not isMoneyCol Then
                                    isDateCol = True
                                End If

                                dataColBody.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight
                                If isDateCol Then
                                    dataColBody.NumberFormat = "yyyy/m/d"
                                    dataColBody.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                                    colTypeSummary.Add(If(headerText.Length > 0, headerText, "列" & j) & "→日期")
                                ElseIf isMoneyCol Then
                                    dataColBody.NumberFormat = "¥#,##0.00"
                                    colTypeSummary.Add(If(headerText.Length > 0, headerText, "列" & j) & "→货币")
                                ElseIf allInteger Then
                                    dataColBody.NumberFormat = "#,##0"
                                    colTypeSummary.Add(If(headerText.Length > 0, headerText, "列" & j) & "→整数")
                                Else
                                    dataColBody.NumberFormat = "#,##0.00_);[红色](#,##0.00)"
                                    colTypeSummary.Add(If(headerText.Length > 0, headerText, "列" & j) & "→数字")
                                End If
                            Else
                                dataColBody.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                            End If
                            dataColBody.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                        Next
                    Catch ex As Exception
                        ' 记录错误但不中断整体流程
                    End Try
                    If colTypeSummary.Count > 0 Then
                        result.Content.Add(New McpTextContent With {
                            .Text = "智能列类型识别：" & String.Join("; ", colTypeSummary)
                        })
                    End If
                End If

                ' ---- 统一 AutoFit 列宽（预设路径已在上方做过一次；显式 format 路径原先漏做）----
                ' excel_format_range 的 intentKeywords 含「列宽」，摘要也常写「列宽自动调整」；
                ' 若仅预设路径 AutoFit，显式美化后覆盖度仍会误报 excel_autofit 未执行。
                If Not appliedDefaultPreset Then
                    Try
                        rng.EntireColumn.AutoFit()
                        result.Content.Add(New McpTextContent With {.Text = "AutoFit列宽"})
                    Catch ex As Exception
                        Try
                            rng.Columns.AutoFit()
                            result.Content.Add(New McpTextContent With {.Text = "AutoFit列宽"})
                        Catch ex2 As Exception
                            result.Content.Add(New McpTextContent With {
                                .Text = "AutoFit列宽失败（已跳过）: " & ex2.Message
                            })
                        End Try
                    End Try
                End If

                If appliedOk.Count > 0 Then
                    result.Content.Add(New McpTextContent With {
                        .Text = "已应用格式项: " & String.Join(", ", appliedOk)
                    })
                End If
                If appliedFail.Count > 0 Then
                    result.Content.Add(New McpTextContent With {
                        .Text = "⚠️ 以下格式项未生效: " & String.Join("; ", appliedFail)
                    })
                    ' 底色/边框是用户肉眼可见的核心诉求——失败则整工具记错误，禁止模型把失败说成"已美化"
                    If appliedFail.Any(Function(f) f.StartsWith("bgColor:", StringComparison.OrdinalIgnoreCase) OrElse
                                                   f.StartsWith("border:", StringComparison.OrdinalIgnoreCase) OrElse
                                                   f.StartsWith("fontColor:", StringComparison.OrdinalIgnoreCase) OrElse
                                                   f.StartsWith("headerFontColor:", StringComparison.OrdinalIgnoreCase)) Then
                        result.IsError = True
                    End If
                End If

                Utils.ComReleaseHelper.Release(rng)
                Utils.ComReleaseHelper.Release(ws)

                result.Content.Add(New McpTextContent With {
                    .Text = If(result.IsError, $"格式化部分失败 {rangeAddr}", $"成功格式化 {rangeAddr}")
                })

            Catch ex As Exception
                ' ============================================
                '  最终成败判定：只要已经完成任何可视化效果，就不算"失败"
                '   - 预设成功项 ≥ 1
                '   - 或者日志里已经写了"已格式化"/"预设已应用"等实际动作的
                '  只有什么都没干就崩了的情况，才标记 IsError=True
                ' ============================================
                Dim anyEffect As Boolean = False
                For Each ct As McpTextContent In result.Content
                    Dim t As String = If(ct Is Nothing, "", ct.Text)
                    If t.IndexOf("已格式化", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                       t.IndexOf("商务蓝预设已应用", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                       t.IndexOf("成功", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                       t.IndexOf("预设", StringComparison.OrdinalIgnoreCase) >= 0 Then
                        anyEffect = True
                        Exit For
                    End If
                Next
                If anyEffect Then
                    ' 部分成功：只追加警告，不改 IsError
                    result.Content.Add(New McpTextContent With {.Text = $"⚠️ 部分优化项跳过（不影响整体美化结果）: {ex.Message}"})
                Else
                    ' 完全没生效才判失败
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"格式化失败: {ex.Message}"})
                End If
            End Try
            Return result
        End Function

        ''' <summary>
        ''' 把参数里的对象/字符串统一成 JSON 文本。CallLocalTool 反序列化后 format 常是 JObject，
        ''' 若走 GetArg(Of String)→.ToString() 对 Dictionary 会得到类型名而非 JSON，导致整份 format 丢失。
        ''' </summary>
        Private Function GetObjectArgAsJson(args As Dictionary(Of String, Object), key As String, defaultValue As String) As String
            If args Is Nothing OrElse Not args.ContainsKey(key) OrElse args(key) Is Nothing Then
                Return defaultValue
            End If
            Dim value = args(key)
            If TypeOf value Is String Then
                Dim s = CStr(value).Trim()
                Return If(s.Length = 0, defaultValue, s)
            End If
            If TypeOf value Is JToken Then
                Return DirectCast(value, JToken).ToString(Formatting.None)
            End If
            Try
                Return JsonConvert.SerializeObject(value)
            Catch
                Return defaultValue
            End Try
        End Function

        ''' <summary>
        ''' 应用边框。模型/硬短路传 borderStyle="thin"|"medium"|"thick"，那是线宽语义不是 XlLineStyle 枚举名；
        ''' 旧逻辑 Enum.Parse("thin") 必失败并被空 Catch 吞掉 → 边框从未画上。
        ''' </summary>
        Private Shared Sub ApplyBorderFormat(rng As Microsoft.Office.Interop.Excel.Range, format As Dictionary(Of String, Object))
            If rng Is Nothing OrElse format Is Nothing Then Return

            Dim lineStyle As Microsoft.Office.Interop.Excel.XlLineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
            Dim weight As Microsoft.Office.Interop.Excel.XlBorderWeight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin

            If format.ContainsKey("borderStyle") Then
                Dim raw = format("borderStyle")
                If IsNumeric(raw) Then
                    lineStyle = CType(CInt(raw), Microsoft.Office.Interop.Excel.XlLineStyle)
                Else
                    Dim s = raw.ToString().Trim().ToLowerInvariant()
                    ' 去掉可选 xl 前缀
                    If s.StartsWith("xl") Then s = s.Substring(2)
                    Select Case s
                        Case "thin", "continuous", "solid", "hair"
                            lineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                            weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                        Case "medium"
                            lineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                            weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlMedium
                        Case "thick"
                            lineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                            weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThick
                        Case "hairline"
                            lineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                            weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlHairline
                        Case "dash", "dashed"
                            lineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDash
                        Case "dot", "dotted"
                            lineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDot
                        Case "dashdot"
                            lineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlDashDot
                        Case "none", "xlnone"
                            lineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone
                        Case Else
                            ' 兼容 xlContinuous 等真枚举名
                            lineStyle = CType([Enum].Parse(GetType(Microsoft.Office.Interop.Excel.XlLineStyle), raw.ToString(), ignoreCase:=True), Microsoft.Office.Interop.Excel.XlLineStyle)
                    End Select
                End If
            End If

            rng.Borders.LineStyle = lineStyle
            If lineStyle <> Microsoft.Office.Interop.Excel.XlLineStyle.xlLineStyleNone Then
                rng.Borders.Weight = weight
            End If
            If format.ContainsKey("borderColor") Then
                rng.Borders.Color = OLE(format("borderColor").ToString())
            End If
        End Sub

    End Class

End Namespace
