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
Imports System.Threading
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ExcelAIVSTOAddIn.Utils


Namespace Services.Mcp

    Partial Public Class ExcelMcpTools

        ''' <summary>共享透视缓存（按 工作簿|源范围 复用，保证同一数据源的所有透视表被任一切片器联动）</summary>
        Private Shared ReadOnly _sharedPivotCaches As New Dictionary(Of String, Microsoft.Office.Interop.Excel.PivotCache)()

        Private Function ExecuteCopySheet(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim sourceName As String = GetArg(Of String)(args, "source", "")
                Dim targetName As String = GetArg(Of String)(args, "target", "")

                If String.IsNullOrEmpty(sourceName) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "源工作表名称不能为空"})
                    Return result
                End If

                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim sourceWs = TryCast(excelApp.ActiveWorkbook.Worksheets(sourceName), Microsoft.Office.Interop.Excel.Worksheet)
                If sourceWs Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"源工作表 {sourceName} 不存在"})
                    Return result
                End If

                Dim newWs As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Try
                    sourceWs.Copy(After:=sourceWs)
                    newWs = TryCast(excelApp.ActiveWorkbook.Worksheets(sourceWs.Index + 1), Microsoft.Office.Interop.Excel.Worksheet)

                    If Not String.IsNullOrEmpty(targetName) Then
                        newWs.Name = targetName
                    End If

                    result.Content.Add(New McpTextContent With {
                        .Text = $"成功复制工作表: {sourceName} -> {newWs.Name}"
                    })
                Finally
                    Utils.ComReleaseHelper.Release(sourceWs)
                    Utils.ComReleaseHelper.Release(newWs)
                End Try

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"复制工作表失败: {ex.Message}"})
            End Try
            Return result
        End Function

        ''' <summary>
        ''' 将范围裁剪到实际有数据的区域（排除底部/右侧空行空列）。
        ''' 用于避免"复制选中数据/格式化表格"把大量空白单元格也纳入操作范围
        ''' （如选区或 UsedRange 为 A1:Z1000，但实际数据只有 A1:G28）。
        ''' </summary>
        Private Function TrimRangeToData(rng As Microsoft.Office.Interop.Excel.Range) As Microsoft.Office.Interop.Excel.Range
            If rng Is Nothing Then Return Nothing
            Try
                Dim lastCell As Microsoft.Office.Interop.Excel.Range = rng.Find(
                    "*", rng.Cells(1, 1),
                    Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues,
                    Microsoft.Office.Interop.Excel.XlLookAt.xlPart,
                    Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
                    Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious)
                If lastCell IsNot Nothing Then
                    Dim trimmed As Microsoft.Office.Interop.Excel.Range = rng.Worksheet.Range(rng.Cells(1, 1), lastCell)
                    If trimmed.Cells.Count < rng.Cells.Count Then
                        Return trimmed
                    End If
                End If
            Catch
            End Try
            Return rng
        End Function

        Private Function ExecuteCopyRange(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Dim excelApp As Microsoft.Office.Interop.Excel.Application = Nothing
            Dim savedScreenUpdating As Boolean = True
            Try
                Dim sourceSheet As String = GetArg(Of String)(args, "sourceSheet", "")
                Dim sourceRange As String = GetArg(Of String)(args, "sourceRange", "")
                Dim destSheet As String = GetArg(Of String)(args, "destinationSheet", "")
                Dim destRange As String = GetArg(Of String)(args, "destinationRange", "A1")

                If String.IsNullOrEmpty(sourceSheet) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "源工作表名称（sourceSheet）不能为空"})
                    Return result
                End If
                If String.IsNullOrEmpty(destSheet) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "目标工作表名称（destinationSheet）不能为空"})
                    Return result
                End If

                excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim wb = excelApp.ActiveWorkbook
                Dim srcWs = TryCast(wb.Worksheets(sourceSheet), Microsoft.Office.Interop.Excel.Worksheet)
                If srcWs Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"源工作表 '{sourceSheet}' 不存在"})
                    Return result
                End If

                Dim dstWs As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Try
                    dstWs = TryCast(wb.Worksheets(destSheet), Microsoft.Office.Interop.Excel.Worksheet)
                Catch
                End Try
                If dstWs Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"目标工作表 '{destSheet}' 不存在，请先创建该工作表"})
                    Return result
                End If

                ' 确定源范围
                Dim srcRng As Microsoft.Office.Interop.Excel.Range = Nothing
                If Not String.IsNullOrEmpty(sourceRange) Then
                    Try
                        srcRng = srcWs.Range(sourceRange)
                    Catch ex As Exception
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = $"源范围 '{sourceRange}' 无效: {ex.Message}"})
                        Return result
                    End Try
                Else
                    ' 默认使用 UsedRange（排除空行/列）
                    Try
                        srcRng = srcWs.UsedRange
                    Catch ex As Exception
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = $"获取源工作表 UsedRange 失败: {ex.Message}"})
                        Return result
                    End Try
                End If

                If srcRng Is Nothing OrElse srcRng.Rows.Count < 1 Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "源范围无有效数据"})
                    Return result
                End If

                ' 裁剪到实际数据区域：避免把选区/UsedRange 中的大量空白单元格一起复制
                srcRng = TrimRangeToData(srcRng)

                ' 执行复制
                Dim dstRng As Microsoft.Office.Interop.Excel.Range = Nothing
                Try
                    savedScreenUpdating = excelApp.ScreenUpdating
                    excelApp.ScreenUpdating = False

                    srcRng.Copy()
                    dstRng = dstWs.Range(destRange)
                    dstWs.Paste(dstRng)

                    Dim rowsCopied = srcRng.Rows.Count
                    Dim colsCopied = srcRng.Columns.Count
                    Dim srcAddr = srcRng.Address(False, False)
                    Dim dstAddr = dstRng.Address(False, False)

                    result.Content.Add(New McpTextContent With {
                        .Text = $"成功复制数据: {sourceSheet}!{srcAddr} → {destSheet}!{dstAddr}（{rowsCopied}行×{colsCopied}列）"
                    })
                Catch ex As Exception
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"复制数据失败: {ex.Message}"})
                Finally
                    If dstRng IsNot Nothing Then Marshal.ReleaseComObject(dstRng)
                    excelApp.ScreenUpdating = savedScreenUpdating
                    Try
                        excelApp.CutCopyMode = Microsoft.Office.Interop.Excel.XlCutCopyMode.xlCopy
                    Catch
                    End Try
                End Try

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"复制数据失败: {ex.Message}"})
            Finally
                ' ⚠️ 不要释放 excelApp！它是从 Globals.ThisAddIn.Application 获取的共享引用，
                ' 释放后会导致后续所有 Excel 工具调用报 "COM 对象与其基础 RCW 分开后就不能再使用"
                ' 仅释放局部创建的 COM 对象（如 dstRng 已在上面释放）
            End Try
            Return result
        End Function

        Private Function ExecuteClearSheet(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Dim excelApp As Microsoft.Office.Interop.Excel.Application = Nothing
            Dim screenUpdating As Boolean = True
            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")

                excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing

                Try
                    If Not String.IsNullOrEmpty(sheetName) Then
                        ws = TryCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                    End If
                    If ws Is Nothing Then
                        ws = TryCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    End If

                    If ws Is Nothing Then
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = "当前活动对象不是工作表，无法清空。"})
                        Return result
                    End If

                    ' 检查工作表是否受保护
                    If ws.ProtectContents Then
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {
                            .Text = $"工作表 ""{ws.Name}"" 受保护，无法清空。请先取消保护。"
                        })
                        Return result
                    End If

                    ' 关闭屏幕更新以加速操作
                    screenUpdating = excelApp.ScreenUpdating
                    excelApp.ScreenUpdating = False

                    ' 第一步：删除所有 Excel 表格（ListObject）
                    ' Table结构会阻止普通Clear生效，必须先删除表格对象本身
                    Dim tableCount As Integer = 0
                    If ws.ListObjects.Count > 0 Then
                        ' 从后往前删除，避免索引变化
                        For i As Integer = ws.ListObjects.Count To 1 Step -1
                            Try
                                Dim tbl As Microsoft.Office.Interop.Excel.ListObject = ws.ListObjects(i)
                                tbl.Delete()
                                tableCount += 1
                            Catch
                                ' 忽略单个表格的删除错误，尝试清除内容
                                Try
                                    Dim tbl2 As Microsoft.Office.Interop.Excel.ListObject = ws.ListObjects(i)
                                    If tbl2.DataBodyRange IsNot Nothing Then
                                        tbl2.DataBodyRange.Delete()
                                    End If
                                Catch : End Try
                            End Try
                        Next
                    End If

                    ' 第二步：清除所有内容和格式
                    ' 使用 Cells.Clear() 清除所有单元格
                    ws.Cells.Clear()

                    ' 第三步：重置UsedRange（清除"幽灵"已用区域）
                    ' 通过删除空行空列来重置UsedRange缓存
                    Try
                        Dim usedRange As Microsoft.Office.Interop.Excel.Range = ws.UsedRange
                        Dim lastRow As Long = usedRange.Row + usedRange.Rows.Count - 1
                        Dim lastCol As Long = usedRange.Column + usedRange.Columns.Count - 1

                        ' 如果UsedRange超过A1，说明有残留，删除多余的行列
                        If lastRow > 1 OrElse lastCol > 1 Then
                            Try
                                ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Delete()
                            Catch : End Try
                        End If
                    Catch : End Try

                    ' 第四步：选中A1，视觉上重置视图
                    Dim firstCell As Microsoft.Office.Interop.Excel.Range = DirectCast(ws.Cells(1, 1), Microsoft.Office.Interop.Excel.Range)
                    firstCell.Select()

                    ' 恢复屏幕更新并强制刷新
                    excelApp.ScreenUpdating = True
                    screenUpdating = True

                    Dim msg As String = $"成功清空工作表 ""{ws.Name}"" 的所有内容和格式"
                    If tableCount > 0 Then
                        msg &= $"（删除 {tableCount} 个表格）"
                    End If

                    result.Content.Add(New McpTextContent With {.Text = msg})
                Catch
                    ' 内部错误由外层 Try 处理
                End Try

            Catch ex As Exception
                Try : excelApp.ScreenUpdating = True : Catch : End Try
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"清空工作表失败: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteAutoFit(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim rangeAddr As String = GetArg(Of String)(args, "range", "")
                Dim fitType As String = GetArg(Of String)(args, "type", "column").ToLower()

                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim rng As Microsoft.Office.Interop.Excel.Range = Nothing
                Dim sheetFallbackNote As String = ""

                Try
                    If Not String.IsNullOrEmpty(sheetName) Then
                        ' 【自愈·2026-10-01】按名取表必须包 Try：表名不存在时 Worksheets(name) 会抛
                        ' DISP_E_BADINDEX(0x8002000B)。旧代码没有 Try → 整个工具报
                        ' "自动调整失败: 无效索引"，用户看到 ✗ 失败（实测 10:01 日志 sheet="订单明细"，
                        ' 而 CSV 打开后真实表名 = 文件名）。改为回退当前活动表并明确告知。
                        Try
                            ws = TryCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                        Catch
                            ws = Nothing
                        End Try
                        If ws Is Nothing Then
                            Dim activeWs = TryCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                            If activeWs IsNot Nothing Then
                                sheetFallbackNote = $"⚠️ 指定的工作表 [{sheetName}] 不存在，已自动改为当前活动工作表 [{activeWs.Name}] 执行。"
                                ws = activeWs
                            End If
                        End If
                    End If
                    If ws Is Nothing Then
                        ws = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    End If

                    If String.IsNullOrEmpty(rangeAddr) Then
                        rng = ws.UsedRange
                    Else
                        rng = ws.Range(rangeAddr)
                    End If

                    If sheetFallbackNote <> "" Then
                        result.Content.Add(New McpTextContent With {.Text = sheetFallbackNote})
                    End If

                    If fitType = "row" Then
                        rng.Rows.AutoFit()
                        result.Content.Add(New McpTextContent With {
                            .Text = $"成功自动调整行高（范围: {rng.Address}）"
                        })
                    Else
                        ' EntireColumn：按整列内容撑开。只对 A1:I1 调 Columns.AutoFit 时只看表头宽度，
                        ' 数据行数字仍可能显示为 ######（用户截图实测）。
                        Try
                            rng.EntireColumn.AutoFit()
                        Catch
                            rng.Columns.AutoFit()
                        End Try
                        result.Content.Add(New McpTextContent With {
                            .Text = $"成功自动调整列宽（范围: {rng.Address}，按整列内容）"
                        })
                    End If
                Finally
                    Utils.ComReleaseHelper.Release(rng)
                    Utils.ComReleaseHelper.Release(ws)
                End Try

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"自动调整失败: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteSaveWorkbook(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim filePath As String = GetArg(Of String)(args, "filePath", "")

                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim wb = excelApp.ActiveWorkbook

                Try
                    If Not String.IsNullOrEmpty(filePath) Then
                        wb.SaveAs(Filename:=filePath)
                        result.Content.Add(New McpTextContent With {
                            .Text = $"工作簿已另存为: {filePath}"
                        })
                    Else
                        wb.Save()
                        result.Content.Add(New McpTextContent With {
                            .Text = $"工作簿已保存: {wb.FullName}"
                        })
                    End If
                Finally
                    Utils.ComReleaseHelper.Release(wb)
                End Try

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"保存工作簿失败: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteProtectSheet(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim password As String = GetArg(Of String)(args, "password", "")

                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing

                Try
                    If Not String.IsNullOrEmpty(sheetName) Then
                        ws = TryCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                    End If
                    If ws Is Nothing Then
                        ws = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    End If

                    If String.IsNullOrEmpty(password) Then
                        ws.Protect()
                        result.Content.Add(New McpTextContent With {
                            .Text = $"工作表 {ws.Name} 已保护（无密码）"
                        })
                    Else
                        ws.Protect(Password:=password)
                        result.Content.Add(New McpTextContent With {
                            .Text = $"工作表 {ws.Name} 已保护（带密码）"
                        })
                    End If
                Finally
                    Utils.ComReleaseHelper.Release(ws)
                End Try

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"保护工作表失败: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteUnprotectSheet(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim password As String = GetArg(Of String)(args, "password", "")

                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing

                Try
                    If Not String.IsNullOrEmpty(sheetName) Then
                        ws = TryCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                    End If
                    If ws Is Nothing Then
                        ws = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    End If

                    If String.IsNullOrEmpty(password) Then
                        ws.Unprotect()
                    Else
                        ws.Unprotect(Password:=password)
                    End If

                    result.Content.Add(New McpTextContent With {
                        .Text = $"工作表 {ws.Name} 已取消保护"
                    })
                Finally
                    Utils.ComReleaseHelper.Release(ws)
                End Try

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"取消保护工作表失败: {ex.Message}"})
            End Try
            Return result
        End Function
        ''' <summary>
        ''' 创建数据透视表
        ''' </summary>
        Private Function ExecuteCreatePivotTable(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Dim targetCell As String = "A3" ' 提到 Try 外，让顶层 Catch 也能引用

            Try
                Dim sourceRange As String = GetArg(Of String)(args, "sourceRange", "")
                Dim targetSheet As String = GetArg(Of String)(args, "targetSheet", "")
                targetCell = GetArg(Of String)(args, "targetCell", "A3")
                Dim rowFieldsJson As String = GetArg(Of String)(args, "rowFields", "[]")
                Dim columnFieldsJson As String = GetArg(Of String)(args, "columnFields", "[]")
                Dim valueFieldsJson As String = GetArg(Of String)(args, "valueFields", "[]")
                Dim valueFunctionsJson As String = GetArg(Of String)(args, "valueFunctions", "[]")

                If String.IsNullOrWhiteSpace(sourceRange) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "源数据范围不能为空。请提供 sourceRange 参数，例如 'A1:J100' 或 'Sheet1!A1:J100'。"})
                    Return result
                End If

                Dim app = Globals.ThisAddIn.Application
                Dim wb = app.ActiveWorkbook
                If wb Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何工作簿"})
                    Return result
                End If

                ' ═══════════════════════════════════════════════════════════════
                ' Step 1: sourceRange 解析 + 合法性校验
                ' Excel COM 对非法范围直接抛 HRESULT 0x800A03EC（不翻译），
                ' Agent 看到裸异常无法自愈。这里前置捕获并给友好错误。
                ' ═══════════════════════════════════════════════════════════════
                Dim sourceWs As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim sourceRng As Microsoft.Office.Interop.Excel.Range = Nothing
                Try
                    If sourceRange.Contains("!") Then
                        Dim parts = sourceRange.Split("!"c)
                        Dim sheetName = parts(0).Trim("'"c) ' 处理带引号的 Sheet 名
                        sourceWs = DirectCast(wb.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                        sourceRng = sourceWs.Range(parts(1))
                    Else
                        sourceWs = DirectCast(app.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                        sourceRng = sourceWs.Range(sourceRange)
                    End If
                Catch rangeEx As Exception
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text =
                        "源数据范围 '" & sourceRange & "' 解析失败。" &
                        "合法格式：'A1:J100'（当前 Sheet）或 'Sheet1!A1:J100'（跨 Sheet）。" &
                        "你可能传了非法列引用（如 '$1' 不是合法列字母，应为 A-XFD）。" &
                        "原始错误: " & rangeEx.Message})
                    Return result
                End Try

                ' Range 非空 + 至少 2 行（表头 + 数据）
                Try
                    Dim rowCount As Long = CLng(sourceRng.Rows.CountLarge)
                    Dim colCount As Long = CLng(sourceRng.Columns.CountLarge)
                    If rowCount < 2 Then
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = "源数据只有 " & rowCount & " 行，至少需要 2 行（表头 + 数据）。"})
                        Utils.ComReleaseHelper.Release(sourceRng)
                        If sourceWs IsNot Nothing Then Utils.ComReleaseHelper.Release(sourceWs)
                        Return result
                    End If
                    If colCount < 1 Then
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = "源数据范围解析为空。请检查 sourceRange 参数。"})
                        Utils.ComReleaseHelper.Release(sourceRng)
                        If sourceWs IsNot Nothing Then Utils.ComReleaseHelper.Release(sourceWs)
                        Return result
                    End If
                Catch sizeEx As Exception
                    ' CountLarge 在边缘场景可能抛异常，降级用 Count
                    Try
                        If sourceRng.Rows.Count < 2 Or sourceRng.Columns.Count < 1 Then
                            result.IsError = True
                            result.Content.Add(New McpTextContent With {.Text = "源数据范围太小。原始错误: " & sizeEx.Message})
                            Utils.ComReleaseHelper.Release(sourceRng)
                            If sourceWs IsNot Nothing Then Utils.ComReleaseHelper.Release(sourceWs)
                            Return result
                        End If
                    Catch
                    End Try
                End Try

                ' ═══════════════════════════════════════════════════════════════
                ' Step 2: existingNames 扫描（异常隔离 + 提前缓存）
                ' ═══════════════════════════════════════════════════════════════
                Dim existingNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                Try
                    Dim shCount As Integer = CInt(wb.Worksheets.Count)
                    For idx As Integer = 1 To shCount
                        Try
                            Dim sh As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(wb.Worksheets(idx), Microsoft.Office.Interop.Excel.Worksheet)
                            Try
                                Dim pts As Microsoft.Office.Interop.Excel.PivotTables = DirectCast(sh.PivotTables, Microsoft.Office.Interop.Excel.PivotTables)
                                If pts IsNot Nothing Then
                                    Dim ptCount As Integer = CInt(pts.Count)
                                    For j As Integer = 1 To ptCount
                                        Try
                                            Dim pt As Microsoft.Office.Interop.Excel.PivotTable = DirectCast(pts(j), Microsoft.Office.Interop.Excel.PivotTable)
                                            existingNames.Add(pt.Name)
                                            Utils.ComReleaseHelper.Release(pt)
                                        Catch
                                            ' 单个 pt 访问异常不阻塞其他
                                        End Try
                                    Next
                                End If
                                Utils.ComReleaseHelper.Release(pts)
                            Catch
                                ' 单个 Sheet 的 PivotTables 访问异常不阻塞其他
                            End Try
                            Utils.ComReleaseHelper.Release(sh)
                        Catch
                            ' 单个 Sheet 访问异常不阻塞其他
                        End Try
                    Next
                Catch scanEx As Exception
                    ' 扫描整体失败 → 退化为保守策略：从 PivotTable1 开始，撞了就加后缀
                    Debug.WriteLine("[ExecuteCreatePivotTable] existingNames 扫描失败: " & scanEx.Message)
                End Try

                Dim pivotName As String = "PivotTable1"
                Dim seq As Integer = 1
                While existingNames.Contains(pivotName)
                    seq += 1
                    pivotName = "PivotTable" & seq.ToString()
                End While

                ' ═══════════════════════════════════════════════════════════════
                ' Step 3: 创建目标工作表
                ' ═══════════════════════════════════════════════════════════════
                Dim destWs As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                If Not String.IsNullOrEmpty(targetSheet) Then
                    Try
                        destWs = DirectCast(wb.Worksheets(targetSheet), Microsoft.Office.Interop.Excel.Worksheet)
                    Catch
                        destWs = DirectCast(wb.Worksheets.Add(), Microsoft.Office.Interop.Excel.Worksheet)
                        Try : destWs.Name = targetSheet : Catch : End Try
                    End Try
                Else
                    destWs = DirectCast(wb.Worksheets.Add(), Microsoft.Office.Interop.Excel.Worksheet)
                    Try : destWs.Name = "数据透视表" : Catch : End Try
                End If

                ' ═══════════════════════════════════════════════════════════════
                ' Step 4: CreatePivotTable（带撞名重试——existingNames 扫描后，
                '                          destWs 创建 / PivotCache 创建可能改变状态）
                ' ═══════════════════════════════════════════════════════════════
                Dim pivotCache As Microsoft.Office.Interop.Excel.PivotCache = Nothing
                Dim pivotTable As Microsoft.Office.Interop.Excel.PivotTable = Nothing
                Dim maxRetry As Integer = 5
                Dim created As Boolean = False

                For attempt As Integer = 1 To maxRetry
                    Try
                        pivotCache = DirectCast(
                            wb.PivotCaches().Create(Microsoft.Office.Interop.Excel.XlPivotTableSourceType.xlDatabase, sourceRng),
                            Microsoft.Office.Interop.Excel.PivotCache)

                        Dim destRange = destWs.Range(targetCell)
                        pivotTable = DirectCast(pivotCache.CreatePivotTable(destRange, pivotName), Microsoft.Office.Interop.Excel.PivotTable)
                        ' 反射调用 ResetFields 清除 Excel 默认自动配置（把第一列设为 Row Field 等），
                        ' 让所有源数据列回到 PivotFields 集合。用反射是因为 PIA 接口上 ResetFields 不在强类型定义里，
                        ' 直接调用会触发 BC30574 Option Strict On 不允许后期绑定。
                        Try
                            pivotTable.GetType().InvokeMember("ResetFields",
                                System.Reflection.BindingFlags.InvokeMethod,
                                Nothing, pivotTable, Nothing)
                        Catch
                        End Try
                        created = True
                        Exit For
                    Catch nameEx As Exception
                        ' 撞名 → 递增序号重试
                        Debug.WriteLine("[ExecuteCreatePivotTable] 尝试 " & attempt & "/" & maxRetry &
                                       " 名称='" & pivotName & "' 失败: " & nameEx.Message)
                        ' 清理上次尝试的 PivotCache（如果已创建）
                        If pivotCache IsNot Nothing Then
                            Try : Utils.ComReleaseHelper.Release(pivotCache) : Catch : End Try
                            pivotCache = Nothing
                        End If
                        ' 跳过错误消息里能看出是名称问题时才重试
                        Dim msgLower = nameEx.Message.ToLowerInvariant()
                        If msgLower.Contains("name") OrElse msgLower.Contains("名称") OrElse
                           msgLower.Contains("already been used") OrElse msgLower.Contains("已被使用") OrElse
                           msgLower.Contains("0x800a03ec") Then
                            ' 递增序号 + 重新扫描 existingNames
                            seq += 1
                            pivotName = "PivotTable" & seq.ToString()
                            existingNames.Add(pivotName) ' 防止同一轮死循环
                            Continue For
                        Else
                            ' 非名称类错误 → 直接抛出去
                            Throw
                        End If
                    End Try
                Next

                If Not created Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text =
                        "创建数据透视表失败（重试 " & maxRetry & " 次均撞名）。" &
                        "当前尝试的名称: " & pivotName & "。" &
                        "工作簿中已存在的透视表名: " & String.Join(", ", existingNames)})
                    Utils.ComReleaseHelper.Release(sourceRng)
                    If sourceWs IsNot Nothing Then Utils.ComReleaseHelper.Release(sourceWs)
                    If destWs IsNot Nothing Then Utils.ComReleaseHelper.Release(destWs)
                    Return result
                End If

                ' 解析字段
                Dim rowFields = JsonConvert.DeserializeObject(Of List(Of String))(rowFieldsJson)
                Dim columnFields = JsonConvert.DeserializeObject(Of List(Of String))(columnFieldsJson)
                Dim valueFields = JsonConvert.DeserializeObject(Of List(Of String))(valueFieldsJson)
                Dim valueFunctions = JsonConvert.DeserializeObject(Of List(Of String))(valueFunctionsJson)

                ' 【可自愈】先收集透视表实际可用字段名，字段不存在时报出清单而不是裸 COM 错误
                ' （实测：弱模型幻觉字段"月份"（表头实为"日期"）→ PivotFields("月份") 抛 COM 异常，
                '  且此时透视表已创建 → 半成品残留，重试被污染。2026-10-06 截图）
                ' 根因：PivotCaches.Create(xlDatabase, range) 不自动识别表头 → PivotField.Name 返回
                ' 空字符串或默认"字段N"（截图里"可用字段"全是"字段2~字段9"就是这个原因）。
                ' 解法：同时读 PivotField.Name/SourceName 和 sourceRng 第一行表头，sourceRng 优先。
                Dim availableFields As New List(Of String)()
                Dim pfAll As Microsoft.Office.Interop.Excel.PivotFields = DirectCast(pivotTable.PivotFields, Microsoft.Office.Interop.Excel.PivotFields)

                ' Step A: 先从 sourceRng 第一行读所有列表头——这是最可靠的字段名来源
                Dim hdrColCount As Integer = CInt(sourceRng.Columns.Count)
                Dim sourceHeaders(hdrColCount - 1) As String
                Try
                    For c As Integer = 1 To hdrColCount
                        Dim hCell As Microsoft.Office.Interop.Excel.Range = DirectCast(sourceRng.Cells.Item(1, c), Microsoft.Office.Interop.Excel.Range)
                        Dim hVal As Object = hCell.Value2
                        sourceHeaders(c - 1) = If(hVal Is Nothing, "", CStr(hVal).Trim())
                        Utils.ComReleaseHelper.Release(hCell)
                    Next
                Catch
                    For c As Integer = 0 To hdrColCount - 1 : sourceHeaders(c) = "" : Next
                End Try

                ' Step B: 遍历 PivotFields，按索引对应 sourceHeaders 表头
                ' ResetFields 后 PivotFields 应包含所有源数据列，索引 1..Count = 源列序
                Dim pfCount As Integer = CInt(pfAll.Count)
                For i0 As Integer = 1 To pfCount
                    Dim f0 As Microsoft.Office.Interop.Excel.PivotField = DirectCast(pfAll(i0), Microsoft.Office.Interop.Excel.PivotField)
                    Try
                        Dim fName As String = ""
                        Try : fName = CStr(f0.Name) : Catch : End Try
                        Dim fSource As String = ""
                        Try : fSource = CStr(f0.SourceName) : Catch : End Try

                        ' 优先级：sourceHeaders(i0-1) > fSourceName > fName
                        Dim idx As Integer = i0 - 1
                        If idx >= 0 AndAlso idx < sourceHeaders.Length AndAlso Not String.IsNullOrWhiteSpace(sourceHeaders(idx)) Then
                            availableFields.Add(sourceHeaders(idx))
                        ElseIf Not String.IsNullOrWhiteSpace(fSource) Then
                            availableFields.Add(fSource)
                        ElseIf Not String.IsNullOrWhiteSpace(fName) Then
                            availableFields.Add(fName)
                        End If
                    Catch
                    End Try
                    Utils.ComReleaseHelper.Release(f0)
                Next
                ' Step C: 如果 sourceHeaders 里还有没被 PivotFields 覆盖的（ResetFields 不完整等边缘情况），补上
                If availableFields.Count < hdrColCount Then
                    For c As Integer = 0 To hdrColCount - 1
                        If Not String.IsNullOrWhiteSpace(sourceHeaders(c)) AndAlso
                           Not availableFields.Contains(sourceHeaders(c), StringComparer.OrdinalIgnoreCase) Then
                            availableFields.Add(sourceHeaders(c))
                        End If
                    Next
                End If

                ' ═══════════════════════════════════════════════════════════════
                ' Phase 1 (dry-run)：全量校验所有字段存在性，不动透视表。
                ' 三个类型独立遍历 → 一次返回全部 missing，让 Agent 一轮自愈。
                ' （旧实现用 AndAlso missingFields.Count=0 提前退出 → 分批暴露错误）
                ' ═══════════════════════════════════════════════════════════════
                Dim missingFields As New List(Of String)()

                If rowFields IsNot Nothing Then
                    For Each name In rowFields
                        If String.IsNullOrWhiteSpace(name) Then Continue For
                        Dim f = ResolvePivotFieldLoose(pfAll, name, sourceHeaders)
                        If f Is Nothing Then missingFields.Add(name) Else Utils.ComReleaseHelper.Release(f)
                    Next
                End If
                If columnFields IsNot Nothing Then
                    For Each name In columnFields
                        If String.IsNullOrWhiteSpace(name) Then Continue For
                        Dim f = ResolvePivotFieldLoose(pfAll, name, sourceHeaders)
                        If f Is Nothing Then missingFields.Add(name) Else Utils.ComReleaseHelper.Release(f)
                    Next
                End If
                If valueFields IsNot Nothing Then
                    For Each name In valueFields
                        If String.IsNullOrWhiteSpace(name) Then Continue For
                        Dim f = ResolvePivotFieldLoose(pfAll, name, sourceHeaders)
                        If f Is Nothing Then missingFields.Add(name) Else Utils.ComReleaseHelper.Release(f)
                    Next
                End If

                If missingFields.Count > 0 Then
                    ' 回滚：彻底删掉刚创建的半成品透视表，避免残留空透视表污染重试/刷新
                    Try
                        pivotTable.TableRange2.Delete()
                    Catch
                        Try : pivotTable.TableRange2.Clear() : Catch : End Try
                    End Try
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "创建数据透视表失败：字段不存在 " & String.Join("、", missingFields) & "。源数据可用字段：" & String.Join("、", availableFields) & "。已回滚本次创建，请改用上述真实字段名重试（也支持列字母 A/B/C…）。"})
                    Utils.ComReleaseHelper.Release(pfAll)
                    Utils.ComReleaseHelper.Release(pivotTable)
                    If pivotCache IsNot Nothing Then Utils.ComReleaseHelper.Release(pivotCache)
                    Utils.ComReleaseHelper.Release(sourceRng)
                    If sourceWs IsNot Nothing Then Utils.ComReleaseHelper.Release(sourceWs)
                    If destWs IsNot Nothing Then Utils.ComReleaseHelper.Release(destWs)
                    Return result
                End If

                ' ═══════════════════════════════════════════════════════════════
                ' Phase 2 (apply)：所有字段校验通过 → 原子配置
                ' ═══════════════════════════════════════════════════════════════
                If rowFields IsNot Nothing Then
                    For i = 0 To rowFields.Count - 1
                        Dim f = ResolvePivotFieldLoose(pfAll, rowFields(i), sourceHeaders)
                        f.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField
                        f.Position = i + 1
                        Utils.ComReleaseHelper.Release(f)
                    Next
                End If
                If columnFields IsNot Nothing Then
                    For i = 0 To columnFields.Count - 1
                        Dim f = ResolvePivotFieldLoose(pfAll, columnFields(i), sourceHeaders)
                        f.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlColumnField
                        f.Position = i + 1
                        Utils.ComReleaseHelper.Release(f)
                    Next
                End If
                If valueFields IsNot Nothing Then
                    For i = 0 To valueFields.Count - 1
                        Dim f = ResolvePivotFieldLoose(pfAll, valueFields(i), sourceHeaders)
                        Dim funcName = If(valueFunctions IsNot Nothing AndAlso i < valueFunctions.Count, valueFunctions(i), "Sum")
                        pivotTable.AddDataField(f, , DirectCast(GetPivotFunctionEnum(funcName), Microsoft.Office.Interop.Excel.XlConsolidationFunction))
                        Utils.ComReleaseHelper.Release(f)
                    Next
                End If
                Utils.ComReleaseHelper.Release(pfAll)
                Utils.ComReleaseHelper.Release(pivotTable)
                If pivotCache IsNot Nothing Then Utils.ComReleaseHelper.Release(pivotCache)
                Utils.ComReleaseHelper.Release(sourceRng)
                If sourceWs IsNot Nothing Then Utils.ComReleaseHelper.Release(sourceWs)
                If destWs IsNot Nothing Then Utils.ComReleaseHelper.Release(destWs)

                result.Content.Add(New McpTextContent With {
                    .Text = $"成功创建数据透视表 {pivotName} 于 {destWs.Name}.{targetCell}。后续 excel_refresh_pivot_table 请带 pivotTableName={pivotName}、sheet={destWs.Name}。"
                })

            Catch ex As Exception
                Dim msgLower = ex.Message.ToLowerInvariant()
                If msgLower.Contains("0x800a03ec") OrElse msgLower.Contains("name") OrElse msgLower.Contains("名称") OrElse msgLower.Contains("已被使用") Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text =
                        "创建数据透视表失败（COM 异常）。原始错误: " & ex.Message &
                        "。可能原因：1) 工作簿中已有同名透视表（已加重试逻辑，仍失败请检查）；" &
                        "2) 目标位置 " & targetCell & " 被占用；" &
                        "3) sourceRange 解析问题。"})
                Else
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "创建数据透视表失败: " & ex.Message})
                End If
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 配置已有数据透视表的字段布局（行/列/值），并可同时更换数据源。
        ''' 【为什么新增】此前透视表只有 create/refresh 两个工具，模型想"调整透视表字段/换数据源"时
        ''' 找不到工具，只能编造 excel_set_pivot_table / excel_move_pivot_table / excel_set_pivot_range
        ''' 等不存在的名字，连环失败十几次（实测截图 2026-10-06）。本工具补上这块能力。
        ''' 字段名不存在时报出全部可用字段，便于模型自纠。
        ''' </summary>
        Private Function ExecuteSetPivotFieldLayout(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()

            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim pivotTableName As String = GetArg(Of String)(args, "pivotTableName", "")
                Dim sourceRange As String = GetArg(Of String)(args, "sourceRange", "")
                Dim rowFieldsJson As String = GetArg(Of String)(args, "rowFields", "[]")
                Dim columnFieldsJson As String = GetArg(Of String)(args, "columnFields", "[]")
                Dim valueFieldsJson As String = GetArg(Of String)(args, "valueFields", "[]")
                Dim valueFunctionsJson As String = GetArg(Of String)(args, "valueFunctions", "[]")
                Dim resetFields As Boolean = GetArg(Of Boolean)(args, "resetFields", True)

                Dim app = Globals.ThisAddIn.Application
                Dim wb = app.ActiveWorkbook
                If wb Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何工作簿"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                If Not String.IsNullOrEmpty(sheetName) Then
                    ws = TryCast(wb.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                Else
                    ws = TryCast(app.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                End If
                If ws Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: {sheetName}"})
                    Return result
                End If

                Dim pts As Microsoft.Office.Interop.Excel.PivotTables = DirectCast(ws.PivotTables, Microsoft.Office.Interop.Excel.PivotTables)
                Dim ptCount As Integer = CInt(pts.Count)
                If ptCount = 0 Then
                    ' 与 refresh 一致：全工作簿找一遍并给出可行动指引（严禁 Release(sh)，sh 可能就是 ws）
                    Dim foundList As New List(Of String)()
                    For Each sh As Microsoft.Office.Interop.Excel.Worksheet In wb.Worksheets
                        Dim pts2 As Microsoft.Office.Interop.Excel.PivotTables = DirectCast(sh.PivotTables, Microsoft.Office.Interop.Excel.PivotTables)
                        If pts2 IsNot Nothing AndAlso CInt(pts2.Count) > 0 Then
                            For i2 As Integer = 1 To CInt(pts2.Count)
                                Dim p2 As Microsoft.Office.Interop.Excel.PivotTable = DirectCast(pts2(i2), Microsoft.Office.Interop.Excel.PivotTable)
                                foundList.Add(sh.Name & " → " & p2.Name)
                                Utils.ComReleaseHelper.Release(p2)
                            Next
                        End If
                        Utils.ComReleaseHelper.Release(pts2)
                    Next
                    If foundList.Count > 0 Then
                        result.Content.Add(New McpTextContent With {.Text = $"工作表『{ws.Name}』上没有数据透视表，但工作簿中存在以下透视表：{String.Join("、", foundList)}。请把 sheet 参数改为对应工作表，必要时带上 pivotTableName 重试。"})
                    Else
                        result.Content.Add(New McpTextContent With {.Text = "工作簿中没有任何数据透视表。请先调用 excel_create_pivot_table 创建。"})
                    End If
                    result.IsError = True
                    Utils.ComReleaseHelper.Release(pts)
                    Utils.ComReleaseHelper.Release(ws)
                    Return result
                End If

                Dim pivot As Microsoft.Office.Interop.Excel.PivotTable = Nothing
                If Not String.IsNullOrEmpty(pivotTableName) Then
                    ' ⚠ PIA 中 PivotTables 索引器参数是 Integer（不接受名称），
                    ' 直接 pts(pivotTableName) 会在 Option Strict On 下报 BC30512（String→Integer）。
                    ' 改为遍历按 Name 匹配（忽略大小写）。
                    Try
                        For iFind As Integer = 1 To ptCount
                            Dim cand As Microsoft.Office.Interop.Excel.PivotTable = DirectCast(pts(iFind), Microsoft.Office.Interop.Excel.PivotTable)
                            Dim candName As String = ""
                            If cand IsNot Nothing Then
                                Try
                                    candName = CStr(cand.Name)
                                Catch
                                    candName = ""
                                End Try
                            End If
                            If String.Equals(candName, pivotTableName, StringComparison.OrdinalIgnoreCase) Then
                                pivot = cand
                                Exit For
                            End If
                            Utils.ComReleaseHelper.Release(cand)
                        Next
                    Catch
                    End Try
                    If pivot Is Nothing Then
                        result.IsError = True
                        Dim knownNames As New List(Of String)()
                        Try
                            For iKnown As Integer = 1 To ptCount
                                Dim cn As Microsoft.Office.Interop.Excel.PivotTable = DirectCast(pts(iKnown), Microsoft.Office.Interop.Excel.PivotTable)
                                If cn IsNot Nothing AndAlso cn.Name IsNot Nothing Then
                                    knownNames.Add(CStr(cn.Name))
                                End If
                                Utils.ComReleaseHelper.Release(cn)
                            Next
                        Catch
                        End Try
                        Dim knownText As String = "(无)"
                        If knownNames.Count > 0 Then knownText = String.Join("、", knownNames)
                        result.Content.Add(New McpTextContent With {.Text = "工作表『" & ws.Name & "』上找不到名为『" & pivotTableName & "』的透视表。该工作表现有透视表：" & knownText})
                        Utils.ComReleaseHelper.Release(pts)
                        Utils.ComReleaseHelper.Release(ws)
                        Return result
                    End If
                End If
                If pivot Is Nothing Then pivot = DirectCast(pts(ptCount), Microsoft.Office.Interop.Excel.PivotTable)

                ' ── 可选：更换数据源 ──
                If Not String.IsNullOrEmpty(sourceRange) Then
                    Dim srcWs As Microsoft.Office.Interop.Excel.Worksheet = ws
                    Dim srcAddr As String = sourceRange
                    If sourceRange.Contains("!") Then
                        Dim parts2 = sourceRange.Split("!"c)
                        srcWs = DirectCast(wb.Worksheets(parts2(0)), Microsoft.Office.Interop.Excel.Worksheet)
                        srcAddr = parts2(1)
                    End If
                    Dim srcRng As Microsoft.Office.Interop.Excel.Range = srcWs.Range(srcAddr)
                    Dim newCache As Microsoft.Office.Interop.Excel.PivotCache = DirectCast(
                        wb.PivotCaches().Create(Microsoft.Office.Interop.Excel.XlPivotTableSourceType.xlDatabase, srcRng),
                        Microsoft.Office.Interop.Excel.PivotCache)
                    pivot.ChangePivotCache(newCache)
                    ' ⚠ 旧 PivotCache 泄漏：Excel COM 中 PivotCache 没有公开 Delete 方法。
                    ' Excel 内部规则是最后一个关联 PivotTable 被删时自动回收，这里只换了 cache，
                    ' 旧 cache 留在 wb.PivotCaches 里成为孤儿。这是 Excel COM 的已知限制，
                    ' 社区无干净解法。单次泄漏可忽略，极高频换数据源时才需要关注。
                    Utils.ComReleaseHelper.Release(srcRng)
                    If srcWs IsNot ws Then Utils.ComReleaseHelper.Release(srcWs)
                End If

                Dim pfColl As Microsoft.Office.Interop.Excel.PivotFields = DirectCast(pivot.PivotFields, Microsoft.Office.Interop.Excel.PivotFields)

                ' 先收集全部可用字段名（重置前），供字段不存在时报错自纠
                Dim available As New List(Of String)()
                For i3 As Integer = 1 To CInt(pfColl.Count)
                    Dim pf0 As Microsoft.Office.Interop.Excel.PivotField = DirectCast(pfColl(i3), Microsoft.Office.Interop.Excel.PivotField)
                    Try
                        If pf0.Name IsNot Nothing Then available.Add(CStr(pf0.Name))
                    Catch
                    End Try
                    Utils.ComReleaseHelper.Release(pf0)
                Next

                Dim rowFields = JsonConvert.DeserializeObject(Of List(Of String))(rowFieldsJson)
                Dim columnFields = JsonConvert.DeserializeObject(Of List(Of String))(columnFieldsJson)
                Dim valueFields = JsonConvert.DeserializeObject(Of List(Of String))(valueFieldsJson)
                Dim valueFunctions = JsonConvert.DeserializeObject(Of List(Of String))(valueFunctionsJson)

                ' ═══════════════════════════════════════════════════════════════
                ' Phase 1 (dry-run)：全量校验字段存在性，不动透视表。
                ' 任一字段 missing 直接返回，透视表保持原状。
                ' 这防止了 resetFields=True 时先清空再失败 → 留下空壳。
                ' ═══════════════════════════════════════════════════════════════
                Dim missing As New List(Of String)()

                If rowFields IsNot Nothing Then
                    For Each name In rowFields
                        If String.IsNullOrWhiteSpace(name) Then Continue For
                        Dim f = ResolvePivotFieldLoose(pfColl, name)
                        If f Is Nothing Then missing.Add(name) Else Utils.ComReleaseHelper.Release(f)
                    Next
                End If
                If columnFields IsNot Nothing Then
                    For Each name In columnFields
                        If String.IsNullOrWhiteSpace(name) Then Continue For
                        Dim f = ResolvePivotFieldLoose(pfColl, name)
                        If f Is Nothing Then missing.Add(name) Else Utils.ComReleaseHelper.Release(f)
                    Next
                End If
                If valueFields IsNot Nothing Then
                    For Each name In valueFields
                        If String.IsNullOrWhiteSpace(name) Then Continue For
                        Dim f = ResolvePivotFieldLoose(pfColl, name)
                        If f Is Nothing Then missing.Add(name) Else Utils.ComReleaseHelper.Release(f)
                    Next
                End If

                If missing.Count > 0 Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "透视表『" & TryGetPivotName(pivot) & "』中没有这些字段：" & String.Join("、", missing) & "。该透视表可用字段：" & String.Join("、", available) & "。透视表布局未被改动，请改用上述真实字段名重试。"})
                    Utils.ComReleaseHelper.Release(pivot)
                    Utils.ComReleaseHelper.Release(pfColl)
                    Utils.ComReleaseHelper.Release(pts)
                    Utils.ComReleaseHelper.Release(ws)
                    Return result
                End If

                ' ═══════════════════════════════════════════════════════════════
                ' Phase 2 (apply)：所有字段都校验通过 → 原子切换
                ' ═══════════════════════════════════════════════════════════════
                If resetFields Then
                    For i4 As Integer = 1 To CInt(pfColl.Count)
                        Dim pf1 As Microsoft.Office.Interop.Excel.PivotField = DirectCast(pfColl(i4), Microsoft.Office.Interop.Excel.PivotField)
                        Try : pf1.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlHidden : Catch : End Try
                        Utils.ComReleaseHelper.Release(pf1)
                    Next
                    Try
                        Dim dfsColl As Microsoft.Office.Interop.Excel.PivotFields = DirectCast(pivot.DataFields, Microsoft.Office.Interop.Excel.PivotFields)
                        For i5 As Integer = CInt(dfsColl.Count) To 1 Step -1
                            Dim df As Microsoft.Office.Interop.Excel.PivotField = DirectCast(dfsColl(i5), Microsoft.Office.Interop.Excel.PivotField)
                            Try : df.Delete() : Catch : End Try
                            Utils.ComReleaseHelper.Release(df)
                        Next
                        Utils.ComReleaseHelper.Release(dfsColl)
                    Catch
                    End Try
                End If

                If rowFields IsNot Nothing Then
                    For i = 0 To rowFields.Count - 1
                        Dim f = ResolvePivotFieldLoose(pfColl, rowFields(i))
                        f.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField
                        f.Position = i + 1
                        Utils.ComReleaseHelper.Release(f)
                    Next
                End If
                If columnFields IsNot Nothing Then
                    For i = 0 To columnFields.Count - 1
                        Dim f = ResolvePivotFieldLoose(pfColl, columnFields(i))
                        f.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlColumnField
                        f.Position = i + 1
                        Utils.ComReleaseHelper.Release(f)
                    Next
                End If
                If valueFields IsNot Nothing Then
                    For i = 0 To valueFields.Count - 1
                        Dim f = ResolvePivotFieldLoose(pfColl, valueFields(i))
                        Dim funcName = If(valueFunctions IsNot Nothing AndAlso i < valueFunctions.Count, valueFunctions(i), "Sum")
                        pivot.AddDataField(f, , DirectCast(GetPivotFunctionEnum(funcName), Microsoft.Office.Interop.Excel.XlConsolidationFunction))
                        Utils.ComReleaseHelper.Release(f)
                    Next
                End If

                pivot.RefreshTable()
                result.Content.Add(New McpTextContent With {.Text = $"已更新透视表 {TryGetPivotName(pivot)}（{ws.Name}）布局：行={If(rowFields Is Nothing, 0, rowFields.Count)} 个、列={If(columnFields Is Nothing, 0, columnFields.Count)} 个、值={If(valueFields Is Nothing, 0, valueFields.Count)} 个" & If(String.IsNullOrEmpty(sourceRange), "", $"，数据源已改为 {sourceRange}")})

                Utils.ComReleaseHelper.Release(pivot)
                Utils.ComReleaseHelper.Release(pfColl)
                Utils.ComReleaseHelper.Release(pts)
                Utils.ComReleaseHelper.Release(ws)

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"配置透视表字段失败: {ex.Message}"})
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 透视表字段宽松解析：① 精确名（忽略大小写）→ ② 包含匹配 → ③ 列字母（A/B/C…）按
        ''' 源数据列序定位。找不到返回 Nothing（调用方负责报出自愈信息）。
        ''' 背景：弱模型高频传列字母或幻觉字段名（实测传"月份"而表头是"日期"，2026-10-06），
        ''' 旧实现 PivotFields("xxx") 直接抛裸 COM 异常且留下半成品透视表。
        ''' 注意：列字母按 PivotFields 的索引 = 源数据列顺序（1 起始）定位。
        ''' 增强：①② 匹配同时检查 f.Name 和 f.SourceName（PivotField.Name 可能为空或"字段N"，
        ''' SourceName 更可靠返回源列名）。
        ''' </summary>
        Private Shared Function ResolvePivotFieldLoose(pfColl As Microsoft.Office.Interop.Excel.PivotFields, spec As String, Optional sourceHeaders() As String = Nothing) As Microsoft.Office.Interop.Excel.PivotField
            If pfColl Is Nothing OrElse String.IsNullOrEmpty(spec) Then Return Nothing
            Dim specTrim = spec.Trim()
            If String.IsNullOrEmpty(specTrim) Then Return Nothing
            Try
                ' ① 精确名（忽略大小写）——同时查 Name 和 SourceName
                For i As Integer = 1 To CInt(pfColl.Count)
                    Dim f As Microsoft.Office.Interop.Excel.PivotField = DirectCast(pfColl(i), Microsoft.Office.Interop.Excel.PivotField)
                    Dim n As String = ""
                    Dim s As String = ""
                    Try : n = CStr(f.Name) : Catch : End Try
                    Try : s = CStr(f.SourceName) : Catch : End Try
                    If String.Equals(n, specTrim, StringComparison.OrdinalIgnoreCase) OrElse
                       String.Equals(s, specTrim, StringComparison.OrdinalIgnoreCase) Then
                        Return f ' 找到：不释放（调用方用完释放）
                    End If
                    Utils.ComReleaseHelper.Release(f)
                Next
                ' ② 包含匹配（互为包含，先到先得）——同时查 Name 和 SourceName
                For i As Integer = 1 To CInt(pfColl.Count)
                    Dim f As Microsoft.Office.Interop.Excel.PivotField = DirectCast(pfColl(i), Microsoft.Office.Interop.Excel.PivotField)
                    Dim n As String = ""
                    Dim s As String = ""
                    Try : n = CStr(f.Name) : Catch : End Try
                    Try : s = CStr(f.SourceName) : Catch : End Try
                    Dim namesToCheck() As String = {n, s}
                    For Each cand In namesToCheck
                        If cand <> "" AndAlso (cand.IndexOf(specTrim, StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                                              specTrim.IndexOf(cand, StringComparison.OrdinalIgnoreCase) >= 0) Then
                            Utils.ComReleaseHelper.Release(f) ' 只释放一次
                            Return f
                        End If
                    Next
                    Utils.ComReleaseHelper.Release(f)
                Next
                ' ③ sourceHeaders fallback（如果调用方传了源列表头，用表头值做精确/包含匹配）
                If sourceHeaders IsNot Nothing Then
                    For i As Integer = 1 To CInt(pfColl.Count)
                        Dim f As Microsoft.Office.Interop.Excel.PivotField = DirectCast(pfColl(i), Microsoft.Office.Interop.Excel.PivotField)
                        Dim idx As Integer = i - 1
                        If idx >= 0 AndAlso idx < sourceHeaders.Length Then
                            Dim h As String = sourceHeaders(idx)
                            If String.Equals(h, specTrim, StringComparison.OrdinalIgnoreCase) Then
                                Return f
                            End If
                        End If
                        Utils.ComReleaseHelper.Release(f)
                    Next
                End If
                ' ④ 列字母 → 源列序号（PivotFields 索引即源数据列顺序）
                If System.Text.RegularExpressions.Regex.IsMatch(specTrim, "^[A-Za-z]{1,3}$") Then
                    Dim idx As Integer = ExcelMcpCoreUtils.ColumnLetterToNumber(specTrim)
                    If idx >= 1 AndAlso idx <= CInt(pfColl.Count) Then
                        Return DirectCast(pfColl(idx), Microsoft.Office.Interop.Excel.PivotField)
                    End If
                End If
            Catch
                ' 异常发生：让调用方收到 Nothing 走 missingFields 流程自愈
                Return Nothing
            End Try
            Return Nothing
        End Function

        ''' <summary>字段查找委托：统一走 ResolvePivotFieldLoose（精确→包含→列字母），
        ''' 保证 CreatePivot / SetLayout / CreateSlicer 三个工具对同一透视表的字段解析语义一致。</summary>
        Private Shared Function FindPivotField(pfColl As Microsoft.Office.Interop.Excel.PivotFields, fieldName As String) As Microsoft.Office.Interop.Excel.PivotField
            Return ResolvePivotFieldLoose(pfColl, fieldName)
        End Function

        ''' <summary>安全读取 PivotTable.Name，COM 对象边缘状态下也不抛异常。</summary>
        Private Shared Function TryGetPivotName(pt As Microsoft.Office.Interop.Excel.PivotTable) As String
            If pt Is Nothing Then Return ""
            Try
                Return CStr(pt.Name)
            Catch
                Return ""
            End Try
        End Function

        Private Function ExecuteCreateSlicer(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Dim progress As New List(Of String)()
            progress.Add("解析参数")
            Try
                Dim sourceRange As String = GetArg(Of String)(args, "sourceRange", "")
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim pivotSheetName As String = GetArg(Of String)(args, "pivotSheet", "")
                Dim slicerDimensionsJson As String = GetArg(Of String)(args, "slicerDimensions", "[]")
                Dim pivotTablesJson As String = GetArg(Of String)(args, "pivotTables", "[]")

                If String.IsNullOrEmpty(sourceRange) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "源数据范围不能为空"})
                    Return result
                End If

                Dim app = Globals.ThisAddIn.Application
                Dim wb = app.ActiveWorkbook
                If wb Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何工作簿"})
                    Return result
                End If

                ' 解析源范围（支持 "Sheet1!A1:G101" 或纯范围）
                Dim sourceWs As Microsoft.Office.Interop.Excel.Worksheet
                Dim sourceAddr As String
                If sourceRange.Contains("!") Then
                    Dim parts = sourceRange.Split("!"c)
                    sourceWs = DirectCast(wb.Worksheets(parts(0)), Microsoft.Office.Interop.Excel.Worksheet)
                    sourceAddr = parts(1)
                Else
                    sourceWs = DirectCast(app.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    sourceAddr = sourceRange
                End If
                Dim sourceRng = sourceWs.Range(sourceAddr)

                ' 目标工作表（切片器所在：看板工作表）
                Dim destWs As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                If Not String.IsNullOrEmpty(sheetName) Then
                    Try
                        destWs = DirectCast(wb.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                    Catch
                        destWs = Nothing
                    End Try
                    ' 传了工作表名但找不到 → 显式报错，禁止静默回退 ActiveSheet（否则切片器会放错表）
                    If destWs Is Nothing Then
                        result.IsError = True
                        Dim _activeName As String = ""
                        Try : _activeName = DirectCast(app.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet).Name : Catch : End Try
                        result.Content.Add(New McpTextContent With {.Text = "找不到工作表: [" & sheetName & "]，请先创建该工作表。当前活动表: [" & _activeName & "]"})
                        Return result
                    End If
                Else
                    destWs = DirectCast(app.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                End If

                ' ═══════════════════════════════════════════════════════════════
                ' 前置清理：在所有 Shapes 创建前彻底清掉看板单元格格式
                ' （源数据可能带黄色高亮/条件格式/边框 → 必须在 shape fill 之前清除！）
                ' ═══════════════════════════════════════════════════════════════
                Try
                    Dim _preClear As Microsoft.Office.Interop.Excel.Range = destWs.Cells
                    _preClear.Interior.Color = Microsoft.Office.Interop.Excel.Constants.xlNone
                    Try : _preClear.FormatConditions.Delete() : Catch : End Try
                    Try : _preClear.Borders.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone : Catch : End Try
                    Try : _preClear.Font.Color = RGB(0, 0, 0) : Catch : End Try
                    progress.Add("  ✅ 看板前置清理: 单元格底色/条件格式/边框已清除")
                Catch
                End Try

                ' 清理旧"数据透视表"残留（旧 ExecuteCreatePivotTable 的默认产物，避免空透视表干扰看板）
                Try
                    Dim legacyPt As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(wb.Worksheets("数据透视表"), Microsoft.Office.Interop.Excel.Worksheet)
                    If legacyPt.Name <> destWs.Name Then
                        app.DisplayAlerts = False
                        legacyPt.Delete()
                        app.DisplayAlerts = True
                    End If
                Catch
                    app.DisplayAlerts = True
                End Try

                ' 清理旧"透视缓存"残留（旧版本将透视表放在独立 sheet，现统一放在看板）
                Try
                    Dim oldCache As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(wb.Worksheets("透视缓存"), Microsoft.Office.Interop.Excel.Worksheet)
                    app.DisplayAlerts = False
                    oldCache.Delete()
                    app.DisplayAlerts = True
                Catch
                    app.DisplayAlerts = True
                End Try

                ' 透视缓存：本次调用内所有透视表共享同一个 cache（切片器联动的必要条件）。
                ' 透视表的目标 sheet：
                '   - 有 pivotSheetName → 独立工作表（不干扰看板布局）
                '   - 没有 → 回退 destWs（看板工作表，兼容旧调用）
                ' 【二次生成必清】旧 _dash_pivot 上残留 Pivot1 会导致 CreatePivotTable 抛 0x800A03EC；
                ' 看板表 Cells.Clear 不会删 PivotTable/Slicer，必须在此彻底重建透视缓存表并清旧切片缓存。
                Dim pivotDestWs As Microsoft.Office.Interop.Excel.Worksheet = destWs
                If Not String.IsNullOrEmpty(pivotSheetName) Then
                    Dim _prevAlertsPt As Boolean = app.DisplayAlerts
                    app.DisplayAlerts = False
                    Try
                        ' 删旧透视缓存表（含残留透视表）
                        For Each _oldName As String In New String() {pivotSheetName, "透视缓存"}
                            Try
                                Dim _oldWs As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(wb.Worksheets(_oldName), Microsoft.Office.Interop.Excel.Worksheet)
                                If _oldWs IsNot Nothing AndAlso _oldWs.Name <> destWs.Name Then
                                    ' 先拆透视表再删表，避免 COM 挂起
                                    Try
                                        Dim _pts As Microsoft.Office.Interop.Excel.PivotTables = DirectCast(_oldWs.PivotTables(), Microsoft.Office.Interop.Excel.PivotTables)
                                        For _pi As Integer = _pts.Count To 1 Step -1
                                            Try
                                                Dim _opt As Microsoft.Office.Interop.Excel.PivotTable = DirectCast(_pts.Item(_pi), Microsoft.Office.Interop.Excel.PivotTable)
                                                _opt.TableRange2.Clear()
                                            Catch
                                            End Try
                                        Next
                                    Catch
                                    End Try
                                    _oldWs.Delete()
                                    progress.Add("  ✅ 已删除旧透视表工作表: " & _oldName)
                                End If
                            Catch
                            End Try
                        Next
                    Finally
                        app.DisplayAlerts = _prevAlertsPt
                    End Try

                    ' 清掉看板侧残留切片器形状 + 工作簿级 SlicerCache（否则二次生成会看到"假切片"）
                    Try
                        Dim _shps As Microsoft.Office.Interop.Excel.Shapes = destWs.Shapes
                        For _si As Integer = _shps.Count To 1 Step -1
                            Try
                                Dim _shp As Microsoft.Office.Interop.Excel.Shape = DirectCast(_shps.Item(_si), Microsoft.Office.Interop.Excel.Shape)
                                If _shp.Type = Microsoft.Office.Core.MsoShapeType.msoSlicer Then _shp.Delete()
                            Catch
                            End Try
                        Next
                    Catch
                    End Try
                    Try
                        For _sci As Integer = wb.SlicerCaches.Count To 1 Step -1
                            Try
                                Dim _osc As Microsoft.Office.Interop.Excel.SlicerCache = DirectCast(wb.SlicerCaches.Item(_sci), Microsoft.Office.Interop.Excel.SlicerCache)
                                Dim _scn As String = ""
                                Try : _scn = _osc.Name : Catch : End Try
                                If _scn.StartsWith("Slicer_", StringComparison.OrdinalIgnoreCase) OrElse _scn.StartsWith("切片器", StringComparison.OrdinalIgnoreCase) Then
                                    _osc.Delete()
                                End If
                            Catch
                            End Try
                        Next
                        progress.Add("  ✅ 已清理旧 SlicerCache / 切片器形状")
                    Catch
                    End Try

                    Try
                        pivotDestWs = DirectCast(wb.Worksheets.Add(), Microsoft.Office.Interop.Excel.Worksheet)
                        pivotDestWs.Name = pivotSheetName
                        ' 创建期保持可见：隐藏承载 PivotTable 的 Sheet 会导致切片器一并消失
                        pivotDestWs.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible
                        progress.Add("  ✅ 新建透视缓存表: " & pivotSheetName)
                    Catch
                        Try
                            pivotDestWs = DirectCast(wb.Worksheets(pivotSheetName), Microsoft.Office.Interop.Excel.Worksheet)
                        Catch
                            pivotDestWs = destWs
                        End Try
                    End Try
                    ' 激活看板表，避免后续形状/图表画到透视缓存表
                    Try : destWs.Activate() : Catch : End Try
                End If

                Dim pivotCache As Microsoft.Office.Interop.Excel.PivotCache = DirectCast(
                    wb.PivotCaches().Create(Microsoft.Office.Interop.Excel.XlPivotTableSourceType.xlDatabase, sourceRng),
                    Microsoft.Office.Interop.Excel.PivotCache)

                ' 创建透视表（每图表一个：行=维度，值=数值列 + 聚合函数）
                Dim firstPivot As Microsoft.Office.Interop.Excel.PivotTable = Nothing
                Dim createdPivots As New List(Of Microsoft.Office.Interop.Excel.PivotTable)()
                Dim pivotInfos As New List(Of Dictionary(Of String, Object))()
                Dim warnings As New List(Of String)()
                Dim jpivs = JArray.Parse(If(String.IsNullOrWhiteSpace(pivotTablesJson), "[]", pivotTablesJson))
                Dim ptIndex = 0
                Dim nextPivotRow As Integer = 1
                Dim ptChartOkCount As Integer = 0
                For Each jp In jpivs
                    ptIndex += 1
                    ' STA 间隙：多个 PivotTable 共享同一 PivotCache 时，
                    ' CreatePivotTable 之间必须让出消息泵让 Excel COM 收敛（VSTO 在 STA 上）。
                    Thread.Sleep(200)
                    Try : Application.DoEvents() : Catch : End Try
                    Dim dimName As String = If(jp("dimension") IsNot Nothing, jp("dimension").ToString(), "")
                    Dim valName As String = If(jp("value") IsNot Nothing, jp("value").ToString(), "")
                    Dim funcName As String = If(jp("func") IsNot Nothing, jp("func").ToString(), "sum")
                    ' ⭐️ 忽略 JS 端硬编码 targetCell（pivotRowStep=20 不知道维度值数量，会导致重叠！）
                    '    用 VB 动态计算的 nextPivotRow，创建完上一张后用 TableRange2 算实际行数
                    Dim targetCell As String = "A" & nextPivotRow
                    Dim ptName As String = "Pivot" & ptIndex
                    If jp("name") IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(jp("name").ToString()) Then
                        ptName = jp("name").ToString()
                    End If

                    progress.Add(String.Format("创建透视表 {0} @ {1}.{2}", ptName, pivotDestWs.Name, targetCell))
                    Dim pt As Microsoft.Office.Interop.Excel.PivotTable = Nothing
                    Try
                        pt = pivotCache.CreatePivotTable(pivotDestWs.Range(targetCell), ptName)
                    Catch exPt As Exception
                        ' 0x800A03EC：常见于同名透视表残留 / 目标格被占用 → 清名后换唯一名重试
                        progress.Add(String.Format("  ⚠ CreatePivotTable 首次失败: {0} → 清残留并重试", exPt.Message))
                        Try
                            For Each _wsPt As Microsoft.Office.Interop.Excel.Worksheet In wb.Worksheets
                                Try
                                    Dim _pts2 As Microsoft.Office.Interop.Excel.PivotTables = DirectCast(_wsPt.PivotTables(), Microsoft.Office.Interop.Excel.PivotTables)
                                    For _pj As Integer = _pts2.Count To 1 Step -1
                                        Try
                                            Dim _existPt As Microsoft.Office.Interop.Excel.PivotTable = DirectCast(_pts2.Item(_pj), Microsoft.Office.Interop.Excel.PivotTable)
                                            If String.Equals(_existPt.Name, ptName, StringComparison.OrdinalIgnoreCase) Then
                                                _existPt.TableRange2.Clear()
                                            End If
                                        Catch
                                        End Try
                                    Next
                                Catch
                                End Try
                            Next
                        Catch
                        End Try
                        Try : pivotDestWs.Range(targetCell).Resize(40, 8).Clear() : Catch : End Try
                        ptName = ptName & "_" & DateTime.Now.ToString("HHmmssfff")
                        Thread.Sleep(150)
                        Try : Application.DoEvents() : Catch : End Try
                        pt = pivotCache.CreatePivotTable(pivotDestWs.Range(targetCell), ptName)
                        progress.Add(String.Format("  ✅ 重试成功，透视表名改为 {0}", ptName))
                    End Try
                    If Not String.IsNullOrEmpty(dimName) Then
                        Dim rowF As Microsoft.Office.Interop.Excel.PivotField = FindPivotField(pt, dimName)
                        If rowF Is Nothing Then
                            warnings.Add(String.Format("透视表 {0}：维度字段 [{1}] 未在数据源中找到，已跳过", ptName, dimName))
                        Else
                            rowF.Orientation = Microsoft.Office.Interop.Excel.XlPivotFieldOrientation.xlRowField
                            rowF.Position = 1
                        End If
                    End If
                    If Not String.IsNullOrEmpty(valName) Then
                        Dim valF As Microsoft.Office.Interop.Excel.PivotField = FindPivotField(pt, valName)
                        If valF Is Nothing Then
                            warnings.Add(String.Format("透视表 {0}：数值字段 [{1}] 未在数据源中找到，已跳过", ptName, valName))
                        Else
                            Dim funcEnum As Integer = GetPivotFunctionEnum(funcName)
                            pt.AddDataField(valF, , DirectCast(funcEnum, Microsoft.Office.Interop.Excel.XlConsolidationFunction))
                        End If
                    End If
                    pt.ShowTableStyleRowStripes = False
                    pt.ManualUpdate = False
                    pt.RefreshTable()
                    Try : System.Windows.Forms.Application.DoEvents() : Catch : End Try
                    createdPivots.Add(pt)
                    If firstPivot Is Nothing Then firstPivot = pt

                    ' ⭐️ 动态计算下一张透视表的起始行（TableRange2 包含筛选字段、表头、数据、合计）
                    '    JS 端硬编码 pivotRowStep=20 会导致重叠（维度值多的部门表可能占 30+ 行）
                    Try
                        Dim ptStartRow As Integer = CInt(pt.TableRange2.Row)
                        Dim ptTotalRows As Integer = CInt(pt.TableRange2.Rows.Count)
                        nextPivotRow = ptStartRow + ptTotalRows + 2  ' 空两行防重叠
                        progress.Add(String.Format("  → 透视表实际占 {0} 行，下一张从 A{1} 开始", ptTotalRows, nextPivotRow))
                    Catch exRow As Exception
                        nextPivotRow += 20  ' 兜底
                        progress.Add(String.Format("  → 行数计算失败（{0}），兜底跳 20 行到 A{1}", exRow.Message.Substring(0, 40), nextPivotRow))
                    End Try

                    ' ═══════════════════════════════════════════════════
                    ' 【一步到位】每张透视表创建后立刻创建 PivotChart
                    ' 用户骂得对：不该拆成 CreateSlicer + 后面 N 次 CreateChart 两步，
                    ' 中间靠 pivotSheet + pivotTable 名字字符串互相对接——太脆弱，只要有一处写错就全崩。
                    ' 现在 pivotTables JSON 里每张透视表 entry 自带 {chartType, title, anchor, ...}，
                    ' VB 拿到 PivotTable 对象后直接 pt.CreatePivotChart()，没有中间环节，100% 不会找不到引用。
                    ' ═══════════════════════════════════════════════════
                    Dim _chartType As String = If(jp("chartType") IsNot Nothing, jp("chartType").ToString(), "xlColumnClustered")
                    Dim _chartTitle As String = If(jp("chartTitle") IsNot Nothing, jp("chartTitle").ToString(), ptName)
                    Dim _anchorAddr As String = If(jp("anchor") IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(jp("anchor").ToString()), jp("anchor").ToString(), "")
                    Dim _anchorSheet As String = If(jp("anchorSheet") IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(jp("anchorSheet").ToString()), jp("anchorSheet").ToString(), "")
                    Dim _chartStyle As Integer = If(jp("chartStyle") IsNot Nothing, CInt(jp("chartStyle")), 201)
                    Dim _hasLegend As Boolean = If(jp("hasLegend") Is Nothing, True, CBool(jp("hasLegend")))
                    Dim _showDataLabels As Boolean = If(jp("showDataLabels") Is Nothing, False, CBool(jp("showDataLabels")))
                    Dim _dataLabelPosition As String = If(jp("dataLabelPosition") IsNot Nothing, jp("dataLabelPosition").ToString(), "")

                    If Not String.IsNullOrEmpty(_anchorAddr) Then
                        Try
                            ' anchorRange 支持 "Sheet!A1:B10" 格式，先剥 sheet 前缀
                            Dim _anchorWs As Microsoft.Office.Interop.Excel.Worksheet = destWs
                            Dim _anchorA1 As String = _anchorAddr
                            If _anchorAddr.Contains("!") Then
                                Dim _parts = _anchorAddr.Split("!"c)
                                Try : _anchorWs = DirectCast(wb.Worksheets(_parts(0)), Microsoft.Office.Interop.Excel.Worksheet) : Catch : _anchorWs = destWs : End Try
                                _anchorA1 = _parts(1)
                            ElseIf Not String.IsNullOrEmpty(_anchorSheet) Then
                                Try : _anchorWs = DirectCast(wb.Worksheets(_anchorSheet), Microsoft.Office.Interop.Excel.Worksheet) : Catch : _anchorWs = destWs : End Try
                            End If

                            Dim _xlChartType As Microsoft.Office.Interop.Excel.XlChartType = MapChartType(_chartType)
                            Dim _chartTypeInt As Integer = CInt(_xlChartType)
                            progress.Add(String.Format("创建透视图 {0} @ {1}.{2}", ptName, _anchorWs.Name, _anchorAddr))

                            ' ════════════════════════════════════════════════════════════
                            ' CreatePivotChart 修复：COM 后绑定对参数类型非常敏感！
                            '
                            ' 方案 A：PivotCache.CreatePivotChart(Worksheet, ChartType, Left, Top, Width, Height)
                            '   枚举参数必须 CInt 转底层 Integer，不能传托管枚举类型
                            '   Double 参数直接传，不要 CDbl 再包一层
                            '   用 CallByName 而非 InvokeMember（COM 后绑定更稳）
                            '
                            ' 方案 B（回退）：PivotTable.CreatePivotChart(ChartType, Missing, XlChartLocation)
                            '   Source 参数传 Type.Missing 让 Excel 自己从 PivotTable 拿数据源
                            ' ════════════════════════════════════════════════════════════
                            Dim _anchorRng As Microsoft.Office.Interop.Excel.Range = _anchorWs.Range(_anchorA1)
                            Dim _pivotChart As Microsoft.Office.Interop.Excel.Chart = Nothing
                            Dim _chartCreated As Boolean = False

                            ' ════════════════════════════════════════════════════════════
                            ' CreatePivotChart — 唯一有效方法：PivotCache.CreatePivotChart
                            ' 反射签名: Shape CreatePivotChart(Object ChartDestination,
                            '   Object XlChartType, Object Left, Object Top, Object Width, Object Height)
                            ' 所有参数都是 Object → Option Strict On 下强类型直接调用 ✓
                            ' PivotTable.CreatePivotChart 在当前 Interop 中不存在 → 方案 B 删除
                            ' ════════════════════════════════════════════════════════════
                            Try
                                Dim _pc As Microsoft.Office.Interop.Excel.PivotCache = pt.PivotCache
                                progress.Add("  [DIAG] CreatePivotChart: pt.PivotCache OK")
                                ' 先强制刷新 PivotCache
                                Try : _pc.Refresh() : Catch : End Try
                                Try : pt.RefreshTable() : Catch : End Try
                                ' ╔══════════════════════════════════════════════════════════╗
                                ' ║ WaitForPivotTableReady：轮询 TableRange2 可访问，       ║
                                ' ║ 替代固定 Sleep(300)。Excel RefreshTable 是异步的，       ║
                                ' ║ TableRange2 没就绪就访问会抛 RPC_E_SERVERFAULT。          ║
                                ' ║ 超时 5 秒 → 兜底继续（ChartWizard 会自行重试）。          ║
                                ' ╚══════════════════════════════════════════════════════════╝
                                Dim _readySw = Stopwatch.StartNew()
                                While _readySw.Elapsed < TimeSpan.FromSeconds(5)
                                    Try
                                        Dim _addrTest = pt.TableRange2.Address(False, False)
                                        If Not String.IsNullOrEmpty(_addrTest) Then Exit While
                                    Catch
                                    End Try
                                    Application.DoEvents()
                                    Thread.Sleep(50)
                                End While

                                ' ═══════ 唯一正确方案（经 PowerShell COM 验证） ═══════
                                ' PivotCache.CreatePivotChart 在当前 Excel Interop 版本中不存在！
                                ' 改用 Shapes.AddChart + ChartWizard(TableRange2) 自动绑定 PivotLayout
                                ' 这是 Excel 官方推荐的 PivotChart 创建方式，且生成的图能被 Slicer 联动
                                Dim _leftObj As Object = CDbl(_anchorRng.Left)
                                Dim _topObj As Object = CDbl(_anchorRng.Top)
                                Dim _wObj As Object = CDbl(_anchorRng.Width)
                                Dim _hObj As Object = CDbl(_anchorRng.Height)

                                Try
                                    ' === 完全复刻 PowerShell 成功路径 ===
                                    progress.Add("  [DIAG] Shapes.AddChart()...")
                                    Dim _chartShape As Microsoft.Office.Interop.Excel.Shape = _anchorWs.Shapes.AddChart()
                                    Dim _chartObj As Microsoft.Office.Interop.Excel.Chart = _chartShape.Chart
                                    progress.Add(String.Format("  [DIAG] AddChart OK, setting ChartType={0} (int={1})...", _xlChartType, _chartTypeInt))
                                    _chartObj.ChartType = _xlChartType
                                    progress.Add("  [DIAG] pt.TableRange2 = " & pt.TableRange2.Address(False, False))
                                    ' === ChartWizard 所有参数用 Object 后绑定 ===
                                    Dim _cwSrc As Object = pt.TableRange2
                                    Dim _cwType As Object = CObj(_chartTypeInt)
                                    CallByName(_chartObj, "ChartWizard", CallType.Method, _cwSrc, _cwType)
                                    progress.Add("  [DIAG] ChartWizard OK, setting coords on Shape...")
                                    ' ⚠ ChartWizard 把图变成 PivotChart 后自动锚定到 PivotTable 位置！
                                    ' 必须先把 Placement 改成 xlFreeFloating(3)，否则坐标设置全部被忽略！
                                    Try : _chartShape.Placement = CType(3, Microsoft.Office.Interop.Excel.XlPlacement) : Catch
                                        Try : _chartShape.GetType().InvokeMember("Placement", System.Reflection.BindingFlags.SetProperty, Nothing, _chartShape, New Object() {3}) : Catch : End Try
                                    End Try
                                    _chartShape.Left = CSng(CDbl(_anchorRng.Left))
                                    _chartShape.Top = CSng(CDbl(_anchorRng.Top))
                                    _chartShape.Width = CSng(CDbl(_anchorRng.Width))
                                    _chartShape.Height = CSng(CDbl(_anchorRng.Height))
                                    progress.Add(String.Format("  [DIAG] Shape coords: L={0}, T={1}, W={2}, H={3}", _chartShape.Left, _chartShape.Top, _chartShape.Width, _chartShape.Height))
                                    ' 验证 PivotLayout
                                    Dim _pl As Object = Nothing
                                    Try : _pl = _chartObj.PivotLayout : Catch : End Try
                                    _pivotChart = _chartObj
                                    _chartCreated = True
                                    ptChartOkCount += 1
                                    progress.Add(String.Format("  ✅ PivotChart OK! PivotLayout={0}", If(_pl IsNot Nothing, "True", "False")))
                                Catch exChart As Exception
                                    warnings.Add(String.Format("  ❌ PivotChart FAIL: {0}", exChart.Message))
                                End Try
                                ' PivotChart 创建后 COM 对象需要一个微秒级的稳定窗口
                                Thread.Sleep(150)
                                Try : Application.DoEvents() : Catch : End Try
                            Catch exPC As Exception
                                warnings.Add("OUTER FAIL: " & exPC.Message)
                            End Try

                            ' 创建完后 DoEvents 让 Excel 处理完 COM 对象就绪
                            Try : System.Windows.Forms.Application.DoEvents() : Catch : End Try

                            If _pivotChart Is Nothing Then
                                warnings.Add("⚠ _pivotChart 为 Nothing — 跳过 PivotLayout 和 UI 美化")
                            Else

                            ' ── PivotChart UI 清理（反射绕过 Option Strict On） ──
                            Try
                                Dim _pl As Microsoft.Office.Interop.Excel.PivotLayout = Nothing
                                Try : _pl = CType(_pivotChart.PivotLayout, Microsoft.Office.Interop.Excel.PivotLayout) : Catch : _pl = Nothing : End Try
                                If _pl IsNot Nothing Then
                                    ' (a) DataFields Caption 去 "求和项:" 等前缀
                                    Try
                                        Dim _dfObj As Object = _pl.DataFields
                                        If _dfObj IsNot Nothing Then
                                            Dim _dfCnt As Integer = CInt(_dfObj.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, Nothing, _dfObj, Nothing))
                                            For _di As Integer = 1 To _dfCnt
                                                Try
                                                    Dim _df As Object = _dfObj.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, Nothing, _dfObj, New Object() {_di})
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

                                    ' (b) 隐藏副标题（= Value Field Caption）
                                    Try : _pivotChart.GetType().InvokeMember("HasSubtitle", System.Reflection.BindingFlags.SetProperty, Nothing, _pivotChart, New Object() {False}) : Catch : End Try

                                    ' (c) Location 后重新隐藏所有字段按钮
                                    Try : _pivotChart.GetType().InvokeMember("ShowAllFieldButtons", System.Reflection.BindingFlags.SetProperty, Nothing, _pivotChart, New Object() {False}) : Catch : End Try
                                    Try : _pl.GetType().InvokeMember("ShowReportFilterFieldButtons", System.Reflection.BindingFlags.SetProperty, Nothing, _pl, New Object() {False}) : Catch : End Try
                                    Try : _pl.GetType().InvokeMember("ShowLegendFieldButtons", System.Reflection.BindingFlags.SetProperty, Nothing, _pl, New Object() {False}) : Catch : End Try
                                    Try : _pl.GetType().InvokeMember("ShowAxisFieldButtons", System.Reflection.BindingFlags.SetProperty, Nothing, _pl, New Object() {False}) : Catch : End Try
                                    Try : _pl.GetType().InvokeMember("ShowValueFieldButtons", System.Reflection.BindingFlags.SetProperty, Nothing, _pl, New Object() {False}) : Catch : End Try
                                End If
                            Catch : End Try

                            ' ── 通用图表美化（非反射） ──
                            Try
                                _pivotChart.HasTitle = True
                                _pivotChart.ChartTitle.Text = _chartTitle
                                _pivotChart.ChartStyle = _chartStyle
                                _pivotChart.HasLegend = _hasLegend
                                If _showDataLabels Then
                                    _pivotChart.HasDataTable = False
                                    ' SeriesCollection 是 COM 集合，Option Strict On 下用反射
                                    Try
                                        Dim _scObj As Object = _pivotChart.GetType().InvokeMember("SeriesCollection", System.Reflection.BindingFlags.GetProperty, Nothing, _pivotChart, Nothing)
                                        If _scObj IsNot Nothing Then
                                            Dim _s1 As Object = _scObj.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, Nothing, _scObj, New Object() {1})
                                            If _s1 IsNot Nothing Then
                                                Try : _s1.GetType().InvokeMember("HasDataLabels", System.Reflection.BindingFlags.SetProperty, Nothing, _s1, New Object() {True}) : Catch : End Try
                                                If Not String.IsNullOrEmpty(_dataLabelPosition) Then
                                                    Try
                                                        Dim _dlObj As Object = _s1.GetType().InvokeMember("DataLabels", System.Reflection.BindingFlags.GetProperty, Nothing, _s1, Nothing)
                                                        If _dlObj IsNot Nothing Then
                                                            _dlObj.GetType().InvokeMember("Position", System.Reflection.BindingFlags.SetProperty, Nothing, _dlObj, New Object() {MapLabelPosition(_dataLabelPosition)})
                                                        End If
                                                    Catch : End Try
                                                End If
                                            End If
                                        End If
                                    Catch : End Try
                                End If
                                ' ⚠️ 不要手动改 ChartArea 透明度/PlotArea！ChartStyle=26 自己管配色！
                            Catch : End Try

                            End If  ' ← 关闭 L876 的 If _pivotChart Is Nothing Then ... Else
                        Catch ex As Exception
                            warnings.Add(String.Format("透视表 {0} 透视图创建失败: {1}", ptName, ex.Message))
                        End Try
                    End If

                    Dim tr2 = pt.TableRange2
                    Try
                        nextPivotRow = tr2.Row + tr2.Rows.Count + 2
                    Catch
                        nextPivotRow += 30
                    End Try
                    pivotInfos.Add(New Dictionary(Of String, Object) From {
                        {"name", ptName},
                        {"dimension", dimName},
                        {"value", valName},
                        {"dataRange", "'" & pivotDestWs.Name & "'!" & tr2.Address(False, False)}
                    })
                Next

                ' 创建切片器（每维度一个；绑定 firstPivot，共享缓存下所有透视表联动）
                ' 【关键修复】确保透视表完全初始化后再创建切片器缓存：
                ' 1) 激活目标工作表（SlicerCaches.Add 在非活动表上偶发失败）
                ' 2) 对每个透视表强制刷新 + DoEvents，让 PivotCache 内部状态就绪
                ' 3) 选中 firstPivot（部分 Excel 版本要求透视表处于选中状态才能创建切片器缓存）
                Try
                    destWs.Activate()
                Catch
                End Try
                For Each p In createdPivots
                    Try
                        p.ManualUpdate = False
                        p.RefreshTable()
                        System.Windows.Forms.Application.DoEvents()
                    Catch
                    End Try
                Next
                If firstPivot IsNot Nothing Then
                    Try
                        firstPivot.TableRange2.Select()
                        System.Windows.Forms.Application.DoEvents()
                    Catch
                    End Try
                End If

                ' 先清理看板上已存在的切片器（避免重名冲突）
                Try
                    Dim slShapes As Microsoft.Office.Interop.Excel.Shapes = destWs.Shapes
                    For i As Integer = slShapes.Count To 1 Step -1
                        Dim shp As Microsoft.Office.Interop.Excel.Shape = DirectCast(slShapes(i), Microsoft.Office.Interop.Excel.Shape)
                        If shp.Type = Microsoft.Office.Core.MsoShapeType.msoSlicer Then
                            shp.Delete()
                        End If
                    Next
                Catch
                End Try

                Dim slicerInfos As New List(Of Dictionary(Of String, Object))()
                Dim jslicers = JArray.Parse(If(String.IsNullOrWhiteSpace(slicerDimensionsJson), "[]", slicerDimensionsJson))
                ' 兜底：若 slicerDimensions 为空，则从 pivotTables 中去重提取所有 dimension 自动创建切片器
                If jslicers.Count = 0 Then
                    For Each jp In jpivs
                        Dim dimName As String = If(jp("dimension") IsNot Nothing, jp("dimension").ToString(), "")
                        If Not String.IsNullOrEmpty(dimName) Then
                            Dim exists As Boolean = False
                            For Each e In jslicers
                                If String.Equals(If(e("header") IsNot Nothing, e("header").ToString(), ""), dimName, StringComparison.OrdinalIgnoreCase) Then exists = True : Exit For
                            Next
                            If Not exists Then jslicers.Add(New JObject(New JProperty("header", dimName)))
                        End If
                    Next
                End If
                Dim slIndex = 0
                Dim anySlicerFailed As Boolean = False
                For Each js In jslicers
                    slIndex += 1
                    ' STA 间隙：Slicers.Add 内部会动 SlicerCache，前后 COM 调用间要让出消息泵
                    Thread.Sleep(300)
                    Try : Application.DoEvents() : Catch : End Try
                    Dim header As String = If(js("header") IsNot Nothing, js("header").ToString(), "")
                    If String.IsNullOrEmpty(header) Then Continue For
                    ' 智能布局：切片器区域放在看板顶部 KPI 区下方
                    ' 每个切片器默认宽 200 × 高 100（多列模式），从 A 列开始横向排列
                    Dim top As Double = 95.0
                    Dim left As Double = 20.0 + (slIndex - 1) * 210.0
                    Dim slWidth As Double = 200.0
                    Dim slHeight As Double = 100.0
                    If js("top") IsNot Nothing Then top = js("top").ToObject(Of Double)()
                    If js("left") IsNot Nothing Then left = js("left").ToObject(Of Double)()
                    If js("width") IsNot Nothing Then slWidth = js("width").ToObject(Of Double)()
                    If js("height") IsNot Nothing Then slHeight = js("height").ToObject(Of Double)()

                    Dim sc As Microsoft.Office.Interop.Excel.SlicerCache = Nothing
                    Dim actualFieldName As String = header
                    If firstPivot IsNot Nothing Then
                        ' 确保透视表完全初始化后再创建切片器缓存
                        Try : System.Windows.Forms.Application.DoEvents() : Catch : End Try
                        progress.Add(String.Format("创建切片器缓存 [{0}]", header))
                        Dim scName = "Slicer_" & header.Replace(" ", "")
                        Try
                            Dim oldSc As Microsoft.Office.Interop.Excel.SlicerCache = DirectCast(
                                wb.SlicerCaches.Item(scName),
                                Microsoft.Office.Interop.Excel.SlicerCache)
                            oldSc.Delete()
                        Catch
                        End Try

                        Dim pfForSlicer As Microsoft.Office.Interop.Excel.PivotField = FindPivotField(firstPivot, header)
                        actualFieldName = If(pfForSlicer IsNot Nothing, pfForSlicer.Name, header)
                        If pfForSlicer Is Nothing Then
                            warnings.Add(String.Format("切片器[{0}]：在透视表字段中未找到匹配项，将按名称 [{1}] 直接尝试", header, actualFieldName))
                        End If

                        Dim slcType As Object = CInt(Microsoft.Office.Interop.Excel.XlSlicerCacheType.xlSlicer)

                        Dim scCreated As Microsoft.Office.Interop.Excel.SlicerCache = Nothing
                        Dim pcForSl As Microsoft.Office.Interop.Excel.PivotCache = firstPivot.PivotCache

                        ' 优先使用 PivotTable 创建切片器缓存（非 OLAP 数据源最可靠的方式）
                        ' 反射签名：Add2(Source, SourceField, Name, SlicerCacheType)
                        '          Add(Source, SourceField, Name)
                        ' SourceField = pfForSlicer（PivotField 对象）— 之前漏掉这个参数导致全部 FALLBACK！
                        If scCreated Is Nothing AndAlso pfForSlicer IsNot Nothing Then
                            Try
                                scCreated = DirectCast(
                                    wb.SlicerCaches.Add2(firstPivot, pfForSlicer, scName, slcType),
                                    Microsoft.Office.Interop.Excel.SlicerCache)
                                progress.Add("  Add2(pt,pf,name,type) OK ✅")
                            Catch exC1 As Exception
                                warnings.Add(String.Format("Add2(pt,pf,name,type) FALLBACK: {0} [0x{1:X8}]", exC1.Message, exC1.HResult))
                                Try
                                    scCreated = DirectCast(
                                        wb.SlicerCaches.Add(firstPivot, pfForSlicer, scName),
                                        Microsoft.Office.Interop.Excel.SlicerCache)
                                    progress.Add("  Add(pt,pf,name) OK ✅")
                                Catch exC2 As Exception
                                    warnings.Add(String.Format("Add(pt,pf,name) FALLBACK: {0} [0x{1:X8}]", exC2.Message, exC2.HResult))
                                End Try
                            End Try
                        End If

                        ' 回退1：使用 PivotCache 创建
                        If scCreated Is Nothing AndAlso pfForSlicer IsNot Nothing Then
                            Try
                                scCreated = DirectCast(
                                    wb.SlicerCaches.Add2(pcForSl, pfForSlicer, scName, slcType),
                                    Microsoft.Office.Interop.Excel.SlicerCache)
                                progress.Add("  Add2(pc,pf,name,type) OK ✅")
                            Catch exC3 As Exception
                                warnings.Add(String.Format("Add2(pc,pf,name,type) FALLBACK: {0} [0x{1:X8}]", exC3.Message, exC3.HResult))
                                Try
                                    scCreated = DirectCast(
                                        wb.SlicerCaches.Add(pcForSl, pfForSlicer, scName),
                                        Microsoft.Office.Interop.Excel.SlicerCache)
                                    progress.Add("  Add(pc,pf,name) OK ✅")
                                Catch exC4 As Exception
                                    warnings.Add(String.Format("Add(pc,pf,name) FALLBACK: {0} [0x{1:X8}]", exC4.Message, exC4.HResult))
                                End Try
                            End Try
                        End If

                        ' 回退2：仅传 SourceField（Source=Nothing，让 Excel 自动从 PivotField 找缓存）
                        If pfForSlicer IsNot Nothing AndAlso scCreated Is Nothing Then
                            Try
                                scCreated = DirectCast(
                                    wb.SlicerCaches.Add(Type.Missing, pfForSlicer, scName),
                                    Microsoft.Office.Interop.Excel.SlicerCache)
                                progress.Add("  Add(missing,pf,name) OK ✅")
                            Catch exPF As Exception
                                warnings.Add(String.Format("Add(missing,pf,name) FALLBACK: {0} [0x{1:X8}]", exPF.Message, exPF.HResult))
                            End Try
                        End If

                        ' 最终回退：XLM 宏 + CallByName 延迟绑定
                        If scCreated Is Nothing Then
                            If pfForSlicer IsNot Nothing Then
                                Try
                                    Dim ptAddr As String = firstPivot.TableRange1.Address(False, False, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1, True)
                                    Dim pfAddr As String = pfForSlicer.DataRange.Address(False, False, Microsoft.Office.Interop.Excel.XlReferenceStyle.xlR1C1, True)
                                    Dim xlmCmd As String = String.Format("=SICREATE(""{0}"",{1},1,{2})", scName, ptAddr, pfAddr)
                                    Dim xlmRes As Object = app.ExecuteExcel4Macro(xlmCmd)
                                    warnings.Add(String.Format("  SICREATE XLM result={0}", If(xlmRes Is Nothing, "nothing", xlmRes.ToString())))
                                    Try : scCreated = DirectCast(wb.SlicerCaches(scName), Microsoft.Office.Interop.Excel.SlicerCache) : Catch : End Try
                                Catch exXLM As Exception
                                    warnings.Add(String.Format("ExecuteExcel4Macro SICREATE FALLBACK: {0} [0x{1:X8}]", exXLM.Message, exXLM.HResult))
                                End Try
                            End If
                            If scCreated Is Nothing AndAlso pfForSlicer IsNot Nothing Then
                                Try
                                    scCreated = DirectCast(Microsoft.VisualBasic.Interaction.CallByName(
                                        wb.SlicerCaches, "Add", Microsoft.VisualBasic.CallType.Method,
                                        New Object() {DirectCast(pfForSlicer, Object), DirectCast(scName, Object)}),
                                        Microsoft.Office.Interop.Excel.SlicerCache)
                                    progress.Add("  CallByName SlicerCaches.Add OK (LateBind)")
                                Catch exLB As Exception
                                    warnings.Add(String.Format("CallByName SlicerCaches.Add FALLBACK: {0} [0x{1:X8}]", exLB.Message, exLB.HResult))
                                End Try
                            End If
                        End If

                        sc = scCreated
                    End If
                    If sc Is Nothing Then
                        warnings.Add(String.Format("切片器 [{0}] 缓存创建失败", header))
                        anySlicerFailed = True
                        Continue For
                    End If
                    Try
                        For Each ptOther In createdPivots
                            If ptOther IsNot firstPivot Then sc.PivotTables.AddPivotTable(ptOther)
                        Next
                    Catch exConn As Exception
                        warnings.Add(String.Format("切片器[{0}]挂接其余透视表回退: {1}", header, exConn.Message))
                    End Try
                    Dim sl As Object = AddSlicerToSheet(destWs, sc, actualFieldName, header, top, left, slWidth, slHeight, warnings)
                    If sl Is Nothing Then
                        anySlicerFailed = True
                    Else
                        slicerInfos.Add(New Dictionary(Of String, Object) From {
                            {"header", header},
                            {"top", top},
                            {"left", left},
                            {"width", slWidth},
                            {"height", slHeight}
                        })
                    End If
                Next

                ' 仅当「一个切片器都没建成」才标 IsError。
                ' 部分维度失败 + 其余成功时仍算成功（warnings 里保留明细）；否则前端会误降级叠静态图。
                If anySlicerFailed AndAlso slicerInfos.Count = 0 Then
                    result.IsError = True
                ElseIf anySlicerFailed Then
                    warnings.Add(String.Format("⚠ 部分切片器未建成（成功 {0}/{1}），已保留可用切片器", slicerInfos.Count, If(jslicers IsNot Nothing, jslicers.Count, 0)))
                End If

                ' 注意：PivotChart 失败不影响 IsError —— PivotTable + Slicer 才是核心
                ' JS 侧会根据 return.ptChartOkCount / slicersCreated 判断是否降级为 dataValues 图表
                If ptChartOkCount = 0 AndAlso jpivs.Count > 0 Then
                    warnings.Add(String.Format("ℹ {0} 张 PivotChart 均未成功创建（A/B 方案均返回异常），但 PivotTable + Slicer 已就绪", jpivs.Count))
                End If

                ' ═══════════════════════════════════════════════════════════════
                ' ⚠ 不要隐藏 _dash_pivot！Excel 在隐藏承载 PivotTable 的 Sheet 时，
                '   会自动把 Slicer 也一起隐藏/删除！这就是切片器消失的根因。
                ' ═══════════════════════════════════════════════════════════════

                ' ═══════════════════════════════════════════════════════════════
                ' 优化 B + C：遍历真正的 PivotTable 对象（createdPivots），隐藏字段按钮 + PivotChart 主题配色
                ' ═══════════════════════════════════════════════════════════════
                Try
                    For Each _pt As Microsoft.Office.Interop.Excel.PivotTable In createdPivots
                        If _pt Is Nothing Then Continue For
                        ' (B) PivotTable 本身也隐藏字段标题
                        Try : _pt.DisplayFieldCaptions = False : Catch : End Try
                        Try : _pt.TableStyle2 = "" : Catch : End Try

                        ' (C) PivotChart 主题 —— 通过 PivotLayout 拿 Chart 对象
                        Dim _cht As Microsoft.Office.Interop.Excel.Chart = Nothing
                        Try : _cht = CType(_pt.GetType().InvokeMember("Chart", System.Reflection.BindingFlags.GetProperty, Nothing, _pt, Nothing), Microsoft.Office.Interop.Excel.Chart) : Catch : _cht = Nothing : End Try
                        If _cht Is Nothing Then Continue For
                        Dim _pl As Microsoft.Office.Interop.Excel.PivotLayout = Nothing
                        Try : _pl = CType(_cht.PivotLayout, Microsoft.Office.Interop.Excel.PivotLayout) : Catch : _pl = Nothing : End Try

                        ' 标题：加粗、深灰、字号 12
                        Try
                            _cht.HasTitle = True
                            _cht.ChartTitle.Font.Bold = True
                            _cht.ChartTitle.Font.Size = 12
                            _cht.ChartTitle.Font.Color = RGB(31, 78, 121)
                        Catch : End Try
                        ' 图例：浅灰字
                        Try
                            If _cht.HasLegend Then
                                _cht.Legend.Font.Size = 9
                                _cht.Legend.Font.Color = RGB(89, 89, 89)
                            End If
                        Catch : End Try
                        ' ChartStyle #26 = 商务蓝专业样式
                        Try : _cht.ChartStyle = 26 : Catch : End Try
                        ' 数值轴：隐藏网格线
                        Try
                            Dim _axObj As Object = _cht.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlValue)
                            Dim _valAx As Microsoft.Office.Interop.Excel.Axis = CType(_axObj, Microsoft.Office.Interop.Excel.Axis)
                            _valAx.HasMajorGridlines = False
                            _valAx.HasMinorGridlines = False
                            _valAx.TickLabels.Font.Size = 9
                            _valAx.TickLabels.Font.Color = RGB(89, 89, 89)
                        Catch : End Try
                        ' 分类轴：字号 9 浅灰
                        Try
                            Dim _axObj2 As Object = _cht.Axes(Microsoft.Office.Interop.Excel.XlAxisType.xlCategory)
                            Dim _catAx As Microsoft.Office.Interop.Excel.Axis = CType(_axObj2, Microsoft.Office.Interop.Excel.Axis)
                            _catAx.TickLabels.Font.Size = 9
                            _catAx.TickLabels.Font.Color = RGB(89, 89, 89)
                        Catch : End Try
                        ' ⚠️ ChartArea/PlotArea 全由 ChartStyle=26 控制，不要手动改！
                    Next
                    If createdPivots.Count > 0 Then progress.Add("  ✅ PivotChart 主题配色已应用（商务蓝 ChartStyle=26）")
                Catch exTheme As Exception
                    warnings.Add("⚠ PivotChart 主题配色部分失败: " & exTheme.Message)
                End Try

                ' ═══════════════════════════════════════════════════════════════
                ' 优化 D：找到看板 Sheet，Trim 尾部空行空列（UsedRange 扩展清理）
                ' ═══════════════════════════════════════════════════════════════
                Try
                    Dim _boardWs As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                    If wb IsNot Nothing Then
                        For Each _ws As Microsoft.Office.Interop.Excel.Worksheet In wb.Worksheets
                            If _ws.Name <> pivotDestWs.Name AndAlso _ws.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible Then
                                _boardWs = _ws
                                Exit For
                            End If
                        Next
                    End If
                    If _boardWs IsNot Nothing Then
                        ' 隐藏网格线（反射绕过后期绑定）
                        Try : _boardWs.GetType().InvokeMember("DisplayGridlines", System.Reflection.BindingFlags.SetProperty, Nothing, _boardWs, New Object() {False}) : Catch : End Try
                        ' 设置 Zoom=100（反射绕过后期绑定）
                        Try
                            Dim _winObj As Object = _boardWs.GetType().InvokeMember("ActiveWindow", System.Reflection.BindingFlags.GetProperty, Nothing, _boardWs, Nothing)
                            If _winObj IsNot Nothing Then
                                _winObj.GetType().InvokeMember("Zoom", System.Reflection.BindingFlags.SetProperty, Nothing, _winObj, New Object() {100})
                            End If
                        Catch : End Try
                        progress.Add("  ✅ 看板 Sheet Trim 完成（网格线已隐藏）")
                    End If
                Catch exBoard As Exception
                    warnings.Add("⚠ 看板 Sheet 清理部分失败: " & exBoard.Message)
                End Try

                ' ═══════════════════════════════════════════════════════════════
                ' 终极自动清理：遍历 Shapes + PivotCharts + Slicers → 算最右/最下边界
                ' Excel UsedRange 不算 Shape/Chart/Slicer，必须手动算！
                ' ═══════════════════════════════════════════════════════════════
                Try
                    ' 直接用 destWs（看板 sheet），别瞎找！
                    Dim _brd As Microsoft.Office.Interop.Excel.Worksheet = destWs
                    If _brd IsNot Nothing Then
                        Dim _maxCol As Integer = 1
                        Dim _maxRow As Integer = 1
                        ' 1) 遍历所有 Shape（含 KPI 卡、图表、切片器浮层）
                        Try
                            For Each _sh As Microsoft.Office.Interop.Excel.Shape In _brd.Shapes
                                Try
                                    Dim _shL As Double = _sh.Left
                                    Dim _shT As Double = _sh.Top
                                    Dim _shR As Double = _shL + _sh.Width
                                    Dim _shB As Double = _shT + _sh.Height
                                    ' 用 _brd.Range 把像素转行列（CDbl 显式转，Option Strict On 需要）
                                    Dim _rCell1 As Microsoft.Office.Interop.Excel.Range = _brd.Range("A1")
                                    Dim _rCellB As Microsoft.Office.Interop.Excel.Range = _rCell1.Offset(0, 1)  ' B1 做参考
                                    Dim _rCell2 As Microsoft.Office.Interop.Excel.Range = _rCell1.Offset(1, 0)  ' A2 做参考
                                    Dim _pxPerCol As Double = CDbl(_rCellB.Left) - CDbl(_rCell1.Left)
                                    Dim _pxPerRow As Double = CDbl(_rCell2.Top) - CDbl(_rCell1.Top)
                                    If _pxPerCol > 0 Then
                                        Dim _cEnd As Integer = CInt(_shR / _pxPerCol) + 1
                                        If _cEnd > _maxCol Then _maxCol = _cEnd
                                    End If
                                    If _pxPerRow > 0 Then
                                        Dim _rEnd As Integer = CInt(_shB / _pxPerRow) + 1
                                        If _rEnd > _maxRow Then _maxRow = _rEnd
                                    End If
                                Catch : End Try
                            Next
                        Catch : End Try
                        ' 2) 兜底：加 5 列 + 20 行安全边距
                        _maxCol = Math.Min(16384, _maxCol + 5)
                        _maxRow = Math.Min(1048576, _maxRow + 20)
                        ' 尾部空列（_maxCol+1 → XFD）
                        If _maxCol < 16384 Then
                            Try : _brd.Range(_brd.Columns.Item(_maxCol + 1), _brd.Columns.Item(16384)).EntireColumn.Hidden = True : Catch : End Try
                        End If
                        ' 尾部空行（_maxRow+1 → 1048576）
                        If _maxRow < 1048576 Then
                            Try : _brd.Range(_brd.Rows.Item(_maxRow + 1), _brd.Rows.Item(1048576)).EntireRow.Hidden = True : Catch : End Try
                        End If
                        ' 统一看板区域列宽（A → _maxCol）
                        If _maxCol >= 1 Then
                            Try : _brd.Range(_brd.Columns.Item(1), _brd.Columns.Item(_maxCol)).ColumnWidth = 9 : Catch : End Try
                        End If
                        ' ⚠ 清除看板区域单元格底色 + 条件格式（源数据可能带黄色高亮/CF）
                        Try
                            Dim _clearRng As Microsoft.Office.Interop.Excel.Range = _brd.Range(_brd.Cells.Item(1, 1), _brd.Cells.Item(_maxRow, _maxCol))
                            _clearRng.Interior.Color = Microsoft.Office.Interop.Excel.Constants.xlNone
                            Try : _clearRng.FormatConditions.Delete() : Catch : End Try
                            Try : _clearRng.Borders.LineStyle = Microsoft.Office.Interop.Excel.Constants.xlNone : Catch : End Try
                        Catch : End Try
                        progress.Add(String.Format("  ✅ 看板自动清理: Shapes推算边界={0}列×{1}行 × 隐藏尾部列={2}+行={3}", _maxCol, _maxRow, 16384 - _maxCol, 1048576 - _maxRow))
                    End If
                Catch exAuto As Exception
                    warnings.Add("⚠ 自动清理部分失败: " & exAuto.Message)
                End Try

                For Each p In createdPivots
                    p.ManualUpdate = False
                    p.RefreshTable()
                Next
                Try
                    app.CalculateFullRebuild()
                Catch
                End Try

                Dim payload = New With {
                    .pivotTables = pivotInfos,
                    .slicers = slicerInfos,
                    .warnings = warnings,
                    .progress = progress,
                    .ptChartOkCount = ptChartOkCount,
                    .pivotTableCount = jpivs.Count,
                    .slicersRequested = If(jslicers IsNot Nothing, jslicers.Count, 0),
                    .slicersCreated = slicerInfos.Count
                }
                ' 人类可读汇总 —— McpTextContent.Text 里直接展示
                Dim summaryLines As New List(Of String)
                summaryLines.Add("=== excel_create_slicer 执行汇总 ===")
                summaryLines.Add(String.Format("PivotTable: {0} 张 | PivotChart: {1}/{0} 成功 | Slicer: {2}/{3}",
                    jpivs.Count, ptChartOkCount, slicerInfos.Count, If(jslicers IsNot Nothing, jslicers.Count, 0)))
                If progress.Count > 0 Then
                    summaryLines.Add("── Progress ──")
                    summaryLines.AddRange(progress)
                End If
                If warnings.Count > 0 Then
                    summaryLines.Add("── Warnings ──")
                    summaryLines.AddRange(warnings)
                End If
                summaryLines.Add("=== Payload JSON ===")
                Dim summaryText = String.Join(vbCrLf, summaryLines) & vbCrLf & JsonConvert.SerializeObject(payload)
                ' ★ 写到临时文件 —— 不依赖任何 UI/VS 输出，100% 能拿到！
                Try
                    Dim _logPath As String = IO.Path.Combine(IO.Path.GetTempPath(), "excel_create_slicer.log")
                    IO.File.WriteAllText(_logPath, summaryText)
                Catch : End Try
                result.Content.Add(New McpTextContent With {
                    .Text = summaryText
                })
            Catch ex As Exception
                result.IsError = True
                Dim failText = String.Format("创建切片器失败 @ [{0}]: {1} | inner: {2}",
                    String.Join(" → ", progress),
                    ex.Message,
                    If(ex.InnerException Is Nothing, "-", ex.InnerException.Message))
                result.Content.Add(New McpTextContent With {.Text = failText})
                ' 失败也落盘，避免只留着上次成功日志误导排查
                Try
                    Dim _logPathFail As String = IO.Path.Combine(IO.Path.GetTempPath(), "excel_create_slicer.log")
                    IO.File.WriteAllText(_logPathFail, failText & vbCrLf & ex.ToString())
                Catch
                End Try
            End Try
            Return result
        End Function
        ''' <summary>PivotTable 重载：从 pt.PivotFields 取集合后委托 ResolvePivotFieldLoose，
        ''' 与 CreatePivot / SetLayout 的字段解析语义完全一致。</summary>
        Private Function FindPivotField(pt As Microsoft.Office.Interop.Excel.PivotTable, name As String) As Microsoft.Office.Interop.Excel.PivotField
            If pt Is Nothing Then Return Nothing
            Try
                Dim allFields As Microsoft.Office.Interop.Excel.PivotFields = DirectCast(pt.PivotFields, Microsoft.Office.Interop.Excel.PivotFields)
                Dim result = ResolvePivotFieldLoose(allFields, name)
                Utils.ComReleaseHelper.Release(allFields)
                Return result
            Catch
                Return Nothing
            End Try
        End Function
        Private Function AddSlicerToSheet(targetWs As Microsoft.Office.Interop.Excel.Worksheet, sc As Microsoft.Office.Interop.Excel.SlicerCache, fieldName As String, caption As String, top As Double, left As Double, width As Double, height As Double, ByRef warnings As List(Of String)) As Object
            If sc Is Nothing Then
                warnings.Add("[SLICER_DIAG] sc Is Nothing")
                Return Nothing
            End If

            Dim slicerName As String = "切片器_" & caption
            Dim missingVal As Object = System.Reflection.Missing.Value
            Dim sl As Microsoft.Office.Interop.Excel.Slicer = Nothing

            ' ====== 强类型方式 ======
            ' PIA Slicers.Add 签名: Add(SlicerDestination, [Level], [Name], [Caption], [Top], [Left], [Width], [Height])
            ' 非 OLAP 数据源：Level 必须是透视缓存中的实际字段名（与 PivotField.Name 一致）
            Try
                Dim slicers As Microsoft.Office.Interop.Excel.Slicers = sc.Slicers
                warnings.Add(String.Format("[SLICER_DIAG] Slicers.Count={0}, fieldName=[{1}]", slicers.Count, fieldName))

                ' 方案 A: 完整版（Level=实际字段名）
                Try
                    sl = slicers.Add(targetWs, fieldName, slicerName, caption, top, left, width, height)
                    If sl IsNot Nothing Then
                        warnings.Add("[SLICER_DIAG] SUCCESS (full 8-arg)")
                        EnsureSlicerVisible(sl, top, left, width, height, warnings)
                        Return sl
                    End If
                Catch exA As Exception
                    warnings.Add(String.Format("[SLICER_DIAG] 8-arg FALLBACK: {0}", exA.Message))
                End Try

                ' 方案 B: 不传 Level（若缓存本身已绑定到特定字段则无需指定）
                Try
                    sl = slicers.Add(targetWs, missingVal, slicerName, caption, top, left, width, height)
                    If sl IsNot Nothing Then
                        warnings.Add("[SLICER_DIAG] SUCCESS (no-Level)")
                        EnsureSlicerVisible(sl, top, left, width, height, warnings)
                        Return sl
                    End If
                Catch exB As Exception
                    warnings.Add(String.Format("[SLICER_DIAG] no-Level FALLBACK: {0}", exB.Message))
                End Try

                ' 方案 C: 最简形式 + 事后定位
                Try
                    sl = slicers.Add(targetWs, missingVal, slicerName)
                    If sl IsNot Nothing Then
                        sl.Top = top
                        sl.Left = left
                        sl.Width = width
                        sl.Height = height
                        warnings.Add("[SLICER_DIAG] SUCCESS (min-3-arg)")
                        EnsureSlicerVisible(sl, top, left, width, height, warnings)
                        Return sl
                    End If
                Catch exC As Exception
                    warnings.Add(String.Format("[SLICER_DIAG] min-3-arg FALLBACK: {0}", exC.Message))
                End Try

            Catch exOuter As Exception
                warnings.Add(String.Format("[SLICER_DIAG] strong-type outer FALLBACK: {0}", exOuter.Message))
            End Try

            ' ====== 备选：反射调用 ======
            Try
                Dim slColl As Object = sc.Slicers
                Dim collType As Type = slColl.GetType()
                Dim addMeth As MethodInfo = collType.GetMethod("Add", BindingFlags.Public Or BindingFlags.Instance)
                If addMeth Is Nothing Then
                    warnings.Add("[SLICER_DIAG] reflection: no Add method found")
                Else
                    Dim pInfos = addMeth.GetParameters()
                    Dim invokeArgs(pInfos.Length - 1) As Object
                    For i As Integer = 0 To pInfos.Length - 1
                        invokeArgs(i) = missingVal
                    Next
                    If pInfos.Length >= 1 Then invokeArgs(0) = targetWs
                    If pInfos.Length >= 2 Then invokeArgs(1) = fieldName
                    If pInfos.Length >= 3 Then invokeArgs(2) = slicerName
                    If pInfos.Length >= 4 Then invokeArgs(3) = caption
                    If pInfos.Length >= 5 Then invokeArgs(4) = top
                    If pInfos.Length >= 6 Then invokeArgs(5) = left
                    If pInfos.Length >= 7 Then invokeArgs(6) = width
                    If pInfos.Length >= 8 Then invokeArgs(7) = height
                    Dim slObj As Object = addMeth.Invoke(slColl, invokeArgs)
                    If slObj IsNot Nothing Then
                        warnings.Add("[SLICER_DIAG] SUCCESS via reflection")
                        Dim slTyped As Microsoft.Office.Interop.Excel.Slicer = TryCast(slObj, Microsoft.Office.Interop.Excel.Slicer)
                        If slTyped IsNot Nothing Then EnsureSlicerVisible(slTyped, top, left, width, height, warnings)
                        Return slObj
                    End If
                End If
            Catch exRefl As Exception
                Dim inner As String = If(exRefl.InnerException Is Nothing, "-", exRefl.InnerException.Message)
                warnings.Add(String.Format("[SLICER_DIAG] reflection FALLBACK: {0} | {1}", exRefl.Message, inner))
            End Try

            ' ====== 最终备选：CallByName ======
            Try
                Dim slObj As Object = Microsoft.VisualBasic.Interaction.CallByName(sc.Slicers, "Add", Microsoft.VisualBasic.CallType.Method, targetWs, fieldName, slicerName, caption, top, left, width, height)
                If slObj IsNot Nothing Then
                    warnings.Add("[SLICER_DIAG] SUCCESS via CallByName")
                    Dim slTyped As Microsoft.Office.Interop.Excel.Slicer = TryCast(slObj, Microsoft.Office.Interop.Excel.Slicer)
                    If slTyped IsNot Nothing Then EnsureSlicerVisible(slTyped, top, left, width, height, warnings)
                    Return slObj
                End If
            Catch exCb As Exception
                warnings.Add(String.Format("[SLICER_DIAG] CallByName FALLBACK: {0}", exCb.Message))
            End Try

            warnings.Add(String.Format("[SLICER_DIAG] *** ALL EXHAUSTED for field=[{0}] caption=[{1}] ***", fieldName, caption))
            Return Nothing
        End Function

        ''' <summary>
        ''' 确保切片器可见：显式设置位置/尺寸、Visible=True、置于顶层（ZOrder msoBringToFront）。
        ''' 部分 Excel 版本创建切片器后默认 Visible=False 或被其他形状遮挡，需主动修正。
        ''' </summary>
        Private Sub EnsureSlicerVisible(sl As Microsoft.Office.Interop.Excel.Slicer, top As Double, left As Double, width As Double, height As Double, ByRef warnings As List(Of String))
            If sl Is Nothing Then Return
            Try
                sl.Top = top
                sl.Left = left
                sl.Width = width
                sl.Height = height
            Catch ex As Exception
                warnings.Add(String.Format("[SLICER_DIAG] set pos/size FALLBACK: {0}", ex.Message))
            End Try
            Try
                Microsoft.VisualBasic.Interaction.CallByName(sl, "Visible", Microsoft.VisualBasic.CallType.Let, True)
            Catch ex As Exception
                warnings.Add(String.Format("[SLICER_DIAG] set Visible FALLBACK: {0}", ex.Message))
            End Try

            ' ════════════════════════════════════════════════════════════════
            ' 切片器视觉美化（统一商务蓝风格）：
            ' 1) 浅色切片器样式（与白底看板协调）
            ' 2) 单列垂直模式（纵向列表，紧凑整齐）
            ' 3) 显示小字号标题（告诉用户筛什么）
            ' 4) 紧凑行高 + 小字号选项
            ' ════════════════════════════════════════════════════════════════
            Try
                ' 浅色商务主题（深色在白底看板上太突兀）
                sl.Style = "SlicerStyleLight1"
            Catch
                Try : sl.Style = "SlicerStyleLight2" : Catch : End Try
            End Try
            Try
                ' 单列垂直模式 — 110px 宽度刚好放下一列选项
                sl.NumberOfColumns = 1
            Catch
            End Try
            Try
                ' 显示标题 + 小字号 + 商务蓝
                Dim slObj As Object = sl
                Dim prop = slObj.GetType().GetProperty("DisplayHeader")
                If prop IsNot Nothing Then prop.SetValue(slObj, True, Nothing)
            Catch
            End Try
            Try
                ' 紧凑行高
                sl.RowHeight = 16
            Catch
            End Try
            Try
                ' 切片器标题字体（反射兜底，PIA Slicer 无 Font 属性）
                Dim slObj2 As Object = sl
                Try : slObj2.GetType().InvokeMember("Font", System.Reflection.BindingFlags.GetProperty, Nothing, slObj2, Nothing) : Catch : End Try
            Catch
            End Try
            Try
                Dim shp As Microsoft.Office.Interop.Excel.Shape = TryCast(Microsoft.VisualBasic.Interaction.CallByName(sl, "Shape", Microsoft.VisualBasic.CallType.Get), Microsoft.Office.Interop.Excel.Shape)
                If shp IsNot Nothing Then
                    shp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                    shp.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringToFront)
                End If
            Catch ex As Exception
                warnings.Add(String.Format("[SLICER_DIAG] shape zorder FALLBACK: {0}", ex.Message))
            End Try
        End Sub

        Private Shared Function MapChartType(chartType As String) As Microsoft.Office.Interop.Excel.XlChartType
            If String.IsNullOrWhiteSpace(chartType) Then Return Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered
            Dim t As String = chartType.Trim().ToLowerInvariant()
            Select Case t
                Case "column", "bar", "柱状图", "柱形图", "柱图", "xlcolumnclustered", "xlcolumnstacked"
                    Return Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered
                Case "barh", "条形图", "横向柱状图", "xlbarclustered"
                    Return Microsoft.Office.Interop.Excel.XlChartType.xlBarClustered
                Case "line", "折线图", "趋势图", "xlline", "xllinemarkers"
                    Return Microsoft.Office.Interop.Excel.XlChartType.xlLine
                Case "pie", "饼图", "占比图", "xlpie"
                    Return Microsoft.Office.Interop.Excel.XlChartType.xlPie
                Case "doughnut", "圆环图", "环形图", "xldoughnut"
                    Return Microsoft.Office.Interop.Excel.XlChartType.xlDoughnut
                Case "area", "面积图", "xlarea"
                    Return Microsoft.Office.Interop.Excel.XlChartType.xlArea
                Case "scatter", "散点图", "xlxyscatter", "xlxyscatterlines"
                    Return Microsoft.Office.Interop.Excel.XlChartType.xlXYScatterLines
                Case "radar", "雷达图", "xlradar"
                    Return Microsoft.Office.Interop.Excel.XlChartType.xlRadar
                Case Else
                    Return Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered
            End Select
        End Function

        Private Shared Function MapLabelPosition(pos As String) As Microsoft.Office.Interop.Excel.XlDataLabelPosition
            If String.IsNullOrWhiteSpace(pos) Then Return Microsoft.Office.Interop.Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd
            Dim t As String = pos.Trim().ToLowerInvariant()
            Select Case t
                Case "outsideend", "xllabelpositionoutsideend"
                    Return Microsoft.Office.Interop.Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd
                Case "insideend", "xllabelpositioninsideend"
                    Return Microsoft.Office.Interop.Excel.XlDataLabelPosition.xlLabelPositionInsideEnd
                Case "insidebase", "xllabelpositioninsidebase"
                    Return Microsoft.Office.Interop.Excel.XlDataLabelPosition.xlLabelPositionInsideBase
                Case "bestfit", "xllabelpositionbestfit"
                    Return Microsoft.Office.Interop.Excel.XlDataLabelPosition.xlLabelPositionBestFit
                Case "above", "xllabelpositionabove"
                    Return Microsoft.Office.Interop.Excel.XlDataLabelPosition.xlLabelPositionAbove
                Case "below", "xllabelpositionbelow"
                    Return Microsoft.Office.Interop.Excel.XlDataLabelPosition.xlLabelPositionBelow
                Case Else
                    Return Microsoft.Office.Interop.Excel.XlDataLabelPosition.xlLabelPositionOutsideEnd
            End Select
        End Function
    End Class

End Namespace
