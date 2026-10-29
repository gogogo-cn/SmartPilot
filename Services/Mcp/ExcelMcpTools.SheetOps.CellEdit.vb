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

        Private Function ExecuteClearRange(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim rangeAddr As String = GetArg(Of String)(args, "range", "")
                ' 按用户意图区分清除粒度：clearContent=仅内容，clearFormat=仅格式。
                ' 默认仅清内容（与 excel_tools.json / ToolDefinitions 描述一致）；
                ' 两者均为 true 时执行"全部清除"（rng.Clear()）。
                Dim clearContent As Boolean = GetArg(Of Boolean)(args, "clearContent", True)
                Dim clearFormat As Boolean = GetArg(Of Boolean)(args, "clearFormat", False)

                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim rng As Microsoft.Office.Interop.Excel.Range = Nothing
                Dim usedRangeAddr As String = ""

                Try
                    If Not String.IsNullOrEmpty(sheetName) Then
                        Try
                            ws = DirectCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                        Catch
                            ws = Nothing
                        End Try
                        If ws Is Nothing Then
                            result.IsError = True
                            result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: [{sheetName}]，请先创建该工作表"})
                            Return result
                        End If
                    Else
                        ws = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    End If

                    If Not String.IsNullOrEmpty(rangeAddr) Then
                        ' 指定了范围地址，直接使用
                        rng = ws.Range(rangeAddr)
                    Else
                        ' 未指定范围：选区决策顺序
                        ' 1) 优先使用上下文选区（最近一次read_range记录的选区）——硬约束，不依赖模型传参
                        Dim ctxAddr As String = ""
                        If RangeContext.TryGet(ws.Name, ctxAddr) Then
                            Try
                                rng = ws.Range(ctxAddr)
                                Debug.WriteLine($"[ExecuteClearRange] 使用上下文选区: {ctxAddr}")
                            Catch ex As Exception
                                Debug.WriteLine($"[ExecuteClearRange] 上下文选区 {ctxAddr} 无效: {ex.Message}")
                                rng = Nothing
                                RangeContext.Clear()
                            End Try
                        End If

                        ' 2) 回退：使用实时Selection
                        If rng Is Nothing Then
                            Try
                                Dim sel = excelApp.Selection
                                If TypeOf sel Is Microsoft.Office.Interop.Excel.Range Then
                                    Dim selRng = DirectCast(sel, Microsoft.Office.Interop.Excel.Range)
                                    Dim selSheetName As String = ""
                                    Try
                                        selSheetName = CStr(selRng.Worksheet.Name)
                                    Catch
                                        selSheetName = ""
                                    End Try
                                    If selSheetName <> "" AndAlso selSheetName = ws.Name Then
                                        rng = selRng
                                        Debug.WriteLine($"[ExecuteClearRange] 使用实时Selection: {rng.Address}")
                                    Else
                                        rng = Nothing
                                    End If
                                End If
                            Catch
                                rng = Nothing
                            End Try
                        End If

                        ' 3) 最终回退：UsedRange
                        If rng Is Nothing Then
                            rng = ws.UsedRange
                            Debug.WriteLine($"[ExecuteClearRange] 回退到UsedRange")
                        End If
                    End If

                    usedRangeAddr = rng.Address(False, False)

                    ' 撤销记录：整表备份（必须在清除之前，撤销时整张表还原）
                    Try
                        If clearContent OrElse clearFormat Then
                            Bridges.BridgeRegistry.Instance.Resolve(Of Bridges.IUndoBridge)().AddSheetBackup(ws.Name)
                        End If
                    Catch undoEx As Exception
                        Debug.WriteLine($"[ExecuteClearRange] 撤销记录失败: {undoEx.Message}")
                    End Try

                    ' 按用户意图执行：仅内容 / 仅格式 / 全部
                    Dim clearedParts As New List(Of String)()
                    If clearFormat AndAlso clearContent Then
                        rng.Clear()
                        clearedParts.Add("内容")
                        clearedParts.Add("格式")
                    ElseIf clearFormat Then
                        rng.ClearFormats()
                        clearedParts.Add("格式")
                    ElseIf clearContent Then
                        rng.ClearContents()
                        clearedParts.Add("内容")
                    End If

                    ' 清空成功后，清除上下文选区（避免后续操作误用）
                    RangeContext.Clear()

                    If clearedParts.Count = 0 Then
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {
                            .Text = $"未执行清除：clearContent 与 clearFormat 均为 False，未清除范围 {usedRangeAddr} 的任何内容。"
                        })
                    Else
                        result.Content.Add(New McpTextContent With {
                            .Text = $"成功清除范围 {usedRangeAddr}（清除{String.Join("、", clearedParts)}）"
                        })
                    End If
                Catch
                    ' 内部错误由外层 Try 处理
                End Try

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"清除范围失败: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteClearFormat(args As Dictionary(Of String, Object)) As ToolsCallResult
            ' 独立"清空格式"工具：本质只清除单元格格式（rng.ClearFormats()），绝不清除内容。
            ' 与 excel_clear_range（清数据）分离，便于 AI 多步骤规划中单独表达"只清格式"的意图。
            Dim result = New ToolsCallResult()
            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim rangeAddr As String = GetArg(Of String)(args, "range", "")

                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim rng As Microsoft.Office.Interop.Excel.Range = Nothing

                Try
                    If Not String.IsNullOrEmpty(sheetName) Then
                        Try
                            ws = DirectCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                        Catch
                            ws = Nothing
                        End Try
                        If ws Is Nothing Then
                            result.IsError = True
                            result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: [{sheetName}]，请先创建该工作表"})
                            Return result
                        End If
                    Else
                        ws = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    End If

                    If Not String.IsNullOrEmpty(rangeAddr) Then
                        rng = ws.Range(rangeAddr)
                    Else
                        ' 选区决策顺序（与 excel_clear_range 一致）：
                        ' 1) 上下文选区（最近一次read_range记录）2) 实时Selection 3) UsedRange
                        Dim ctxAddr As String = ""
                        If RangeContext.TryGet(ws.Name, ctxAddr) Then
                            Try
                                rng = ws.Range(ctxAddr)
                                Debug.WriteLine($"[ExecuteClearFormat] 使用上下文选区: {ctxAddr}")
                            Catch ex As Exception
                                Debug.WriteLine($"[ExecuteClearFormat] 上下文选区 {ctxAddr} 无效: {ex.Message}")
                                rng = Nothing
                                RangeContext.Clear()
                            End Try
                        End If

                        If rng Is Nothing Then
                            Try
                                Dim sel = excelApp.Selection
                                If TypeOf sel Is Microsoft.Office.Interop.Excel.Range Then
                                    Dim selRng = DirectCast(sel, Microsoft.Office.Interop.Excel.Range)
                                    Dim selSheetName As String = ""
                                    Try
                                        selSheetName = CStr(selRng.Worksheet.Name)
                                    Catch
                                        selSheetName = ""
                                    End Try
                                    If selSheetName <> "" AndAlso selSheetName = ws.Name Then
                                        rng = selRng
                                        Debug.WriteLine($"[ExecuteClearFormat] 使用实时Selection: {rng.Address}")
                                    Else
                                        rng = Nothing
                                    End If
                                End If
                            Catch
                                rng = Nothing
                            End Try
                        End If

                        If rng Is Nothing Then
                            rng = ws.UsedRange
                            Debug.WriteLine($"[ExecuteClearFormat] 回退到UsedRange")
                        End If
                    End If

                    Dim usedRangeAddr = rng.Address(False, False)

                    ' 撤销记录：整表备份（必须在清除格式之前）
                    Try
                        Bridges.BridgeRegistry.Instance.Resolve(Of Bridges.IUndoBridge)().AddSheetBackup(ws.Name)
                    Catch undoEx As Exception
                        Debug.WriteLine($"[ExecuteClearFormat] 撤销记录失败: {undoEx.Message}")
                    End Try

                    ' 仅清除格式，内容保持不变
                    rng.ClearFormats()

                    ' 清空成功后，清除上下文选区（避免后续操作误用）
                    RangeContext.Clear()

                    result.Content.Add(New McpTextContent With {
                        .Text = $"成功清除范围 {usedRangeAddr} 的格式（内容保持不变）"
                    })
                Finally
                    Utils.ComReleaseHelper.Release(rng)
                    Utils.ComReleaseHelper.Release(ws)
                End Try

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"清除格式失败: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteSetCellValue(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim cellAddr As String = GetArg(Of String)(args, "cell", "")
                Dim value As String = GetArg(Of String)(args, "value", "")

                If String.IsNullOrEmpty(cellAddr) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "单元格地址不能为空"})
                    Return result
                End If

                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim rng As Microsoft.Office.Interop.Excel.Range = Nothing

                Try
                    If Not String.IsNullOrEmpty(sheetName) Then
                        Try
                            ws = DirectCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                        Catch
                            ws = Nothing
                        End Try
                        If ws Is Nothing Then
                            result.IsError = True
                            result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: [{sheetName}]，请先创建该工作表"})
                            Return result
                        End If
                    Else
                        ws = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    End If

                    rng = ws.Range(cellAddr)

                    ' 撤销记录：整表备份（必须在写入之前，撤销时整张表还原）
                    Try
                        Bridges.BridgeRegistry.Instance.Resolve(Of Bridges.IUndoBridge)().AddSheetBackup(ws.Name)
                    Catch undoEx As Exception
                        Debug.WriteLine($"[ExecuteSetCellValue] 撤销记录失败: {undoEx.Message}")
                    End Try

                    rng.Value2 = value

                    result.Content.Add(New McpTextContent With {
                        .Text = $"成功设置单元格 {cellAddr} 的值为: {value}"
                    })
                Finally
                    Utils.ComReleaseHelper.Release(rng)
                    Utils.ComReleaseHelper.Release(ws)
                End Try

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"设置单元格值失败: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteGetCellValue(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim cellAddr As String = GetArg(Of String)(args, "cell", "")

                If String.IsNullOrEmpty(cellAddr) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "单元格地址不能为空"})
                    Return result
                End If

                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim rng As Microsoft.Office.Interop.Excel.Range = Nothing

                Try
                    If Not String.IsNullOrEmpty(sheetName) Then
                        Try
                            ws = DirectCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                        Catch
                            ws = Nothing
                        End Try
                        If ws Is Nothing Then
                            result.IsError = True
                            result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: [{sheetName}]，请先创建该工作表"})
                            Return result
                        End If
                    Else
                        ws = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    End If

                    rng = ws.Range(cellAddr)
                    Dim cellValue = If(rng.Value2 IsNot Nothing, rng.Value2.ToString(), "")

                    result.Content.Add(New McpTextContent With {
                        .Text = JsonConvert.SerializeObject(New With {
                            .cell = cellAddr,
                            .value = cellValue
                        }, Formatting.Indented)
                    })
                Finally
                    Utils.ComReleaseHelper.Release(rng)
                    Utils.ComReleaseHelper.Release(ws)
                End Try

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"获取单元格值失败: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteAddFormula(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim cellAddr As String = GetArg(Of String)(args, "cell", "")
                Dim formula As String = GetArg(Of String)(args, "formula", "")

                If String.IsNullOrEmpty(cellAddr) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "单元格地址不能为空"})
                    Return result
                End If

                If String.IsNullOrEmpty(formula) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "公式内容不能为空"})
                    Return result
                End If

                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim rng As Microsoft.Office.Interop.Excel.Range = Nothing

                Try
                    If Not String.IsNullOrEmpty(sheetName) Then
                        Try
                            ws = DirectCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                        Catch
                            ws = Nothing
                        End Try
                        If ws Is Nothing Then
                            result.IsError = True
                            result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: [{sheetName}]，请先创建该工作表"})
                            Return result
                        End If
                    Else
                        ws = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    End If

                    rng = ws.Range(cellAddr)

                    ' 撤销记录：整表备份（必须在写入公式之前）
                    Try
                        Bridges.BridgeRegistry.Instance.Resolve(Of Bridges.IUndoBridge)().AddSheetBackup(ws.Name)
                    Catch undoEx As Exception
                        Debug.WriteLine($"[ExecuteAddFormula] 撤销记录失败: {undoEx.Message}")
                    End Try

                    rng.Formula = If(formula.StartsWith("="), formula, "=" & formula)

                    result.Content.Add(New McpTextContent With {
                        .Text = $"成功在单元格 {cellAddr} 添加公式: {rng.Formula}"
                    })
                Finally
                    Utils.ComReleaseHelper.Release(rng)
                    Utils.ComReleaseHelper.Release(ws)
                End Try

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"添加公式失败: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteInsertRow(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim rowNum As Integer = GetArg(Of Integer)(args, "row", 0)
                Dim count As Integer = GetArg(Of Integer)(args, "count", 1)

                If rowNum <= 0 Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "行号必须大于0"})
                    Return result
                End If

                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim rng As Microsoft.Office.Interop.Excel.Range = Nothing

                Try
                    If Not String.IsNullOrEmpty(sheetName) Then
                        Try
                            ws = DirectCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                        Catch
                            ws = Nothing
                        End Try
                        If ws Is Nothing Then
                            result.IsError = True
                            result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: [{sheetName}]，请先创建该工作表"})
                            Return result
                        End If
                    Else
                        ws = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    End If

                    ' 撤销记录：整表备份（必须在插入之前，撤销时整张表还原）
                    Try
                        Bridges.BridgeRegistry.Instance.Resolve(Of Bridges.IUndoBridge)().AddSheetBackup(ws.Name)
                    Catch undoEx As Exception
                        Debug.WriteLine($"[ExecuteInsertRow] 撤销记录失败: {undoEx.Message}")
                    End Try

                    rng = DirectCast(ws.Rows(rowNum), Microsoft.Office.Interop.Excel.Range)
                    rng.Insert(Shift:=Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, CopyOrigin:=Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
                    Utils.ComReleaseHelper.Release(rng)
                    rng = Nothing

                    If count > 1 Then
                        For i As Integer = 1 To count - 1
                            rng = DirectCast(ws.Rows(rowNum + i), Microsoft.Office.Interop.Excel.Range)
                            rng.Insert(Shift:=Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, CopyOrigin:=Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
                            Utils.ComReleaseHelper.Release(rng)
                            rng = Nothing
                        Next
                    End If

                    result.Content.Add(New McpTextContent With {
                        .Text = $"成功在第 {rowNum} 行位置插入 {count} 行"
                    })
                Finally
                    Utils.ComReleaseHelper.Release(rng)
                    Utils.ComReleaseHelper.Release(ws)
                End Try

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"插入行失败: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteInsertColumn(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim colNum As Integer = GetArg(Of Integer)(args, "column", 0)
                Dim count As Integer = GetArg(Of Integer)(args, "count", 1)

                If colNum <= 0 Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "列号必须大于0"})
                    Return result
                End If

                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim rng As Microsoft.Office.Interop.Excel.Range = Nothing

                Try
                    If Not String.IsNullOrEmpty(sheetName) Then
                        Try
                            ws = DirectCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                        Catch
                            ws = Nothing
                        End Try
                        If ws Is Nothing Then
                            result.IsError = True
                            result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: [{sheetName}]，请先创建该工作表"})
                            Return result
                        End If
                    Else
                        ws = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    End If

                    ' 撤销记录：整表备份（必须在插入之前）
                    Try
                        Bridges.BridgeRegistry.Instance.Resolve(Of Bridges.IUndoBridge)().AddSheetBackup(ws.Name)
                    Catch undoEx As Exception
                        Debug.WriteLine($"[ExecuteInsertColumn] 撤销记录失败: {undoEx.Message}")
                    End Try

                    rng = DirectCast(ws.Columns(colNum), Microsoft.Office.Interop.Excel.Range)
                    rng.Insert(Shift:=Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
                    Utils.ComReleaseHelper.Release(rng)
                    rng = Nothing

                    If count > 1 Then
                        For i As Integer = 1 To count - 1
                            rng = DirectCast(ws.Columns(colNum + i), Microsoft.Office.Interop.Excel.Range)
                            rng.Insert(Shift:=Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftToRight, CopyOrigin:=Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
                            Utils.ComReleaseHelper.Release(rng)
                            rng = Nothing
                        Next
                    End If

                    result.Content.Add(New McpTextContent With {
                        .Text = $"成功在第 {colNum} 列位置插入 {count} 列"
                    })
                Finally
                    Utils.ComReleaseHelper.Release(rng)
                    Utils.ComReleaseHelper.Release(ws)
                End Try

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"插入列失败: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteDeleteRow(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim rowNum As Integer = GetArg(Of Integer)(args, "row", 0)
                Dim count As Integer = GetArg(Of Integer)(args, "count", 1)

                If rowNum <= 0 Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "行号必须大于0"})
                    Return result
                End If

                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim rng As Microsoft.Office.Interop.Excel.Range = Nothing

                Try
                    If Not String.IsNullOrEmpty(sheetName) Then
                        Try
                            ws = DirectCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                        Catch
                            ws = Nothing
                        End Try
                        If ws Is Nothing Then
                            result.IsError = True
                            result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: [{sheetName}]，请先创建该工作表"})
                            Return result
                        End If
                    Else
                        ws = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    End If

                    excelApp.DisplayAlerts = False
                    rng = ws.Range(ws.Rows(rowNum), ws.Rows(rowNum + count - 1))

                    ' 撤销记录：整表备份（必须在删除之前，撤销时整张表还原）
                    Try
                        Bridges.BridgeRegistry.Instance.Resolve(Of Bridges.IUndoBridge)().AddSheetBackup(ws.Name)
                    Catch undoEx As Exception
                        Debug.WriteLine($"[ExecuteDeleteRow] 撤销记录失败: {undoEx.Message}")
                    End Try

                    rng.Delete(Shift:=Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp)
                    excelApp.DisplayAlerts = True

                    result.Content.Add(New McpTextContent With {
                        .Text = $"成功删除从第 {rowNum} 行开始的 {count} 行"
                    })
                Finally
                    Utils.ComReleaseHelper.Release(rng)
                    Utils.ComReleaseHelper.Release(ws)
                    excelApp.DisplayAlerts = True
                End Try

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"删除行失败: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteDeleteColumn(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim colNum As Integer = GetArg(Of Integer)(args, "column", 0)
                Dim count As Integer = GetArg(Of Integer)(args, "count", 1)

                If colNum <= 0 Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "列号必须大于0"})
                    Return result
                End If

                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim rng As Microsoft.Office.Interop.Excel.Range = Nothing

                Try
                    If Not String.IsNullOrEmpty(sheetName) Then
                        Try
                            ws = DirectCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                        Catch
                            ws = Nothing
                        End Try
                        If ws Is Nothing Then
                            result.IsError = True
                            result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: [{sheetName}]，请先创建该工作表"})
                            Return result
                        End If
                    Else
                        ws = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    End If

                    excelApp.DisplayAlerts = False
                    rng = ws.Range(ws.Columns(colNum), ws.Columns(colNum + count - 1))

                    ' 撤销记录：整表备份（必须在删除之前，撤销时整张表还原）
                    Try
                        Bridges.BridgeRegistry.Instance.Resolve(Of Bridges.IUndoBridge)().AddSheetBackup(ws.Name)
                    Catch undoEx As Exception
                        Debug.WriteLine($"[ExecuteDeleteColumn] 撤销记录失败: {undoEx.Message}")
                    End Try

                    rng.Delete(Shift:=Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftToLeft)
                    excelApp.DisplayAlerts = True

                    result.Content.Add(New McpTextContent With {
                        .Text = $"成功删除从第 {colNum} 列开始的 {count} 列"
                    })
                Finally
                    Utils.ComReleaseHelper.Release(rng)
                    Utils.ComReleaseHelper.Release(ws)
                    excelApp.DisplayAlerts = True
                End Try

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"删除列失败: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteMergeCells(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim rangeAddr As String = GetArg(Of String)(args, "range", "")

                If String.IsNullOrEmpty(rangeAddr) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "范围不能为空"})
                    Return result
                End If

                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim rng As Microsoft.Office.Interop.Excel.Range = Nothing

                Try
                    If Not String.IsNullOrEmpty(sheetName) Then
                        Try
                            ws = DirectCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                        Catch
                            ws = Nothing
                        End Try
                        If ws Is Nothing Then
                            result.IsError = True
                            result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: [{sheetName}]，请先创建该工作表"})
                            Return result
                        End If
                    Else
                        ws = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    End If

                    rng = ws.Range(rangeAddr)

                    ' mergeType：merge=仅合并；mergeAndCenter=合并后文字居中（与 Excel 的"合并后居中"一致）
                    Dim mergeType As String = GetArg(Of String)(args, "mergeType", "merge").ToLower()

                    ' 关键：关闭 DisplayAlerts，否则"合并单元格时仅保留左上角值"会弹模态警告阻塞后续全部步骤
                    Dim prevAlerts As Boolean = excelApp.DisplayAlerts
                    Try
                        excelApp.DisplayAlerts = False
                        rng.Merge()
                        If mergeType = "mergeandcenter" Then
                            rng.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                        End If
                    Finally
                        excelApp.DisplayAlerts = prevAlerts
                    End Try

                    result.Content.Add(New McpTextContent With {
                        .Text = $"成功合并范围 {rangeAddr}"
                    })
                Finally
                    Utils.ComReleaseHelper.Release(rng)
                    Utils.ComReleaseHelper.Release(ws)
                End Try

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"合并单元格失败: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteUnmergeCells(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim rangeAddr As String = GetArg(Of String)(args, "range", "")

                If String.IsNullOrEmpty(rangeAddr) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "范围不能为空"})
                    Return result
                End If

                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim rng As Microsoft.Office.Interop.Excel.Range = Nothing

                Try
                    If Not String.IsNullOrEmpty(sheetName) Then
                        Try
                            ws = DirectCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                        Catch
                            ws = Nothing
                        End Try
                        If ws Is Nothing Then
                            result.IsError = True
                            result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: [{sheetName}]，请先创建该工作表"})
                            Return result
                        End If
                    Else
                        ws = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    End If

                    rng = ws.Range(rangeAddr)

                    ' fillAfterUnmerge=true：取消合并后，把原来合并区域的值展开填充到整个区域
                    ' （原生 UnMerge() 仅保留左上角值，其余单元格会变空）
                    ' 说明：合并范围内 rng.Value2 只返回左上角值（标量）；若选择区域非单一合并块则返回二维数组，
                    ' 此时跳过填充以避免误写，退化为仅取消合并。仅用静态类型成员，兼容 Option Strict On。
                    Dim fillAfterUnmerge As Boolean = GetArg(Of Boolean)(args, "fillAfterUnmerge", False)
                    Dim spreadValue As Object = Nothing
                    If fillAfterUnmerge Then
                        Try
                            spreadValue = rng.Value2
                        Catch ex As Exception
                            spreadValue = Nothing
                        End Try
                    End If

                    rng.UnMerge()

                    Dim didSpread As Boolean = (fillAfterUnmerge AndAlso spreadValue IsNot Nothing AndAlso Not (TypeOf spreadValue Is System.Array))
                    If didSpread Then
                        rng.Value2 = spreadValue
                    End If

                    If didSpread Then
                        result.Content.Add(New McpTextContent With {
                            .Text = $"成功取消合并范围 {rangeAddr}，并已将原来合并区域的值展开填充到整个区域"
                        })
                    Else
                        result.Content.Add(New McpTextContent With {
                            .Text = $"成功取消合并范围 {rangeAddr}"
                        })
                    End If
                Finally
                    Utils.ComReleaseHelper.Release(rng)
                    Utils.ComReleaseHelper.Release(ws)
                End Try

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"取消合并单元格失败: {ex.Message}"})
            End Try
            Return result
        End Function

    End Class

End Namespace
