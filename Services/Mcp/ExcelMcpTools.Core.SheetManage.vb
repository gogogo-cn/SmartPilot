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

        Private Function ExecuteGetSheets(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "Excel 应用不可用"})
                    Return result
                End If

                ' 支持按路径打开指定工作簿，否则使用当前活动工作簿
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

                Dim sheets = New List(Of Dictionary(Of String, Object))()
                Dim sheetCount = wb.Worksheets.Count
                For i As Integer = 1 To sheetCount
                    ' 【关键】必须用索引循环，For Each + ReleaseComObject 会破坏 COM 枚举器导致重复/遗漏
                    Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                    Try
                        ws = DirectCast(wb.Worksheets(i), Microsoft.Office.Interop.Excel.Worksheet)
                        ' 跳过隐藏工作表（与 Excel UI 显示一致）
                        If ws.Visible <> Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible Then
                            Continue For
                        End If
                        Dim usedRangeAddr As String = ""
                        If ws.UsedRange IsNot Nothing Then
                            usedRangeAddr = ws.UsedRange.Address(False, False)
                        End If
                        Dim info = New Dictionary(Of String, Object) From {
                            {"name", ws.Name},
                            {"index", ws.Index},
                            {"usedRange", usedRangeAddr}
                        }
                        sheets.Add(info)
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine($"[ExecuteGetSheets] 读取第 {i} 个工作表失败: {ex.Message}")
                    Finally
                        Utils.ComReleaseHelper.Release(ws)
                    End Try
                Next
                
                ' 如果是按路径打开的，读取完后关闭（不保存）
                If openedByPath Then
                    wb.Close(SaveChanges:=False)
                    Utils.ComReleaseHelper.Release(wb)
                End If

                result.Content.Add(New McpTextContent With {
                    .Text = JsonConvert.SerializeObject(sheets, Formatting.Indented)
                })

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"获取工作表列表失败: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteCreateSheet(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim name As String = GetArg(Of String)(args, "name", "")
                If String.IsNullOrEmpty(name) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "工作表名称不能为空"})
                    Return result
                End If

                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Dim prevSheet As String = ""
                Try
                    Dim activeSheet = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    If activeSheet IsNot Nothing Then
                        prevSheet = activeSheet.Name
                        Marshal.ReleaseComObject(activeSheet)
                    End If
                Catch
                End Try

                Dim prevAlerts As Boolean = excelApp.DisplayAlerts
                excelApp.DisplayAlerts = False
                Try
                    Dim createErrorMsg As String = ""
                    ' 幂等处理：同名表已存在则"就地清空复用"；任一环节失败也不再外抛（见下方兜底）
                    Try
                        Dim oldWs As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                        Try
                            oldWs = DirectCast(excelApp.ActiveWorkbook.Worksheets(name), Microsoft.Office.Interop.Excel.Worksheet)
                        Catch
                            oldWs = Nothing  ' sheet 不存在，正常情况，走新建分支
                        End Try
                        If oldWs IsNot Nothing Then
                            Try
                                Dim cObjs As Microsoft.Office.Interop.Excel.ChartObjects = DirectCast(oldWs.ChartObjects, Microsoft.Office.Interop.Excel.ChartObjects)
                                While cObjs IsNot Nothing AndAlso CInt(cObjs.Count) > 0
                                    Try
                                        Dim co As Microsoft.Office.Interop.Excel.ChartObject = DirectCast(cObjs.Item(1), Microsoft.Office.Interop.Excel.ChartObject)
                                        If co IsNot Nothing Then co.Delete()
                                    Catch : End Try
                                End While
                                Utils.ComReleaseHelper.Release(cObjs)
                            Catch : End Try
                            ' 清透视表：Cells.Clear 不会移除 PivotTable，残留同名会导致后续 CreatePivotTable 0x800A03EC
                            Try
                                Dim pts As Microsoft.Office.Interop.Excel.PivotTables = DirectCast(oldWs.PivotTables(), Microsoft.Office.Interop.Excel.PivotTables)
                                For pi As Integer = pts.Count To 1 Step -1
                                    Try
                                        Dim opt As Microsoft.Office.Interop.Excel.PivotTable = DirectCast(pts.Item(pi), Microsoft.Office.Interop.Excel.PivotTable)
                                        opt.TableRange2.Clear()
                                    Catch : End Try
                                Next
                            Catch : End Try
                            ' 清空所有形状（含切片器/KPI 浮层），与"就地清空复用"语义一致
                            Try
                                Dim shps As Microsoft.Office.Interop.Excel.Shapes = oldWs.Shapes
                                For si As Integer = shps.Count To 1 Step -1
                                    Try
                                        DirectCast(shps.Item(si), Microsoft.Office.Interop.Excel.Shape).Delete()
                                    Catch : End Try
                                Next
                            Catch : End Try
                            Try : oldWs.Cells.Clear() : Catch : End Try
                            Try : oldWs.Cells.ClearFormats() : Catch : End Try
                            ws = oldWs
                        Else
                            
                            Try
                                ws = DirectCast(excelApp.ActiveWorkbook.Worksheets.Add(), Microsoft.Office.Interop.Excel.Worksheet)
                                ws.Name = name
                            Catch ex As Exception
                                createErrorMsg = ex.Message
                                ws = Nothing
                            End Try
                        End If
                    Catch
                        ws = Nothing
                    End Try
                    ' 【修复】去掉静默兜底！创建失败就直接报错，
                    ' 否则 agent 以为成功了，后续形状/图表全画到销售数据上
                    If ws Is Nothing Then
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = $"创建工作表 [{name}] 失败: {createErrorMsg}"})
                        Return result
                    End If
                    Dim usedName As String = ""
                    If ws IsNot Nothing Then
                        Try : usedName = ws.Name : Catch : End Try
                        ' 【关键修复】创建/定位看板工作表后必须显式激活，
                        ' 否则后续 excel_add_shape / excel_create_chart 等工具
                        ' 未传 sheet 参数时会默认画在 ActiveSheet（原始数据表）上，
                        ' 导致看板内容错位到销售数据 sheet。
                        Try : ws.Activate() : Catch : End Try
                    End If
                    Dim sheetMode As String = ""
                    ' 判断操作模式：新建/复用成功 / 兜底回退
                    If usedName = name Then
                        sheetMode = "新建/复用成功"
                    Else
                        sheetMode = "兜底回退(原活动表)"
                    End If
                    Dim output = New With {
                        .worksheet = usedName,
                        .previousSheet = prevSheet,
                        .mode = sheetMode,
                        .message = If(ws Is Nothing, "未能定位可用工作表", $"成功定位工作台: [{usedName}]（{sheetMode}），已激活为当前活动表")
                    }
                                        result.Content.Add(New McpTextContent With {
                        .Text = JsonConvert.SerializeObject(output, Formatting.Indented)
                    })
                Finally
                    excelApp.DisplayAlerts = prevAlerts
                    Utils.ComReleaseHelper.Release(ws)
                End Try

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"创建工作表失败: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteDeleteSheet(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim name As String = GetArg(Of String)(args, "name", "")
                Dim ignoreNotExists As Boolean = GetArg(Of Boolean)(args, "ignoreNotExists", False)
                If String.IsNullOrEmpty(name) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "工作表名称不能为空"})
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
                    ws = TryCast(excelApp.ActiveWorkbook.Worksheets(name), Microsoft.Office.Interop.Excel.Worksheet)
                Catch exIdx As Exception
                    ' DISP_E_BADINDEX (0x8002000B) = sheet 不存在
                    If ignoreNotExists Then
                        result.Content.Add(New McpTextContent With {.Text = $"工作表 {name} 不存在（已忽略）"})
                        Return result
                    End If
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"删除工作表失败: {exIdx.Message}"})
                    Return result
                End Try
                If ws Is Nothing Then
                    If ignoreNotExists Then
                        result.Content.Add(New McpTextContent With {.Text = $"工作表 {name} 不存在（已忽略）"})
                        Return result
                    End If
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"工作表 {name} 不存在"})
                    Return result
                End If

                ' 撤销记录：删除前捕获工作表快照
                Try
                    Dim snap = Bridges.UndoBridge.Instance.CaptureSheetSnapshot(ws)
                    Bridges.BridgeRegistry.Instance.Resolve(Of Bridges.IUndoBridge)().AddUndoAction(New Bridges.UndoAction(
                        "delete_sheet", "", Nothing, Nothing, name, snap))
                Catch undoEx As Exception
                    Debug.WriteLine($"[ExecuteDeleteSheet] 撤销记录失败: {undoEx.Message}")
                End Try

                excelApp.DisplayAlerts = False
                ws.Delete()
                excelApp.DisplayAlerts = True

                result.Content.Add(New McpTextContent With {
                    .Text = $"成功删除工作表: {name}"
                })

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"删除工作表失败: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteRenameSheet(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim oldName As String = GetArg(Of String)(args, "oldName", "")
                Dim newName As String = GetArg(Of String)(args, "newName", "")
                If String.IsNullOrEmpty(oldName) OrElse String.IsNullOrEmpty(newName) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "新旧工作表名称都不能为空"})
                    Return result
                End If

                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim ws = TryCast(excelApp.ActiveWorkbook.Worksheets(oldName), Microsoft.Office.Interop.Excel.Worksheet)
                If ws Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"工作表 {oldName} 不存在"})
                    Return result
                End If

                ws.Name = newName

                ' 撤销记录：记录新旧表名
                Try
                    Dim payload = JsonConvert.SerializeObject(New With {.oldName = oldName, .newName = newName})
                    Bridges.BridgeRegistry.Instance.Resolve(Of Bridges.IUndoBridge)().AddUndoAction(New Bridges.UndoAction(
                        "rename_sheet", "", Nothing, Nothing, newName, payload))
                Catch undoEx As Exception
                    Debug.WriteLine($"[ExecuteRenameSheet] 撤销记录失败: {undoEx.Message}")
                End Try

                result.Content.Add(New McpTextContent With {
                    .Text = $"成功重命名工作表: {oldName} -> {newName}"
                })

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"重命名工作表失败: {ex.Message}"})
            End Try
            Return result
        End Function

    End Class

End Namespace










