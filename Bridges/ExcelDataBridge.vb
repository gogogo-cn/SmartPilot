' TECHNICAL DEBT: 本文件大量使用 Object 类型进行 Excel COM Interop（Sheets/Range/Cells/Borders 的返回值均为 Object）
' Option Strict On 需要对几乎每个 COM 调用加 DirectCast，严重影响可读性；优先在应用层/UI层严格类型安全。
Option Strict Off
Option Explicit On
Option Infer On

Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Diagnostics
Imports Newtonsoft.Json
Imports Microsoft.Office.Interop.Excel
Imports ExcelAIVSTOAddIn.Utils

Namespace Bridges

    ''' <summary>
    ''' ExcelDataBridge - 从 ExcelBridge 中提取的 Excel 数据操作相关 API 方法
    ''' 提供 COM 可见的方法供前端 JavaScript 调用
    ''' </summary>
    <ComVisible(True)>
    <ClassInterface(ClassInterfaceType.AutoDispatch)>
    <ProgId("ExcelAIVSTOAddIn.ExcelDataBridge")>
    Public Class ExcelDataBridge
        Implements IBridge, IExcelDataReader

        ' 单例实例 - 使用 Lazy(Of T) 确保线程安全
        Private Shared ReadOnly _lazyInstance As New Lazy(Of ExcelDataBridge)(Function() New ExcelDataBridge(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

#Region "单例模式"

        ''' <summary>
        ''' 获取 ExcelDataBridge 单例实例（线程安全 Lazy 初始化）
        ''' </summary>
        Public Shared ReadOnly Property Instance As ExcelDataBridge
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        ''' <summary>
        ''' 私有构造函数（单例模式）
        ''' </summary>
        Private Sub New()
            ' 统一生命周期管理：实例化即登记，供 CustomSkillExecutor 通过 IExcelDataReader 解析读取 Excel 数据。
            BridgeRegistry.Instance.Register(Me)
        End Sub

#End Region

        ''' <summary>
        ''' 在 Excel 主线程（UI 线程）上执行有返回值的函数。
        ''' Excel COM 对象必须在创建它们的 STA 线程上访问。
        ''' 使用 TaskCompletionSource 实现异步等待，避免死锁。
        ''' </summary>
        Private Shared Function ExecuteOnExcelUIThread(Of T)(func As Func(Of T)) As T
            Dim tcs As New TaskCompletionSource(Of T)()

            ThisAddIn.RunOnUIThread(Sub()
                Try
                    Dim result = func()
                    tcs.SetResult(result)
                Catch ex As Exception
                    tcs.SetException(ex)
                End Try
            End Sub)

            Dim task = tcs.Task
            If task.Wait(10000) Then
                Return task.Result
            Else
                Throw New TimeoutException("Excel UI 线程操作超时")
            End If
        End Function

        ' 调用频率限制
        Private _dataWriteLimiter As Utils.RateLimiter = Utils.RateLimiter.ForDataWrite

#Region "Excel 数据操作 API"

        ''' <summary>
        ''' 智能填入数据到Excel（前端调用）
        ''' </summary>
        <ComVisible(True)>
        Public Function FillDataToExcel(dataJson As String, dataType As String) As String
            Try
                ' 频率限制检查
                If Not _dataWriteLimiter.TryAcquire() Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "数据写入频率过高，请稍后再试"})
                End If

                Dim excelApp As Application = Globals.ThisAddIn.Application
                If excelApp Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "Excel应用不可用"})
                End If

                ' 使用共享单例（无状态工具类，实例复用）
                Dim excelService = Services.ExcelService.Instance

                ' 首先尝试解析为数组格式
                Dim dataArray As List(Of Dictionary(Of String, Object)) = Nothing
                Dim dataObject As Dictionary(Of String, Object) = Nothing

                Try
                    ' 尝试解析为数组
                    dataArray = JsonConvert.DeserializeObject(Of List(Of Dictionary(Of String, Object)))(dataJson)
                Catch
                    ' 如果失败，尝试解析为单个对象
                    dataObject = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(dataJson)
                End Try

                If dataArray IsNot Nothing AndAlso dataArray.Count > 0 Then
                    ' 数组格式：调用数组版本的方法
                    excelService.FillDataToExcel(dataArray, dataType, "")
                    Return JsonConvert.SerializeObject(New With {.success = True, .message = "数据已成功填入Excel，共" & dataArray.Count & "行"})
                ElseIf dataObject IsNot Nothing AndAlso dataObject.Count > 0 Then
                    ' 单个对象格式：包装成数组后调用
                    dataArray = New List(Of Dictionary(Of String, Object)) From {dataObject}
                    excelService.FillDataToExcel(dataArray, dataType, "")
                    Return JsonConvert.SerializeObject(New With {.success = True, .message = "数据已成功填入Excel，共1行"})
                Else
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "数据为空或格式不正确"})
                End If

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("FillDataToExcel 错误: " & ex.Message)
                System.Diagnostics.Debug.WriteLine("堆栈: " & ex.StackTrace)
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 获取当前选择
        ''' </summary>
        <ComVisible(True)>
        Public Function GetSelection() As String
            Return ExecuteOnExcelUIThread(Function() GetSelectionInternal())
        End Function

        ''' <summary>
        ''' 获取所有当前选中的区域（支持多选区）
        ''' 返回结构：{ areaCount: N, areas: [{address, fullAddress, sheetName, workbookName, rowCount, colCount}, ...] }
        ''' </summary>
        <ComVisible(True)>
        Public Function GetMultiSelection() As String
            Return ExecuteOnExcelUIThread(Function() GetMultiSelectionInternal())
        End Function

        ''' <summary>
        ''' 获取所有选区内部实现
        ''' </summary>
        Private Function GetMultiSelectionInternal() As String
            Try
                Dim excelApp As Application = Globals.ThisAddIn.GetExcelApp()
                If excelApp Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.areaCount = 0, .areas = New List(Of Object)()})
                End If

                Dim selection As Object = Nothing
                Dim areasList As New List(Of Object)()

                Try
                    selection = excelApp.Selection
                    If selection Is Nothing Then
                        Return JsonConvert.SerializeObject(New With {.areaCount = 0, .areas = New List(Of Object)()})
                    End If

                    Dim selectionRange As Range = CType(selection, Range)
                    Dim areas As Areas = selectionRange.Areas
                    Dim areaCount = areas.Count

                    For i = 1 To areaCount
                        Dim area = areas(i)
                        Dim areaRange As Range = Nothing
                        Try
                            areaRange = CType(area, Range)

                            Dim sheetName As String = ""
                            Dim workbookName As String = ""
                            Try
                                If areaRange.Worksheet IsNot Nothing Then
                                    sheetName = areaRange.Worksheet.Name
                                    If areaRange.Worksheet.Parent IsNot Nothing Then
                                        workbookName = areaRange.Worksheet.Parent.Name
                                    End If
                                End If
                            Catch
                            End Try

                            areasList.Add(New With {
                                .address = areaRange.Address(False, False),
                                .fullAddress = areaRange.Address(True, True),
                                .sheetName = sheetName,
                                .workbookName = workbookName,
                                .rowCount = areaRange.Rows.Count,
                                .colCount = areaRange.Columns.Count
                            })
                        Finally
                            If areaRange IsNot Nothing Then Utils.ComReleaseHelper.Release(areaRange)
                        End Try
                    Next

                    Return JsonConvert.SerializeObject(New With {
                        .areaCount = areaCount,
                        .areas = areasList
                    })
                Finally
                    If selection IsNot Nothing Then Utils.ComReleaseHelper.Release(selection)
                End Try
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.error = ex.Message, .areaCount = 0, .areas = New List(Of Object)()})
            End Try
        End Function

        ''' <summary>
        ''' 获取当前选择（必须在 Excel UI 线程上调用）
        ''' </summary>
        Private Function GetSelectionInternal() As String
            Try
                Dim excelApp As Application = Globals.ThisAddIn.GetExcelApp()
                If excelApp Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.error = "Excel 未运行"})
                End If

                Dim selection As Object = Nothing
                Try
                    selection = excelApp.Selection
                    If selection Is Nothing Then
                        Return JsonConvert.SerializeObject(New With {.address = "", .text = ""})
                    End If

                    Dim selectionRange As Range = CType(selection, Range)
                    Dim worksheetName As String = ""
                    Try
                        If selectionRange.Worksheet IsNot Nothing Then
                            worksheetName = selectionRange.Worksheet.Name
                        End If
                    Catch
                        worksheetName = ""
                    End Try
                    Dim workbookName As String = ""
                    Try
                        If selectionRange.Worksheet IsNot Nothing AndAlso selectionRange.Worksheet.Parent IsNot Nothing Then
                            workbookName = selectionRange.Worksheet.Parent.Name
                        End If
                    Catch
                        workbookName = ""
                    End Try

                    Dim result = JsonConvert.SerializeObject(New With {
                        .address = selectionRange.Address,
                        .sheetName = worksheetName,
                        .workbookName = workbookName,
                        .text = If(selectionRange.Text Is Nothing, "", selectionRange.Text.ToString())
                    })
                    Return result
                Finally
                    ' 【修复】释放COM对象防止泄漏
                    If selection IsNot Nothing Then
                        Utils.ComReleaseHelper.Release(selection)
                    End If
                End Try
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 获取当前选中区域的结构化数据（含表头和数据行）
        ''' </summary>
        <ComVisible(True)>
        Public Function GetSelectedRangeData() As String Implements IExcelDataReader.GetSelectedRangeData
            Try
                Dim excelApp As Application = Globals.ThisAddIn.GetExcelApp()
                If excelApp Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "Excel 未运行"})
                End If

                Dim selection As Object = Nothing
                Dim rng As Range = Nothing
                Try
                    selection = excelApp.Selection
                    If selection Is Nothing Then
                        Return JsonConvert.SerializeObject(New With {.success = False, .error = "未选择任何区域"})
                    End If

                    rng = CType(selection, Range)
                    
                    ' 【修复】尝试向上扩展一行以获取表格标题
                    Dim titleRow As Integer = 1
                    Dim hasTitleRow As Boolean = False
                    Dim tableTitle As String = ""
                    
                    Try
                        ' 检查选中区域上方是否有一行（可能包含表格标题）
                        If rng.Row > 1 Then
                            Dim titleCell As Range = Nothing
                            Try
                                titleCell = CType(rng.Worksheet.Cells(rng.Row - 1, rng.Column), Range)
                                Dim titleValue = If(titleCell.Value Is Nothing, "", titleCell.Value.ToString().Trim())
                                
                                ' 判断是否为表格标题：非空、合并单元格或跨越多列
                                If Not String.IsNullOrEmpty(titleValue) Then
                                    ' 检查该单元格是否合并或跨越选中区域的列数
                                    If titleCell.MergeCells OrElse titleCell.MergeArea.Columns.Count >= rng.Columns.Count Then
                                        tableTitle = titleValue
                                        hasTitleRow = True
                                        titleRow = rng.Row - 1
                                    End If
                                End If
                            Finally
                                Utils.ComReleaseHelper.Release(titleCell)
                            End Try
                        End If
                    Catch
                        ' 如果检查失败，忽略标题行
                    End Try
                    
                    Dim rowCount As Integer = rng.Rows.Count
                    Dim colCount As Integer = rng.Columns.Count

                    ' 第一行作为表头
                    Dim headers As New List(Of String)()
                    For c As Integer = 1 To colCount
                        Dim headerCell As Object = Nothing
                        Try
                            headerCell = rng.Cells(1, c)
                            Dim headerValue = If(headerCell.Value Is Nothing, "列" & c, headerCell.Value.ToString())
                            headers.Add(headerValue)
                        Finally
                            Utils.ComReleaseHelper.Release(headerCell)
                        End Try
                    Next

                    ' 预读每列格式，缓存供数据读取使用（避免逐单元格 COM 调用）
                    Dim columnFormats = ReadColumnFormats(rng, colCount, 2)

                    ' 剩余行作为数据：保留原始类型，文本格式→字符串，日期→ISO格式，数字→保留Double
                    Dim rows As New List(Of Dictionary(Of String, Object))()
                    For r As Integer = 2 To rowCount
                        Dim rowDict As New Dictionary(Of String, Object)()
                        For c As Integer = 1 To colCount
                            Dim cell As Object = Nothing
                            Dim cellValue As Object = Nothing
                            Try
                                cell = rng.Cells(r, c)
                                cellValue = cell.Value2
                                rowDict(headers(c - 1)) = NormalizeCellValue(cell, cellValue, columnFormats(c - 1))
                            Finally
                                Utils.ComReleaseHelper.Release(cell)
                            End Try
                        Next
                        rows.Add(rowDict)
                    Next

                    ' 【选区范围判断·2026-10-22】补 sheetName/workbookName：此前本方法只返回 address，
                    ' ContextEngine 又优先走它 → turn.context.selection 结构性缺表名/工作簿名，
                    ' 下游（prompt-builder/agent-loop/text-pipeline）拿到的全是空串。
                    ' 与 GetUsedRangeData/GetRangeData 的返回结构对齐（取值失败不影响主流程）。
                    Dim selSheetName As String = ""
                    Dim selWorkbookName As String = ""
                    Try
                        selSheetName = CStr(rng.Worksheet.Name)
                        selWorkbookName = CStr(rng.Worksheet.Parent.Name)
                    Catch
                    End Try

                    Return JsonConvert.SerializeObject(New With {
                        .success = True,
                        .title = tableTitle,
                        .hasTitle = hasTitleRow,
                        .headers = headers,
                        .rows = rows,
                        .rowCount = rowCount - 1,
                        .colCount = colCount,
                        .address = rng.Address,
                        .sheetName = selSheetName,
                        .workbookName = selWorkbookName
                    })
                Finally
                    Utils.ComReleaseHelper.Release(rng)
                    If selection IsNot Nothing Then Utils.ComReleaseHelper.Release(selection)
                End Try
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 获取当前活动工作表的已用范围数据（用于用户指代"当前工作表"但无显式选区时）
        ''' </summary>
        <ComVisible(True)>
        Public Function GetUsedRangeData() As String Implements IExcelDataReader.GetUsedRangeData
            Return ExecuteOnExcelUIThread(Function() GetUsedRangeDataInternal())
        End Function

        Private Function GetUsedRangeDataInternal() As String
            Try
                Dim excelApp As Application = Globals.ThisAddIn.GetExcelApp()
                If excelApp Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "Excel 未运行"})
                End If

                Dim usedRange As Object = Nothing
                Dim rng As Range = Nothing
                Dim sheet As Object = Nothing
                Try
                    sheet = excelApp.ActiveSheet
                    If sheet Is Nothing Then
                        Return JsonConvert.SerializeObject(New With {.success = False, .error = "没有活动工作表"})
                    End If

                    usedRange = CType(sheet, Worksheet).UsedRange
                    If usedRange Is Nothing Then
                        Return JsonConvert.SerializeObject(New With {.success = False, .error = "当前工作表没有数据"})
                    End If

                    rng = CType(usedRange, Range)
                    Dim rowCount As Integer = rng.Rows.Count
                    Dim colCount As Integer = rng.Columns.Count

                    ' 【修复】自动检测合并标题行：如果第一行是跨多列的合并标题，则表头从第二行开始。
                    Dim headerRow As Integer = 1
                    Dim firstCell As Object = Nothing
                    Dim mergeArea As Object = Nothing
                    Try
                        firstCell = rng.Cells(1, 1)
                        mergeArea = firstCell.MergeArea
                        If mergeArea IsNot Nothing AndAlso CInt(mergeArea.Columns.Count) >= colCount Then
                            headerRow = 2
                        End If
                    Catch
                    Finally
                        Utils.ComReleaseHelper.Release(mergeArea)
                        Utils.ComReleaseHelper.Release(firstCell)
                    End Try

                    ' 表头行
                    Dim headers As New List(Of String)()
                    For c As Integer = 1 To colCount
                        Dim headerCell As Object = Nothing
                        Try
                            headerCell = rng.Cells(headerRow, c)
                            Dim headerValue = If(headerCell.Value Is Nothing, "列" & c, headerCell.Value.ToString())
                            headers.Add(headerValue)
                        Finally
                            Utils.ComReleaseHelper.Release(headerCell)
                        End Try
                    Next

                    ' 预读每列格式，缓存供数据读取使用（避免逐单元格 COM 调用）
                    Dim columnFormats = ReadColumnFormats(rng, colCount, headerRow + 1)

                    ' 剩余行作为数据：保留原始类型，文本格式→字符串，日期→ISO格式，数字→保留Double
                    Dim rows As New List(Of Dictionary(Of String, Object))()
                    For r As Integer = headerRow + 1 To rowCount
                        Dim rowDict As New Dictionary(Of String, Object)()
                        For c As Integer = 1 To colCount
                            Dim cell As Object = Nothing
                            Dim cellValue As Object = Nothing
                            Try
                                cell = rng.Cells(r, c)
                                cellValue = cell.Value2
                                rowDict(headers(c - 1)) = NormalizeCellValue(cell, cellValue, columnFormats(c - 1))
                            Finally
                                Utils.ComReleaseHelper.Release(cell)
                            End Try
                        Next
                        rows.Add(rowDict)
                    Next

                    Return JsonConvert.SerializeObject(New With {
                        .success = True,
                        .headers = headers,
                        .rows = rows,
                        .rowCount = rowCount - headerRow,
                        .colCount = colCount,
                        .address = rng.Address,
                        .sheetName = CType(sheet, Worksheet).Name,
                        .workbookName = CType(sheet, Worksheet).Parent.Name
                    })
                Finally
                    Utils.ComReleaseHelper.Release(rng)
                    Utils.ComReleaseHelper.Release(usedRange)
                    Utils.ComReleaseHelper.Release(sheet)
                End Try
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 按地址获取 Excel 区域数据（支持引用区域功能）
        ''' 地址格式: [工作簿名.xlsx]工作表名!$A$1:$H$10 或 工作表名!$A$1:$H$10
        ''' </summary>
        <ComVisible(True)>
        Public Function GetRangeData(rangeAddress As String) As String Implements IExcelDataReader.GetRangeData
            If String.IsNullOrWhiteSpace(rangeAddress) Then
                Return JsonConvert.SerializeObject(New With {.success = False, .error = "地址不能为空"})
            End If

            Return ExecuteOnExcelUIThread(Function() GetRangeDataInternal(rangeAddress))
        End Function

        Private Shared Function GetRangeDataInternal(rangeAddress As String) As String
            Try
                Dim excelApp As Application = Globals.ThisAddIn.GetExcelApp()
                If excelApp Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "Excel 未运行"})
                End If

                System.Diagnostics.Debug.WriteLine("[ExcelDataBridge.GetRangeData] 地址: " & rangeAddress)

                ' 解析地址格式: [workbook.xlsx]Sheet1!A1:B10 或 Sheet1!A1:B10
                Dim workbookName As String = ""
                Dim sheetAndRange As String = rangeAddress
                If rangeAddress.StartsWith("[") Then
                    Dim closeIdx As Integer = rangeAddress.IndexOf("]")
                    If closeIdx > 0 Then
                        workbookName = rangeAddress.Substring(1, closeIdx - 1)
                        sheetAndRange = rangeAddress.Substring(closeIdx + 1)
                    End If
                End If

                ' 解析 sheet!range 部分
                Dim sheetName As String = ""
                Dim targetRangeAddress As String = sheetAndRange
                Dim exclIdx As Integer = sheetAndRange.IndexOf("!"c)
                If exclIdx >= 0 Then
                    sheetName = sheetAndRange.Substring(0, exclIdx).Trim()
                    targetRangeAddress = sheetAndRange.Substring(exclIdx + 1).Trim()
                End If

                Dim rng As Range = Nothing
                Dim ws As Worksheet = Nothing
                Try
                    If Not String.IsNullOrEmpty(workbookName) Then
                        ' 显式从指定工作簿获取 Range
                        Dim wb As Object = Nothing
                        Try
                            wb = excelApp.Workbooks(workbookName)
                            If wb Is Nothing Then
                                Return JsonConvert.SerializeObject(New With {.success = False, .error = "未找到工作簿: " & workbookName})
                            End If
                            If String.IsNullOrEmpty(sheetName) Then
                                rng = CType(wb.Range(targetRangeAddress), Range)
                            Else
                                ws = CType(wb.Worksheets(sheetName), Worksheet)
                                rng = CType(ws.Range(targetRangeAddress), Range)
                            End If
                        Finally
                            Utils.ComReleaseHelper.Release(wb)
                        End Try
                    Else
                        rng = CType(excelApp.Range(sheetAndRange), Range)
                    End If

                    If rng Is Nothing Then
                        Return JsonConvert.SerializeObject(New With {.success = False, .error = "无法解析地址: " & rangeAddress})
                    End If

                    Dim rowCount As Integer = rng.Rows.Count
                    Dim colCount As Integer = rng.Columns.Count

                    ' 第一行作为表头
                    Dim headers As New List(Of String)()
                    For c As Integer = 1 To colCount
                        Dim headerCell As Object = Nothing
                        Try
                            headerCell = rng.Cells(1, c)
                            Dim headerValue = If(headerCell.Value Is Nothing, "列" & c, headerCell.Value.ToString())
                            headers.Add(headerValue)
                        Finally
                            Utils.ComReleaseHelper.Release(headerCell)
                        End Try
                    Next

                    ' 预读每列格式，缓存供数据读取使用（避免逐单元格 COM 调用）
                    Dim columnFormats = ReadColumnFormats(rng, colCount, 2)

                    ' 剩余行作为数据：保留原始类型，文本格式→字符串，日期→ISO格式，数字→保留Double
                    Dim rows As New List(Of Dictionary(Of String, Object))()
                    For r As Integer = 2 To rowCount
                        Dim rowDict As New Dictionary(Of String, Object)()
                        For c As Integer = 1 To colCount
                            Dim cell As Object = Nothing
                            Dim cellValue As Object = Nothing
                            Try
                                cell = rng.Cells(r, c)
                                cellValue = cell.Value2
                                rowDict(headers(c - 1)) = NormalizeCellValue(cell, cellValue, columnFormats(c - 1))
                            Finally
                                Utils.ComReleaseHelper.Release(cell)
                            End Try
                        Next
                        rows.Add(rowDict)
                    Next

                    Return JsonConvert.SerializeObject(New With {
                        .success = True,
                        .headers = headers,
                        .rows = rows,
                        .rowCount = rowCount - 1,
                        .colCount = colCount,
                        .address = rng.Address,
                        .sheetName = rng.Worksheet.Name,
                        .workbookName = rng.Worksheet.Parent.Name
                    })
                Finally
                    Utils.ComReleaseHelper.Release(rng)
                    Utils.ComReleaseHelper.Release(ws)
                End Try
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 获取单元格值
        ''' </summary>
        <ComVisible(True)>
        Public Function GetCellValue(address As String) As String
            Dim cell As Object = Nothing
            Try
                Dim excelApp As Application = Globals.ThisAddIn.GetExcelApp()
                If excelApp Is Nothing Then
                    Return ""
                End If

                cell = excelApp.Range(address)
                If cell Is Nothing Then
                    Return ""
                End If

                Dim cellRange As Range = Nothing
                Try
                    cellRange = CType(cell, Range)
                    Dim value = If(cellRange.Value Is Nothing, "", cellRange.Value.ToString())
                    Return value
                Finally
                    Utils.ComReleaseHelper.Release(cellRange)
                End Try
            Catch ex As Exception
                Return ""
            Finally
                ' 【修复】释放COM对象
                If cell IsNot Nothing Then Utils.ComReleaseHelper.Release(cell)
            End Try
        End Function

        ''' <summary>
        ''' 设置单元格值
        ''' </summary>
        <ComVisible(True)>
        Public Function SetCellValue(address As String, value As String) As String
            Dim cell As Object = Nothing
            Dim activeSheet As Object = Nothing
            Dim cellWorksheet As Object = Nothing

            Try
                Dim excelApp As Application = Globals.ThisAddIn.GetExcelApp()
                If excelApp Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "Excel 未运行"})
                End If

                cell = excelApp.Range(address)
                If cell Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "无效单元格地址"})
                End If

                Dim cellRange As Range = Nothing
                Try
                    cellRange = CType(cell, Range)
                    ' 记录旧值用于撤销
                    Dim oldValue As Object = cellRange.Value
                    Dim sheetName As String = ""
                    Try
                        activeSheet = excelApp.ActiveSheet
                        If activeSheet IsNot Nothing Then
                            sheetName = CType(activeSheet, Worksheet).Name
                        Else
                            cellWorksheet = cellRange.Worksheet
                            sheetName = If(cellWorksheet IsNot Nothing, CType(cellWorksheet, Worksheet).Name, "Unknown")
                        End If
                    Catch
                        sheetName = "Unknown"
                    End Try

                    ' 添加到撤销栈：整表备份（必须在写入之前，撤销时整张表还原）
                    Try
                        BridgeRegistry.Instance.Resolve(Of IUndoBridge)()?.AddSheetBackup(sheetName)
                    Catch
                    End Try

                    cellRange.NumberFormatLocal = "@"
                    cellRange.Value2 = SanitizeCellValue(value)

                    Return JsonConvert.SerializeObject(New With {.success = True})
                Finally
                    Utils.ComReleaseHelper.Release(cellRange)
                End Try
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            Finally
                ' 【修复】释放COM对象防止泄漏
                If cellWorksheet IsNot Nothing Then Utils.ComReleaseHelper.Release(cellWorksheet)
                If activeSheet IsNot Nothing Then Utils.ComReleaseHelper.Release(activeSheet)
                If cell IsNot Nothing Then Utils.ComReleaseHelper.Release(cell)
            End Try
        End Function

        ''' <summary>
        ''' 删除指定工作表的行
        ''' </summary>
        ''' <param name="sheetName">工作表名称</param>
        ''' <param name="rowIndex">行号（1-based）</param>
        ''' <param name="count">删除行数</param>
        <ComVisible(True)>
        Public Function DeleteRows(sheetName As String, rowIndex As Integer, count As Integer) As String
            Dim ws As Worksheet = Nothing
            Dim rows As Range = Nothing
            Try
                Dim excelApp As Application = Globals.ThisAddIn.GetExcelApp()
                If excelApp Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "Excel应用不可用"})
                End If

                If String.IsNullOrWhiteSpace(sheetName) Then
                    ws = CType(excelApp.ActiveSheet, Worksheet)
                Else
                    ws = CType(excelApp.ActiveWorkbook.Worksheets(sheetName), Worksheet)
                End If
                sheetName = ws.Name

                If rowIndex < 1 OrElse count < 1 Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "行号或行数无效"})
                End If

                rows = ws.Rows($"{rowIndex}:{rowIndex + count - 1}")

                ' 捕获被删除行的数据用于撤销
                Dim undoPayload As String = ""
                Try
                    Dim lastCol As Integer = ws.UsedRange.Columns.Count
                    If lastCol < 1 Then lastCol = 1
                    Dim dataRange = ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex + count - 1, lastCol))
                    Dim dataObj = dataRange.Value
                    Dim dataArr As Object()() = Nothing
                    If dataObj IsNot Nothing Then
                        Dim arr = DirectCast(dataObj, Object(,))
                        Dim lb1 = arr.GetLowerBound(0), ub1 = arr.GetUpperBound(0)
                        Dim lb2 = arr.GetLowerBound(1), ub2 = arr.GetUpperBound(1)
                        Dim rc = ub1 - lb1 + 1, cc = ub2 - lb2 + 1
                        dataArr = New Object(rc - 1)() {}
                        For r = 0 To rc - 1
                            Dim rowArr(cc - 1) As Object
                            For c = 0 To cc - 1
                                rowArr(c) = arr(lb1 + r, lb2 + c)
                            Next
                            dataArr(r) = rowArr
                        Next
                    End If
                    Utils.ComReleaseHelper.Release(dataRange)
                    undoPayload = JsonConvert.SerializeObject(New With {.index = rowIndex, .count = count, .data = dataArr})
                Catch
                    undoPayload = ""
                End Try

                ' 撤销：整表备份（必须在删除之前）
                Try
                    BridgeRegistry.Instance.Resolve(Of IUndoBridge)()?.AddSheetBackup(sheetName)
                Catch
                End Try

                rows.Delete()

                Return JsonConvert.SerializeObject(New With {.success = True})
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            Finally
                Utils.ComReleaseHelper.Release(rows)
                Utils.ComReleaseHelper.Release(ws)
            End Try
        End Function

        ''' <summary>
        ''' 在指定工作表插入行
        ''' </summary>
        ''' <param name="sheetName">工作表名称</param>
        ''' <param name="rowIndex">插入位置行号（1-based）</param>
        ''' <param name="count">插入行数</param>
        <ComVisible(True)>
        Public Function InsertRows(sheetName As String, rowIndex As Integer, count As Integer) As String
            Dim ws As Worksheet = Nothing
            Dim rows As Range = Nothing
            Try
                Dim excelApp As Application = Globals.ThisAddIn.GetExcelApp()
                If excelApp Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "Excel应用不可用"})
                End If

                If String.IsNullOrWhiteSpace(sheetName) Then
                    ws = CType(excelApp.ActiveSheet, Worksheet)
                Else
                    ws = CType(excelApp.ActiveWorkbook.Worksheets(sheetName), Worksheet)
                End If
                sheetName = ws.Name

                If rowIndex < 1 OrElse count < 1 Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "行号或行数无效"})
                End If

                rows = ws.Rows($"{rowIndex}:{rowIndex + count - 1}")

                ' 撤销：整表备份（必须在插入之前）
                Try
                    BridgeRegistry.Instance.Resolve(Of IUndoBridge)()?.AddSheetBackup(sheetName)
                Catch
                End Try

                rows.Insert()

                Return JsonConvert.SerializeObject(New With {.success = True})
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            Finally
                Utils.ComReleaseHelper.Release(rows)
                Utils.ComReleaseHelper.Release(ws)
            End Try
        End Function

        ''' <summary>
        ''' 获取活动工作表名称
        ''' </summary>
        <ComVisible(True)>
        Public Function GetActiveSheetName() As String
            Return ExecuteOnExcelUIThread(Function() GetActiveSheetNameInternal())
        End Function

        Private Function GetActiveSheetNameInternal() As String
            Try
                Dim excelApp As Application = Globals.ThisAddIn.GetExcelApp()
                If excelApp Is Nothing OrElse excelApp.ActiveSheet Is Nothing Then
                    Return ""
                End If
                Return CType(excelApp.ActiveSheet, Worksheet).Name
            Catch ex As Exception
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' 获取当前活动工作簿中所有【可见】工作表的名称列表（JSON 字符串数组）。
        ''' 供前端把"真实工作表清单"注入 LLM 上下文，从根上避免模型编造表名（如"订单"）。
        ''' 用索引循环 + ComReleaseHelper，与 ExcelMcpTools.Core.SheetManage 一致，避免破坏 COM 枚举器。
        ''' </summary>
        <ComVisible(True)>
        Public Function GetSheets() As String
            Try
                Return ExecuteOnExcelUIThread(Function() GetSheetsInternal())
            Catch ex As Exception
                Return "[]"
            End Try
        End Function

        Private Function GetSheetsInternal() As String
            Dim names As New List(Of String)()
            Try
                Dim excelApp As Application = Globals.ThisAddIn.GetExcelApp()
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    Return "[]"
                End If
                Dim wb = excelApp.ActiveWorkbook
                Dim cnt = wb.Worksheets.Count
                For i As Integer = 1 To cnt
                    Dim ws As Worksheet = Nothing
                    Try
                        ws = DirectCast(wb.Worksheets(i), Worksheet)
                        If ws.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible Then
                            names.Add(ws.Name)
                        End If
                    Catch
                        ' 单个表读取失败不阻塞整体
                    Finally
                        If ws IsNot Nothing Then Utils.ComReleaseHelper.Release(ws)
                    End Try
                Next
                Return Newtonsoft.Json.JsonConvert.SerializeObject(names)
            Catch ex As Exception
                Return "[]"
            End Try
        End Function

        ''' <summary>
        ''' 获取用户在 Excel 标签栏当前选中的多个工作表名称（JSON 字符串数组）。
        ''' 用户按住 Ctrl/Shift 点击标签多选时，这些"选中态"表名需要注入 LLM 上下文，
        ''' 让模型知道"合并/拆分操作应该作用于这几张表"，而不是只读活动工作表。
        ''' 单选或未多选时返回空数组（前端据此判定退回到"全部表"或"活动工作表"语义）。
        ''' </summary>
        <ComVisible(True)>
        Public Function GetSelectedSheetNames() As String
            Try
                Return ExecuteOnExcelUIThread(Function() GetSelectedSheetNamesInternal())
            Catch ex As Exception
                Return "[]"
            End Try
        End Function

        Private Function GetSelectedSheetNamesInternal() As String
            Dim names As New List(Of String)()
            Try
                Dim excelApp As Application = Globals.ThisAddIn.GetExcelApp()
                If excelApp Is Nothing OrElse excelApp.ActiveWindow Is Nothing Then
                    Return "[]"
                End If
                Dim aw = excelApp.ActiveWindow
                Dim sel = TryCast(aw.SelectedSheets, Microsoft.Office.Interop.Excel.Sheets)
                If sel Is Nothing OrElse sel.Count <= 1 Then
                    ' 没多选或只选了一张（活动工作表）→ 返回空数组，前端知道"没有显式多选"
                    Utils.ComReleaseHelper.Release(aw)
                    If sel IsNot Nothing Then Utils.ComReleaseHelper.Release(sel)
                    Return "[]"
                End If
                For i As Integer = 1 To sel.Count
                    Dim ws As Worksheet = Nothing
                    Try
                        ws = DirectCast(sel(i), Worksheet)
                        If ws IsNot Nothing AndAlso ws.Visible = Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible Then
                            names.Add(ws.Name)
                        End If
                    Catch
                    Finally
                        If ws IsNot Nothing Then Utils.ComReleaseHelper.Release(ws)
                    End Try
                Next
                Utils.ComReleaseHelper.Release(aw)
                If sel IsNot Nothing Then Utils.ComReleaseHelper.Release(sel)
                Return Newtonsoft.Json.JsonConvert.SerializeObject(names)
            Catch ex As Exception
                Return "[]"
            End Try
        End Function

        ''' <summary>
        ''' 获取活动单元格地址
        ''' </summary>
        <ComVisible(True)>
        Public Function GetActiveCell() As String
            Return ExecuteOnExcelUIThread(Function() GetActiveCellInternal())
        End Function

        Private Function GetActiveCellInternal() As String
            Try
                Dim excelApp As Application = Globals.ThisAddIn.GetExcelApp()
                If excelApp Is Nothing Then Return ""
                Dim activeCell As Object = excelApp.ActiveCell
                If activeCell Is Nothing Then Return ""
                Dim activeCellRange As Range = CType(activeCell, Range)
                Dim result = activeCellRange.Address
                Utils.ComReleaseHelper.Release(activeCell)
                Return result
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[GetActiveCell] 错误: {ex.Message}")
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' 获取活动工作簿名称
        ''' </summary>
        <ComVisible(True)>
        Public Function GetActiveWorkbookName() As String
            Return ExecuteOnExcelUIThread(Function() GetActiveWorkbookNameInternal())
        End Function

        Private Function GetActiveWorkbookNameInternal() As String
            Try
                Dim excelApp As Application = Globals.ThisAddIn.GetExcelApp()
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    Return ""
                End If
                Return excelApp.ActiveWorkbook.Name
            Catch ex As Exception
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' 按文件名搜索已打开的工作簿，返回完整路径（供前端文件选择器获取路径）
        ''' </summary>
        <ComVisible(True)>
        Public Function GetWorkbookPath(fileName As String) As String
            Return ExecuteOnExcelUIThread(Function() GetWorkbookPathInternal(fileName))
        End Function

        Private Function GetWorkbookPathInternal(fileName As String) As String
            Try
                Dim excelApp As Application = Globals.ThisAddIn.GetExcelApp()
                If excelApp Is Nothing Then Return ""
                Dim wbCount = excelApp.Workbooks.Count
                For i As Integer = 1 To wbCount
                    Dim wb As Workbook = Nothing
                    Try
                        wb = excelApp.Workbooks(i)
                        If String.Equals(wb.Name, fileName, StringComparison.OrdinalIgnoreCase) Then
                            Return wb.FullName
                        End If
                    Finally
                        If wb IsNot Nothing Then Utils.ComReleaseHelper.Release(wb)
                    End Try
                Next
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[GetWorkbookPath] 错误: {ex.Message}")
            End Try
            Return ""
        End Function

        ''' <summary>
        ''' 获取当前 Excel 中所有打开的工作簿列表（名称、路径、工作表数）
        ''' </summary>
        <ComVisible(True)>
        Public Function GetOpenWorkbooks() As String
            Return ExecuteOnExcelUIThread(Function() GetOpenWorkbooksInternal())
        End Function

        Private Function GetOpenWorkbooksInternal() As String
            Try
                Dim excelApp As Application = Globals.ThisAddIn.GetExcelApp()
                If excelApp Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .message = "Excel 应用程序未启动"})
                End If

                Dim wbCount = excelApp.Workbooks.Count
                Dim workbooks = New List(Of Object)()
                For i As Integer = 1 To wbCount
                    Dim wb As Workbook = Nothing
                    Try
                        wb = excelApp.Workbooks(i)
                        Dim visibleSheetCount = 0
                        For j As Integer = 1 To wb.Worksheets.Count
                            Dim ws As Worksheet = Nothing
                            Try
                                ws = DirectCast(wb.Worksheets(j), Worksheet)
                                If ws.Visible = XlSheetVisibility.xlSheetVisible Then visibleSheetCount += 1
                            Finally
                                If ws IsNot Nothing Then Utils.ComReleaseHelper.Release(ws)
                            End Try
                        Next

                        workbooks.Add(New With {
                            .name = wb.Name,
                            .fullName = wb.FullName,
                            .path = wb.Path,
                            .sheetCount = visibleSheetCount,
                            .isActive = (excelApp.ActiveWorkbook IsNot Nothing AndAlso String.Equals(excelApp.ActiveWorkbook.Name, wb.Name))
                        })
                    Finally
                        If wb IsNot Nothing Then Utils.ComReleaseHelper.Release(wb)
                    End Try
                Next

                Return JsonConvert.SerializeObject(New With {.success = True, .count = workbooks.Count, .workbooks = workbooks})
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[GetOpenWorkbooks] 错误: {ex.Message}")
                Return JsonConvert.SerializeObject(New With {.success = False, .message = $"获取打开的工作簿失败: {ex.Message}"})
            End Try
        End Function

        ''' <summary>
        ''' 激活（切换到）指定名称/路径的工作簿，使其变为 ActiveWorkbook
        ''' </summary>
        <ComVisible(True)>
        Public Function ActivateWorkbook(name As String) As String
            Return ExecuteOnExcelUIThread(Function() ActivateWorkbookInternal(name))
        End Function

        Private Function ActivateWorkbookInternal(name As String) As String
            Try
                Dim excelApp As Application = Globals.ThisAddIn.GetExcelApp()
                If excelApp Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .message = "Excel 应用程序未启动"})
                End If
                If String.IsNullOrEmpty(name) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .message = "工作簿名称不能为空"})
                End If

                ' 先只解析名称：循环里 Release 后再 Activate，避免「找到即 Release」导致
                ' target.Activate() 打到已释放 RCW（切换成功但后续 GetSheets 读到空/旧表）。
                Dim matchedName As String = Nothing
                For i As Integer = 1 To excelApp.Workbooks.Count
                    Dim wb As Workbook = Nothing
                    Try
                        wb = excelApp.Workbooks(i)
                        If String.Equals(wb.Name, name, StringComparison.OrdinalIgnoreCase) _
                            OrElse String.Equals(wb.FullName, name, StringComparison.OrdinalIgnoreCase) Then
                            matchedName = wb.Name
                            Exit For
                        End If
                    Finally
                        If wb IsNot Nothing Then Utils.ComReleaseHelper.Release(wb)
                    End Try
                Next
                If String.IsNullOrEmpty(matchedName) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .message = "未找到工作簿: " & name})
                End If

                Dim target As Workbook = Nothing
                Try
                    target = excelApp.Workbooks(matchedName)
                    target.Activate()

                    Dim sheetNames As New List(Of String)()
                    Dim activeSheetName As String = ""
                    Dim cnt = target.Worksheets.Count
                    For j As Integer = 1 To cnt
                        Dim ws As Worksheet = Nothing
                        Try
                            ws = DirectCast(target.Worksheets(j), Worksheet)
                            If ws.Visible = XlSheetVisibility.xlSheetVisible Then
                                sheetNames.Add(ws.Name)
                            End If
                        Finally
                            If ws IsNot Nothing Then Utils.ComReleaseHelper.Release(ws)
                        End Try
                    Next
                    Try
                        If target.ActiveSheet IsNot Nothing Then
                            activeSheetName = CType(target.ActiveSheet, Worksheet).Name
                        End If
                    Catch
                        If sheetNames.Count > 0 Then activeSheetName = sheetNames(0)
                    End Try

                    ' 带上 sheets，前端可直接刷新下拉，不必再打一次 GetSheets（避免竞态）
                    Return JsonConvert.SerializeObject(New With {
                        .success = True,
                        .name = matchedName,
                        .activeSheet = activeSheetName,
                        .sheetCount = sheetNames.Count,
                        .sheets = sheetNames
                    })
                Finally
                    If target IsNot Nothing Then Utils.ComReleaseHelper.Release(target)
                End Try
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[ActivateWorkbook] 错误: {ex.Message}")
                Return JsonConvert.SerializeObject(New With {.success = False, .message = $"激活工作簿失败: {ex.Message}"})
            End Try
        End Function

        ''' <summary>
        ''' 激活（切换到）当前活动工作簿中指定名称的工作表
        ''' 前端 excel-connection-widget.js 的 switchSheet() 调用此方法
        ''' </summary>
        <ComVisible(True)>
        Public Function ActivateSheet(sheetName As String) As String
            Return ExecuteOnExcelUIThread(Function() ActivateSheetInternal(sheetName))
        End Function

        Private Function ActivateSheetInternal(sheetName As String) As String
            Try
                Dim excelApp As Application = Globals.ThisAddIn.GetExcelApp()
                If excelApp Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .message = "Excel 应用程序未启动"})
                End If
                If String.IsNullOrEmpty(sheetName) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .message = "工作表名称不能为空"})
                End If
                Dim activeWb As Workbook = excelApp.ActiveWorkbook
                If activeWb Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .message = "没有活动的工作簿"})
                End If
                Dim targetSheet As Worksheet = Nothing
                For i As Integer = 1 To activeWb.Worksheets.Count
                    Dim ws As Worksheet = Nothing
                    Try
                        ws = CType(activeWb.Worksheets(i), Worksheet)
                        If String.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase) Then
                            targetSheet = ws
                            Exit For
                        End If
                    Finally
                        If ws IsNot Nothing Then Utils.ComReleaseHelper.Release(ws)
                    End Try
                Next
                If targetSheet Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .message = "未找到工作表: " & sheetName})
                End If
                targetSheet.Activate()
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .sheetName = targetSheet.Name
                })
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[ActivateSheet] 错误: {ex.Message}")
                Return JsonConvert.SerializeObject(New With {.success = False, .message = $"激活工作表失败: {ex.Message}"})
            End Try
        End Function

#End Region

#Region "格式化与渲染 API"

        ''' <summary>
        ''' 将Markdown文本渲染到Excel，每行文本独立成行显示
        ''' </summary>
        ''' <param name="markdownText">Markdown格式文本</param>
        ''' <param name="startCell">起始单元格地址，如"A1"</param>
        ''' <returns>操作结果JSON</returns>
        <ComVisible(True)>
        Public Function RenderFormattedText(markdownText As String, Optional startCell As String = "") As String
            Dim renderer As New ExcelMarkdownRenderer(Globals.ThisAddIn.Application)
            Return renderer.RenderFormattedText(markdownText, startCell)
        End Function

        ''' <summary>
        ''' 插入表格数据到Excel（保持原始表格布局）
        ''' </summary>
        ''' <param name="rowsJson">JSON格式的二维数组，包含表头和数据行</param>
        ''' <returns>操作结果JSON</returns>
        <ComVisible(True)>
        Public Function InsertTableData(rowsJson As String) As String
            Dim workbook As Object = Nothing
            Dim sheet As Object = Nothing
            Dim activeCell As Object = Nothing
            Dim rangeCell As Object = Nothing
            Dim colRange As Object = Nothing

            Try
                ' 频率限制检查
                If Not _dataWriteLimiter.TryAcquire() Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "数据写入频率过高，请稍后再试"})
                End If

                Dim excelApp As Application = Globals.ThisAddIn.Application
                If excelApp Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "Excel应用不可用"})
                End If

                ' 解析JSON数据（二维数组格式）
                Dim rows As List(Of List(Of String)) = JsonConvert.DeserializeObject(Of List(Of List(Of String)))(rowsJson)

                If rows Is Nothing OrElse rows.Count = 0 Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "表格数据为空"})
                End If

                ' 获取当前工作表
                workbook = excelApp.ActiveWorkbook
                If workbook Is Nothing Then
                    workbook = excelApp.Workbooks.Add()
                End If

                sheet = workbook.ActiveSheet
                If sheet Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "无法获取活动工作表"})
                End If

                ' 确定起始位置（有选中区域则从选中区域左上角开始，否则从A1开始）
                Dim startRow As Integer = 1
                Dim startCol As Integer = 1
                activeCell = excelApp.ActiveCell
                If activeCell IsNot Nothing Then
                    Dim activeCellRange As Range = CType(activeCell, Range)
                    startRow = activeCellRange.Row
                    startCol = activeCellRange.Column
                End If

                ' 写入表格数据（保持原始布局，同时防护公式注入）
                For rowIndex = 0 To rows.Count - 1
                    Dim row = rows(rowIndex)
                    For colIndex = 0 To row.Count - 1
                        rangeCell = CType(excelApp.Cells(startRow + rowIndex, startCol + colIndex), Range)
                        CType(rangeCell, Range).NumberFormatLocal = "@"
                        CType(rangeCell, Range).Value2 = SanitizeCellValue(row(colIndex))
                    Next
                Next

                ' 自动调整列宽
                Dim maxCols = rows.Max(Function(r) r.Count)
                Dim sheetWorksheet As Worksheet = CType(sheet, Worksheet)
                For colIndex = 0 To maxCols - 1
                    colRange = sheetWorksheet.Columns(startCol + colIndex)
                    CType(colRange, Range).AutoFit()
                Next

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .message = $"表格数据已插入，共{rows.Count}行"
                })

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("InsertTableData 错误: " & ex.Message)
                System.Diagnostics.Debug.WriteLine("堆栈: " & ex.StackTrace)
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            Finally
                ' 【修复】释放COM对象防止泄漏
                If colRange IsNot Nothing Then Utils.ComReleaseHelper.Release(colRange)
                If rangeCell IsNot Nothing Then Utils.ComReleaseHelper.Release(rangeCell)
                If activeCell IsNot Nothing Then Utils.ComReleaseHelper.Release(activeCell)
                If sheet IsNot Nothing Then Utils.ComReleaseHelper.Release(sheet)
                If workbook IsNot Nothing Then Utils.ComReleaseHelper.Release(workbook)
            End Try
        End Function

        ''' <summary>
        ''' 智能插入数据（支持多种数据格式）
        ''' </summary>
        ''' <param name="payloadJson">包含records和mode的JSON</param>
        ''' <returns>操作结果JSON</returns>
        <ComVisible(True)>
        Public Function InsertDataSmart(payloadJson As String) As String
            Dim workbook As Object = Nothing
            Dim sheet As Object = Nothing
            Dim activeCell As Object = Nothing
            Dim rangeCell As Object = Nothing

            Try
                ' 频率限制检查（已恢复）
                If Not _dataWriteLimiter.TryAcquire() Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "数据写入频率过高，请稍后再试"})
                End If

                Dim excelApp As Application = Globals.ThisAddIn.Application
                If excelApp Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "Excel应用不可用"})
                End If

                ' 解析JSON数据
                Dim payload As Dictionary(Of String, Object) = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(payloadJson)

                If Not payload.ContainsKey("records") Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "数据格式不正确，缺少records字段"})
                End If

                Dim records = payload("records")
                Dim mode = If(payload.ContainsKey("mode"), payload("mode").ToString(), "smart")

                ' 获取当前工作表
                workbook = excelApp.ActiveWorkbook
                If workbook Is Nothing Then
                    workbook = excelApp.Workbooks.Add()
                End If

                sheet = workbook.ActiveSheet
                If sheet Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "无法获取活动工作表"})
                End If

                ' 确定起始位置（有选中区域则从选中区域左上角开始，否则从A1开始）
                Dim startRow As Integer = 1
                Dim startCol As Integer = 1
                activeCell = excelApp.ActiveCell
                If activeCell IsNot Nothing Then
                    Dim activeCellRange As Range = CType(activeCell, Range)
                    startRow = activeCellRange.Row
                    startCol = activeCellRange.Column
                End If

                ' 处理不同类型的数据
                If TypeOf records Is Newtonsoft.Json.Linq.JArray Then
                    Dim recordsArray = CType(records, Newtonsoft.Json.Linq.JArray)

                    If recordsArray.Count = 0 Then
                        Return JsonConvert.SerializeObject(New With {.success = False, .error = "数据为空"})
                    End If

                    ' 检查第一条记录类型
                    Dim firstRecord = recordsArray(0)

                    If TypeOf firstRecord Is Newtonsoft.Json.Linq.JArray Then
                        ' 二维数组格式（表格数据）
                        Dim rows As List(Of List(Of String)) = recordsArray.ToObject(Of List(Of List(Of String)))()

                        For rowIndex = 0 To rows.Count - 1
                            Dim row = rows(rowIndex)
                            For colIndex = 0 To row.Count - 1
                                rangeCell = CType(excelApp.Cells(startRow + rowIndex, startCol + colIndex), Range)
                                CType(rangeCell, Range).NumberFormatLocal = "@"
                                CType(rangeCell, Range).Value2 = SanitizeCellValue(row(colIndex))
                            Next
                        Next

                        Return JsonConvert.SerializeObject(New With {
                            .success = True,
                            .message = $"表格数据已插入，共{rows.Count}行"
                        })
                    Else
                        ' 对象数组格式（结构化数据）
                        Dim dataArray As List(Of Dictionary(Of String, Object)) = recordsArray.ToObject(Of List(Of Dictionary(Of String, Object)))()

                        If dataArray.Count > 0 Then
                            ' 获取所有列名
                            Dim columns = dataArray(0).Keys.ToList()

                            ' 写入表头
                            For colIndex = 0 To columns.Count - 1
                                rangeCell = CType(excelApp.Cells(startRow, startCol + colIndex), Range)
                                CType(rangeCell, Range).Value = columns(colIndex)
                            Next

                            ' 写入数据
                            For rowIndex = 0 To dataArray.Count - 1
                                Dim row = dataArray(rowIndex)
                                For colIndex = 0 To columns.Count - 1
                                    Dim colName = columns(colIndex)
                                    Dim value = ""
                                    If row.ContainsKey(colName) AndAlso row(colName) IsNot Nothing Then
                                        value = row(colName).ToString()
                                    End If
                                    rangeCell = CType(excelApp.Cells(startRow + 1 + rowIndex, startCol + colIndex), Range)
                                    CType(rangeCell, Range).NumberFormatLocal = "@"
                                    CType(rangeCell, Range).Value2 = SanitizeCellValue(value)
                                Next
                            Next

                            Return JsonConvert.SerializeObject(New With {
                                .success = True,
                                .message = $"数据已插入，共{dataArray.Count}行"
                            })
                        End If
                    End If
                End If

                Return JsonConvert.SerializeObject(New With {.success = False, .error = "不支持的数据格式"})

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("InsertDataSmart 错误: " & ex.Message)
                System.Diagnostics.Debug.WriteLine("堆栈: " & ex.StackTrace)
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            Finally
                ' 【修复】释放COM对象防止泄漏
                If rangeCell IsNot Nothing Then Utils.ComReleaseHelper.Release(rangeCell)
                If activeCell IsNot Nothing Then Utils.ComReleaseHelper.Release(activeCell)
                If sheet IsNot Nothing Then Utils.ComReleaseHelper.Release(sheet)
                If workbook IsNot Nothing Then Utils.ComReleaseHelper.Release(workbook)
            End Try
        End Function

#End Region

#Region "私有辅助方法"

        ''' <summary>
        ''' 规范化单元格值：基于 Value2（底层原始值），按列格式判断类型
        ''' Value2 不做类型推断，数字/日期均返回 Double（日期为 OADate 序列号）
        ''' 只有列格式为日期时才将 Double 转为日期字符串，其余保持原始类型
        ''' </summary>
        ''' <param name="cellValue">Value2 读取的原始值</param>
        ''' <param name="columnFormat">列的 NumberFormatLocal 字符串</param>
        Private Shared Function NormalizeCellValue(cell As Object, cellValue As Object, columnFormat As String) As Object
            If cellValue Is Nothing OrElse cellValue Is DBNull.Value Then
                Return ""
            End If

            ' 错误值（#DIV/0! / #N/A / #REF! 等）：保留原生错误文本，使通用读表透明携带错误，
            ' 供上层（LLM / Skill / 前端）识别，避免在 JSON 序列化时丢失。
            ' Value2 读取的错误值有时是 COM 内部错误码整数（如 -2146826281），
            ' Information.IsError 可能识别不了，需用 XlCVError 枚举兜底判断。
            If Microsoft.VisualBasic.Information.IsError(cellValue) Then
                Try
                    Return CStr(Microsoft.VisualBasic.Interaction.CallByName(cell, "Text", CallType.Get))
                Catch
                    Return "#ERROR"
                End Try
            End If
            Try
                Dim errCode As Integer = 0
                If Integer.TryParse(cellValue.ToString(), errCode) AndAlso [Enum].IsDefined(GetType(Microsoft.Office.Interop.Excel.XlCVError), errCode) Then
                    Try
                        Return CStr(Microsoft.VisualBasic.Interaction.CallByName(cell, "Text", CallType.Get))
                    Catch
                        Return "#ERROR"
                    End Try
                End If
            Catch
            End Try

            ' 文本格式(@) → 始终返回字符串，保留前导零等原始显示
            If Not String.IsNullOrEmpty(columnFormat) AndAlso
               (columnFormat = "@"c.ToString() OrElse columnFormat.StartsWith("@")) Then
                Return cellValue.ToString()
            End If

            ' Value2 读取的 Double：需要根据列格式区分是日期还是数字
            If TypeOf cellValue Is Double Then
                Dim dblVal = CDbl(cellValue)
                If Not String.IsNullOrEmpty(columnFormat) AndAlso IsDateFormatFormat(columnFormat) Then
                    ' 日期格式列：OADate → DateTime → ISO 格式字符串
                    If dblVal >= 0 AndAlso dblVal <= 2958465.9999 Then
                        Try
                            Dim dt = DateTime.FromOADate(dblVal)
                            If dt.Hour = 0 AndAlso dt.Minute = 0 AndAlso dt.Second = 0 Then
                                Return dt.ToString("yyyy/MM/dd")
                            End If
                            Return dt.ToString("yyyy/MM/dd HH:mm:ss")
                        Catch
                            Return dblVal
                        End Try
                    End If
                End If
                Return dblVal
            End If

            ' Decimal 类型（罕见，通常来自特殊单元格）
            If TypeOf cellValue Is Decimal Then
                Dim decVal = CDbl(cellValue)
                If Not String.IsNullOrEmpty(columnFormat) AndAlso IsDateFormatFormat(columnFormat) Then
                    Try
                        Dim dt = DateTime.FromOADate(decVal)
                        If dt.Hour = 0 AndAlso dt.Minute = 0 AndAlso dt.Second = 0 Then
                            Return dt.ToString("yyyy/MM/dd")
                        End If
                        Return dt.ToString("yyyy/MM/dd HH:mm:ss")
                    Catch
                        Return cellValue
                    End Try
                End If
                Return cellValue
            End If

            ' 字符串 → 直接返回
            If TypeOf cellValue Is String Then
                Return cellValue
            End If

            ' 整数、布尔等其他类型 → 保留原始类型
            Return cellValue
        End Function

        ''' <summary>
        ''' 判断列的 NumberFormatLocal 是否为日期/时间格式
        ''' 检测 y/m/d/h/s 等日期格式标记以及中文年月日字符
        ''' </summary>
        Private Shared Function IsDateFormatFormat(formatStr As String) As Boolean
            If String.IsNullOrEmpty(formatStr) Then Return False

            ' 明确的数字/文本格式，排除
            If formatStr.StartsWith("#") OrElse
               formatStr.StartsWith("0") OrElse
               formatStr.StartsWith("?") OrElse
               formatStr.StartsWith("*") OrElse
               formatStr.StartsWith("@") OrElse
               formatStr.Equals("General", StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If

            ' 检测日期/时间格式标记
            If formatStr.Contains("yyyy") OrElse
               formatStr.Contains("yy") OrElse
               formatStr.Contains("dddd") OrElse
               formatStr.Contains("ddd") OrElse
               formatStr.Contains("dd") OrElse
               formatStr.Contains("hhhh") OrElse
               formatStr.Contains("hh") OrElse
               formatStr.Contains("ss") OrElse
               formatStr.Contains("年") OrElse
               formatStr.Contains("月") OrElse
               formatStr.Contains("日") Then
                Return True
            End If

            ' 检测 m 作为月份/分钟标记（需结合上下文，但简单判断即可）
            ' 排除 mmmm/mmm 已在上方覆盖
            ' 单 m 出现在日期格式中（如 "m/d/yyyy"）
            Return False
        End Function

        ''' <summary>
        ''' 读取指定区域每列的 NumberFormatLocal，缓存为字符串数组供 NormalizeCellValue 使用
        ''' 从数据起始行（startRow）读取，若读取失败则回退到表头行
        ''' </summary>
        Private Shared Function ReadColumnFormats(rng As Range, colCount As Integer, startRow As Integer) As String()
            Dim formats(colCount - 1) As String
            For c As Integer = 1 To colCount
                formats(c - 1) = ""
                Try
                    Dim fmtCell As Range = Nothing
                    Try
                        fmtCell = rng.Cells(startRow, c)
                        formats(c - 1) = CStr(fmtCell.NumberFormatLocal)
                    Finally
                        If fmtCell IsNot Nothing Then Utils.ComReleaseHelper.Release(fmtCell)
                    End Try
                Catch
                    ' 回退：尝试从表头行读取
                    Try
                        Dim headerCell As Range = Nothing
                        Try
                            headerCell = rng.Cells(1, c)
                            formats(c - 1) = CStr(headerCell.NumberFormatLocal)
                        Finally
                            If headerCell IsNot Nothing Then Utils.ComReleaseHelper.Release(headerCell)
                        End Try
                    Catch
                        ' 保持空字符串，后续按类型判断
                    End Try
                End Try
            Next
            Return formats
        End Function

        ''' <summary>
        ''' 添加撤销动作到栈（经接口解析委托给 UndoBridge 统一管理，避免依赖具体单例）
        ''' </summary>
        Private Sub AddUndoAction(action As UndoAction)
            Dim undo = BridgeRegistry.Instance.Resolve(Of IUndoBridge)()
            If undo IsNot Nothing Then undo.AddUndoAction(action)
        End Sub

        ''' <summary>
        ''' 清理单元格值，防止公式注入攻击，并强制身份证号/长数字串作为文本保留
        ''' </summary>
        ''' <param name="value">原始值</param>
        ''' <returns>转义后的安全值</returns>
        Private Function SanitizeCellValue(value As String) As String
            If String.IsNullOrEmpty(value) Then
                Return value
            End If

            Dim trimmedValue = value.Trim()

            ' 身份证号或长数字串应作为文本写入，避免科学计数法或精度丢失
            If ShouldPreserveAsText(trimmedValue) Then
                Return "'" & value
            End If

            ' 检查是否以危险字符开头
            trimmedValue = value.TrimStart()
            If trimmedValue.Length > 0 Then
                Dim firstChar = trimmedValue(0)
                ' Excel公式危险字符: = + - @
                If firstChar = "="c OrElse firstChar = "+"c OrElse firstChar = "-"c OrElse firstChar = "@"c Then
                    ' 在开头添加单引号,使其成为文本而非公式
                    Return "'" & value
                End If
            End If

            Return value
        End Function

        ''' <summary>
        ''' 判断字符串是否应作为文本保留（避免 Excel 科学计数法或精度丢失）
        ''' </summary>
        Private Function ShouldPreserveAsText(value As String) As Boolean
            If String.IsNullOrEmpty(value) Then Return False

            ' 18位身份证号：17位数字 + 1位数字或X/x
            If System.Text.RegularExpressions.Regex.IsMatch(value, "^\d{17}[\dXx]$") Then Return True

            ' 超过15位的纯数字（Excel 精度极限）
            If value.Length > 15 AndAlso System.Text.RegularExpressions.Regex.IsMatch(value, "^\d+$") Then Return True

            ' 有前导零的数字串（如订单号、区号等）
            If value.Length > 1 AndAlso value.StartsWith("0") AndAlso System.Text.RegularExpressions.Regex.IsMatch(value, "^\d+$") Then Return True

            Return False
        End Function

        ''' <summary>
        ''' 判断两个字段名是否匹配（支持同义词）
        ''' </summary>
        Private Function IsFieldMatch(header As String, field As String) As Boolean
            If String.IsNullOrWhiteSpace(header) OrElse String.IsNullOrWhiteSpace(field) Then Return False

            Dim h = header.Trim().ToLower()
            Dim f = field.Trim().ToLower()

            If h = f Then Return True
            If h.Contains(f) OrElse f.Contains(h) Then Return True

            ' 常见同义词组
            Dim synonymGroups As New List(Of List(Of String)) From {
                New List(Of String) From {"名称", "企业名称", "公司名称", "单位名称", "姓名", "名字", "营业执照名称", "证照名称"},
                New List(Of String) From {"统一社会信用代码", "信用代码", "社会信用代码", "税号", "纳税人识别号", "税务登记号"},
                New List(Of String) From {"法定代表人", "法人", "负责人", "代表人"},
                New List(Of String) From {"注册资本", "注册资金", "资本", "出资额"},
                New List(Of String) From {"成立日期", "成立时间", "注册日期", "设立日期", "开业日期"},
                New List(Of String) From {"营业期限", "经营期限", "有效期", "营业期限至"},
                New List(Of String) From {"经营范围", "业务范围", "经营业务", "主营业务"},
                New List(Of String) From {"地址", "住所", "所在地", "经营场所", "注册地址", "企业地址"},
                New List(Of String) From {"电话", "联系电话", "手机号", "手机号码", "联系方式", "传真"},
                New List(Of String) From {"金额", "价税合计", "总额", "合计", "总金额", "金额小写", "金额大写"},
                New List(Of String) From {"发票代码", "代码", "发票代码"},
                New List(Of String) From {"发票号码", "号码", "发票号码", "发票号码"},
                New List(Of String) From {"开票日期", "日期", "开票时间"},
                New List(Of String) From {"车票", "车次", "出发站", "到达站", "出发地", "目的地", "始发站", "终到站"},
                New List(Of String) From {"票价", "金额", "价格", "票价金额"},
                New List(Of String) From {"座位", "座位号", "席位", "席位号", "车厢座位"},
                New List(Of String) From {"席别", "座位类型", "席位类型", "舱位", "铺位"},
                New List(Of String) From {"日期", "乘车日期", "发车日期", "出行日期", "日期时间"},
                New List(Of String) From {"时间", "发车时间", "乘车时间", "时刻", "发车时刻"},
                New List(Of String) From {"类型", "公司类型", "企业类型", "主体类型"},
                New List(Of String) From {"登记机关", "注册机关", "工商局", "市场监管", "登记机关名称"},
                New List(Of String) From {"营业执照", "证照名称", "证件名称"},
                New List(Of String) From {"企业信息公示系统链接", "公示链接", "查询链接"},
                New List(Of String) From {"工商行政管理机构", "管理机构", "登记机构"}
            }

            For Each group In synonymGroups
                Dim hInGroup = group.Any(Function(t) h.Contains(t))
                Dim fInGroup = group.Any(Function(t) f.Contains(t))
                If hInGroup AndAlso fInGroup Then Return True
            Next

            Return False
        End Function

        ''' <summary>
        ''' 智能插入解析后的数据（检测现有表格，匹配表头后追加数据）
        ''' </summary>
        ''' <param name="dataJson">JSON格式的数据数组</param>
        ''' <returns>操作结果JSON</returns>
        <ComVisible(True)>
        Public Function InsertParsedData(dataJson As String) As String
            Dim workbook As Object = Nothing
            Dim sheet As Object = Nothing
            Dim activeCell As Object = Nothing
            Dim columnMapping As Dictionary(Of String, Integer) = Nothing

            Try
                ' 频率限制检查
                If Not _dataWriteLimiter.TryAcquire() Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "数据写入频率过高，请稍后再试"})
                End If

                Dim excelApp As Application = Globals.ThisAddIn.Application
                If excelApp Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "Excel应用不可用"})
                End If

                ' 解析JSON数据
                Dim dataArray As List(Of Dictionary(Of String, Object)) = JsonConvert.DeserializeObject(Of List(Of Dictionary(Of String, Object)))(dataJson)
                If dataArray Is Nothing OrElse dataArray.Count = 0 Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "数据为空"})
                End If

                ' 获取活动工作表
                sheet = excelApp.ActiveSheet
                If sheet Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "没有活动工作表"})
                End If

                ' 获取数据的所有列名
                Dim newColumns = dataArray(0).Keys.ToList()

                ' 尝试查找现有表格（从第1行开始扫描）
                Dim sheetWorksheet As Worksheet = CType(sheet, Worksheet)
                Dim startRow As Integer = 1
                Dim startCol As Integer = 1
                Dim matchedHeaders As List(Of String) = Nothing
                Dim insertRow As Integer = 2 ' 默认从第2行开始插入数据

                ' 扫描前10行查找最佳匹配表头
                Dim bestMatchRow As Integer = -1
                Dim bestMatchRate As Double = 0
                Dim bestMatchCount As Integer = 0
                Dim bestColumnMapping As Dictionary(Of String, Integer) = Nothing
                Dim bestMatchedHeaders As List(Of String) = Nothing

                For scanRow = 1 To 10
                    Dim headers As New List(Of String)()
                    Dim headerCount As Integer = 0

                    ' 读取当前行的所有非空单元格
                    For col = 1 To 50 ' 最多检查50列
                        Dim rangeCell As Object = Nothing
                        Try
                            rangeCell = CType(excelApp.Cells(scanRow, col), Range)
                            Dim cellValue = rangeCell.Value
                            If cellValue IsNot Nothing AndAlso Not String.IsNullOrEmpty(cellValue.ToString()) Then
                                headers.Add(cellValue.ToString().Trim())
                                headerCount += 1
                            Else
                                Exit For ' 遇到空单元格停止
                            End If
                        Finally
                            Utils.ComReleaseHelper.Release(rangeCell)
                        End Try
                    Next

                    ' 如果找到表头（至少有1个列），检查是否匹配
                    If headerCount >= 1 AndAlso headers.Count > 0 Then
                        ' 计算匹配的列数
                        Dim matchCount = 0
                        Dim currentMapping As New Dictionary(Of String, Integer)() ' 新列名 -> 现有列索引

                        For i = 0 To headers.Count - 1
                            Dim header = headers(i)
                            ' 检查是否有匹配的列
                            For Each newCol In newColumns
                                If IsFieldMatch(header, newCol) Then
                                    matchCount += 1
                                    If Not currentMapping.ContainsKey(newCol) Then
                                        currentMapping(newCol) = i + 1 ' 列索引从1开始
                                    End If
                                    Exit For
                                End If
                            Next
                        Next

                        ' 计算匹配率，选择最佳匹配
                        Dim matchRate = matchCount / Math.Max(headers.Count, newColumns.Count)
                        If matchRate > bestMatchRate OrElse
                           (matchRate = bestMatchRate AndAlso matchCount > bestMatchCount) Then
                            bestMatchRate = matchRate
                            bestMatchCount = matchCount
                            bestMatchRow = scanRow
                            bestColumnMapping = currentMapping
                            bestMatchedHeaders = headers
                        End If
                    End If
                Next

                ' 如果最佳匹配率超过阈值（至少匹配1列且匹配率≥20%），认为是同一个表格
                If bestMatchRate >= 0.2 AndAlso bestMatchCount >= 1 Then
                    matchedHeaders = bestMatchedHeaders
                    columnMapping = bestColumnMapping
                    startRow = bestMatchRow
                    startCol = 1
                    insertRow = bestMatchRow + 1 ' 在表头下一行开始插入
                    System.Diagnostics.Debug.WriteLine($"找到匹配表格：第{bestMatchRow}行，匹配率{bestMatchRate:P0}")

                    ' 找到已有数据的最后一行，在其后追加，避免覆盖
                    Dim dataRow As Integer = insertRow
                    Dim lastNonEmptyRow As Integer = bestMatchRow
                    Dim maxColToCheck As Integer = Math.Max(matchedHeaders.Count, newColumns.Count)
                    While dataRow <= bestMatchRow + 1000 ' 最多扫描1000行
                        Dim hasData As Boolean = False
                        For checkCol = 1 To maxColToCheck
                            Dim checkCell As Object = Nothing
                            Try
                                checkCell = CType(excelApp.Cells(dataRow, checkCol), Range)
                                If checkCell.Value IsNot Nothing AndAlso Not String.IsNullOrEmpty(checkCell.Value.ToString()) Then
                                    hasData = True
                                    Exit For
                                End If
                            Finally
                                Utils.ComReleaseHelper.Release(checkCell)
                            End Try
                        Next
                        If Not hasData Then Exit While
                        lastNonEmptyRow = dataRow
                        dataRow += 1
                    End While
                    insertRow = lastNonEmptyRow + 1
                    System.Diagnostics.Debug.WriteLine($"从第{insertRow}行追加数据")
                End If

                ' 如果没有找到匹配的表格，从活动单元格开始
                If matchedHeaders Is Nothing Then
                    activeCell = excelApp.ActiveCell
                    If activeCell IsNot Nothing Then
                        Dim activeCellRange As Range = CType(activeCell, Range)
                        startRow = activeCellRange.Row
                        startCol = activeCellRange.Column
                        Utils.ComReleaseHelper.Release(activeCellRange)
                    End If

                    ' 写入新表头并设置样式
                    For colIndex = 0 To newColumns.Count - 1
                        Dim rangeCell As Object = Nothing
                        Try
                            rangeCell = CType(excelApp.Cells(startRow, startCol + colIndex), Range)
                            Dim headerCell As Range = CType(rangeCell, Range)
                            headerCell.Value = newColumns(colIndex)
                            headerCell.Font.Bold = True
                            headerCell.Interior.Color = RGB(70, 130, 180) ' 钢蓝色
                            headerCell.Font.Color = RGB(255, 255, 255) ' 白色字体
                            headerCell.HorizontalAlignment = XlHAlign.xlHAlignCenter
                            headerCell.VerticalAlignment = XlVAlign.xlVAlignCenter
                        Finally
                            Utils.ComReleaseHelper.Release(rangeCell)
                        End Try
                    Next
                    insertRow = startRow + 1
                    System.Diagnostics.Debug.WriteLine("未找到匹配表格，创建新表头")
                End If

                ' 插入数据行
                Dim insertedCount = 0
                For Each rowData In dataArray
                    For colIndex = 0 To newColumns.Count - 1
                        Dim colName = newColumns(colIndex)
                        Dim value = ""

                        If rowData.ContainsKey(colName) AndAlso rowData(colName) IsNot Nothing Then
                            value = rowData(colName).ToString()
                        End If

                        ' 确定插入的列：如果有匹配表头，使用映射的列索引；否则按顺序插入
                        Dim targetCol As Integer
                        If matchedHeaders IsNot Nothing AndAlso columnMapping.ContainsKey(colName) Then
                            targetCol = columnMapping(colName)
                        Else
                            targetCol = startCol + colIndex
                        End If

                        Dim rangeCell As Object = Nothing
                        Try
                            rangeCell = CType(excelApp.Cells(insertRow, targetCol), Range)
                            CType(rangeCell, Range).NumberFormatLocal = "@"
                            CType(rangeCell, Range).Value2 = SanitizeCellValue(value)
                        Finally
                            Utils.ComReleaseHelper.Release(rangeCell)
                        End Try
                    Next
                    insertedCount += 1
                    insertRow += 1
                Next

                ' 自动调整列宽
                For colIndex = 0 To newColumns.Count - 1
                    Dim targetCol = If(matchedHeaders IsNot Nothing AndAlso columnMapping.ContainsKey(newColumns(colIndex)), columnMapping(newColumns(colIndex)), startCol + colIndex)
                    Dim colRange As Object = Nothing
                    Try
                        colRange = sheetWorksheet.Columns(targetCol)
                        CType(colRange, Range).AutoFit()
                    Finally
                        Utils.ComReleaseHelper.Release(colRange)
                    End Try
                Next

                ' 如果创建了新表格，应用整体边框样式
                If matchedHeaders Is Nothing Then
                    Try
                        Dim totalRows As Integer = dataArray.Count + 1 ' 表头 + 数据行
                        Dim totalCols As Integer = newColumns.Count
                        Dim startCell As Object = Nothing
                        Dim endCell As Object = Nothing
                        Dim tableRange As Object = Nothing
                        Try
                            startCell = sheetWorksheet.Cells(startRow, startCol)
                            endCell = sheetWorksheet.Cells(startRow + totalRows - 1, startCol + totalCols - 1)
                            tableRange = sheetWorksheet.Range(startCell, endCell)
                            CType(tableRange, Range).Borders.LineStyle = XlLineStyle.xlContinuous
                            CType(tableRange, Range).Borders.Weight = XlBorderWeight.xlThin
                            CType(tableRange, Range).Borders.Color = RGB(200, 200, 200)
                        Finally
                            Utils.ComReleaseHelper.Release(tableRange)
                            Utils.ComReleaseHelper.Release(endCell)
                            Utils.ComReleaseHelper.Release(startCell)
                        End Try
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine("应用表格样式失败: " & ex.Message)
                    End Try
                End If

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .message = If(matchedHeaders IsNot Nothing, $"已插入{insertedCount}行数据到现有表格", $"已创建新表格并插入{insertedCount}行数据"),
                    .insertedRows = insertedCount,
                    .foundExistingTable = matchedHeaders IsNot Nothing
                })

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("InsertParsedData 错误: " & ex.Message)
                System.Diagnostics.Debug.WriteLine("堆栈: " & ex.StackTrace)
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            Finally
                ' 释放COM对象
                Utils.ComReleaseHelper.Release(activeCell)
                Utils.ComReleaseHelper.Release(sheet)
                Utils.ComReleaseHelper.Release(workbook)
            End Try
        End Function

#End Region

#Region "IBridge 生命周期"

        Public ReadOnly Property BridgeName As String Implements IBridge.BridgeName
            Get
                Return "ExcelDataBridge"
            End Get
        End Property

        Private _disposed As Boolean = False
        Public Sub Dispose() Implements IBridge.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not _disposed Then
                ' ExcelDataBridge 无外部资源需要释放
                _disposed = True
            End If
        End Sub

#End Region

    End Class

End Namespace