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

        ' ========== 【重写】按工作表拆分工作簿（统一版） ==========
        Private Function ExecuteSplitWorkbook(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Dim createdFiles As New List(Of String)
            Dim failedSheets As New List(Of String)

            Dim sourceWb As Microsoft.Office.Interop.Excel.Workbook = Nothing
            Dim tempWb As Microsoft.Office.Interop.Excel.Workbook = Nothing
            Dim app As Microsoft.Office.Interop.Excel.Application = Globals.ThisAddIn.Application
            Dim displayAlertsOriginal As Boolean = app.DisplayAlerts

            Try
                ' ========== 参数解析 ==========
                Dim outputDir As String = GetArg(Of String)(args, "outputDir", "")
                Dim namingPattern As String = GetArg(Of String)(args, "namingPattern", "{original}_{sheet}")
                Dim fileFormat As String = GetArg(Of String)(args, "fileFormat", "xlsx").ToLower()
                Dim includeHidden As Boolean = GetArg(Of Boolean)(args, "includeHidden", False)

                ' 校验格式
                Dim formatMap As New Dictionary(Of String, Microsoft.Office.Interop.Excel.XlFileFormat) From {
                    {"xlsx", Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook},
                    {"xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel8},
                    {"csv", Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV}
                }
                If Not formatMap.ContainsKey(fileFormat) Then
                    Throw New ArgumentException($"不支持的格式: {fileFormat}，支持 xlsx/xls/csv")
                End If

                ' ========== 确定源工作簿 ==========
                sourceWb = app.ActiveWorkbook
                If sourceWb Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿"})
                    Return result
                End If

                Dim sourcePath As String = sourceWb.FullName
                Dim originalName As String = IO.Path.GetFileNameWithoutExtension(sourcePath)

                ' ========== 确定输出目录 ==========
                If String.IsNullOrWhiteSpace(outputDir) Then
                    outputDir = IO.Path.GetDirectoryName(sourcePath)
                End If
                ' 【防幻觉路径】净化模型编造的输出目录，避免 CreateDirectory 造出垃圾目录
                outputDir = Utils.PathHelper.SanitizeUserDirectory(outputDir)
                If Not IO.Directory.Exists(outputDir) Then
                    IO.Directory.CreateDirectory(outputDir)
                End If

                ' ========== 遍历拆分 ==========
                Dim sheetIndex As Integer = 1
                Dim sheetCount As Integer = sourceWb.Worksheets.Count

                For i As Integer = 1 To sheetCount
                    Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                    Try
                        ws = DirectCast(sourceWb.Worksheets(i), Microsoft.Office.Interop.Excel.Worksheet)

                        ' 跳过隐藏表
                        If Not includeHidden AndAlso ws.Visible <> Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible Then
                            Continue For
                        End If

                        ' 生成文件名
                        Dim safeSheetName = SanitizeFileName(ws.Name)
                        Dim fileName As String
                        If Not namingPattern.Contains("{"c) Then
                            fileName = $"{namingPattern}{safeSheetName}"
                        Else
                            fileName = namingPattern _
                                .Replace("{original}", originalName) _
                                .Replace("{sheet}", safeSheetName) _
                                .Replace("{index}", sheetIndex.ToString("D2"))
                        End If

                        ' 确保唯一路径
                        Dim fullPath = GetUniqueFilePath(outputDir, fileName, fileFormat)

                        ' 创建新工作簿并复制
                        tempWb = app.Workbooks.Add()
                        ws.Copy(Before:=tempWb.Sheets(1))

                        ' 删除新建时自带的空白 Sheet
                        If tempWb.Sheets.Count > 1 Then
                            app.DisplayAlerts = False
                            Dim defaultSheet As Microsoft.Office.Interop.Excel.Worksheet =
                                DirectCast(tempWb.Sheets(2), Microsoft.Office.Interop.Excel.Worksheet)
                            defaultSheet.Delete()
                            app.DisplayAlerts = displayAlertsOriginal
                            Utils.ComReleaseHelper.Release(defaultSheet)
                        End If

                        ' 保存
                        tempWb.SaveAs(Filename:=fullPath, FileFormat:=formatMap(fileFormat))
                        tempWb.Close(SaveChanges:=False)
                        createdFiles.Add(fullPath)

                        Utils.ComReleaseHelper.Release(tempWb)
                        tempWb = Nothing
                        sheetIndex += 1

                    Catch ex As Exception
                        failedSheets.Add($"{ws.Name}: {ex.Message}")
                        If tempWb IsNot Nothing Then
                            Try : tempWb.Close(SaveChanges:=False) : Catch : End Try
                            Utils.ComReleaseHelper.Release(tempWb)
                            tempWb = Nothing
                        End If
                    Finally
                        Utils.ComReleaseHelper.Release(ws)
                    End Try
                Next

                ' ========== 组装结果 ==========
                Dim success As Boolean = (failedSheets.Count = 0)
                Dim summary As String = $"成功拆分 {createdFiles.Count} 个工作表到: {outputDir}"
                If failedSheets.Count > 0 Then
                    summary &= $"，{failedSheets.Count} 个失败"
                End If

                ' CSV 警告
                If fileFormat = "csv" AndAlso createdFiles.Count > 1 Then
                    summary &= " [注意：CSV 格式仅保留单 Sheet，每个文件已分别保存]"
                End If

                Dim resultObj As New Dictionary(Of String, Object) From {
                    {"success", success},
                    {"files_created", createdFiles},
                    {"summary", summary},
                    {"failed", failedSheets},
                    {"error", If(failedSheets.Count > 0, String.Join("; ", failedSheets), "")}
                }

                result.Content.Add(New McpTextContent With {
                    .Text = JsonConvert.SerializeObject(resultObj, Formatting.Indented)
                })

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"拆分工作簿失败: {ex.Message}"})

            Finally
                app.DisplayAlerts = displayAlertsOriginal
                If tempWb IsNot Nothing Then
                    Try : tempWb.Close(SaveChanges:=False) : Catch : End Try
                    Utils.ComReleaseHelper.Release(tempWb)
                End If
            End Try

            Return result
        End Function
        ''' <summary>
        ''' 按指定列的唯一值拆分数据
        ''' </summary>
        Private Function ExecuteSplitByColumn(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Dim app = Globals.ThisAddIn.Application
            Dim displayAlertsOriginal = app.DisplayAlerts
            Dim pr As New StepRecorder()

            Try
                Dim column As String = GetArg(Of String)(args, "column", "")
                Dim outputMode As String = GetArg(Of String)(args, "outputMode", "sheets")
                Dim outputDir As String = GetArg(Of String)(args, "outputDir", "")
                Dim namingPattern As String = GetArg(Of String)(args, "namingPattern", "{original}_{value}")
                ' onConflict: 目标表名已存在时的策略
                '   skip（默认）→ 跳过该组，不建 _2/_3 副本（避免"拆两次"叠表）
                '   suffix     → 旧行为：递增后缀 _2/_3…
                '   replace    → 清空并覆写已存在表
                Dim onConflict As String = GetArg(Of String)(args, "onConflict", "skip")
                If String.IsNullOrWhiteSpace(onConflict) Then onConflict = "skip"
                onConflict = onConflict.Trim().ToLowerInvariant()
                If onConflict <> "suffix" AndAlso onConflict <> "replace" Then onConflict = "skip"
                ' sourceSheets: JSON 数组字符串，指定要合并后再拆分的多表（用户在标签栏多选的场景）
                ' 不传/空 → 只拆分活动工作表（旧行为）；传了 → 多表纵向拼接后按 column 拆分
                Dim sourceSheetsJson As String = GetArg(Of String)(args, "sourceSheets", "")
                Dim sourceNames As New List(Of String)()
                If Not String.IsNullOrWhiteSpace(sourceSheetsJson) Then
                    Try
                        sourceNames = Newtonsoft.Json.JsonConvert.DeserializeObject(Of List(Of String))(sourceSheetsJson)
                    Catch
                        sourceNames = New List(Of String)()
                    End Try
                End If

                If String.IsNullOrEmpty(column) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "拆分列不能为空"})
                    Return result
                End If

                Dim wb = app.ActiveWorkbook
                If wb Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何工作簿"})
                    Return result
                End If

                ' ===== 确定数据源：单表 vs 多表纵向拼接 =====
                Dim rawArr(,) As Object = Nothing
                Dim fmtsArr As String() = Nothing
                Dim sourceSheetName As String = ""
                Dim startCol As Integer = 0  ' 拼接数组列从 0 基开始；单表模式会覆盖为实际 startCol
                Dim startRow As Integer = 0  ' 多表模式不使用；单表模式会覆盖
                If sourceNames.Count > 0 Then
                    ' 多表模式：纵向拼接所有 sourceSheets 的数据
                    Dim sheetList As New List(Of Object(,))()
                    Dim headerArr As Object() = Nothing
                    Dim maxCols As Integer = 0
                    Dim totalSourceRows As Integer = 0
                    For Each sname In sourceNames
                        Dim srcWs As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                        Try
                            srcWs = DirectCast(wb.Worksheets(sname), Microsoft.Office.Interop.Excel.Worksheet)
                        Catch
                            Continue For
                        End Try
                        If srcWs Is Nothing Then Continue For
                        Dim srcUsed = srcWs.UsedRange
                        Dim srcRange = DirectCast(srcWs.Range(srcWs.Cells(srcUsed.Row, srcUsed.Column),
                                                               srcWs.Cells(srcUsed.Row + srcUsed.Rows.Count - 1, srcUsed.Column + srcUsed.Columns.Count - 1)),
                                                               Microsoft.Office.Interop.Excel.Range)
                        Dim localFmts As String() = Nothing
                        Dim localArr = ReadRangeTo0BasedWithFormats(srcRange, localFmts)
                        Utils.ComReleaseHelper.Release(srcRange)
                        Utils.ComReleaseHelper.Release(srcUsed)
                        localArr = TrimPhantomRanges(localArr)
                        If localArr Is Nothing OrElse ArrayIsAllBlank(localArr) Then
                            Utils.ComReleaseHelper.Release(srcWs)
                            Continue For
                        End If
                        If headerArr Is Nothing Then
                            ' 第一张表的表头作为合并基准
                            Dim localCols = localArr.GetLength(1)
                            Dim hdrArr(localCols - 1) As Object
                            For c = 0 To localCols - 1
                                hdrArr(c) = localArr(0, c)
                            Next c
                            headerArr = hdrArr
                            maxCols = localCols
                        End If
                        ' 跳过表头行，追加数据
                        Dim dataRows = localArr.GetLength(0) - 1
                        Dim dataCols = localArr.GetLength(1)
                        Dim dataOnly(dataRows - 1, dataCols - 1) As Object
                        For r = 1 To localArr.GetLength(0) - 1
                            For c = 0 To dataCols - 1
                                dataOnly(r - 1, c) = localArr(r, c)
                            Next c
                        Next r
                        sheetList.Add(dataOnly)
                        totalSourceRows += localArr.GetLength(0) - 1
                        pr.ReadSource(sname, localArr.GetLength(0) - 1)
                        Utils.ComReleaseHelper.Release(srcWs)
                    Next
                    If sheetList.Count = 0 Then
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = "指定的源工作表都没有数据或不存在"})
                        Return result
                    End If
                    ' 合并：表头 + 所有数据块
                    rawArr = New Object(totalSourceRows, maxCols - 1) {}
                    For c = 0 To maxCols - 1
                        rawArr(0, c) = If(headerArr IsNot Nothing AndAlso c < headerArr.Length, headerArr(c), Nothing)
                    Next
                    Dim writeRow = 1
                    For Each block In sheetList
                        For r = 0 To block.GetLength(0) - 1
                            For c = 0 To block.GetLength(1) - 1
                                rawArr(writeRow, c) = block(r, c)
                            Next
                            writeRow += 1
                        Next
                    Next
                    sourceSheetName = String.Join("+", sourceNames)
                Else
                    ' 单表模式（默认）：未指定表时优先"主数据表"——可见表中已用行数最多者。
                    ' 【修复 2026-10-10】旧逻辑盲目用活动表：活动表若是残损/小表（实测：只剩 1 行的
                    ' 大客户部被当拆分源 → 只拆出 1 组、建了 1 张"大客户部_2"），用户"按部门拆分成
                    ' 不同Sheet"的意图显然指向主数据表。仅当其它可见表行数 ≥ 活动表 5 倍时才覆盖，
                    ' 避免误抢用户明确想拆的小表；产物表（合并结果/聚合结果）不作为拆分源。
                    Dim ws As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(app.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    Dim activeRows As Integer = 0
                    Try
                        Dim aUr = ws.UsedRange
                        activeRows = aUr.Row + aUr.Rows.Count - 1
                        Utils.ComReleaseHelper.Release(aUr)
                    Catch
                    End Try
                    Dim bestWs As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                    Dim bestRows As Integer = 0
                    For Each s As Microsoft.Office.Interop.Excel.Worksheet In wb.Worksheets
                        Dim sName As String = String.Empty
                        Try
                            sName = s.Name
                            If s.Visible <> Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible Then
                                Utils.ComReleaseHelper.Release(s)
                                Continue For
                            End If
                        Catch
                            Utils.ComReleaseHelper.Release(s)
                            Continue For
                        End Try
                        ' 排除本产品的产物表，避免把"合并结果/聚合结果"当拆分源
                        If sName = "合并结果" OrElse sName = "聚合结果" Then
                            Utils.ComReleaseHelper.Release(s)
                            Continue For
                        End If
                        Dim rCount As Integer = 0
                        Try
                            Dim ur = s.UsedRange
                            rCount = ur.Row + ur.Rows.Count - 1
                            Utils.ComReleaseHelper.Release(ur)
                        Catch
                        End Try
                        If rCount > bestRows Then
                            If bestWs IsNot Nothing Then Utils.ComReleaseHelper.Release(bestWs)
                            bestWs = s
                            bestRows = rCount
                        Else
                            Utils.ComReleaseHelper.Release(s)
                        End If
                    Next
                    If bestWs IsNot Nothing Then
                        If bestRows > activeRows AndAlso activeRows < bestRows \ 5 Then
                            Utils.ComReleaseHelper.Release(ws)
                            ws = bestWs
                        Else
                            Utils.ComReleaseHelper.Release(bestWs)
                        End If
                    End If
                    Dim usedRange As Microsoft.Office.Interop.Excel.Range = ws.UsedRange
                    startRow = usedRange.Row      ' ← 1 基：UsedRange 起始行（Excel COM 索引从 1 开始，0 会抛 0x800A03EC）
                    startCol = usedRange.Column   ' ← 1 基：UsedRange 起始列

                    ' 一次性读入整个已用区域并裁剪幽灵空行/空列，避免幻影区域被当成真实数据（否则会生成大量空白拆分表）
                    Dim rawRange = DirectCast(ws.Range(ws.Cells(startRow, startCol),
                                                       ws.Cells(startRow + usedRange.Rows.Count - 1, startCol + usedRange.Columns.Count - 1)),
                                                       Microsoft.Office.Interop.Excel.Range)
                    rawArr = ReadRangeTo0BasedWithFormats(rawRange, fmtsArr)
                    Utils.ComReleaseHelper.Release(rawRange)
                    Utils.ComReleaseHelper.Release(usedRange)
                    rawArr = TrimPhantomRanges(rawArr)
                    If rawArr Is Nothing OrElse ArrayIsAllBlank(rawArr) Then
                        Utils.ComReleaseHelper.Release(ws)
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = "当前工作表没有数据"})
                        Return result
                    End If
                    sourceSheetName = ws.Name
                    Utils.ComReleaseHelper.Release(ws)
                End If

                Dim totalRows = rawArr.GetLength(0)
                Dim totalCols = rawArr.GetLength(1)

                ' 确定拆分列索引（0 基，相对读入区域）
                Dim splitColIndex As Integer = -1
                If column.Length = 1 AndAlso Char.IsLetter(column(0)) Then
                    splitColIndex = ExcelMcpCoreUtils.ColumnLetterToNumber(column.ToUpper()) - startCol
                Else
                    ' 按表头名称查找
                    For c = 0 To totalCols - 1
                        Dim hv = rawArr(0, c)
                        If hv IsNot Nothing AndAlso hv.ToString() = column Then
                            splitColIndex = c
                            Exit For
                        End If
                    Next
                End If

                If splitColIndex < 0 OrElse splitColIndex >= totalCols Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"找不到列: {column}"})
                    Return result
                End If

                ' 读取表头
                Dim headers(totalCols - 1) As Object
                For c = 0 To totalCols - 1
                    headers(c) = rawArr(0, c)
                Next

                ' 按唯一值分组数据（跳过整行全空的幽灵行）
                Dim groups As New Dictionary(Of String, List(Of Object()))()
                For r = 1 To totalRows - 1
                    Dim blank As Boolean = True
                    For c = 0 To totalCols - 1
                        If rawArr(r, c) IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(rawArr(r, c).ToString()) Then
                            blank = False
                            Exit For
                        End If
                    Next
                    If blank Then Continue For
                    Dim cellVal = rawArr(r, splitColIndex)
                    Dim key = If(cellVal IsNot Nothing, cellVal.ToString(), "空值")
                    If Not groups.ContainsKey(key) Then groups(key) = New List(Of Object())()
                    Dim rowData = New Object(totalCols - 1) {}
                    For c = 0 To totalCols - 1 : rowData(c) = NormalizeCell(rawArr(r, c)) : Next
                    groups(key).Add(rowData)
                Next

                Dim createdSheets As New List(Of String)()
                Dim skippedSheets As New List(Of String)()
                Dim replacedSheets As New List(Of String)()
                Dim originalName = IO.Path.GetFileNameWithoutExtension(wb.FullName)
                pr.SplitIntoGroups(column, groups.Count, groups.Keys)

                If outputMode = "sheets" Then
                    ' 拆分到同一工作簿的新工作表
                    For Each kvp In groups
                        Dim sheetName = SanitizeFileName(kvp.Key)
                        If sheetName.Length > 31 Then sheetName = sheetName.Substring(0, 31)

                        Dim existingWs As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                        Try
                            existingWs = DirectCast(wb.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                        Catch
                            existingWs = Nothing
                        End Try

                        If existingWs IsNot Nothing Then
                            If onConflict = "skip" Then
                                ' 默认幂等：同名表已存在 → 视为上次拆分产物，不建 _2 副本
                                skippedSheets.Add(sheetName)
                                pr.Add("工作表 '" & sheetName & "' 已存在，跳过重复拆分（" & kvp.Value.Count.ToString() & " 行）")
                                Utils.ComReleaseHelper.Release(existingWs)
                                Continue For
                            ElseIf onConflict = "replace" Then
                                ' 覆写：清空后复用该表写入
                                Try
                                    existingWs.Cells.Clear()
                                Catch
                                End Try
                                WriteSplitGroupToSheet(existingWs, headers, kvp.Value, totalCols, fmtsArr)
                                replacedSheets.Add(sheetName)
                                createdSheets.Add(sheetName)
                                pr.Add("覆写工作表 '" & sheetName & "'，写入表头 + " & kvp.Value.Count.ToString() & " 行数据")
                                Utils.ComReleaseHelper.Release(existingWs)
                                Continue For
                            Else
                                ' suffix：递增 _2/_3…（旧行为，仅显式要求时使用）
                                Utils.ComReleaseHelper.Release(existingWs)
                                Dim uniqueName = sheetName
                                Dim sfx = 2
                                Do
                                    Dim exists As Boolean = False
                                    For Each s As Microsoft.Office.Interop.Excel.Worksheet In wb.Worksheets
                                        If String.Compare(s.Name, uniqueName, True) = 0 Then
                                            exists = True
                                            Utils.ComReleaseHelper.Release(s)
                                            Exit For
                                        End If
                                        Utils.ComReleaseHelper.Release(s)
                                    Next
                                    If Not exists Then Exit Do
                                    Dim tag = "_" & sfx.ToString()
                                    uniqueName = sheetName.Substring(0, Math.Min(sheetName.Length, 31 - tag.Length)) & tag
                                    sfx += 1
                                Loop While sfx <= 50
                                sheetName = uniqueName
                            End If
                        End If

                        Dim newWs As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(wb.Worksheets.Add(), Microsoft.Office.Interop.Excel.Worksheet)
                        Try
                            newWs.Name = sheetName
                        Catch
                            ' 表名冲突/非法：删除刚新增的空白表，避免残留空白 Sheet
                            Try : newWs.Delete() : Catch : End Try
                            Utils.ComReleaseHelper.Release(newWs)
                            Continue For
                        End Try
                        pr.NewSheet(sheetName, kvp.Value.Count, True)
                        WriteSplitGroupToSheet(newWs, headers, kvp.Value, totalCols, fmtsArr)
                        createdSheets.Add(sheetName)
                        Utils.ComReleaseHelper.Release(newWs)
                    Next
                Else
                    ' 拆分到独立文件
                    If String.IsNullOrWhiteSpace(outputDir) Then
                        outputDir = IO.Path.GetDirectoryName(wb.FullName)
                    End If
                    ' 【防幻觉路径】净化模型编造的输出目录，避免 CreateDirectory 造出垃圾目录
                    outputDir = Utils.PathHelper.SanitizeUserDirectory(outputDir)
                    If Not IO.Directory.Exists(outputDir) Then
                        IO.Directory.CreateDirectory(outputDir)
                    End If

                    For Each kvp In groups
                        Dim fileName = namingPattern _
                            .Replace("{original}", originalName) _
                            .Replace("{value}", SanitizeFileName(kvp.Key))
                        Dim fullPath = GetUniqueFilePath(outputDir, fileName, "xlsx")

                        Dim newWb As Microsoft.Office.Interop.Excel.Workbook = DirectCast(app.Workbooks.Add(), Microsoft.Office.Interop.Excel.Workbook)
                        ' 新建工作簿自带多个默认空白 Sheet，删除多余空白表，避免每个输出文件都带一堆空白 Sheet
                        While newWb.Worksheets.Count > 1
                            Dim extraWs = DirectCast(newWb.Worksheets(newWb.Worksheets.Count), Microsoft.Office.Interop.Excel.Worksheet)
                            extraWs.Delete()
                            Utils.ComReleaseHelper.Release(extraWs)
                        End While
                        Dim newWs As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(newWb.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)

                        If kvp.Value.Count > 0 Then
                            Dim outRows = kvp.Value.Count + 1
                            Dim block = New Object(outRows - 1, Math.Max(totalCols, 1) - 1) {}
                            For c = 0 To Math.Max(totalCols, 1) - 1
                                If c < totalCols Then block(0, c) = headers(c) Else block(0, c) = Nothing
                            Next
                            Dim dr = 1
                            For Each srcRow In kvp.Value
                                For c = 0 To Math.Max(totalCols, 1) - 1
                                    If srcRow IsNot Nothing AndAlso c < srcRow.Length Then block(dr, c) = srcRow(c) Else block(dr, c) = Nothing
                                Next
                                dr += 1
                            Next
                            Dim rng = DirectCast(newWs.Range(newWs.Cells(1, 1), newWs.Cells(outRows, Math.Max(totalCols, 1))), Microsoft.Office.Interop.Excel.Range)
                            rng.Value2 = block
                            Utils.ComReleaseHelper.Release(rng)
                            ' 还原每列数字格式（日期/百分比/货币等），避免 Value2 写回后日期变成序列号
                            ApplyColumnFormats(newWs, fmtsArr, 2, outRows - 1, Math.Max(totalCols, 1))
                        End If

                        newWb.SaveAs(Filename:=fullPath, FileFormat:=Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook)
                        newWb.Close(SaveChanges:=False)
                        createdSheets.Add(fullPath)
                        pr.GenerateFile(IO.Path.GetFileName(fullPath), kvp.Value.Count, True)
                        Utils.ComReleaseHelper.Release(newWs)
                        Utils.ComReleaseHelper.Release(newWb)
                    Next
                End If

                ' 单表模式下活动工作表已在 Else 分支里释放；多表模式各 srcWs 也在循环里释放了，这里统一兜底
                ' （单表模式的 ws 变量作用域在 Else 内，外面访问不到，跳过即可）

                Dim summaryParts As New List(Of String)()
                If createdSheets.Count > 0 AndAlso replacedSheets.Count = 0 Then
                    summaryParts.Add("新建 " & createdSheets.Count.ToString() & " 个工作表：" & String.Join("、", createdSheets))
                ElseIf replacedSheets.Count > 0 Then
                    summaryParts.Add("覆写 " & replacedSheets.Count.ToString() & " 个工作表：" & String.Join("、", replacedSheets))
                End If
                If skippedSheets.Count > 0 Then
                    If createdSheets.Count = 0 AndAlso replacedSheets.Count = 0 Then
                        summaryParts.Add("拆分表已存在（" & String.Join("、", skippedSheets) & "），未重复创建。如需覆盖请说「重新拆分」或「覆盖拆分」")
                    Else
                        summaryParts.Add("已跳过同名表 " & skippedSheets.Count.ToString() & " 个：" & String.Join("、", skippedSheets))
                    End If
                End If
                If summaryParts.Count = 0 Then
                    summaryParts.Add("按列拆分完成（无新建工作表）")
                End If

                Dim resultObj = New Dictionary(Of String, Object) From {
                    {"success", True},
                    {"groups", groups.Count},
                    {"created", createdSheets},
                    {"skipped", skippedSheets},
                    {"replaced", replacedSheets},
                    {"summary", String.Join("；", summaryParts)},
                    {"steps", pr.Steps}
                }
                result.Content.Add(New McpTextContent With {
                    .Text = JsonConvert.SerializeObject(resultObj, Formatting.Indented)
                })

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"按列拆分失败: {ex.Message}"})
            Finally
                app.DisplayAlerts = displayAlertsOriginal
            End Try

            Return result
        End Function

        ''' <summary>把一组拆分数据（表头+行）写入目标工作表。</summary>
        Private Sub WriteSplitGroupToSheet(ws As Microsoft.Office.Interop.Excel.Worksheet,
                                           headers As Object(),
                                           rows As List(Of Object()),
                                           totalCols As Integer,
                                           fmtsArr As String())
            If ws Is Nothing Then Return
            Dim outRows = If(rows Is Nothing, 0, rows.Count) + 1
            If rows IsNot Nothing AndAlso rows.Count > 0 Then
                Dim block = New Object(outRows - 1, Math.Max(totalCols, 1) - 1) {}
                For c = 0 To Math.Max(totalCols, 1) - 1
                    If c < totalCols AndAlso headers IsNot Nothing AndAlso c < headers.Length Then
                        block(0, c) = headers(c)
                    Else
                        block(0, c) = Nothing
                    End If
                Next
                Dim dr = 1
                For Each srcRow In rows
                    For c = 0 To Math.Max(totalCols, 1) - 1
                        If srcRow IsNot Nothing AndAlso c < srcRow.Length Then
                            block(dr, c) = srcRow(c)
                        Else
                            block(dr, c) = Nothing
                        End If
                    Next
                    dr += 1
                Next
                Dim rng = DirectCast(ws.Range(ws.Cells(1, 1), ws.Cells(outRows, Math.Max(totalCols, 1))), Microsoft.Office.Interop.Excel.Range)
                rng.Value2 = block
                Utils.ComReleaseHelper.Release(rng)
                ApplyColumnFormats(ws, fmtsArr, 2, outRows - 1, Math.Max(totalCols, 1))
            ElseIf headers IsNot Nothing AndAlso headers.Length > 0 Then
                Dim hdrBlock = New Object(0, Math.Max(totalCols, 1) - 1) {}
                For c = 0 To Math.Max(totalCols, 1) - 1
                    If c < totalCols AndAlso c < headers.Length Then hdrBlock(0, c) = headers(c) Else hdrBlock(0, c) = Nothing
                Next
                Dim hr = DirectCast(ws.Range(ws.Cells(1, 1), ws.Cells(1, Math.Max(totalCols, 1))), Microsoft.Office.Interop.Excel.Range)
                hr.Value2 = hdrBlock
                Utils.ComReleaseHelper.Release(hr)
            End If
        End Sub

        ''' <summary>
        ''' 合并多个工作簿
        ''' </summary>
        Private Function ExecuteMergeWorkbooks(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Dim app = Globals.ThisAddIn.Application
            Dim displayAlertsOriginal = app.DisplayAlerts

            Try
                Dim filePathsJson As String = GetArg(Of String)(args, "filePaths", "[]")
                Dim targetPath As String = GetArg(Of String)(args, "targetPath", "")
                Dim sheetPrefix As String = GetArg(Of String)(args, "sheetPrefix", "")

                Dim filePaths = JsonConvert.DeserializeObject(Of List(Of String))(filePathsJson)
                If filePaths Is Nothing OrElse filePaths.Count = 0 Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "文件列表不能为空"})
                    Return result
                End If

                ' 创建新工作簿
                Dim targetWb As Microsoft.Office.Interop.Excel.Workbook = DirectCast(app.Workbooks.Add(), Microsoft.Office.Interop.Excel.Workbook)
                Dim mergedSheets As New List(Of String)()

                ' 删除默认空白表
                app.DisplayAlerts = False
                If targetWb.Sheets.Count > 0 Then
                    Dim defaultSheet = targetWb.Sheets(1)
                    ' 先复制第一个文件的工作表，再删除默认表
                End If

                Dim isFirstFile = True
                For Each filePath In filePaths
                    If Not IO.File.Exists(filePath) Then
                        Continue For
                    End If

                    Dim sourceWb = app.Workbooks.Open(filePath)
                    Dim fileName = IO.Path.GetFileNameWithoutExtension(filePath)
                    Dim prefix = If(String.IsNullOrEmpty(sheetPrefix), $"[{fileName}] ", sheetPrefix.Replace("{file}", fileName))

                    For i = 1 To sourceWb.Worksheets.Count
                        Dim srcWs As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(sourceWb.Worksheets(i), Microsoft.Office.Interop.Excel.Worksheet)
                        Dim destName = SanitizeFileName($"{prefix}{srcWs.Name}")
                        If destName.Length > 31 Then destName = destName.Substring(0, 31)

                        If isFirstFile Then
                            ' 第一个文件的第一个工作表重命名默认表
                            If i = 1 Then
                                DirectCast(targetWb.Sheets(1), Microsoft.Office.Interop.Excel.Worksheet).Name = destName
                                srcWs.UsedRange.Copy(DirectCast(targetWb.Sheets(1), Microsoft.Office.Interop.Excel.Worksheet).Range("A1"))
                            Else
                                srcWs.Copy(After:=targetWb.Sheets(targetWb.Sheets.Count))
                                DirectCast(targetWb.Sheets(targetWb.Sheets.Count), Microsoft.Office.Interop.Excel.Worksheet).Name = destName
                            End If
                        Else
                            srcWs.Copy(After:=targetWb.Sheets(targetWb.Sheets.Count))
                            DirectCast(targetWb.Sheets(targetWb.Sheets.Count), Microsoft.Office.Interop.Excel.Worksheet).Name = destName
                        End If
                        mergedSheets.Add(destName)
                        Utils.ComReleaseHelper.Release(srcWs)
                    Next

                    isFirstFile = False
                    sourceWb.Close(SaveChanges:=False)
                    Utils.ComReleaseHelper.Release(sourceWb)
                Next

                ' 删除残留的默认空白表（Workbooks.Add 自带的 Sheet1/2/…）。
                ' 按"是否属于合并结果"的表名集合来删，从后往前删，避免误删已命名的数据表。
                Try
                    Dim mergedNames As New HashSet(Of String)(mergedSheets, StringComparer.OrdinalIgnoreCase)
                    For i = targetWb.Sheets.Count To 1 Step -1
                        Dim s As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(targetWb.Sheets(i), Microsoft.Office.Interop.Excel.Worksheet)
                        Dim nm = s.Name
                        If Not mergedNames.Contains(nm) Then
                            s.Delete()
                        End If
                        Utils.ComReleaseHelper.Release(s)
                    Next
                Catch
                End Try

                app.DisplayAlerts = displayAlertsOriginal

                ' 保存
                If Not String.IsNullOrEmpty(targetPath) Then
                    targetWb.SaveAs(Filename:=targetPath, FileFormat:=Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook)
                End If

                result.Content.Add(New McpTextContent With {
                    .Text = $"成功合并 {filePaths.Count} 个文件，共 {mergedSheets.Count} 个工作表"
                })

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"合并工作簿失败: {ex.Message}"})
            Finally
                app.DisplayAlerts = displayAlertsOriginal
            End Try

            Return result
        End Function
        ''' <summary>
        ''' 安全清理工作簿中"遗留的默认空白表"：仅删除 名为默认模式(^Sheet\d+$) 且 确认为空 的表。
        ''' 用于消除历史操作(Workbooks.Add/Sheets.Add 未清理)累积的 Sheet2/3/... 空白表，绝不触碰有数据的表。
        ''' </summary>
        Private Function DeleteEmptyDefaultSheets(wb As Microsoft.Office.Interop.Excel.Workbook,
                                                  excelApp As Microsoft.Office.Interop.Excel.Application) As Integer
            Dim removed As Integer = 0
            Dim alerts = excelApp.DisplayAlerts
            excelApp.DisplayAlerts = False
            Try
                For i = wb.Worksheets.Count To 1 Step -1
                    Dim ws As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(wb.Worksheets(i), Microsoft.Office.Interop.Excel.Worksheet)
                    Try
                        Dim nm = ws.Name
                        ' 只处理默认命名，避免误删用户表
                        If Not System.Text.RegularExpressions.Regex.IsMatch(nm, "^Sheet\d+$") Then
                            Utils.ComReleaseHelper.Release(ws)
                            Continue For
                        End If
                        Dim isEmpty As Boolean = False
                        Try
                            Dim used = ws.UsedRange
                            If used IsNot Nothing Then
                                Dim bottom = used.Row + used.Rows.Count - 1
                                Dim right = used.Column + used.Columns.Count - 1
                                isEmpty = (bottom <= 1 AndAlso right <= 1) AndAlso
                                          (used.Cells(1, 1) Is Nothing OrElse DirectCast(used.Cells(1, 1), Microsoft.Office.Interop.Excel.Range).Value2 Is Nothing)
                                Utils.ComReleaseHelper.Release(used)
                            End If
                        Catch
                            isEmpty = False
                        End Try
                        If isEmpty Then
                            ws.Delete()
                            removed += 1
                        End If
                        Utils.ComReleaseHelper.Release(ws)
                    Catch
                        Utils.ComReleaseHelper.Release(ws)
                    End Try
                Next
            Finally
                excelApp.DisplayAlerts = alerts
            End Try
            Return removed
        End Function
        ''' <summary>
        ''' 合并同一工作簿内多个工作表的数据到一张表（按行追加）
        ''' </summary>
        Private Function ExecuteMergeSheets(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Dim app = Globals.ThisAddIn.Application
            Dim displayAlertsOriginal = app.DisplayAlerts
            Try
                Dim sourceSheetsJson As String = GetArg(Of String)(args, "sourceSheets", "[]")
                ' 可选落地：targetSheet 留空时自动写 '合并结果' 表（确保有实体产出，不再只返回计算数据）
                Dim targetSheetName As String = GetArg(Of String)(args, "targetSheet", "")
                If String.IsNullOrWhiteSpace(targetSheetName) Then targetSheetName = "合并结果"
                Dim includeHeaders As Boolean = GetArg(Of Boolean)(args, "includeHeaders", True)
                ' 用户明确要求"合并所有Sheet"时，忽略当前选中的工作表标签，强制合并全部可见表
                Dim mergeAll As Boolean = GetArg(Of Boolean)(args, "mergeAll", False)

                Dim wb = app.ActiveWorkbook
                If wb Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何工作簿"})
                    Return result
                End If

                ' 解析用户指定的源表
                Dim sourceNames As List(Of String) = Nothing
                Try
                    sourceNames = JsonConvert.DeserializeObject(Of List(Of String))(sourceSheetsJson)
                Catch
                    sourceNames = Nothing
                End Try
                If sourceNames Is Nothing Then sourceNames = New List(Of String)()

                ' 用户未在消息中指定源表 → 尝试读取 Excel 中当前选中的多个工作表标签（尊重"选中哪些表"）
                ' mergeAll=true 时跳过选中表逻辑，直接合并全部可见表，避免"合并所有"只合并了选中的几张表
                Dim usedSelectedSheets As Boolean = False
                If sourceNames.Count = 0 AndAlso Not mergeAll Then
                    Try
                        Dim aw = app.ActiveWindow
                        If aw IsNot Nothing Then
                            Dim sel = TryCast(aw.SelectedSheets, Microsoft.Office.Interop.Excel.Sheets)
                            If sel IsNot Nothing AndAlso sel.Count > 1 Then
                                For i = 1 To sel.Count
                                    Dim s = TryCast(sel(i), Microsoft.Office.Interop.Excel.Worksheet)
                                    If s IsNot Nothing AndAlso Not String.Equals(s.Name, targetSheetName, StringComparison.OrdinalIgnoreCase) Then
                                        sourceNames.Add(s.Name)
                                    End If
                                    Utils.ComReleaseHelper.Release(s)
                                Next
                                usedSelectedSheets = (sourceNames.Count > 0)
                            End If
                            Utils.ComReleaseHelper.Release(aw)
                        End If
                    Catch
                    End Try
                End If

                ' 读取源表数据（0 基二维数组）；保留首个源表的每列数字格式用于还原日期/百分比等显示类型
                Dim arrays As New List(Of Object(,))()
                Dim usedSourceNames As New List(Of String)()
                Dim firstFormats As String() = Nothing

                If sourceNames.Count > 0 Then
                    For Each sn In sourceNames
                        Dim sws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                        Try
                            sws = DirectCast(wb.Worksheets(sn), Microsoft.Office.Interop.Excel.Worksheet)
                            If sws.Visible <> Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible Then Continue For
                            Dim fmts As String() = Nothing
                            Dim t = ReadRangeTo0BasedWithFormats(sws.UsedRange, fmts)
                            ' 裁剪 UsedRange 幽灵空行/空列，并跳过整表空白的表
                            Dim t2 = TrimPhantomRanges(t)
                            If t2 IsNot Nothing AndAlso Not ArrayIsAllBlank(t2) Then
                                If firstFormats Is Nothing Then firstFormats = fmts
                                arrays.Add(t2)
                                usedSourceNames.Add(sws.Name)
                            End If
                        Catch
                        Finally
                            Utils.ComReleaseHelper.Release(sws)
                        End Try
                    Next
                Else
                    For i = 1 To wb.Worksheets.Count
                        Dim sws = DirectCast(wb.Worksheets(i), Microsoft.Office.Interop.Excel.Worksheet)
                        If sws.Visible <> Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible Then Continue For
                        If String.Equals(sws.Name, targetSheetName, StringComparison.OrdinalIgnoreCase) Then
                            Utils.ComReleaseHelper.Release(sws)
                            Continue For
                        End If
                        Dim fmts As String() = Nothing
                        Dim t = ReadRangeTo0BasedWithFormats(sws.UsedRange, fmts)
                        Dim t2 = TrimPhantomRanges(t)
                        If t2 IsNot Nothing AndAlso Not ArrayIsAllBlank(t2) Then
                            If firstFormats Is Nothing Then firstFormats = fmts
                            arrays.Add(t2)
                            usedSourceNames.Add(sws.Name)
                        End If
                        Utils.ComReleaseHelper.Release(sws)
                    Next
                End If

                ' 【兜底】显式指定/选中态解析出的源表均无可读数据（首次调用选中态可能被折叠、或传入的旧表名已失效）
                ' → 回退为合并"除目标表外所有可见非空表"，避免误报"没有找到可合并的工作表"。
                If arrays.Count = 0 Then
                    For i = 1 To wb.Worksheets.Count
                        Dim sws = DirectCast(wb.Worksheets(i), Microsoft.Office.Interop.Excel.Worksheet)
                        Try
                            If sws.Visible <> Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible Then Continue For
                            If String.Equals(sws.Name, targetSheetName, StringComparison.OrdinalIgnoreCase) Then Continue For
                            If usedSourceNames.Contains(sws.Name) Then Continue For ' 已尝试过，跳过
                            Dim fmts As String() = Nothing
                            Dim t = ReadRangeTo0BasedWithFormats(sws.UsedRange, fmts)
                            Dim t2 = TrimPhantomRanges(t)
                            If t2 IsNot Nothing AndAlso Not ArrayIsAllBlank(t2) Then
                                If firstFormats Is Nothing Then firstFormats = fmts
                                arrays.Add(t2)
                                usedSourceNames.Add(sws.Name)
                            End If
                        Catch
                        Finally
                            Utils.ComReleaseHelper.Release(sws)
                        End Try
                    Next
                End If

                If arrays.Count = 0 Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "没有找到可合并的工作表"})
                    Return result
                End If

                Dim maxCols As Integer = 0
                For Each arr In arrays
                    If arr.GetLength(1) > maxCols Then maxCols = arr.GetLength(1)
                Next
                ' 格式数组与最大列数对齐（不足部分留空 → 默认 General，不影响显示）
                If firstFormats IsNot Nothing AndAlso firstFormats.Length < maxCols Then
                    Dim grown = New String(maxCols - 1) {}
                    Array.Copy(firstFormats, grown, firstFormats.Length)
                    firstFormats = grown
                End If

                ' —— 步骤规划（操作型任务逐步执行）——
                Dim pr As New StepRecorder()
                pr.ReviewSources(usedSourceNames.Count, usedSourceNames)
                pr.PrepareTarget(If(String.IsNullOrEmpty(targetSheetName), "合并结果", targetSheetName))

                ' 逐源表裁剪数据并去重表头：仅保留首源表首行为表头写入一次，其余源表剔重后作为数据
                ' 便于后续"逐个插入"：每张源表单独一批写进目标表末尾，天然不会复制表头。
                Dim prepared As New List(Of Tuple(Of String, List(Of List(Of Object))))()
                Dim canonicalHdr As Object() = If(arrays.Count > 0, RowOf(arrays(0), 0), Nothing)

                For idx = 0 To arrays.Count - 1
                    Dim arr = arrays(idx)
                    Dim srcName = If(idx < usedSourceNames.Count, usedSourceNames(idx), "源表" & (idx + 1))
                    Dim dataRows = New List(Of List(Of Object))()
                    Dim headerAlreadySeen As Boolean = (Not includeHeaders) OrElse canonicalHdr Is Nothing
                    For r = 0 To arr.GetLength(0) - 1
                        Dim srcRow As Object() = RowOf(arr, r)
                        If Not headerAlreadySeen Then
                            If RowIsBlank(srcRow) Then Continue For
                            If HeaderRowsMatch(srcRow, canonicalHdr) Then
                                headerAlreadySeen = True
                                Continue For
                            End If
                        End If
                        headerAlreadySeen = True
                        Dim row = New List(Of Object)()
                        For c = 0 To maxCols - 1
                            If c < arr.GetLength(1) Then
                                row.Add(NormalizeCell(arr(r, c)))
                            Else
                                row.Add(Nothing)
                            End If
                        Next
                        dataRows.Add(row)
                    Next
                    prepared.Add(Tuple.Create(srcName, dataRows))
                Next

                ' ----- 组装合并结果到内存：表头一次 + 各源表数据行 -----
                Dim mergedData As New List(Of List(Of Object))()
                If includeHeaders AndAlso canonicalHdr IsNot Nothing Then
                    Dim hdrRow = New List(Of Object)()
                    For c = 0 To maxCols - 1
                        If c < canonicalHdr.Length Then
                            hdrRow.Add(NormalizeCell(canonicalHdr(c)))
                        Else
                            hdrRow.Add(Nothing)
                        End If
                    Next
                    mergedData.Add(hdrRow)
                End If
                For Each p In prepared
                    If p.Item2 IsNot Nothing Then
                        For Each r In p.Item2
                            mergedData.Add(r)
                        Next
                    End If
                Next
                Dim totalDataRows As Integer = mergedData.Count - (If(includeHeaders AndAlso canonicalHdr IsNot Nothing, 1, 0))

                ' ----- 落地：默认写 '合并结果' 表（自动新建或清空复用），确保有实体产出 -----
                Dim landedSheetName As String = Nothing
                Dim cleanedSheets As Integer = 0
                ' targetSheetName 此时保证非空（上方已填默认值），总是执行写表逻辑
                app.DisplayAlerts = False
                Dim destWs As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Try
                    destWs = DirectCast(wb.Worksheets(targetSheetName), Microsoft.Office.Interop.Excel.Worksheet)
                    destWs.Cells.Clear()
                Catch
                    Dim added = DirectCast(wb.Worksheets.Add(), Microsoft.Office.Interop.Excel.Worksheet)
                    Try
                        added.Name = targetSheetName
                        destWs = added
                    Catch
                        Try : added.Delete() : Catch : End Try
                        Utils.ComReleaseHelper.Release(added)
                        Throw
                    End Try
                End Try

                WriteBlockRows(destWs, mergedData, 1, maxCols)
                pr.WriteHeader(1)

                ' 还原日期/百分比等数字格式掩码（数据区域从第 2 行开始，若无双表头则从第 1 行开始）
                ApplyColumnFormats(destWs, firstFormats, If(includeHeaders, 2, 1), totalDataRows, maxCols)

                ' 清理工作簿中遗留的默认空白表(^Sheet\d+$ 且为空)，并报告清理数量
                Try
                    cleanedSheets = DeleteEmptyDefaultSheets(wb, app)
                Catch
                    cleanedSheets = 0
                End Try
                app.DisplayAlerts = displayAlertsOriginal
                Utils.ComReleaseHelper.Release(destWs)
                landedSheetName = targetSheetName

                ' 【根治 2026-10-09】不再回传完整 data 矩阵。数据已经 WriteBlockRows 落到
                ' "合并结果"工作表，回传全量矩阵只有坏处：回喂 LLM 后被"复述"成无逗号裸 JSON
                ' 倒进气泡（用户连续 9 次实测），且白白膨胀桥接传输。只回 10 行预览供 UI 展示。
                Dim previewData As New List(Of List(Of Object))()
                For i = 0 To Math.Min(10, mergedData.Count) - 1
                    previewData.Add(mergedData(i))
                Next

                Dim resultObj = New Dictionary(Of String, Object) From {
                    {"success", True},
                    {"sourceSheets", usedSourceNames},
                    {"mergedRows", totalDataRows},
                    {"dataRows", totalDataRows},
                    {"headerPreserved", includeHeaders},
                    {"usedSelectedSheets", usedSelectedSheets},
                    {"previewRows", previewData},
                    {"previewTruncated", totalDataRows + If(includeHeaders AndAlso canonicalHdr IsNot Nothing, 1, 0) > previewData.Count},
                    {"steps", pr.Steps}
                }
                If landedSheetName IsNot Nothing Then
                    resultObj("writtenTo") = landedSheetName
                    resultObj("targetSheet") = landedSheetName
                    resultObj("cleanedEmptySheets") = cleanedSheets
                End If
                result.Content.Add(New McpTextContent With {.Text = JsonConvert.SerializeObject(resultObj, Formatting.Indented)})

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"合并工作表失败: {ex.Message}"})
            Finally
                app.DisplayAlerts = displayAlertsOriginal
            End Try
            Return result
        End Function

    End Class

End Namespace