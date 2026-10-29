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
        ''' 原子能力：向指定工作表写入数据行（插入/追加）。
        ''' rows 为二维数组（JSON），从 startRow(1基) 写入；append=true 时自动取目标表下一个空行追加。
        ''' 供合并/拆分等"操作型任务"分步调用（模型规划：读源表 → 定位目标行 → 逐表插入）。
        ''' </summary>
        Private Function ExecuteInsertRows(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Dim app = Globals.ThisAddIn.Application
            Dim displayAlertsOriginal = app.DisplayAlerts
            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim rowsJson As String = GetArg(Of String)(args, "rows", "[]")
                Dim startRow As Integer = GetArg(Of Integer)(args, "startRow", 1)
                Dim append As Boolean = GetArg(Of Boolean)(args, "append", True)
                Dim createIfMissing As Boolean = GetArg(Of Boolean)(args, "createIfMissing", True)

                If String.IsNullOrWhiteSpace(sheetName) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "缺少目标工作表名(sheet)"})
                    Return result
                End If

                ' 解析二维数组（Object[][] 的 JSON 形式）
                Dim rawRows As Object = Nothing
                Try
                    rawRows = JsonConvert.DeserializeObject(Of Object())(rowsJson)
                Catch
                    rawRows = Nothing
                End Try
                If rawRows Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "rows 必须是二维 JSON 数组，如 [[""a"",1],[""b"",2]]"})
                    Return result
                End If

                ' 归一化为 Object(,) 0 基
                Dim rowsObj = DirectCast(rawRows, Object())
                Dim maxCols As Integer = 0
                Dim rc As Integer = 0
                For i = 0 To rowsObj.Length - 1
                    Dim row = rowsObj(i)
                    If TypeOf row Is Object() Then row = CType(row, Object())
                    If TypeOf row Is Newtonsoft.Json.Linq.JArray Then row = CType(row, Newtonsoft.Json.Linq.JArray).ToObject(Of Object())()
                    If row Is Nothing Then Continue For
                    Dim rowArr = TryCast(row, Object())
                    If rowArr Is Nothing Then Continue For
                    If rowArr.Length > maxCols Then maxCols = rowArr.Length
                    rc += 1
                Next
                If rc = 0 OrElse maxCols = 0 Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "rows 为空"})
                    Return result
                End If

                Dim data As Object(,) = Nothing
                Dim dr2 = 0
                For i = 0 To rowsObj.Length - 1
                    Dim rowObj = rowsObj(i)
                    If TypeOf rowObj Is Newtonsoft.Json.Linq.JArray Then rowObj = CType(rowObj, Newtonsoft.Json.Linq.JArray).ToObject(Of Object())()
                    Dim rowArr = TryCast(rowObj, Object())
                    If rowArr Is Nothing Then Continue For
                    If data Is Nothing Then data = New Object(rc - 1, maxCols - 1) {}
                    For c = 0 To maxCols - 1
                        If c < rowArr.Length Then data(dr2, c) = NormalizeCell(rowArr(c)) Else data(dr2, c) = Nothing
                    Next
                    dr2 += 1
                Next
                If data Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "rows 解析失败"})
                    Return result
                End If

                Dim wb = app.ActiveWorkbook
                If wb Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何工作簿"})
                    Return result
                End If

                ' 定位目标表（不存在则按 createIfMissing 新建；新建失败重命名失败要删除刚新增的空白表）
                Dim destWs As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                Try
                    destWs = DirectCast(wb.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                Catch
                    If Not createIfMissing Then
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = $"目标工作表不存在: {sheetName}"})
                        Return result
                    End If
                    app.DisplayAlerts = False
                    Dim added = DirectCast(wb.Worksheets.Add(), Microsoft.Office.Interop.Excel.Worksheet)
                    Try
                        added.Name = sheetName
                        destWs = added
                    Catch
                        Try : added.Delete() : Catch : End Try
                        Utils.ComReleaseHelper.Release(added)
                        Throw
                    End Try
                End Try

                ' 确定实际起始行（1 基，已用区域下一空行）
                Dim targetStart = startRow
                If append Then
                    Dim usedRng = destWs.UsedRange
                    Dim usedRows = 0
                    If usedRng IsNot Nothing Then usedRows = usedRng.Row + usedRng.Rows.Count - 1
                    Utils.ComReleaseHelper.Release(usedRng)
                    ' UsedRange 可能含格式残留，向前搜索首个真正空行
                    targetStart = Math.Max(1, usedRows + 1)
                End If

                Dim writeRng = DirectCast(destWs.Range(destWs.Cells(targetStart, 1), destWs.Cells(targetStart + rc - 1, maxCols)), Microsoft.Office.Interop.Excel.Range)
                writeRng.Value2 = data
                Utils.ComReleaseHelper.Release(writeRng)

                app.DisplayAlerts = displayAlertsOriginal
                Utils.ComReleaseHelper.Release(destWs)

                Dim resultObj = New Dictionary(Of String, Object) From {
                    {"success", True},
                    {"sheet", sheetName},
                    {"startRow", targetStart},
                    {"rowsInserted", rc},
                    {"columns", maxCols},
                    {"nextRow", targetStart + rc}
                }
                result.Content.Add(New McpTextContent With {.Text = JsonConvert.SerializeObject(resultObj, Formatting.Indented)})

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"插入数据行失败: {ex.Message}"})
            Finally
                app.DisplayAlerts = displayAlertsOriginal
            End Try
            Return result
        End Function
        ''' <summary>
        ''' 按行数拆分当前工作表数据
        ''' </summary>
        Private Function ExecuteSplitByRows(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Dim app = Globals.ThisAddIn.Application
            Dim displayAlertsOriginal = app.DisplayAlerts
            Dim pr As New StepRecorder()
            Try
                ' 兼容两种参数命名：rowsPerFile / rowsPerChunk；includeHeaders / keepHeader
                Dim rowsPerFile As Integer = GetArg(Of Integer)(args, "rowsPerFile", GetArg(Of Integer)(args, "rowsPerChunk", 500))
                Dim includeHeaders As Boolean = GetArg(Of Boolean)(args, "includeHeaders", GetArg(Of Boolean)(args, "keepHeader", True))
                Dim outputMode As String = GetArg(Of String)(args, "outputMode", "sheets")
                Dim outputDir As String = GetArg(Of String)(args, "outputDir", "")
                Dim namingPattern As String = GetArg(Of String)(args, "namingPattern", "{original}_第{i}份")

                If rowsPerFile < 1 Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "每份行数必须大于 0"})
                    Return result
                End If

                Dim wb = app.ActiveWorkbook
                If wb Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何工作簿"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet = TryCast(app.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                Dim wsName = ws.Name
                Dim srcFormats As String() = Nothing
                Dim arr = ReadRangeTo0BasedWithFormats(ws.UsedRange, srcFormats)
                Utils.ComReleaseHelper.Release(ws)
                ' 裁剪 UsedRange 幽灵空行/空列，避免 104 万行的幽灵区域被当成真实数据，从而生成大量空白分块表
                arr = TrimPhantomRanges(arr)
                If arr Is Nothing OrElse ArrayIsAllBlank(arr) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前工作表没有数据"})
                    Return result
                End If
                pr.ReadSource(wsName, arr.GetLength(0))

                Dim totalCols = arr.GetLength(1)
                Dim headerRow As Object() = Nothing
                Dim dataRows As New List(Of Integer)()
                If includeHeaders Then
                    Dim h = New Object(totalCols - 1) {}
                    For c = 0 To totalCols - 1 : h(c) = NormalizeCell(arr(0, c)) : Next
                    headerRow = h
                    For r = 1 To arr.GetLength(0) - 1 : dataRows.Add(r) : Next
                Else
                    For r = 0 To arr.GetLength(0) - 1 : dataRows.Add(r) : Next
                End If

                Dim totalData = dataRows.Count
                If totalData = 0 Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前工作表没有数据（除表头外）"})
                    Return result
                End If

                Dim originalName = IO.Path.GetFileNameWithoutExtension(wb.FullName)
                Dim created As New List(Of String)()
                Dim chunkCount As Integer = CInt(Math.Ceiling(totalData / CDbl(rowsPerFile)))
                pr.SplitByRowCount(rowsPerFile, chunkCount, totalData)

                app.DisplayAlerts = False
                For chunk = 1 To chunkCount
                    Dim startIdx = (chunk - 1) * rowsPerFile
                    Dim endIdx = Math.Min(startIdx + rowsPerFile, totalData) - 1
                    Dim rowsInChunk = endIdx - startIdx + 1

                    ' 组装该份数据（0 基，含表头则首行为表头）
                    Dim outArr = New Object(rowsInChunk - (If(headerRow IsNot Nothing, 1, 0)), totalCols - 1) {}
                    Dim outR = 0
                    If headerRow IsNot Nothing Then
                        For c = 0 To totalCols - 1 : outArr(outR, c) = headerRow(c) : Next
                        outR += 1
                    End If
                    For rr = startIdx To endIdx
                        Dim srcIdx = dataRows(rr)
                        For c = 0 To totalCols - 1 : outArr(outR, c) = NormalizeCell(arr(srcIdx, c)) : Next
                        outR += 1
                    Next

                    If outputMode = "files" Then
                        If String.IsNullOrWhiteSpace(outputDir) Then
                            outputDir = IO.Path.GetDirectoryName(wb.FullName)
                        End If
                        ' 【防幻觉路径】净化模型编造的输出目录，避免 CreateDirectory 造出垃圾目录
                        outputDir = Utils.PathHelper.SanitizeUserDirectory(outputDir)
                        If Not String.IsNullOrWhiteSpace(outputDir) AndAlso Not IO.Directory.Exists(outputDir) Then
                            IO.Directory.CreateDirectory(outputDir)
                        End If
                        Dim fileName = ExcelMcpCoreUtils.BuildChunkName(namingPattern, originalName, chunk, dataRows(startIdx) + 1, dataRows(endIdx) + 1)
                        fileName = SanitizeFileName(fileName)
                        Dim fullPath = GetUniqueFilePath(outputDir, fileName, "xlsx")

                        Dim newWb = app.Workbooks.Add()
                        ' 新建工作簿自带多个默认空白 Sheet，删除多余空白表，避免每个输出文件都出现一堆空白 Sheet
                        While newWb.Worksheets.Count > 1
                            Dim extraWs = DirectCast(newWb.Worksheets(newWb.Worksheets.Count), Microsoft.Office.Interop.Excel.Worksheet)
                            extraWs.Delete()
                            Utils.ComReleaseHelper.Release(extraWs)
                        End While
                        Dim newWs As Microsoft.Office.Interop.Excel.Worksheet = DirectCast(newWb.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)
                        WriteArrayWithFormats(newWs, outArr, headerRow IsNot Nothing, srcFormats)
                        newWb.SaveAs(Filename:=fullPath, FileFormat:=Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook)
                        newWb.Close(SaveChanges:=False)
                        created.Add(fullPath)
                        pr.GenerateFile(IO.Path.GetFileName(fullPath), rowsInChunk, includeHeaders)
                        Utils.ComReleaseHelper.Release(newWs)
                        Utils.ComReleaseHelper.Release(newWb)
                    Else
                        Dim sheetName = ExcelMcpCoreUtils.BuildChunkName(namingPattern, originalName, chunk, dataRows(startIdx) + 1, dataRows(endIdx) + 1)
                        If sheetName.Length > 31 Then sheetName = sheetName.Substring(0, 31)
                        Dim newWs = DirectCast(wb.Worksheets.Add(), Microsoft.Office.Interop.Excel.Worksheet)
                        newWs.Name = sheetName
                        WriteArrayWithFormats(newWs, outArr, headerRow IsNot Nothing, srcFormats)
                        created.Add(sheetName)
                        pr.NewSheet(sheetName, rowsInChunk, includeHeaders)
                        Utils.ComReleaseHelper.Release(newWs)
                    End If
                Next
                app.DisplayAlerts = displayAlertsOriginal

                Dim resultObj = New Dictionary(Of String, Object) From {
                    {"success", True},
                    {"chunks", chunkCount},
                    {"rowsPerChunk", rowsPerFile},
                    {"created", created},
                    {"steps", pr.Steps}
                }
                result.Content.Add(New McpTextContent With {.Text = JsonConvert.SerializeObject(resultObj, Formatting.Indented)})

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"按行数拆分失败: {ex.Message}"})
            Finally
                app.DisplayAlerts = displayAlertsOriginal
            End Try
            Return result
        End Function

    End Class

End Namespace