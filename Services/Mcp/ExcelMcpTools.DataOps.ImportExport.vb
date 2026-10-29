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
        ''' 导入 CSV 文件
        ''' </summary>
        Private Function ExecuteImportCsv(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()

            Try
                Dim filePath As String = GetArg(Of String)(args, "filePath", "")
                Dim delimiter As String = GetArg(Of String)(args, "delimiter", ",")
                Dim encoding As String = GetArg(Of String)(args, "encoding", "utf8")
                Dim targetSheet As String = GetArg(Of String)(args, "targetSheet", "")

                If String.IsNullOrEmpty(filePath) OrElse Not IO.File.Exists(filePath) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"CSV 文件不存在: {filePath}"})
                    Return result
                End If

                Dim app = Globals.ThisAddIn.Application
                Dim wb = app.ActiveWorkbook
                If wb Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何工作簿"})
                    Return result
                End If

                ' 确定目标工作表
                Dim ws As Microsoft.Office.Interop.Excel.Worksheet
                If Not String.IsNullOrEmpty(targetSheet) Then
                    Try
                        ws = DirectCast(wb.Worksheets(targetSheet), Microsoft.Office.Interop.Excel.Worksheet)
                    Catch
                        ws = DirectCast(wb.Worksheets.Add(), Microsoft.Office.Interop.Excel.Worksheet)
                        ws.Name = targetSheet
                    End Try
                Else
                    ws = DirectCast(wb.Worksheets.Add(), Microsoft.Office.Interop.Excel.Worksheet)
                    ws.Name = IO.Path.GetFileNameWithoutExtension(filePath)
                End If

                ' 使用 QueryTable 导入 CSV
                Dim enc = System.Text.Encoding.UTF8
                Select Case encoding.ToLower()
                    Case "gbk" : enc = System.Text.Encoding.GetEncoding("GBK")
                    Case "ansi" : enc = System.Text.Encoding.Default
                End Select

                ' 读取 CSV 内容
                Dim lines = IO.File.ReadAllLines(filePath, enc)
                If lines.Length = 0 Then
                    result.Content.Add(New McpTextContent With {.Text = "CSV 文件为空"})
                    Return result
                End If

                ' 解析并写入
                Dim delimChar As Char = If(delimiter = "\t", Convert.ToChar(9), delimiter(0))
                For rowIdx = 0 To lines.Length - 1
                    Dim fields = lines(rowIdx).Split(New Char() {delimChar})
                    For colIdx = 0 To fields.Length - 1
                        DirectCast(ws.Cells(rowIdx + 1, colIdx + 1), Microsoft.Office.Interop.Excel.Range).Value2 = fields(colIdx).Trim(""""c)
                    Next
                Next

                Utils.ComReleaseHelper.Release(ws)

                result.Content.Add(New McpTextContent With {
                    .Text = $"成功导入 CSV: {lines.Length} 行到工作表 {ws.Name}"
                })

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"导入 CSV 失败: {ex.Message}"})
            End Try

            Return result
        End Function
        ''' <summary>
        ''' 导出为 CSV
        ''' </summary>
        Private Function ExecuteExportCsv(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()

            Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Dim usedRange As Microsoft.Office.Interop.Excel.Range = Nothing

            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim filePath As String = GetArg(Of String)(args, "filePath", "")
                Dim delimiter As String = GetArg(Of String)(args, "delimiter", ",")
                Dim includeHeader As Boolean = GetArg(Of Boolean)(args, "includeHeader", True)

                Dim app = Globals.ThisAddIn.Application

                If String.IsNullOrEmpty(filePath) Then
                    ' filePath 可选：留空自动保存到桌面，文件名为 工作表名称_时间戳.csv
                    Dim desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                    Dim ts = DateTime.Now.ToString("yyyyMMdd_HHmmss")
                    Dim safeSheet = If(String.IsNullOrEmpty(sheetName), "", sheetName)
                    If String.IsNullOrEmpty(safeSheet) AndAlso app IsNot Nothing AndAlso app.ActiveSheet IsNot Nothing Then
                        Try
                            safeSheet = DirectCast(app.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet).Name
                        Catch
                        End Try
                    End If
                    If String.IsNullOrEmpty(safeSheet) Then safeSheet = "export"
                    safeSheet = PathHelper.GetSafeFileName(safeSheet)
                    filePath = IO.Path.Combine(desktop, $"{safeSheet}_{ts}.csv")
                End If

                ' 【防幻觉路径】模型常把用户名当占位符（例：C:\Users\Yourusername\Desktop\output.csv）。
                ' 该值非空 → 前端不会走默认路径推断 → 若不净化，下面的 CreateDirectory 会在磁盘上
                ' 真造出这个编造的用户目录（已实测发生）。
                Dim rawFilePath As String = filePath
                filePath = PathHelper.SanitizeUserOutputPath(filePath)
                If Not String.Equals(rawFilePath, filePath, StringComparison.OrdinalIgnoreCase) Then
                    System.Diagnostics.Debug.WriteLine($"[ExportCsv] 输出路径已净化: {rawFilePath} -> {filePath}")
                End If

                If app.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何工作簿"})
                    Return result
                End If

                If Not String.IsNullOrEmpty(sheetName) Then
                    ws = DirectCast(app.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                Else
                    ws = TryCast(app.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                End If

                If IO.Directory.Exists(filePath) Then
                    Dim safeSheetName = PathHelper.GetSafeFileName(ws.Name)
                    filePath = IO.Path.Combine(filePath, safeSheetName & ".csv")
                End If

                Dim directory = IO.Path.GetDirectoryName(filePath)
                If Not String.IsNullOrEmpty(directory) AndAlso Not IO.Directory.Exists(directory) Then
                    IO.Directory.CreateDirectory(directory)
                End If

                usedRange = ws.UsedRange
                Dim startRow = usedRange.Row
                Dim startCol = usedRange.Column
                Dim totalRows = usedRange.Rows.Count
                Dim totalCols = usedRange.Columns.Count

                Dim delimChar As String = If(delimiter = "\t", vbTab, delimiter)
                Dim sb As New System.Text.StringBuilder()

                Dim value2Data = usedRange.Value2

                If value2Data Is Nothing Then
                    IO.File.WriteAllText(filePath, sb.ToString(), System.Text.Encoding.UTF8)
                    result.Content.Add(New McpTextContent With {
                        .Text = $"成功导出到 CSV: {filePath} (空工作表)"
                    })
                    Return result
                End If

                Dim dataArray As Object(,) = Nothing
                If DirectCast(usedRange.Cells(), Microsoft.Office.Interop.Excel.Range).Count = 1 Then
                    dataArray = New Object(1, 1) {}
                    dataArray(1, 1) = value2Data
                Else
                    dataArray = DirectCast(value2Data, Object(,))
                End If

                If includeHeader Then
                    Dim headers As New List(Of String)()
                    For c = 1 To totalCols
                        headers.Add(EscapeCsvField(dataArray(startRow - usedRange.Row + 1, c)))
                    Next
                    sb.AppendLine(String.Join(delimChar, headers))
                End If

                Dim dataStartRow = If(includeHeader, startRow + 1, startRow)
                For r = dataStartRow To startRow + totalRows - 1
                    Dim fields As New List(Of String)()
                    For c = 1 To totalCols
                        fields.Add(EscapeCsvField(dataArray(r - usedRange.Row + 1, c)))
                    Next
                    sb.AppendLine(String.Join(delimChar, fields))
                Next

                IO.File.WriteAllText(filePath, sb.ToString(), System.Text.Encoding.UTF8)

                result.Content.Add(New McpTextContent With {
                    .Text = $"成功导出到 CSV: {filePath}"
                })
                result.TaskCompleted = True

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"导出 CSV 失败: {ex.Message}"})
            Finally
                Utils.ComReleaseHelper.Release(usedRange)
                Utils.ComReleaseHelper.Release(ws)
            End Try

            Return result
        End Function
        Private Function EscapeCsvField(value As Object) As String
            If value Is Nothing Then
                Return ""
            End If

            Dim strValue As String = value.ToString()

            If strValue.Contains(""""c) OrElse strValue.Contains(","c) OrElse strValue.Contains(vbCr) OrElse strValue.Contains(vbLf) Then
                Return """" & strValue.Replace("""", """""") & """"
            End If

            Return strValue
        End Function

    End Class

End Namespace