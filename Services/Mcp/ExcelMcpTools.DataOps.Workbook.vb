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
        ''' 打开指定路径的 Excel 文件
        ''' </summary>
        Private Function ExecuteOpenWorkbook(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()

            Try
                Dim filePath As String = GetArg(Of String)(args, "filePath", "")

                If String.IsNullOrEmpty(filePath) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "文件路径不能为空"})
                    Return result
                End If

                If Not IO.File.Exists(filePath) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"文件不存在: {filePath}"})
                    Return result
                End If

                Dim app = Globals.ThisAddIn.Application
                If app Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "Excel 应用程序未就绪"})
                    Return result
                End If

                ' 检查是否已打开
                If IsWorkbookOpen(app, filePath) Then
                    result.Content.Add(New McpTextContent With {.Text = $"工作簿已打开: {filePath}"})
                    Return result
                End If

                Dim wb = app.Workbooks.Open(filePath)
                Dim sheetCount = wb.Worksheets.Count

                result.Content.Add(New McpTextContent With {
                    .Text = $"成功打开工作簿: {filePath}，共 {sheetCount} 个工作表"
                })

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"打开工作簿失败: {ex.Message}"})
            End Try

            Return result
        End Function
        ''' <summary>
        ''' 关闭指定工作簿
        ''' </summary>
        Private Function ExecuteCloseWorkbook(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()

            Try
                Dim filePath As String = GetArg(Of String)(args, "filePath", "")
                Dim saveChanges As Boolean = GetArg(Of Boolean)(args, "saveChanges", False)

                Dim app = Globals.ThisAddIn.Application
                If app Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "Excel 应用程序未就绪"})
                    Return result
                End If

                Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing

                If String.IsNullOrEmpty(filePath) Then
                    ' 关闭当前活动工作簿
                    wb = app.ActiveWorkbook
                Else
                    ' 按路径查找工作簿
                    Dim targetName = IO.Path.GetFileName(filePath)
                    Dim count = app.Workbooks.Count
                    For i = 1 To count
                        Dim tempWb = app.Workbooks(i)
                        If String.Equals(IO.Path.GetFileName(tempWb.Name), targetName, StringComparison.OrdinalIgnoreCase) OrElse
                           String.Equals(tempWb.FullName, filePath, StringComparison.OrdinalIgnoreCase) Then
                            wb = tempWb
                            Exit For
                        End If
                        Utils.ComReleaseHelper.Release(tempWb)
                    Next
                End If

                If wb Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"找不到工作簿: {If(filePath, "当前活动工作簿")}"})
                    Return result
                End If

                Dim wbName = wb.Name
                wb.Close(SaveChanges:=saveChanges)
                Utils.ComReleaseHelper.Release(wb)

                result.Content.Add(New McpTextContent With {
                    .Text = $"成功关闭工作簿: {wbName}"
                })

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"关闭工作簿失败: {ex.Message}"})
            End Try

            Return result
        End Function
        ''' <summary>
        ''' 获取工作表的已使用范围
        ''' </summary>
        Private Function ExecuteGetUsedRange(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()

            Try
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")

                Dim app = Globals.ThisAddIn.Application
                If app Is Nothing OrElse app.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何工作簿"})
                    Return result
                End If

                Dim ws As Microsoft.Office.Interop.Excel.Worksheet
                If Not String.IsNullOrEmpty(sheetName) Then
                    ws = TryCast(app.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                Else
                    ws = TryCast(app.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                End If

                If ws Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: {sheetName}"})
                    Return result
                End If

                Dim usedRange = ws.UsedRange
                Dim rangeAddr = usedRange.Address(False, False)
                Dim rowCount = usedRange.Rows.Count
                Dim colCount = usedRange.Columns.Count
                Dim startRow = usedRange.Row
                Dim startCol = usedRange.Column

                Utils.ComReleaseHelper.Release(usedRange)
                Utils.ComReleaseHelper.Release(ws)

                Dim resultObj = New Dictionary(Of String, Object) From {
                    {"sheet", ws.Name},
                    {"range", rangeAddr},
                    {"startRow", startRow},
                    {"startCol", startCol},
                    {"rowCount", rowCount},
                    {"colCount", colCount}
                }

                result.Content.Add(New McpTextContent With {
                    .Text = JsonConvert.SerializeObject(resultObj, Formatting.Indented)
                })

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"获取已使用范围失败: {ex.Message}"})
            End Try

            Return result
        End Function
        ''' <summary>
        ''' 检查工作簿是否已打开（使用索引循环，避免 COM 枚举器崩溃）
        ''' </summary>
        Private Function IsWorkbookOpen(excelApp As Microsoft.Office.Interop.Excel.Application, fileName As String) As Boolean
            Dim targetName = IO.Path.GetFileName(fileName)
            Dim count As Integer = excelApp.Workbooks.Count
            For i As Integer = 1 To count
                Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
                Try
                    wb = excelApp.Workbooks(i)
                    If String.Equals(IO.Path.GetFileName(wb.Name), targetName, StringComparison.OrdinalIgnoreCase) Then
                        Utils.ComReleaseHelper.Release(wb)
                        Return True
                    End If
                Finally
                    Utils.ComReleaseHelper.Release(wb)
                End Try
            Next
            Return False
        End Function
        ''' <summary>
        ''' 清理文件名非法字符
        ''' </summary>
        Private Function SanitizeFileName(name As String) As String
            Dim invalid As String = New String(IO.Path.GetInvalidFileNameChars())
            For Each c As Char In invalid
                name = name.Replace(c, "_"c)
            Next
            name = System.Text.RegularExpressions.Regex.Replace(name, "_+", "_")
            Return name.Trim().Trim("_"c)
        End Function
        ''' <summary>
        ''' 获取唯一文件路径，避免覆盖
        ''' </summary>
        Private Function GetUniqueFilePath(dir As String, baseName As String, ext As String) As String
            Dim fullPath = IO.Path.Combine(dir, $"{baseName}.{ext}")
            If Not IO.File.Exists(fullPath) Then Return fullPath
            Dim counter As Integer = 1
            Do
                fullPath = IO.Path.Combine(dir, $"{baseName}_{counter:D2}.{ext}")
                counter += 1
            Loop While IO.File.Exists(fullPath)
            Return fullPath
        End Function

    End Class

End Namespace