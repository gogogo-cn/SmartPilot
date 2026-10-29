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

        Private Function ExecuteSearch(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim query As String = GetArg(Of String)(args, "query", "")
                Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
                Dim rangeAddr As String = GetArg(Of String)(args, "range", "")

                If String.IsNullOrEmpty(query) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "搜索关键词不能为空"})
                    Return result
                End If

                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                    Return result
                End If

                Dim foundCells = New List(Of Dictionary(Of String, Object))()
                Dim firstFind As Microsoft.Office.Interop.Excel.Range = Nothing
                Dim currentFind As Microsoft.Office.Interop.Excel.Range = Nothing
                Dim searchRange As Microsoft.Office.Interop.Excel.Range = Nothing
                Dim searchWs As Microsoft.Office.Interop.Excel.Worksheet = Nothing

                Try
                    ' 确定搜索范围
                    If Not String.IsNullOrEmpty(sheetName) AndAlso Not String.IsNullOrEmpty(rangeAddr) Then
                        searchWs = TryCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                        If searchWs IsNot Nothing Then
                            searchRange = searchWs.Range(rangeAddr)
                        End If
                    ElseIf Not String.IsNullOrEmpty(sheetName) Then
                        searchWs = TryCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                        If searchWs IsNot Nothing Then
                            searchRange = searchWs.UsedRange
                        End If
                    ElseIf Not String.IsNullOrEmpty(rangeAddr) Then
                        searchRange = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet).Range(rangeAddr)
                    Else
                        searchRange = DirectCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet).UsedRange
                    End If

                    If searchRange IsNot Nothing Then
                        currentFind = DirectCast(searchRange.Find(query, Type.Missing,
                            Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues,
                            Microsoft.Office.Interop.Excel.XlLookAt.xlPart,
                            Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
                            Microsoft.Office.Interop.Excel.XlSearchDirection.xlNext,
                            False), Microsoft.Office.Interop.Excel.Range)

                        If currentFind IsNot Nothing Then
                            firstFind = currentFind
                            Do
                                foundCells.Add(New Dictionary(Of String, Object) From {
                                    {"address", currentFind.Address(False, False)},
                                    {"value", If(currentFind.Value2 IsNot Nothing, currentFind.Value2.ToString(), "")},
                                    {"sheet", currentFind.Worksheet.Name}
                                })
                                Utils.ComReleaseHelper.Release(currentFind)
                                currentFind = Nothing
                                currentFind = DirectCast(searchRange.FindNext(If(firstFind IsNot Nothing, firstFind, Nothing)), Microsoft.Office.Interop.Excel.Range)
                                If currentFind Is Nothing Then Exit Do
                            Loop While currentFind.Address <> firstFind.Address
                        End If
                    End If

                    result.Content.Add(New McpTextContent With {
                        .Text = JsonConvert.SerializeObject(New With {
                            .query = query,
                            .count = foundCells.Count,
                            .results = foundCells
                        }, Formatting.Indented)
                    })

                Catch ex As Exception
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = $"搜索失败: {ex.Message}"})
                Finally
                    Utils.ComReleaseHelper.Release(currentFind)
                    Utils.ComReleaseHelper.Release(firstFind)
                    Utils.ComReleaseHelper.Release(searchRange)
                    Utils.ComReleaseHelper.Release(searchWs)
                End Try

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"搜索失败: {ex.Message}"})
            End Try
            Return result
        End Function

    End Class

End Namespace