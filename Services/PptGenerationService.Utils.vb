Option Explicit On
Option Infer On
Option Strict Off

Imports System
Imports System.IO
Imports System.Text
Imports Newtonsoft.Json.Linq
Imports ExcelAIVSTOAddIn.Utils

Namespace Services

    Partial Public Class PptGenerationService

        Private Shared Function GetWorkbookDirectory() As String
            Try
                Dim excelApp As Object = Globals.ThisAddIn.GetExcelApp()
                If excelApp IsNot Nothing Then
                    Dim wb As Object = Nothing
                    Try
                        wb = excelApp.ActiveWorkbook
                        If wb IsNot Nothing Then
                            Dim p As String = wb.Path.ToString()
                            If Not String.IsNullOrWhiteSpace(p) AndAlso Directory.Exists(p) Then
                                Return p
                            End If
                        End If
                    Finally
                        ' ActiveWorkbook 是宿主共享对象，只读 Path，不释放
                        Try
                            If wb IsNot Nothing Then ComReleaseHelper.Release(wb)
                        Catch
                        End Try
                    End Try
                End If
            Catch
            End Try
            Try
                Return Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            Catch
                Return System.IO.Path.GetTempPath()
            End Try
        End Function

        ' ------------------------------------------------------------
        Private Shared Function SanitizeFileName(name As String) As String
            Dim invalid As Char() = Path.GetInvalidFileNameChars()
            Dim sb As New System.Text.StringBuilder()
            For Each ch As Char In (If(name, "报告"))
                If Array.IndexOf(invalid, ch) < 0 Then sb.Append(ch)
            Next
            Dim result As String = sb.ToString().Trim()
            If String.IsNullOrWhiteSpace(result) Then result = "数据分析报告"
            Return result
        End Function

        Private Shared Function JsonSuccess(message As String, path As String) As String
            Dim jo As New JObject()
            jo("success") = True
            jo("message") = message
            jo("path") = path
            Return jo.ToString()
        End Function

        Private Shared Function JsonError(message As String) As String
            Dim jo As New JObject()
            jo("success") = False
            jo("error") = message
            Return jo.ToString()
        End Function

        ''' <summary>
        ''' 生成历史日志：每次 PPT 生成成功/失败追加一条记录到「我的文档\PPT生成历史.log」
        ''' </summary>
        Private Shared Sub AppendGenerationLog(success As Boolean, reportTitle As String, themeName As String,
                                               pageCount As Integer, filePath As String, errorMsg As String)
            Try
                Dim logPath As String = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "PPT生成历史.log")
                Dim line As String = String.Format(
                    "[{0}] {1} | 标题: {2} | 主题: {3} | 页数: {4} | 文件: {5}{6}",
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    If(success, "成功", "失败"),
                    If(String.IsNullOrEmpty(reportTitle), "-", reportTitle),
                    If(String.IsNullOrEmpty(themeName), "-", themeName),
                    pageCount.ToString(),
                    If(String.IsNullOrEmpty(filePath), "-", filePath),
                    If(String.IsNullOrEmpty(errorMsg), "", " | 错误: " & errorMsg))
                File.AppendAllText(logPath, line & vbCrLf, System.Text.Encoding.UTF8)
            Catch
            End Try
        End Sub
    End Class
End Namespace
