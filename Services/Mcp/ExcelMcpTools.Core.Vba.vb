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

        Private Function ExecuteVBA(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim code As String = GetArg(Of String)(args, "code", "")
                If String.IsNullOrEmpty(code) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "VBA 代码不能为空"})
                    Return result
                End If

                ' 通过 ExcelBridge 执行 VBA（包含安全检查）。撤销已经在 VbaExecutionBridge.ExecuteVBA 统一登记，
                ' 此处不再登记，避免与直接通道（前端 bridge.ExecuteVBA）产生双重登记。
                Dim vbaResult = DirectCast(_bridge, ExcelAIVSTOAddIn.ExcelBridge).ExecuteVBA(code)
                result.Content.Add(New McpTextContent With {.Text = vbaResult})

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"VBA 执行失败: {ex.Message}"})
            End Try
            Return result
        End Function
        Private Function ExecuteReadVbaMacros(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim moduleName As String = GetArg(Of String)(args, "moduleName", "")
                Dim includeCode As Boolean = GetArg(Of Boolean)(args, "includeCode", True)

                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "Excel 应用程序未就绪"})
                    Return result
                End If
                If excelApp.ActiveWorkbook Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿"})
                    Return result
                End If

                Dim wb = excelApp.ActiveWorkbook
                Dim vbProj As Object = Nothing
                Dim vbComps As Object = Nothing

                Try
                    vbProj = wb.VBProject
                    If vbProj Is Nothing Then
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {
                            .Text = "无法访问 VBProject。请在 Excel 中启用：文件→选项→信任中心→信任中心设置→宏设置→勾选'信任对 VBA 项目对象模型的访问'"
                        })
                        Return result
                    End If

                    vbComps = Microsoft.VisualBasic.Interaction.CallByName(vbProj, "VBComponents", Microsoft.VisualBasic.CallType.Get)
                    If vbComps Is Nothing Then
                        result.IsError = True
                        result.Content.Add(New McpTextContent With {.Text = "无法访问 VBComponents"})
                        Return result
                    End If

                    Dim workbookName As String = wb.Name
                    Dim compCount As Integer = CInt(Microsoft.VisualBasic.Interaction.CallByName(vbComps, "Count", Microsoft.VisualBasic.CallType.Get))
                    Dim modules = New List(Of Dictionary(Of String, Object))()

                    For i As Integer = 1 To compCount
                        Dim vbComp As Object = Nothing
                        Try
                            vbComp = Microsoft.VisualBasic.Interaction.CallByName(vbComps, "Item", Microsoft.VisualBasic.CallType.Method, i)
                            Dim compName As String = ""
                            Dim compType As String = ""
                            Dim codeLines As Integer = 0
                            Dim codeContent As String = ""

                            Try
                                compName = CStr(Microsoft.VisualBasic.Interaction.CallByName(vbComp, "Name", Microsoft.VisualBasic.CallType.Get))
                            Catch
                                compName = "未知模块"
                            End Try

                            Try
                                Dim typeVal As Integer = CInt(Microsoft.VisualBasic.Interaction.CallByName(vbComp, "Type", Microsoft.VisualBasic.CallType.Get))
                                Select Case typeVal
                                    Case 0
                                        compType = "标准模块 (Standard Module)"
                                    Case 1
                                        compType = "类模块 (Class Module)"
                                    Case 2
                                        compType = "窗体 (UserForm)"
                                    Case 3, 100
                                        compType = "文档模块 (Document)"
                                    Case Else
                                        compType = "其他 (Type=" & typeVal.ToString() & ")"
                                End Select
                            Catch
                                compType = "未知类型"
                            End Try

                            If Not String.IsNullOrEmpty(moduleName) Then
                                If Not String.Equals(compName, moduleName, StringComparison.OrdinalIgnoreCase) Then
                                    Continue For
                                End If
                            End If

                            Try
                                Dim codeModule As Object = Microsoft.VisualBasic.Interaction.CallByName(vbComp, "CodeModule", Microsoft.VisualBasic.CallType.Get)
                                    If codeModule IsNot Nothing Then
                                        codeLines = CInt(Microsoft.VisualBasic.Interaction.CallByName(codeModule, "CountOfLines", Microsoft.VisualBasic.CallType.Get))
                                        If includeCode AndAlso codeLines > 0 Then
                                            codeContent = CStr(Microsoft.VisualBasic.Interaction.CallByName(codeModule, "Lines", Microsoft.VisualBasic.CallType.Method, 1, codeLines))
                                        End If
                                    End If
                            Catch
                                codeLines = 0
                                codeContent = ""
                            End Try

                            Dim modInfo = New Dictionary(Of String, Object) From {
                                {"name", compName},
                                {"type", compType},
                                {"lineCount", codeLines}
                            }
                            If includeCode AndAlso Not String.IsNullOrEmpty(codeContent) Then
                                modInfo("code") = codeContent
                            End If
                            modules.Add(modInfo)

                        Catch ex As Exception
                            System.Diagnostics.Debug.WriteLine($"[ReadVbaMacros] 模块 {i} 读取失败: {ex.Message}")
                        Finally
                            Utils.ComReleaseHelper.Release(vbComp)
                        End Try
                    Next

                    Dim output = New With {
                        .workbook = workbookName,
                        .totalModules = compCount,
                        .returnedModules = modules.Count,
                        .includeCode = includeCode,
                        .modules = modules
                    }

                    result.Content.Add(New McpTextContent With {
                        .Text = JsonConvert.SerializeObject(output, Formatting.Indented)
                    })

                Catch ex As Exception
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {
                        .Text = $"读取 VBA 宏失败: {ex.Message}。请确认已在信任中心启用'信任对 VBA 项目对象模型的访问'。"
                    })
                Finally
                    Utils.ComReleaseHelper.Release(vbComps)
                    Utils.ComReleaseHelper.Release(vbProj)
                End Try

            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"读取 VBA 宏异常: {ex.Message}"})
            End Try
            Return result
        End Function

    End Class

End Namespace