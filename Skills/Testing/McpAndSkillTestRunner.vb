Option Strict On
Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Threading.Tasks

Namespace Skills.Testing

    ''' <summary>
    ''' [STUB] 统一编排器:MCP工具 3 级体检 + Skill 4 级自动化回归
    ''' 当前为占位类(确保 ThisAddIn.vb 引用合法),后续填充真实实现。
    ''' </summary>
    Public Class McpAndSkillTestRunner

        ''' <summary>进度通知 POCO(不用 ValueTuple,避免 Option Strict 后期绑定)</summary>
        Public Class RegressionProgress
            Public Property Percent As Double
            Public Property Stage As String
            Public Property Message As String
        End Class

        ''' <summary>报告输出 DTO(占位)</summary>
        Public Class CombinedTestReport
            Public Property SummaryLine As String
            Public Property CombinedReportMarkdownPath As String
            Public Property McpReportPath As String
            Public Property SkillReportPath As String
            Public Property GeneratedAt As DateTime = DateTime.Now
        End Class

        ''' <summary>
        ''' [STUB] 统一入口:后台异步执行 MCP 工具全量体检 + Skill 全量回归
        ''' 当前版本直接返回空报告(不报错),等 McpToolCatalogValidator.vb 独立文件实现后再填充。
        ''' </summary>
        Public Shared Async Function RunAllAsync(projectRoot As String,
                                                  reportDir As String,
                                                  minSkillLevel As SkillTestLevel,
                                                  maxSkillLevel As SkillTestLevel,
                                                  Optional progress As IProgress(Of RegressionProgress) = Nothing) As Task(Of CombinedTestReport)
            Await Task.Delay(10)

            Directory.CreateDirectory(reportDir)
            Dim mdPath As String = Path.Combine(reportDir, "CombinedReport_Stub.md")
            File.WriteAllText(mdPath, "# [STUB] MCP + Skill 联合测试报告" & Environment.NewLine &
                                      "真实实现待 McpToolCatalogValidator.vb 实现后填充。" & Environment.NewLine &
                                      "生成时间:" & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))

            Return New CombinedTestReport With {
                .SummaryLine = "[STUB] 0/0 通过(占位实现,待填充)",
                .CombinedReportMarkdownPath = mdPath,
                .McpReportPath = "",
                .SkillReportPath = ""
            }
        End Function

    End Class

End Namespace