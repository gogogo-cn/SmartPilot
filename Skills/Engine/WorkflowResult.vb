' ============================================================================
' WorkflowResult - 工作流执行结果（纯 POCO）
'
' 【为什么拆出来】
' 本类原本定义在 WorkflowEngine.vb 内部，而 WorkflowEngine.vb 强依赖
' Microsoft.Office.Interop.Excel（与 SkillManager / SkillExecutionRouter 等 16 个
' 技能引擎文件一样）。结果是：只要 ParseScriptManager 用到 WorkflowResult，
' 就会把整套 Excel 耦合的技能引擎级联进来，PPT/Word 根本无法承载。
'
' 本类只用到 Dictionary / List / TimeSpan，与 Excel 完全无关，因此拆成独立文件后：
'   - ParseScriptManager 仅需 SkillDefinition + 本类即可进入公共核心库 SmartPilot.Core，
'     供 Excel / PowerPoint / Word 三个宿主共用；
'   - WorkflowEngine 不受影响（两者同属 Skills 命名空间，编译期照常解析）。
' ============================================================================
Option Strict Off
Option Explicit On

Imports System
Imports System.Collections.Generic

Namespace Skills

    Public Class WorkflowResult
        Public Property Success As Boolean
        Public Property Outputs As Dictionary(Of String, Object)
        Public Property ExecutionLog As List(Of String)
        Public Property ErrorMessage As String
        Public Property TotalExecutionTime As TimeSpan
        Public Property Metadata As Dictionary(Of String, Object)

        Public Property Message As String
            Get
                If ErrorMessage IsNot Nothing AndAlso ErrorMessage.Length > 0 Then
                    Return ErrorMessage
                End If
                If Outputs IsNot Nothing AndAlso Outputs.ContainsKey("message") Then
                    Return Outputs("message")?.ToString()
                End If
                Return If(Success, "执行成功", "执行失败")
            End Get
            Set(value As String)
                If Outputs Is Nothing Then
                    Outputs = New Dictionary(Of String, Object)
                End If
                Outputs("message") = value
            End Set
        End Property

        Public Sub New()
            Outputs = New Dictionary(Of String, Object)
            ExecutionLog = New List(Of String)
            Metadata = New Dictionary(Of String, Object)()
        End Sub
    End Class

End Namespace
