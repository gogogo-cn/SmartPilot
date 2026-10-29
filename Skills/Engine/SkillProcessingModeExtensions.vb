Imports System.Runtime.CompilerServices

Namespace Skills

    ''' <summary>
    ''' SkillProcessingMode 枚举扩展方法（路由收敛：GetProcessingModeName 单一事实源）
    ''' 收编 SkillExecutionRouter.vb 与 SkillManager.vb 中重复的 Private GetProcessingModeName。
    ''' </summary>
    Public Module SkillProcessingModeExtensions

        ''' <summary>
        ''' 将 SkillProcessingMode 枚举转换为显示名称
        ''' </summary>
        <Extension>
        Public Function ToDisplayName(mode As SkillProcessingMode) As String
            Select Case mode
                Case SkillProcessingMode.VBACode : Return "VBACode"
                Case SkillProcessingMode.MultimodalLLM : Return "MultimodalLLM"
                Case SkillProcessingMode.Hybrid : Return "Hybrid"
                Case SkillProcessingMode.PythonCode : Return "PythonCode"
                Case SkillProcessingMode.LocalOCR : Return "LocalOCR"
                Case SkillProcessingMode.SkillSpec : Return "SkillSpec"
                Case Else : Return mode.ToString()
            End Select
        End Function

    End Module

End Namespace
