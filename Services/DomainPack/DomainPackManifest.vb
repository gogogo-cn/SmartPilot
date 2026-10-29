Imports System.Collections.Generic

Namespace Services.DomainPack

    ''' <summary>
    ''' 专项识别模块（DomainPack）的清单模型。
    '''
    ''' 背景：机械制图原先被硬编码在 UI 复选框与 OcrServiceManager.IsMechanicalMode 的
    ''' 字符串比对里，新增建筑图之类的专项每改一处都要动此项目源码并重新编译。
    '''
    ''' 现在每个专项 = 一个 .md 清单文件 + 可选的执行体（JS / Python / 内置 VB）：
    '''   pipeline: text    → 走现有 ParseScript 引擎（前端 JS），只能处理 OCR 文本
    '''   pipeline: python  → 调用同名 Python 模块，可做几何 / 符号 / 图像处理
    '''   pipeline: builtin → 编译在本程序集里的识别器（如机械制图 GDT），零回归复用
    '''
    ''' 清单放在 %AppData%\ExcelAIVSTOAddIn\DomainPacks 下即被识别，无需重新编译宿主。
    ''' </summary>
    Public Enum DomainPackPipeline
        Text = 0
        Python = 1
        BuiltIn = 2
    End Enum

    ''' <summary>模块自带的开关项，安装后由前端自动渲染（不必再改 HTML）。</summary>
    Public Class DomainPackOption

        ''' <summary>存取值用的键，在模块内唯一。</summary>
        Public Property Key As String

        ''' <summary>界面上显示的中文名。</summary>
        Public Property Label As String

        ''' <summary>bool（勾选）| select（下拉）| number（数字）。</summary>
        Public Property Type As String

        ''' <summary>默认值字符串；bool 用 true/false。</summary>
        Public Property DefaultValue As String

        ''' <summary>type=select 时的候选项。</summary>
        Public Property Values As New List(Of String)()
    End Class

    Public Class DomainPackManifest

        ''' <summary>唯一标识，小写字母数字与下划线，等于文件名。</summary>
        Public Property Id As String

        Public Property Name As String
        Public Property Description As String
        Public Property Icon As String
        Public Property Category As String
        Public Property Version As String
        Public Property Author As String

        Public Property Pipeline As DomainPackPipeline = DomainPackPipeline.Text

        ''' <summary>
        ''' 执行体引用：builtin 填识别器代号（GdtSymbolRecognizer）；
        ''' python 填脚本文件（相对模块目录）；text 可留空。
        ''' </summary>
        Public Property ModuleRef As String

        ''' <summary>python 流水线需要的 pip 包（逗号分隔写在清单里）。</summary>
        Public Property Requires As New List(Of String)()

        ''' <summary>text 流水线对应的 ParseScript id，留空则用同 id。</summary>
        Public Property ParseScriptId As String

        Public Property Options As New List(Of DomainPackOption)()

        ''' <summary>清单文件绝对路径；内置模块为空。</summary>
        Public Property FilePath As String

        ''' <summary>内置模块（随程序发布，不可卸载）。</summary>
        Public Property IsBuiltIn As Boolean

        ''' <summary>是否被用户停用。</summary>
        Public Property Enabled As Boolean = True

        ''' <summary>清单解析失败时的原因，界面上提示用。</summary>
        Public Property ErrorText As String

        Public Function PipelineName() As String
            Select Case Pipeline
                Case DomainPackPipeline.BuiltIn
                    Return "builtin"
                Case DomainPackPipeline.Python
                    Return "python"
                Case Else
                    Return "text"
            End Select
        End Function
    End Class
End Namespace
