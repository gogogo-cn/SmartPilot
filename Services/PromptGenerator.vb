Option Strict On
Option Infer On
Imports System.IO
Imports System.Threading
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services

    ''' <summary>
    ''' 提示词生成器 - 根据思考模式动态生成提示词
    ''' 【DI 双轨模式】Instance 属性优先从 AppServiceProvider 解析，未初始化时回退到经典 Lazy 单例
    ''' </summary>
    Public Class PromptGenerator

        Private Shared ReadOnly _legacyInstance As New Lazy(Of PromptGenerator)(Function() New PromptGenerator(), System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)
        Private _promptsConfig As JObject

        ''' <summary>
        ''' 全局访问入口。DI 容器已配置则从容器取（便于 Mock 和生命周期管理），否则回退到 Lazy 单例。
        ''' </summary>
        Public Shared ReadOnly Property Instance As PromptGenerator
            Get
                If Composition.AppServiceProvider.IsConfigured Then
                    Dim fromDi = Composition.AppServiceProvider.Instance.TryGet(Of PromptGenerator)()
                    If fromDi IsNot Nothing Then Return fromDi
                End If
                Return _legacyInstance.Value
            End Get
        End Property

        ''' <summary>
        ''' 公共构造函数 — 支持 DI 容器注入和单元测试时直接 New
        ''' （原构造函数为 Private，DI 迁移后改为 Public 以便 ActivatorUtilities.CreateInstance）
        ''' </summary>
        Public Sub New()
            LoadPromptsConfig()
        End Sub

        ''' <summary>
        ''' 加载提示词配置（从 Markdown 文件）
        ''' </summary>
        Private Sub LoadPromptsConfig()
            Try
                Dim configPath As String = Path.Combine(
                    Utils.PathHelper.ConfigDirectory,
                    "ThinkingPrompts.md"
                )

                If File.Exists(configPath) Then
                    Dim markdown As String = File.ReadAllText(configPath)
                    _promptsConfig = ParseMarkdownPrompts(markdown)
                    System.Diagnostics.Debug.WriteLine("[PromptGenerator] 成功加载 ThinkingPrompts.md")
                Else
                    System.Diagnostics.Debug.WriteLine("[PromptGenerator] 警告: ThinkingPrompts.md 不存在，使用默认配置")
                    _promptsConfig = CreateDefaultPromptsConfig()
                End If

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[PromptGenerator] 加载提示词配置失败: " & ex.Message)
                _promptsConfig = CreateDefaultPromptsConfig()
            End Try
        End Sub

        ''' <summary>
        ''' 解析 Markdown 格式的提示词配置
        ''' </summary>
        Private Function ParseMarkdownPrompts(markdown As String) As JObject
            Dim config As New JObject()
            Dim thinkingModes As New JObject()

            ' 解析快速模式
            Dim quickMatch As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(
                markdown,
                "### 1\. 快速模式.*?\*\*系统提示词\*\*:[^\n]*\n```([\s\S]*?)```",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase
            )
            If quickMatch.Success Then
                thinkingModes("quick") = New JObject From {
                    {"systemPrompt", quickMatch.Groups(1).Value.Trim()},
                    {"responseFormat", "direct"},
                    {"maxSteps", 5},
                    {"riskCheck", False},
                    {"userConfirm", False}
                }
            End If

            ' 解析标准模式
            Dim standardMatch As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(
                markdown,
                "### 2\. 标准模式.*?\*\*系统提示词\*\*:[^\n]*\n```([\s\S]*?)```",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase
            )
            If standardMatch.Success Then
                thinkingModes("standard") = New JObject From {
                    {"systemPrompt", standardMatch.Groups(1).Value.Trim()},
                    {"responseFormat", "thinking + steps"},
                    {"maxSteps", 15},
                    {"riskCheck", True},
                    {"userConfirm", False}
                }
            End If

            ' 解析深度模式
            Dim deepMatch As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(
                markdown,
                "### 3\. 深度模式.*?\*\*系统提示词\*\*:[^\n]*\n```([\s\S]*?)```",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase
            )
            If deepMatch.Success Then
                thinkingModes("deep") = New JObject From {
                    {"systemPrompt", deepMatch.Groups(1).Value.Trim()},
                    {"responseFormat", "thinking + detailed_steps + validation"},
                    {"maxSteps", 50},
                    {"riskCheck", True},
                    {"userConfirm", True}
                }
            End If

            ' 解析专家模式
            Dim expertMatch As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(
                markdown,
                "### 4\. 专家模式.*?\*\*系统提示词\*\*:[^\n]*\n```([\s\S]*?)```",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase
            )
            If expertMatch.Success Then
                thinkingModes("expert") = New JObject From {
                    {"systemPrompt", expertMatch.Groups(1).Value.Trim()},
                    {"responseFormat", "editable_thinking + simulation + user_confirm"},
                    {"maxSteps", 50},
                    {"riskCheck", True},
                    {"userConfirm", True},
                    {"editableThinking", True}
                }
            End If

            ' 解析核心能力
            Dim capabilitiesMatch As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(
                markdown,
                "## 核心能力([\s\S]*?)---",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase
            )
            If capabilitiesMatch.Success Then
                config("capabilities") = capabilitiesMatch.Groups(1).Value.Trim()
            End If

            config("thinkingModes") = thinkingModes
            config("version") = "3.0"
            config("source") = "ThinkingPrompts.md"

            Return config
        End Function

        ''' <summary>
        ''' 创建默认提示词配置
        ''' </summary>
        Private Function CreateDefaultPromptsConfig() As JObject
            Dim config As New JObject()

            Dim thinkingModes As New JObject()

            ' v3.0 新模式：Quick
            thinkingModes("quick") = New JObject From {
                {"systemPrompt", "【快速模式】你是一个Excel智能助手。用户要求简单操作，直接执行即可。" & vbCrLf & "最多5步，不需要风险检查。"},
                {"responseFormat", "direct"}
            }

            ' v3.0 新模式：Standard
            thinkingModes("standard") = New JObject From {
                {"systemPrompt", "【标准模式】你是一个Excel智能助手。用户要求多步骤操作，需要平衡效率与准确性。" & vbCrLf & "最多15步，需要风险检查。"},
                {"responseFormat", "thinking + steps"}
            }

            ' v3.0 新模式：Deep
            thinkingModes("deep") = New JObject From {
                {"systemPrompt", "【深度模式】你是一个Excel智能助手。用户要求复杂任务，需要系统化分析。" & vbCrLf & "最多50步，全面风险检查。"},
                {"responseFormat", "thinking + detailed_steps + validation"}
            }

            ' v3.0 新模式：Expert
            thinkingModes("expert") = New JObject From {
                {"systemPrompt", "【专家模式】你是一个Excel智能助手。用户是高级用户，可以介入思考过程。" & vbCrLf & "思考过程可编辑，用户可修改步骤。"},
                {"responseFormat", "editable_thinking + simulation + user_confirm"}
            }

            config("thinkingModes") = thinkingModes

            Return config
        End Function

        ''' <summary>
        ''' 根据思考模式生成完整提示词
        ''' </summary>
        Public Function GeneratePrompt(
            mode As ThinkingMode,
            userMessage As String,
            Optional hasImage As Boolean = False,
            Optional contextHistory As String = ""
        ) As String

            ' 获取基础提示词
            Dim basePrompt As String = GetThinkingModePrompt(mode)

            ' 构建完整提示词
            Dim promptParts As New List(Of String)()

            ' 1. 系统提示
            promptParts.Add(basePrompt)

            ' 2. 上下文信息(如果有)
            ' 【修复】防御性 null 检查，防止 Substring 抛出 NullReferenceException
            If Not String.IsNullOrEmpty(contextHistory) Then
                promptParts.Add(vbCrLf & "---" & vbCrLf)
                promptParts.Add("对话历史:")
                Dim maxLen As Integer = Math.Min(2000, contextHistory.Length)
                promptParts.Add(contextHistory.Substring(0, maxLen))
            End If

            ' 3. 图片提示(如果有)
            If hasImage Then
                promptParts.Add(vbCrLf & "---" & vbCrLf)
                promptParts.Add("[用户提供了图片,请结合图片内容进行分析]")
            End If

            ' 4. 用户消息
            promptParts.Add(vbCrLf & "---" & vbCrLf)
            promptParts.Add("用户消息: " & userMessage)

            ' 5. 强调格式要求
            promptParts.Add(vbCrLf & "请严格按照上述格式输出,确保包含所有必需的标签。")

            Return String.Join(vbCrLf, promptParts)
        End Function

        ''' <summary>
        ''' 构建系统提示词（ExcelBridge调用）
        ''' </summary>
        Public Function BuildSystemPrompt(excelInfoJson As String, Optional mcpCapabilitiesText As String = "") As String
            ' 精简：MCP定义由工具列表自我说明，不再单独堆砌token
            Dim basePrompt As String = "你是一个专业的Excel AI 助手，具备以下核心能力:" & vbCrLf & vbCrLf &
                "**1. Excel 操作能力**" & vbCrLf &
                "- 读写Excel单元格、创建/删除工作表、格式化数据、生成图表、排序筛选" & vbCrLf &
                "- 通过VBA代码执行复杂的数据处理和自动化操作" & vbCrLf & vbCrLf

            If Not String.IsNullOrEmpty(mcpCapabilitiesText) Then
                basePrompt &= "**2. MCP 可用工具**" & vbCrLf & mcpCapabilitiesText & vbCrLf
            Else
                basePrompt &= "**2. MCP 工具**" & vbCrLf &
                    "- 当前未检测到可用的MCP工具" & vbCrLf
            End If

            basePrompt &= vbCrLf & "**3. 数据处理能力**" & vbCrLf &
                "- 批量处理多个文件和大量数据" & vbCrLf &
                "- 将结果写入当前工作表的单元格中供用户查看" & vbCrLf &
                "- 当用户问'有多少MCP工具'时，根据上方第2节的工具列表回答，不要编造或误解MCP含义"

            If Not String.IsNullOrEmpty(excelInfoJson) AndAlso excelInfoJson <> "{}" Then
                basePrompt &= vbCrLf & vbCrLf & "当前Excel环境: " & excelInfoJson
            End If

            Return basePrompt
        End Function

        ''' <summary>
        ''' 构建用户提示词（ExcelBridge调用）
        ''' </summary>
        Public Function BuildUserPrompt(
            message As String,
            chatHistoryJson As String,
            excelInfoJson As String,
            selectedCellsJson As String
        ) As String
            Dim userPrompt As String = message

            If Not String.IsNullOrEmpty(chatHistoryJson) AndAlso chatHistoryJson <> "[]" Then
                userPrompt &= vbCrLf & vbCrLf & "对话历史: " & chatHistoryJson
            End If

            If Not String.IsNullOrEmpty(selectedCellsJson) AndAlso selectedCellsJson <> "{}" Then
                userPrompt &= vbCrLf & vbCrLf & "选中的单元格: " & selectedCellsJson
            End If

            Return userPrompt
        End Function

        ''' <summary>
        ''' 获取思考模式基础提示词
        ''' </summary>
        Private Function GetThinkingModePrompt(mode As ThinkingMode) As String
            Try
                Dim modeName As String = mode.ToString()
                Dim thinkingModes As JToken = _promptsConfig("thinkingModes")
                If thinkingModes Is Nothing Then
                    Return "你是一个Excel AI助手。请帮助用户解决问题。"
                End If
                Dim modeConfig As JToken = thinkingModes(modeName)

                If modeConfig IsNot Nothing Then
                    Dim prompt As JToken = modeConfig("systemPrompt")
                    If prompt IsNot Nothing Then
                        Return prompt.ToString()
                    End If
                End If

                ' 回退到默认（chainOfThought 不存在则用 standard）
                Dim fallback As JToken = thinkingModes("chainOfThought")
                If fallback Is Nothing Then fallback = thinkingModes("standard")
                If fallback IsNot Nothing Then
                    Dim fbPrompt As JToken = fallback("systemPrompt")
                    If fbPrompt IsNot Nothing Then Return fbPrompt.ToString()
                End If
                Return "你是一个Excel AI助手。请帮助用户解决问题。"

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("获取提示词失败: " & ex.Message)
                Return "你是一个Excel AI助手。请帮助用户解决问题。"
            End Try
        End Function

        ''' <summary>
        ''' 获取响应格式说明
        ''' </summary>
        Public Function GetResponseFormat(mode As ThinkingMode) As String
            Try
                Dim modeName As String = mode.ToString()
                Dim thinkingModes As JToken = _promptsConfig("thinkingModes")
                If thinkingModes Is Nothing Then Return "thinking + answer"
                Dim modeToken As JToken = thinkingModes(modeName)
                If modeToken Is Nothing Then Return "thinking + answer"
                Dim formatToken As JToken = modeToken("responseFormat")
                If formatToken Is Nothing Then Return "thinking + answer"
                Return formatToken.ToString()
            Catch ex As Exception  ' 
                Return "thinking + answer"
            End Try
        End Function

    End Class

End Namespace