Option Strict On
Option Explicit On
Option Infer On

Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services

    ''' <summary>
    ''' 思考模式枚举（v3.0）
    ''' </summary>
    Public Enum ThinkingMode
        Quick       ' 快速模式（单操作，<5步，无风险检查）
        Standard    ' 标准模式（多操作组合，5-15步，有风险检查）
        Deep        ' 深度模式（复杂任务，15-50步，系统化分析）
        Expert      ' 专家模式（用户可介入，可编辑思考过程）
    End Enum

    ''' <summary>
    ''' v3.0决策JSON模型（模型决策层输出）
    ''' </summary>
    Public Class DecisionJson
        Public Property mode As String          ' quick|standard|deep|expert
        Public Property skill As String         ' 从当前可用 Skill 列表中选择的 skill 名
        Public Property complexity As String    ' simple|medium|complex
        Public Property parameters As JObject   ' 动态参数

        Public Sub New()
            parameters = New JObject()
        End Sub
    End Class

    ''' <summary>
    ''' 思考步骤（v3.0模板化）
    ''' 对齐三阶段消息协议：Thinking/Step/Final
    ''' </summary>
    Public Class ThinkingStep
        Public Property StepId As Integer
        Public Property Phase As String         ' 理解/约束/方案/执行/验证（后端）| analyze/plan/decide/reflect（前端协议）
        Public Property StepAction As String    ' 动作名称
        Public Property Detail As String
        Public Property Status As String        ' pending/running/completed/failed（对齐前端协议）
        Public Property Result As String
        Public Property Timestamp As DateTime
        
        ''' <summary>
        ''' 映射到前端协议的 phase 标签
        ''' </summary>
        Public Function GetProtocolPhase() As String
            Select Case Phase
                Case "理解", "分析"
                    Return "analyze"
                Case "方案", "计划"
                    Return "plan"
                Case "约束", "决策", "选择"
                    Return "decide"
                Case "验证", "总结", "反思"
                    Return "reflect"
                Case Else
                    Return "analyze"
            End Select
        End Function
        
        ''' <summary>
        ''' 转换为前端协议的 Step 消息格式
        ''' </summary>
        Public Function ToProtocolStep() As Dictionary(Of String, Object)
            Return New Dictionary(Of String, Object) From {
                {"type", "step"},
                {"stepId", "vb-step-" & StepId},
                {"toolName", Phase & If(Not String.IsNullOrEmpty(StepAction), "-" & StepAction, "")},
                {"status", Status},
                {"detail", Detail},
                {"result", If(String.IsNullOrEmpty(Result), Nothing, Result)}
            }
        End Function
    End Class

    ''' <summary>
    ''' 执行动作
    ''' </summary>
    Public Class ActionItem
        Public Property ActionType As String  ' excel_operation/code_generation/data_extraction
        Public Property Parameters As Dictionary(Of String, Object)
        Public Property Confidence As Double
        Public Property Description As String

        Public Sub New()
            Parameters = New Dictionary(Of String, Object)
        End Sub
    End Class

    ''' <summary>
    ''' 思考响应
    ''' </summary>
    Public Class ThinkingResponse
        Public Property Mode As ThinkingMode
        Public Property ThinkingSteps As List(Of ThinkingStep)
        Public Property FinalAnswer As String
        Public Property Actions As List(Of ActionItem)
        Public Property Confidence As Double
        Public Property ProcessingTime As TimeSpan
        Public Property ConversationId As String

        Public Sub New()
            ThinkingSteps = New List(Of ThinkingStep)
            Actions = New List(Of ActionItem)
        End Sub
    End Class

    ''' <summary>
    ''' 任务特征档案
    ''' </summary>
    Public Class TaskProfile
        Public Property OriginalMessage As String   ' 原始用户消息（用于完整复杂度评估）
        Public Property ComplexityScore As Integer  ' 0-10
        Public Property TaskType As String          ' greeting/question/code_generation/data_extraction/multi_step/creative
        Public Property RequiresAction As Boolean
        Public Property HasImage As Boolean
        Public Property MessageLength As Integer
        Public Property Keywords As List(Of String)

        Public Sub New()
            Keywords = New List(Of String)
        End Sub
    End Class

#Region "思考步骤结果模型"

    ''' <summary>
    ''' 思考步骤结果基类
    ''' </summary>
    Public MustInherit Class StepResult
        Public MustOverride Function GetSummary() As String

        Public Overrides Function ToString() As String
            Return JsonConvert.SerializeObject(Me)
        End Function
    End Class

    ''' <summary>
    ''' 理解阶段结果
    ''' </summary>
    Public Class UnderstandingResult
        Inherits StepResult

        Public Property TaskType As String
        Public Property Summary As String
        Public Property Analysis As String

        Public Overrides Function GetSummary() As String
            Return If(Not String.IsNullOrEmpty(Analysis), Analysis, Summary)
        End Function
    End Class

    ''' <summary>
    ''' 约束阶段结果
    ''' </summary>
    Public Class ConstraintResult
        Inherits StepResult

        Public Property CanProceed As Boolean
        Public Property HasRisk As Boolean
        Public Property Warning As String

        Public Overrides Function GetSummary() As String
            If HasRisk Then
                Return If(Not String.IsNullOrEmpty(Warning), Warning, "存在风险")
            End If
            Return If(CanProceed, "约束检查通过", "存在阻塞性约束")
        End Function
    End Class

    ''' <summary>
    ''' 方案阶段结果
    ''' </summary>
    Public Class SolutionResult
        Inherits StepResult

        Public Property Plan As String
        Public Property Steps As String

        Public Overrides Function GetSummary() As String
            Return If(Not String.IsNullOrEmpty(Plan), Plan, Steps)
        End Function
    End Class

    ''' <summary>
    ''' 执行阶段结果
    ''' </summary>
    Public Class ExecutionResult
        Inherits StepResult

        Public Property Success As Boolean
        Public Property Message As String
        Public Property Data As Object

        Public Overrides Function GetSummary() As String
            Return Message
        End Function
    End Class

    ''' <summary>
    ''' 验证阶段结果
    ''' </summary>
    Public Class VerificationResult
        Inherits StepResult

        Public Property Success As Boolean
        Public Property Tips As String

        Public Overrides Function GetSummary() As String
            Return Tips
        End Function
    End Class

#End Region

    ''' <summary>
    ''' 思考引擎 - 核心决策与思考流程管理
    ''' </summary>
    Public Class ThinkingEngine

        Private _promptGenerator As PromptGenerator
        Private _llmEngine As MultimodalLLMEngine
        Private _excelBridge As ExcelAIVSTOAddIn.ExcelBridge
        Private _actionEngine As ActionEngine
        Private _memoryService As ConversationMemoryService
        Private _skillRecognizer As Skills.SkillRecognizer   ' 【规则兜底】LLM 决策失败/无效时的规则识别器
        Private Shared ReadOnly _instance As New Lazy(Of ThinkingEngine)(Function() New ThinkingEngine(), System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        Public Shared ReadOnly Property Instance As ThinkingEngine
            Get
                Return _instance.Value
            End Get
        End Property

        Private Sub New()
            _promptGenerator = PromptGenerator.Instance
            _llmEngine = MultimodalLLMEngine.SharedInstance
            _memoryService = ConversationMemoryService.Instance
        End Sub

        Public Sub Initialize(bridge As ExcelAIVSTOAddIn.ExcelBridge)
            _excelBridge = bridge
            If _actionEngine Is Nothing Then
                _actionEngine = ActionEngine.Instance
                _actionEngine.Initialize(bridge)
            End If
            ' 启动时同步加载 LLM 配置，否则 MultimodalLLMEngine._config 恒为空，
            ' 导致 Describe/图表等走 LLM 的请求报"LLM Engine 未初始化配置"。
            Try
                _llmEngine.LoadFromConfiguration(Utils.ConfigurationManager.Instance)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[ThinkingEngine] LLM 配置加载失败: " & ex.Message)
            End Try
            System.Diagnostics.Debug.WriteLine("[ThinkingEngine] 会话记忆服务已连接, ID: " & _memoryService.CurrentConversationId)
        End Sub

        Public ReadOnly Property Memory As ConversationMemoryService
            Get
                Return _memoryService
            End Get
        End Property

        ''' <summary>
        ''' 主入口:分析并执行思考流程（v3.0架构）
        ''' 流程：用户输入 → 模型决策（输出JSON） → 代码加载思考模板 → 按模板执行
        ''' </summary>
        Public Async Function AnalyzeAndThink(
            userMessage As String,
            Optional hasImage As Boolean = False,
            Optional contextHistory As String = "",
            Optional progress As IProgress(Of ThinkingStep) = Nothing,
            Optional externalMode As String = "AUTO"
        ) As Task(Of ThinkingResponse)

            Dim stopwatch = System.Diagnostics.Stopwatch.StartNew()
            Dim mode As ThinkingMode = ThinkingMode.Standard
            Dim taskProfile As TaskProfile = Nothing
            Dim decisionJson As DecisionJson = Nothing

            Try
                If (String.IsNullOrWhiteSpace(externalMode) OrElse externalMode.ToUpper() = "AUTO") AndAlso
                   Not String.IsNullOrWhiteSpace(_memoryService.Preferences.DefaultThinkingMode) AndAlso
                   _memoryService.Preferences.DefaultThinkingMode.ToUpper() <> "AUTO" Then
                    externalMode = _memoryService.Preferences.DefaultThinkingMode
                    System.Diagnostics.Debug.WriteLine("[ThinkingEngine] 从偏好恢复模式: " & externalMode)
                End If

                If Not String.IsNullOrWhiteSpace(externalMode) AndAlso externalMode.ToUpper() <> "AUTO" Then
                    mode = ParseThinkingMode(externalMode.ToLower())
                    System.Diagnostics.Debug.WriteLine("[ThinkingEngine] 使用模式: " & mode.ToString())
                End If

                Dim memoryContext As String = _memoryService.GetContextSummary(1500)
                Dim mergedContext As String = contextHistory
                If Not String.IsNullOrWhiteSpace(memoryContext) Then
                    If Not String.IsNullOrWhiteSpace(mergedContext) Then
                        mergedContext = memoryContext & "; " & mergedContext
                    Else
                        mergedContext = memoryContext
                    End If
                End If

                Dim effectiveMessage As String = userMessage
                If Not String.IsNullOrWhiteSpace(_memoryService.Preferences.CustomPromptPrefix) Then
                    effectiveMessage = _memoryService.Preferences.CustomPromptPrefix & vbCrLf & userMessage
                End If

                ' ═══════════════════════════════════════════════════════════
                ' 前置短路：用 McpCapabilityRegistry.Match() 自动匹配 MCP 工具
                ' 加新工具零维护 — 只要注册 McpTool 带 keywords 就能自动被 AI 驱动
                ' ═══════════════════════════════════════════════════════════
                Dim autoMcp = Await TryAutoMcpRoute(effectiveMessage)
                If autoMcp IsNot Nothing Then
                    autoMcp.Mode = ThinkingMode.Quick
                    autoMcp.ProcessingTime = stopwatch.Elapsed
                    autoMcp.ConversationId = _memoryService.CurrentConversationId
                    Return autoMcp
                End If

                taskProfile = AnalyzeTaskProfile(effectiveMessage, hasImage, mergedContext)

                If String.IsNullOrWhiteSpace(externalMode) OrElse externalMode.ToUpper() = "AUTO" Then
                    mode = SelectThinkingMode(taskProfile)
                End If

                Dim decisionPrompt = BuildDecisionPrompt(effectiveMessage, mode)

                ' 【规则兜底】LLM 决策失败（异常/超时/解析失败）不直接抛给 ErrorRecovery，先尝试规则识别器兜底
                Try
                    decisionJson = Await GetModelDecision(decisionPrompt, hasImage)
                Catch exDec As Exception
                    System.Diagnostics.Debug.WriteLine("[ThinkingEngine] LLM 决策失败，尝试规则识别兜底: " & exDec.Message)
                    decisionJson = Nothing
                End Try

                If decisionJson IsNot Nothing Then
                    If String.IsNullOrWhiteSpace(decisionJson.mode) OrElse
                       Not {"quick", "standard", "deep", "expert"}.Contains(decisionJson.mode.ToLower()) Then
                        decisionJson.mode = mode.ToString().ToLower()
                    End If

                    ' 【规则兜底】LLM 选出的 skill 无效（为空/幻觉/不在注册列表）→ 规则识别器兜底
                    If Not IsValidRegisteredSkill(decisionJson.skill) Then
                        Dim ruleSkill = TryRuleSkillFallback(effectiveMessage, mergedContext)
                        If Not String.IsNullOrWhiteSpace(ruleSkill) Then
                            decisionJson.skill = ruleSkill
                            System.Diagnostics.Debug.WriteLine("[ThinkingEngine] 规则识别兜底命中 skill: " & ruleSkill)
                        Else
                            System.Diagnostics.Debug.WriteLine("[ThinkingEngine] LLM 决策 skill 无效且规则未命中，走通用模板: " & decisionJson.skill)
                        End If
                    End If
                Else
                    ' 【规则兜底】LLM 决策失败 → 用规则识别构造最小决策
                    Dim ruleSkill = TryRuleSkillFallback(effectiveMessage, mergedContext)
                    decisionJson = New DecisionJson With {
                        .mode = mode.ToString().ToLower(),
                        .complexity = MapComplexityScore(taskProfile),
                        .skill = If(ruleSkill, "")
                    }
                    If Not String.IsNullOrWhiteSpace(ruleSkill) Then
                        System.Diagnostics.Debug.WriteLine("[ThinkingEngine] LLM 决策失败，规则识别兜底命中 skill: " & ruleSkill)
                    End If
                End If

                mode = ParseThinkingMode(decisionJson.mode)

                Dim template = LoadThinkingTemplate(decisionJson, taskProfile)
                Dim response = Await ExecuteTemplate(decisionJson, template, progress)
                response.Mode = mode
                response.ProcessingTime = stopwatch.Elapsed
                response.ConversationId = _memoryService.CurrentConversationId

                Try
                    _memoryService.AddMessage("user", userMessage, New Dictionary(Of String, Object) From {
                        {"hasImage", hasImage},
                        {"mode", mode.ToString()},
                        {"complexity", taskProfile.ComplexityScore}
                    })
                    If Not String.IsNullOrWhiteSpace(response.FinalAnswer) Then
                        _memoryService.AddMessage("assistant", response.FinalAnswer, New Dictionary(Of String, Object) From {
                            {"skill", If(decisionJson IsNot Nothing AndAlso decisionJson.skill IsNot Nothing, decisionJson.skill, "")},
                            {"confidence", response.Confidence},
                            {"steps", response.ThinkingSteps.Count}
                        })
                    End If
                Catch exMem As Exception
                    System.Diagnostics.Debug.WriteLine("[ThinkingEngine] 写入记忆失败: " & exMem.Message)
                End Try

                Return response

            Catch ex As Exception
                Dim recoveryContext = New Dictionary(Of String, Object) From {
                    {"mode", mode},
                    {"taskProfile", taskProfile}
                }

                Dim recoveryResult = ErrorRecovery.Instance.Handle(ex, recoveryContext).Result

                If CBool(recoveryResult("success")) = True Then
                    Return New ThinkingResponse With {
                        .Mode = ThinkingMode.Standard,
                        .FinalAnswer = recoveryResult("message").ToString(),
                        .Confidence = CDbl(recoveryResult("confidence")),
                        .ProcessingTime = stopwatch.Elapsed,
                        .ConversationId = _memoryService.CurrentConversationId
                    }
                Else
                    Return New ThinkingResponse With {
                        .Mode = mode,
                        .FinalAnswer = "执行思考流程时遇到错误: " & ex.Message,
                        .Confidence = 0.0,
                        .ProcessingTime = stopwatch.Elapsed,
                        .ConversationId = _memoryService.CurrentConversationId
                    }
                End If
            End Try
        End Function

        ''' <summary>
        ''' 执行思考流程（ExcelBridge调用）
        ''' </summary>
        Public Async Function ExecuteThinkingProcess(
            thinkingMode As ThinkingMode,
            taskProfile As TaskProfile,
            systemPrompt As String,
            userPrompt As String
        ) As Task(Of ThinkingResponse)

            Dim stopwatch = System.Diagnostics.Stopwatch.StartNew()

            Try
                ' 0. 将系统提示词注入到 LLM 引擎配置中
                If Not String.IsNullOrEmpty(systemPrompt) Then
                    Dim currentConfig = _llmEngine.GetCurrentConfig()
                    If currentConfig IsNot Nothing Then
                        currentConfig.SystemPrompt = systemPrompt
                        _llmEngine.Initialize(currentConfig)
                        System.Diagnostics.Debug.WriteLine(String.Format("[ThinkingEngine] systemPrompt 已注入 LLM 引擎, 长度: {0}", systemPrompt.Length))
                    Else
                        System.Diagnostics.Debug.WriteLine("[ThinkingEngine] 警告: GetCurrentConfig() 返回 Nothing, 无法注入 systemPrompt")
                    End If
                Else
                    System.Diagnostics.Debug.WriteLine("[ThinkingEngine] 警告: systemPrompt 为空, 未注入")
                End If

                ' 1. 调用LLM执行思考
                Dim llmResponse = Await _llmEngine.ProcessTextAsync(userPrompt)

                ' 2. 检查LLM响应是否成功
                If Not llmResponse.Success Then
                    Return New ThinkingResponse With {
                        .Mode = thinkingMode,
                        .FinalAnswer = "LLM调用失败: " & llmResponse.ErrorMessage,
                        .Confidence = 0.0,
                        .ProcessingTime = stopwatch.Elapsed
                    }
                End If

                ' 3. 解析思考结果
                Dim response = ParseThinkingResponse(llmResponse.Content, thinkingMode)
                response.Mode = thinkingMode
                response.ProcessingTime = stopwatch.Elapsed

                Return response

            Catch ex As Exception
                Return New ThinkingResponse With {
                    .Mode = thinkingMode,
                    .FinalAnswer = "执行思考流程时遇到错误: " & ex.Message,
                    .Confidence = 0.0,
                    .ProcessingTime = stopwatch.Elapsed
                }
            End Try
        End Function

        ''' <summary>
        ''' 分析任务特征
        ''' </summary>
        Public Function AnalyzeTask(
            message As String,
            Optional contextMessages As List(Of String) = Nothing,
            Optional hasImages As Boolean = False,
            Optional taskType As String = "auto"
        ) As TaskProfile

            Dim contextHistory = ""
            If contextMessages IsNot Nothing Then
                contextHistory = String.Join("; ", contextMessages)
            End If

            Return AnalyzeTaskProfile(message, hasImages, contextHistory)
        End Function

        ''' <summary>
        ''' 分析任务特征
        ''' </summary>
        Private Function AnalyzeTaskProfile(
            message As String,
            hasImage As Boolean,
            contextHistory As String
        ) As TaskProfile

            Dim profile As New TaskProfile
            profile.OriginalMessage = message
            profile.MessageLength = message.Length
            profile.HasImage = hasImage

            ' 复杂度评分 (0-10)
            profile.ComplexityScore = CalculateComplexity(message, hasImage, contextHistory)

            ' 任务类型识别
            profile.TaskType = IdentifyTaskType(message)

            ' 是否需要执行操作
            profile.RequiresAction = DetectActionRequirement(message)

            ' 提取关键词
            profile.Keywords = ExtractKeywords(message)

            Return profile
        End Function

        ''' <summary>
        ''' 计算任务复杂度 (0-10)
        ''' </summary>
        Private Function CalculateComplexity(
            message As String,
            hasImage As Boolean,
            contextHistory As String
        ) As Integer

            Dim score As Integer = 0

            ' 因素1: 消息长度
            If message.Length > 50 Then score += 1
            If message.Length > 200 Then score += 2

            ' 因素2: 是否有图片
            If hasImage Then score += 2

            ' 因素3: 复杂关键词
            Dim complexKeywords = {"分析", "优化", "比较", "设计", "规划", "生成", "自动化", "批量"}
            For Each keyword In complexKeywords
                If message.Contains(keyword) Then
                    score += 1
                End If
            Next

            ' 因素4: 上下文长度
            If contextHistory.Length > 1000 Then score += 1
            If contextHistory.Length > 5000 Then score += 1

            Return Math.Min(score, 10)
        End Function

        ''' <summary>
        ''' 识别任务类型
        ''' </summary>
        Private Function IdentifyTaskType(message As String) As String
            Dim lowerMessage = message.ToLower()

            ' 问候
            Dim greetings = {"你好", "您好", "嗨", "hello", "hi", "在吗"}
            If greetings.Any(Function(g) lowerMessage.Contains(g)) Then
                Return "greeting"
            End If

            ' 代码生成
            Dim codeKeywords = {"代码", "vba", "宏", "生成", "编写", "脚本", "文件", "文件夹", "目录", "盘", "列出"}
            If codeKeywords.Any(Function(k) lowerMessage.Contains(k)) Then
                Return "code_generation"
            End If

            ' 数据提取
            ' 数据提取 [已弃用 - 识别类Skill已移除] 移除了"身份证", "名片", "发票", "表格"
            Dim extractionKeywords = {"识别", "提取", "填入", "录入"}
            If extractionKeywords.Any(Function(k) lowerMessage.Contains(k)) Then
                Return "data_extraction"
            End If

            ' 多步骤任务
            Dim multiStepKeywords = {"拆分", "合并", "转换", "处理", "批量", "所有"}
            If multiStepKeywords.Any(Function(k) lowerMessage.Contains(k)) Then
                Return "multi_step"
            End If

            ' 创意/分析
            Dim creativeKeywords = {"如何", "怎么", "优化", "改进", "建议", "分析"}
            If creativeKeywords.Any(Function(k) lowerMessage.Contains(k)) Then
                Return "creative"
            End If

            Return "question"
        End Function

        ''' <summary>
        ''' 检测是否需要执行操作
        ''' </summary>
        Private Function DetectActionRequirement(message As String) As Boolean
            Dim actionKeywords = {
                "填入", "执行", "创建", "删除", "修改", "保存",
                "拆分", "合并", "运行", "生成", "提取", "识别",
                "列出", "读取", "浏览", "导出"
            }
            Return actionKeywords.Any(Function(k) message.Contains(k))
        End Function

        ''' <summary>
        ''' 提取关键词
        ''' </summary>
        Private Function ExtractKeywords(message As String) As List(Of String)
            Dim keywords = New List(Of String)
            ' [已弃用 - 识别类Skill已移除] 移除了"身份证", "名片", "发票", "表格"
            Dim importantWords = {
                "识别", "提取", "数据",
                "拆分", "合并", "VBA", "代码", "宏",
                "Excel", "工作簿", "工作表", "单元格"
            }

            For Each word In importantWords
                If message.Contains(word) Then
                    keywords.Add(word)
                End If
            Next

            Return keywords
        End Function

        ''' <summary>
        ''' 评估任务复杂度（v3.0新逻辑）
        ''' </summary>
        Private Function AssessComplexity(userMessage As String, hasImage As Boolean) As String
            Dim complexKeywords = {"报表", "自动化", "跨表", "透视表", "宏", "VBA", "数据透视", "多表合并", "批量处理"}
            Dim mediumKeywords = {"美化", "计算", "排序", "筛选", "图表", "格式化", "公式", "合计", "分析"}

            ' 短消息（<30字）+ 简单分析关键词 → 降级为 simple
            ' 例如"分析一下选中内容"不应判为 complex
            If userMessage.Length < 30 Then
                Dim simpleAnalysisKeywords = {"分析", "看看", "总结", "概括"}
                If simpleAnalysisKeywords.Any(Function(kw) userMessage.Contains(kw)) Then
                    ' 检查是否有真正的复杂关键词
                    If Not complexKeywords.Any(Function(kw) userMessage.Contains(kw)) Then
                        Return "simple"
                    End If
                End If
            End If

            ' 关键词匹配
            If complexKeywords.Any(Function(kw) userMessage.Contains(kw)) Then
                Return "complex"
            End If
            If mediumKeywords.Any(Function(kw) userMessage.Contains(kw)) Then
                Return "medium"
            End If

            ' 消息长度判断
            If userMessage.Length > 50 Then
                Return "medium"
            End If

            Return "simple"
        End Function

        ''' <summary>
        ''' 选择思考模式（v3.0新逻辑）
        ''' </summary>
        Public Function SelectThinkingMode(profile As TaskProfile) As ThinkingMode
            Dim complexity = AssessComplexity(profile.OriginalMessage, profile.HasImage)

            Select Case complexity
                Case "simple"
                    Return ThinkingMode.Quick
                Case "medium"
                    Return ThinkingMode.Standard
                Case "complex"
                    ' 检查是否需要用户确认
                    If profile.RequiresAction AndAlso profile.ComplexityScore >= 8 Then
                        Return ThinkingMode.Expert
                    Else
                        Return ThinkingMode.Deep
                    End If
                Case Else
                    Return ThinkingMode.Standard
            End Select
        End Function

        ''' <summary>
        ''' 获取模式的System Prompt模板（v3.0）
        ''' 不再强制固定 5 阶段输出格式，而是根据任务实际需要引导模型动态推理。
        ''' </summary>
        Private Function GetModePromptTemplate(mode As ThinkingMode) As String
            Select Case mode
                Case ThinkingMode.Quick
                    Return "【快速模式 - 严格约束】" & vbCrLf &
                           "1. 禁止输出任何思考过程、分析推理、意图猜测" & vbCrLf &
                           "2. 禁止说'嗯，用户发来了...'、'让我分析一下...'等废话" & vbCrLf &
                           "3. 直接执行操作或直接给出结果" & vbCrLf &
                           "4. 如果用户要求分析数据，直接读取数据并给出分析结论" & vbCrLf &
                           "5. 回复不超过3句话"

                Case ThinkingMode.Standard
                    Return "【标准模式】用户请求可能需要多个步骤。" & vbCrLf &
                           "请根据任务实际需要，自行决定需要哪些思考环节（例如理解需求、检查约束、比较方案、规划步骤、评估风险等）。" & vbCrLf &
                           "每一步思考都应服务于准确完成任务，不要生搬硬套固定格式。"

                Case ThinkingMode.Deep
                    Return "【深度模式】用户请求较复杂。" & vbCrLf &
                           "请进行充分、系统化的思考：理解显式和隐含需求，检查约束与风险，必要时对比多个方案，制定详细执行计划。" & vbCrLf &
                           "思考结构应根据任务特点自然展开，避免为了填满固定章节而输出冗余内容。"

                Case ThinkingMode.Expert
                    Return "【专家模式】用户可能需要介入思考过程。" & vbCrLf &
                           "请在深度分析的基础上，在关键决策点给出可选方案或确认点，让用户能够修改或选择。" & vbCrLf &
                           "思考应围绕任务本身展开，不要拘泥于固定模板。"

                Case Else
                    Return ""
            End Select
        End Function

#Region "v3.0 核心方法"

        ''' <summary>
        ''' 构建决策提示词（让模型输出JSON）
        ''' </summary>
        Private Function BuildDecisionPrompt(userInput As String, mode As ThinkingMode) As String
            Dim excelContext = GetExcelContext()
            Dim skillOptions = BuildSkillOptionsSection()

            Dim prompt As String = ""
            prompt &= "你是一个Excel助手决策器。" & vbCrLf & vbCrLf
            
            ' 【关键修复】注入模式提示词，让模型知道当前模式
            Dim modePrompt = GetModePromptTemplate(mode)
            If Not String.IsNullOrEmpty(modePrompt) Then
                prompt &= modePrompt & vbCrLf & vbCrLf
            End If
            
            prompt &= "【任务】" & vbCrLf
            prompt &= "根据用户指令和当前Excel状态，从下方【skill列表】中选择最匹配的一个，输出决策JSON。" & vbCrLf & vbCrLf
            prompt &= "【输出格式】（严格JSON，不要任何解释文字）" & vbCrLf
            prompt &= "{" & vbCrLf
            prompt &= "  ""mode"": ""quick|standard|deep|expert""," & vbCrLf
            prompt &= "  ""skill"": ""<从【可选skill名】中选择一个>""," & vbCrLf
            prompt &= "  ""complexity"": ""simple|medium|complex""," & vbCrLf
            prompt &= "  ""parameters"": {}" & vbCrLf
            prompt &= "}" & vbCrLf & vbCrLf
            prompt &= skillOptions & vbCrLf & vbCrLf
            prompt &= "【parameters提取规则】" & vbCrLf
            prompt &= "- range: 从用户指令或当前选区提取" & vbCrLf
            prompt &= "- operations: 提取具体操作类型" & vbCrLf
            prompt &= "- 只提取明确提到的参数，不要推测" & vbCrLf & vbCrLf
            prompt &= "【示例1】" & vbCrLf
            prompt &= "输入：""把表头加粗""" & vbCrLf
            prompt &= "输出：{""mode"":""quick"",""skill"":""format_table"",""complexity"":""simple"",""parameters"":{""operations"":[""bold""]}}}" & vbCrLf & vbCrLf
            prompt &= "【示例2】" & vbCrLf
            prompt &= "输入：""帮我做一个销售统计表，包含日期、产品、数量、单价、金额、销售员，还要能自动计算金额和合计""" & vbCrLf
            prompt &= "输出：{""mode"":""deep"",""skill"":""format_table"",""complexity"":""complex"",""parameters"":{""operations"":[""create_structure"",""formula"",""format"",""calculate_total""]}}}" & vbCrLf & vbCrLf
            prompt &= "【示例3】" & vbCrLf
            prompt &= "输入：""A2到A10求和""" & vbCrLf
            prompt &= "输出：{""mode"":""quick"",""skill"":""calculate"",""complexity"":""simple"",""parameters"":{""range"":""A2:A10"",""operations"":[""sum""]}}}" & vbCrLf & vbCrLf
            prompt &= "【输入】" & vbCrLf
            prompt &= "用户指令：" & userInput & vbCrLf
            prompt &= "当前选区：" & If(excelContext.ContainsKey("selection"), excelContext("selection")?.ToString(), "") & vbCrLf
            prompt &= "工作表：" & If(excelContext.ContainsKey("sheetName"), excelContext("sheetName")?.ToString(), "") & vbCrLf
            prompt &= "数据特征：" & JsonConvert.SerializeObject(If(excelContext.ContainsKey("dataFeatures"), excelContext("dataFeatures"), Nothing)) & vbCrLf & vbCrLf
            
            ' 【关键修复】快速模式下，强化JSON输出约束
            If mode = ThinkingMode.Quick Then
                prompt &= "【重要】当前为快速模式，请直接输出JSON，不要输出任何思考过程、分析或解释文字。" & vbCrLf
                prompt &= "错误示例：嗯，用户发来了一个指令...（这是错误的！）" & vbCrLf
                prompt &= "正确示例：{""mode"":""quick"",""skill"":""...""}" & vbCrLf & vbCrLf
            End If
            
            prompt &= "【输出】"

            Return prompt
        End Function

        ''' <summary>
        ''' 获取Excel上下文
        ''' 优先从 ExcelBridge 读取真实状态，未注入时返回兜底值。
        ''' </summary>
        Private Function GetExcelContext() As Dictionary(Of String, Object)
            Dim context As New Dictionary(Of String, Object)
            Dim rowCount As Integer = 0
            Dim colCount As Integer = 0
            Dim headers As New List(Of String)
            Dim sampleRows As New List(Of List(Of String))

            Try
                If _excelBridge IsNot Nothing Then
                    context("selection") = _excelBridge.GetSelection()
                    context("sheetName") = _excelBridge.GetActiveSheetName()

                    Dim rangeDataJson = _excelBridge.GetSelectedRangeData()
                    If Not String.IsNullOrEmpty(rangeDataJson) Then
                        Dim rangeData = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(rangeDataJson)
                        If rangeData IsNot Nothing Then
                            If rangeData.ContainsKey("rowCount") Then
                                rowCount = Convert.ToInt32(rangeData("rowCount"))
                            End If
                            If rangeData.ContainsKey("colCount") Then
                                colCount = Convert.ToInt32(rangeData("colCount"))
                            End If

                            ' 提取表头（最多前 20 列，避免提示词过长）
                            If rangeData.ContainsKey("headers") AndAlso TypeOf rangeData("headers") Is JArray Then
                                Dim hdrArray = CType(rangeData("headers"), JArray)
                                headers = hdrArray.Take(20).Select(Function(t) t.ToString()).ToList()
                            End If

                            ' 提取前 3 行样本（最多前 20 列）
                            If rangeData.ContainsKey("rows") AndAlso TypeOf rangeData("rows") Is JArray Then
                                Dim rowsArray = CType(rangeData("rows"), JArray)
                                sampleRows = rowsArray.Take(3).Select(Function(r)
                                    If TypeOf r Is JArray Then
                                        Return CType(r, JArray).Take(20).Select(Function(c) c.ToString()).ToList()
                                    Else
                                        Return New List(Of String)
                                    End If
                                End Function).ToList()
                            End If
                        End If
                    End If
                Else
                    context("selection") = "A1"
                    context("sheetName") = "Sheet1"
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[ThinkingEngine] 获取Excel上下文失败: {ex.Message}")
                context("selection") = "A1"
                context("sheetName") = "Sheet1"
            End Try

            context("dataFeatures") = New Dictionary(Of String, Object) From {
                {"rowCount", rowCount},
                {"colCount", colCount},
                {"headers", headers},
                {"sampleRows", sampleRows}
            }
            Return context
        End Function

        ''' <summary>
        ''' 将 Excel 上下文格式化为提示词文本
        ''' </summary>
        Private Function FormatExcelContextForPrompt(context As Object) As String
            If TypeOf context IsNot Dictionary(Of String, Object) Then
                Return ""
            End If

            Dim ctx = CType(context, Dictionary(Of String, Object))
            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("【当前 Excel 上下文】")

            If ctx.ContainsKey("sheetName") Then
                sb.AppendLine($"工作表: {ctx("sheetName")}")
            End If
            If ctx.ContainsKey("selection") Then
                sb.AppendLine($"当前选区: {ctx("selection")}")
            End If

            If ctx.ContainsKey("dataFeatures") AndAlso TypeOf ctx("dataFeatures") Is Dictionary(Of String, Object) Then
                Dim features = CType(ctx("dataFeatures"), Dictionary(Of String, Object))
                If features.ContainsKey("rowCount") AndAlso features.ContainsKey("colCount") Then
                    sb.AppendLine($"选区规模: {features("rowCount")} 行 x {features("colCount")} 列")
                End If

                If features.ContainsKey("headers") AndAlso TypeOf features("headers") Is List(Of String) Then
                    Dim headers = CType(features("headers"), List(Of String))
                    If headers.Count > 0 Then
                        sb.AppendLine($"表头: {String.Join(", ", headers)}")
                    End If
                End If

                If features.ContainsKey("sampleRows") AndAlso TypeOf features("sampleRows") Is List(Of List(Of String)) Then
                    Dim sampleRows = CType(features("sampleRows"), List(Of List(Of String)))
                    If sampleRows.Count > 0 Then
                        sb.AppendLine("样本数据（前3行）:")
                        For Each row In sampleRows
                            sb.AppendLine("  " & String.Join(" | ", row))
                        Next
                    End If
                End If
            End If

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' 从 ExcelBridge 加载可用 Skill 元数据，并生成决策提示词可用的 skill 选项文本。
        ''' 若无法获取，则回退到内置的基础 skill 列表。
        ''' </summary>
        Private Function BuildSkillOptionsSection() As String
            Dim sb As New System.Text.StringBuilder()
            Dim skillNames As New List(Of String)

            Try
                If _excelBridge IsNot Nothing Then
                    Dim skillsJson = _excelBridge.GetAllSkills()
                    If Not String.IsNullOrEmpty(skillsJson) Then
                        Dim resultObj = JsonConvert.DeserializeObject(Of JObject)(skillsJson)
                        If resultObj IsNot Nothing AndAlso resultObj.ContainsKey("skills") AndAlso TypeOf resultObj("skills") Is JArray Then
                            Dim skillsArray = CType(resultObj("skills"), JArray)
                            sb.AppendLine("【skill列表】（根据当前可用 Skill 动态生成）")
                            For Each skillToken In skillsArray
                                Dim name = skillToken("name")?.ToString()
                                Dim displayName = skillToken("displayName")?.ToString()
                                Dim description = skillToken("description")?.ToString()
                                Dim examples = If(skillToken("examples")?.ToObject(Of String())(), New String() {})

                                If String.IsNullOrWhiteSpace(name) Then Continue For
                                skillNames.Add(name)

                                sb.Append($"- {name}")
                                If Not String.IsNullOrWhiteSpace(displayName) Then
                                    sb.Append($"（{displayName}）")
                                End If
                                If Not String.IsNullOrWhiteSpace(description) Then
                                    sb.Append($": {description}")
                                End If
                                sb.AppendLine()

                                If examples.Length > 0 Then
                                    Dim sampleExamples = examples.Take(3).Select(Function(e) e.Trim()).Where(Function(e) Not String.IsNullOrEmpty(e))
                                    If sampleExamples.Any() Then
                                        sb.AppendLine($"  示例：{String.Join("；", sampleExamples)}")
                                    End If
                                End If
                            Next
                            sb.AppendLine()
                        End If
                    End If
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[BuildSkillOptionsSection] 加载 Skill 列表失败: {ex.Message}")
            End Try

            ' 兜底：没有任何 skill 时返回基础列表，保证提示词可用
            If skillNames.Count = 0 Then
                sb.AppendLine("【skill列表】")
                sb.AppendLine("- format_table: 格式化表格")
                sb.AppendLine("- calculate: 计算")
                sb.AppendLine("- create_chart: 创建图表")
                sb.AppendLine("- data_analysis: 数据分析")
                sb.AppendLine("- pivot_table: 数据透视表")
                sb.AppendLine()
                skillNames.AddRange({"format_table", "calculate", "create_chart", "data_analysis", "pivot_table"})
            End If

            ' 同时返回可供 JSON 格式引用的 pipe 分隔 skill 名
            Dim pipeList = String.Join("|", skillNames)
            Return $"{sb.ToString().Trim()}" & vbCrLf & vbCrLf & $"【可选skill名】{pipeList}"
        End Function

        ''' <summary>
        ''' 调用模型获取决策JSON
        ''' </summary>
        Private Async Function GetModelDecision(prompt As String, hasImage As Boolean) As Task(Of DecisionJson)
            Dim contents = New List(Of MultimodalContent) From {
                MultimodalContent.CreateText(prompt)
            }

            Dim response = Await _llmEngine.ProcessContentAsync(contents)

            If response.Success Then
                Return ParseDecisionJson(response.Content)
            Else
                Throw New Exception("LLM调用失败: " & response.ErrorMessage)
            End If
        End Function

        ''' <summary>
        ''' 解析决策JSON
        ''' </summary>
        Private Function ParseDecisionJson(response As String) As DecisionJson
            Try
                ' 提取JSON部分（可能包含在markdown代码块中）
                Dim jsonStr = response
                Dim jsonMatch = System.Text.RegularExpressions.Regex.Match(response, "```(?:json)?\s*([\s\S]*?)```", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                If jsonMatch.Success Then
                    jsonStr = jsonMatch.Groups(1).Value.Trim()
                End If

                Dim decision = JsonConvert.DeserializeObject(Of DecisionJson)(jsonStr)
                Return decision
            Catch ex As Exception
                ' 备用决策：根据关键词自动选择
                Return CreateFallbackDecision(response)
            End Try
        End Function

        ''' <summary>
        ''' 创建备用决策（解析失败时）
        ''' </summary>
        Private Function CreateFallbackDecision(response As String) As DecisionJson
            Dim decision = New DecisionJson()
            Dim lowerResponse = response.ToLower()

            ' 根据关键词判断skill
            If lowerResponse.Contains("加粗") Or lowerResponse.Contains("格式化") Or lowerResponse.Contains("美化") Then
                decision.skill = "format_table"
            ElseIf lowerResponse.Contains("求和") Or lowerResponse.Contains("计算") Or lowerResponse.Contains("平均") Then
                decision.skill = "calculate"
            ElseIf lowerResponse.Contains("图表") Or lowerResponse.Contains("图") Then
                decision.skill = "create_chart"
            Else
                decision.skill = "format_table"
            End If

            ' 简单判断mode
            If lowerResponse.Length < 20 Then
                decision.mode = "quick"
            ElseIf lowerResponse.Contains("分析") Or lowerResponse.Contains("报表") Then
                decision.mode = "deep"
            Else
                decision.mode = "standard"
            End If

            decision.complexity = "simple"
            Return decision
        End Function

        ''' <summary>
        ''' 将字符串解析为思考模式枚举
        ''' </summary>
        Private Function ParseThinkingMode(modeStr As String) As ThinkingMode
            Select Case modeStr?.ToLower()
                Case "quick"
                    Return ThinkingMode.Quick
                Case "standard"
                    Return ThinkingMode.Standard
                Case "deep"
                    Return ThinkingMode.Deep
                Case "expert"
                    Return ThinkingMode.Expert
                Case Else
                    Return ThinkingMode.Standard
            End Select
        End Function

        ''' <summary>
        ''' 加载思考模板（根据决策和任务特征动态构建）
        ''' </summary>
        Private Function LoadThinkingTemplate(decision As DecisionJson, profile As TaskProfile) As Dictionary(Of String, Object)
            Return BuildThinkingTemplate(decision, profile)
        End Function

        ''' <summary>
        ''' 创建单一步骤定义
        ''' </summary>
        Private Function CreateStep(id As Integer, phase As String, action As String, detail As String) As Dictionary(Of String, Object)
            Return New Dictionary(Of String, Object) From {
                {"id", id},
                {"phase", phase},
                {"action", action},
                {"detail", detail}
            }
        End Function

        ''' <summary>
        ''' 根据任务决策和特征动态构建思考模板
        ''' 不再为所有任务硬编码 5 阶段流水线，而是按需生成阶段。
        ''' </summary>
        Private Function BuildThinkingTemplate(decision As DecisionJson, profile As TaskProfile) As Dictionary(Of String, Object)
            Dim steps As New List(Of Dictionary(Of String, Object))()
            Dim stepId As Integer = 0

            ' 【关键修复】快速模式：跳过思考步骤，直接执行
            If decision.mode = "quick" Then
                ' 只需要一个执行步骤
                If profile.RequiresAction Then
                    stepId += 1
                    steps.Add(CreateStep(stepId, "执行", "执行操作",
                        String.Format("执行 '{0}' 对应的具体操作", decision.skill)))
                Else
                    ' 无需操作的任务，直接返回结果
                    stepId += 1
                    steps.Add(CreateStep(stepId, "执行", "直接回答",
                        String.Format("直接回答用户关于 '{0}' 的问题", decision.skill)))
                End If
                
                Return New Dictionary(Of String, Object) From {
                    {"name", "quick模式"},
                    {"description", "快速模式：跳过思考，直接执行"},
                    {"maxSteps", 1},
                    {"allowBranch", False},
                    {"riskCheck", False},
                    {"userConfirm", False},
                    {"steps", steps}
                }
            End If

            ' 标准/深度/专家模式：按复杂度动态裁剪步骤
            ' 【优化】simple 复杂度精简为 理解+执行（或仅理解）
            If decision.complexity = "simple" Then
                ' 简单任务：理解 + 执行（如果有操作）
                stepId += 1
                steps.Add(CreateStep(stepId, "理解", "分析用户需求",
                    String.Format("解析用户指令，识别 '{0}' 任务需求", decision.skill)))

                If profile.RequiresAction Then
                    stepId += 1
                    steps.Add(CreateStep(stepId, "执行", "执行操作",
                        String.Format("执行 '{0}' 对应的具体操作", decision.skill)))
                End If
            Else
                ' medium/complex 复杂度：保留完整逻辑
                ' 1. 理解：几乎所有任务都需要先理解需求
                stepId += 1
                steps.Add(CreateStep(stepId, "理解", "分析用户需求",
                    String.Format("解析用户指令，识别 '{0}' 任务的显式与隐含需求", decision.skill)))

                ' 2. 约束：涉及实际操作或复杂度较高时才需要专门检查
                If profile.RequiresAction OrElse profile.ComplexityScore >= 4 Then
                    stepId += 1
                    steps.Add(CreateStep(stepId, "约束", "检查执行条件",
                        "检查选区、数据量、覆盖风险、可撤销性等实际约束"))
                End If

                ' 3. 方案：中/复杂任务才需要制定方案
                If decision.complexity = "medium" OrElse decision.complexity = "complex" Then
                    stepId += 1
                    steps.Add(CreateStep(stepId, "方案", "制定执行方案",
                        String.Format("制定 '{0}' 的执行方案与关键步骤", decision.skill)))
                End If

                ' 4. 执行：只有需要操作的任务才进入执行阶段
                If profile.RequiresAction Then
                    stepId += 1
                    steps.Add(CreateStep(stepId, "执行", "执行操作",
                        String.Format("执行 '{0}' 对应的具体操作", decision.skill)))
                End If

                ' 5. 验证/总结：最后一步，根据是否有操作选择合适描述
                stepId += 1
                If profile.RequiresAction Then
                    steps.Add(CreateStep(stepId, "验证", "验证执行结果",
                        "确认操作结果是否符合预期，检查关键数据和副作用"))
                Else
                    steps.Add(CreateStep(stepId, "验证", "复核结论",
                        "确认回答准确完整，覆盖用户问题的核心要点"))
                End If
            End If

            Dim isComplex = decision.complexity = "complex"
            Dim isExpert = decision.mode = "expert"

            Return New Dictionary(Of String, Object) From {
                {"name", decision.mode & "模式"},
                {"description", "根据任务特征动态生成的思考流程"},
                {"maxSteps", Math.Max(steps.Count * 3, 5)},
                {"allowBranch", isComplex},
                {"riskCheck", profile.RequiresAction OrElse profile.ComplexityScore >= 4},
                {"userConfirm", isExpert OrElse (isComplex AndAlso profile.RequiresAction)},
                {"steps", steps}
            }
        End Function

        ''' <summary>
        ''' 执行思考模板
        ''' </summary>
        Private Async Function ExecuteTemplate(decision As DecisionJson, template As Dictionary(Of String, Object), Optional progress As IProgress(Of ThinkingStep) = Nothing) As Task(Of ThinkingResponse)
            Dim response = New ThinkingResponse()
            Dim context = New Dictionary(Of String, Object) From {
                {"decision", decision}
            }

            Dim steps = CType(template("steps"), List(Of Dictionary(Of String, Object)))

            For Each stepDef In steps
                Dim thinkingStep As New ThinkingStep()
                thinkingStep.StepId = CInt(stepDef("id"))
                thinkingStep.Phase = stepDef("phase").ToString()
                thinkingStep.StepAction = stepDef("action").ToString()
                thinkingStep.Detail = stepDef("detail").ToString() ' 初始使用模板描述
                thinkingStep.Status = "running"
                thinkingStep.Timestamp = DateTime.Now

                response.ThinkingSteps.Add(thinkingStep)

                ' 【流式推送】立即推送开始执行的步骤
                If progress IsNot Nothing Then
                    progress.Report(thinkingStep)
                    Await Task.Delay(100)
                End If

                Try
                    ' 执行步骤逻辑
                    Dim resultObj = Await ExecuteStep(stepDef, decision, context)
                    thinkingStep.Result = resultObj.ToString()
                    thinkingStep.Status = "completed"
                    context(thinkingStep.Phase) = resultObj
                    
                    ' 【关键修复】根据执行结果动态更新步骤Detail
                    thinkingStep.Detail = GenerateStepDetail(thinkingStep.Phase, resultObj, decision)
                    
                    ' 【流式推送】推送完成状态（含更新的Detail）
                    If progress IsNot Nothing Then
                        progress.Report(thinkingStep)
                        Await Task.Delay(100)
                    End If
                    
                    ' 如果是执行阶段,将执行结果转换为 ActionItem
                    If thinkingStep.Phase = "执行" AndAlso TypeOf resultObj Is ExecutionResult Then
                        Dim execResult = CType(resultObj, ExecutionResult)
                        Dim action = New ActionItem()
                        action.ActionType = "skill_execution"
                        action.Description = execResult.Message
                        action.Confidence = 0.9
                        action.Parameters = New Dictionary(Of String, Object) From {
                            {"success", execResult.Success}
                        }

                        ' 尝试从 Data 中提取 skill 和 parameters
                        If TypeOf execResult.Data Is Dictionary(Of String, Object) Then
                            Dim dataDict = CType(execResult.Data, Dictionary(Of String, Object))
                            If dataDict.ContainsKey("skill") Then action.Parameters("skill") = dataDict("skill")
                            If dataDict.ContainsKey("parameters") Then action.Parameters("parameters") = dataDict("parameters")
                        Else
                            action.Parameters("data") = execResult.Data
                        End If

                        response.Actions.Add(action)
                    End If

                Catch ex As Exception
                    thinkingStep.Status = "failed"
                    thinkingStep.Result = "错误: " & ex.Message
                    
                    ' 【流式推送】推送失败状态
                    If progress IsNot Nothing Then
                        progress.Report(thinkingStep)
                    End If
                End Try
            Next

            response.FinalAnswer = GenerateFinalAnswer(response)
            response.Confidence = CalculateConfidence(response)

            Return response
        End Function
        
        ''' <summary>
        ''' 根据执行结果动态生成步骤详情
        ''' 使用强类型结果对象，避免反射和硬编码截断。
        ''' </summary>
        Private Function GenerateStepDetail(phase As String, resultObj As Object, decision As DecisionJson) As String
            Try
                Dim result As StepResult = TryCast(resultObj, StepResult)
                If result IsNot Nothing Then
                    Dim summary = result.GetSummary()
                    If Not String.IsNullOrEmpty(summary) Then
                        Return summary
                    End If
                End If

                ' 兜底描述
                Select Case phase
                    Case "理解"
                        Return $"分析{decision.skill}任务需求"
                    Case "约束"
                        Return "检查执行约束"
                    Case "方案"
                        Return $"制定{decision.skill}执行方案"
                    Case "执行"
                        Return $"执行{decision.skill}操作"
                    Case "验证"
                        Return "验证执行结果"
                    Case Else
                        Return $"{phase}阶段"
                End Select
            Catch ex As Exception
                Return $"{phase}阶段执行中"
            End Try
        End Function

        ''' <summary>
        ''' 执行单个步骤
        ''' </summary>
        Private Async Function ExecuteStep(stepDef As Dictionary(Of String, Object), decision As DecisionJson, context As Dictionary(Of String, Object)) As Task(Of Object)
            Dim phase = stepDef("phase").ToString()
            Dim stepId = CInt(stepDef("id"))

            Select Case phase
                Case "理解"
                    ' 根据任务类型动态生成理解内容
                    Dim understanding = Await GenerateUnderstanding(decision, stepId)
                    Return understanding
                    
                Case "约束"
                    Return Await CheckConstraints(decision)
                    
                Case "方案"
                    Return Await GenerateSolution(decision)
                    
                Case "执行"
                    Return Await ExecuteSkill(decision)
                    
                Case "验证"
                    Return New VerificationResult With {
                        .Success = True,
                        .Tips = "操作已完成"
                    }
                    
                Case Else
                    Return New With {.status = "unknown"}
            End Select
        End Function
        
        ''' <summary>
        ''' 生成任务理解（动态）
        ''' 基于真实 Excel 上下文分析用户需求，避免脱离实际工作簿状态。
        ''' 【优化】仅 complex 复杂度调用 LLM，simple/medium 用代码生成
        ''' </summary>
        Private Async Function GenerateUnderstanding(decision As DecisionJson, stepId As Integer) As Task(Of UnderstandingResult)
            Dim result As New UnderstandingResult()
            result.TaskType = decision.skill
            Dim excelContext = FormatExcelContextForPrompt(GetExcelContext())

            ' 【优化】simple/medium 复杂度：直接返回基本信息，不调用 LLM
            If decision.complexity = "simple" OrElse decision.complexity = "medium" Then
                result.Summary = $"执行{decision.skill}操作"
                result.Analysis = If(String.IsNullOrEmpty(excelContext), result.Summary, excelContext & vbCrLf & "任务: " & result.Summary)
                Return result
            End If

            ' complex 复杂度：调用LLM生成详细理解
            Dim prompt = $"分析用户任务并输出理解结果。" & vbCrLf & vbCrLf &
                         $"任务类型：{decision.skill}" & vbCrLf &
                         $"复杂度：{decision.complexity}" & vbCrLf &
                         $"决策参数：{decision.parameters?.ToString()}" & vbCrLf & vbCrLf &
                         excelContext & vbCrLf &
                         "要求：结合当前 Excel 上下文理解用户真正意图，不要脱离选区和数据规模做假设。输出简洁的任务理解。"

            Try
                Dim contents = New List(Of MultimodalContent) From {
                    MultimodalContent.CreateText(prompt)
                }
                Dim response = Await _llmEngine.ProcessContentAsync(contents)

                If response.Success Then
                    result.Analysis = response.Content
                Else
                    result.Analysis = $"分析{decision.skill}任务需求"
                End If
            Catch ex As Exception
                result.Analysis = $"理解用户需求：{decision.skill}"
            End Try

            Return result
        End Function
        
        ''' <summary>
        ''' 生成解决方案（动态）
        ''' 基于真实 Excel 上下文制定执行方案，避免覆盖错误范围或处理空选区。
        ''' </summary>
        Private Async Function GenerateSolution(decision As DecisionJson) As Task(Of SolutionResult)
            Dim result As New SolutionResult()
            Dim skillType = decision.skill
            Dim excelContext = FormatExcelContextForPrompt(GetExcelContext())

            Try
                Dim prompt = $"为 {skillType} 任务生成执行方案。" & vbCrLf & vbCrLf &
                             $"复杂度：{decision.complexity}" & vbCrLf &
                             $"决策参数：{decision.parameters?.ToString()}" & vbCrLf & vbCrLf &
                             excelContext & vbCrLf &
                             "要求：方案必须考虑当前选区、数据规模和表头。如果选区为空或不确定，先说明需要用户指定范围。输出具体可执行步骤，不要泛泛而谈。"
                Dim contents = New List(Of MultimodalContent) From {
                    MultimodalContent.CreateText(prompt)
                }
                Dim response = Await _llmEngine.ProcessContentAsync(contents)

                If response.Success Then
                    result.Plan = response.Content
                    result.Steps = "按步骤执行"
                Else
                    result.Plan = $"执行{skillType}操作"
                    result.Steps = "默认步骤"
                End If
            Catch ex As Exception
                result.Plan = $"执行{skillType}任务"
                result.Steps = "标准流程"
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 检查约束
        ''' 基于真实 Excel 上下文和决策参数进行基础约束与风险评估。
        ''' </summary>
        Private Function CheckConstraints(decision As DecisionJson) As Task(Of ConstraintResult)
            Dim result As New ConstraintResult With {
                .CanProceed = True,
                .HasRisk = False,
                .Warning = ""
            }

            Try
                Dim context = GetExcelContext()
                Dim selection As String = ""
                If context.ContainsKey("selection") AndAlso context("selection") IsNot Nothing Then
                    selection = context("selection").ToString()
                End If
                Dim dataFeatures As Dictionary(Of String, Object) = Nothing
                If context.ContainsKey("dataFeatures") AndAlso TypeOf context("dataFeatures") Is Dictionary(Of String, Object) Then
                    dataFeatures = DirectCast(context("dataFeatures"), Dictionary(Of String, Object))
                Else
                    dataFeatures = New Dictionary(Of String, Object) From {{"rowCount", 0}, {"colCount", 0}}
                End If
                Dim rowCount = Convert.ToInt32(dataFeatures("rowCount"))
                Dim colCount = Convert.ToInt32(dataFeatures("colCount"))
                Dim cellCount = rowCount * colCount

                ' 检查 1：需要范围但没有明确范围或选区为空
                Dim hasExplicitRange = decision.parameters IsNot Nothing AndAlso decision.parameters.ContainsKey("range")
                Dim selectionEmpty = String.IsNullOrEmpty(selection) OrElse selection = "A1" OrElse cellCount = 0

                If Not hasExplicitRange AndAlso selectionEmpty Then
                    If decision.skill = "format_table" OrElse decision.skill = "calculate" Then
                        result.HasRisk = True
                        result.Warning = "未指定操作范围且当前选区为空，执行可能无法生效。"
                    End If
                End If

                ' 检查 2：大数据量风险
                If cellCount > 100000 Then
                    result.HasRisk = True
                    result.Warning = If(String.IsNullOrEmpty(result.Warning), "", result.Warning & " ") &
                        String.Format("当前选区包含约 {0} 个单元格，操作可能耗时较长。", cellCount)
                End If

                ' 检查 3：可能覆盖或改变现有数据/格式
                If decision.skill = "format_table" AndAlso Not selectionEmpty Then
                    result.HasRisk = True
                    result.Warning = If(String.IsNullOrEmpty(result.Warning), "", result.Warning & " ") &
                        "格式化操作会影响选区内的现有格式或数据，请确认。"
                End If

                ' 检查 4：创建图表/透视表会修改工作表结构
                If decision.skill = "create_chart" OrElse decision.skill = "pivot_table" Then
                    result.HasRisk = True
                    result.Warning = If(String.IsNullOrEmpty(result.Warning), "", result.Warning & " ") &
                        "该操作会在工作表中创建新对象，请确认目标位置合适。"
                End If

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("[CheckConstraints] 约束检查异常: {0}", ex.Message))
            End Try

            Return Task.FromResult(result)
        End Function

        ''' <summary>
        ''' 【规则兜底】校验 skill 名是否存在于已注册技能列表（防止 LLM 幻觉/拼错）
        ''' </summary>
        Private Function IsValidRegisteredSkill(skillName As String) As Boolean
            If String.IsNullOrWhiteSpace(skillName) Then Return False
            Try
                Dim skill = Skills.SkillManager.Instance.GetSkill(skillName)
                Return skill IsNot Nothing
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[ThinkingEngine] 校验 skill 失败: " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 【规则兜底】用 SkillRecognizer 从用户输入识别 skill（规则语义匹配，无 LLM 调用）
        ''' </summary>
        Private Function TryRuleSkillFallback(input As String, context As String) As String
            Try
                If _skillRecognizer Is Nothing Then
                    _skillRecognizer = New Skills.SkillRecognizer(Skills.SkillManager.Instance)
                End If
                Dim rec = _skillRecognizer.Recognize(input, context)
                If rec IsNot Nothing AndAlso rec.IsMatch AndAlso rec.MatchedSkill IsNot Nothing Then
                    Return rec.MatchedSkill.Name
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[ThinkingEngine] 规则识别兜底失败: " & ex.Message)
            End Try
            Return Nothing
        End Function

        ''' <summary>
        ''' 【规则兜底】将任务复杂度评分(0-10)映射为决策 complexity
        ''' </summary>
        Private Function MapComplexityScore(profile As TaskProfile) As String
            If profile Is Nothing Then Return "medium"
            If profile.ComplexityScore < 4 Then Return "simple"
            If profile.ComplexityScore < 7 Then Return "medium"
            Return "complex"
        End Function

        ''' <summary>
        ''' 执行Skill
        ''' 将决策的 skill 映射为具体 ActionItem，通过 ActionEngine 真正执行。
        ''' 优先使用 excel_operation 直接完成；不满足内置操作条件时降级为 code_generation。
        ''' </summary>
        Private Async Function ExecuteSkill(decision As DecisionJson) As Task(Of ExecutionResult)
            If _actionEngine Is Nothing Then
                Return New ExecutionResult With {
                    .Success = False,
                    .Message = "动作引擎未初始化，无法执行 skill: " & decision.skill,
                    .Data = New Dictionary(Of String, Object) From {
                        {"type", "skill_execution"},
                        {"skill", decision.skill},
                        {"parameters", decision.parameters}
                    }
                }
            End If

            Dim actionType = DetermineActionType(decision.skill, decision.parameters)
            Dim action = BuildActionItem(decision, actionType)

            Dim operationResult = Await _actionEngine.ExecuteSingleAction(action)

            Return New ExecutionResult With {
                .Success = operationResult.Success,
                .Message = operationResult.Message,
                .Data = operationResult.Data
            }
        End Function

        ''' <summary>
        ''' 根据 skill 和参数决定动作类型：优先内置操作，必要时生成代码。
        ''' </summary>
        Private Function DetermineActionType(skill As String, parameters As JObject) As String
            Select Case skill?.ToLower()
                Case "format_table"
                    ' 格式化可直接通过 excel_operation 完成
                    Return "excel_operation"
                Case "calculate"
                    ' 简单读取统计可直读；涉及公式写入时仍走代码生成
                    Dim ops = GetOperationsList(parameters)
                    If ops.Count > 0 AndAlso ops.All(Function(o) {"read", "sum", "average", "count", "max", "min"}.Contains(o.ToLower())) Then
                        Return "excel_operation"
                    End If
                    Return "code_generation"
                Case Else
                    ' 图表、数据分析、透视表等当前内置操作未覆盖，走代码生成
                    Return "code_generation"
            End Select
        End Function

        ''' <summary>
        ''' 构建 ActionItem，根据动作类型填充对应参数。
        ''' </summary>
        Private Function BuildActionItem(decision As DecisionJson, actionType As String) As ActionItem
            Dim description = BuildSkillActionDescription(decision.skill, decision.parameters)
            Dim action = New ActionItem With {
                .ActionType = actionType,
                .Description = description,
                .Confidence = 0.8,
                .Parameters = New Dictionary(Of String, Object) From {
                    {"description", description}
                }
            }

            If actionType = "excel_operation" Then
                Select Case decision.skill?.ToLower()
                    Case "format_table"
                        action.Parameters("operation") = "format_cells"
                        If decision.parameters IsNot Nothing AndAlso decision.parameters.ContainsKey("range") Then
                            action.Parameters("range") = decision.parameters("range").ToString()
                        Else
                            Dim ctxFmt = GetExcelContext()
                            action.Parameters("range") = If(ctxFmt.ContainsKey("selection") AndAlso ctxFmt("selection") IsNot Nothing, ctxFmt("selection").ToString(), "A1")
                        End If
                        action.Parameters("format") = BuildFormatParameters(decision.parameters)
                    Case "calculate"
                        action.Parameters("operation") = "read_data"
                        If decision.parameters IsNot Nothing AndAlso decision.parameters.ContainsKey("range") Then
                            action.Parameters("range") = decision.parameters("range").ToString()
                        Else
                            Dim ctxCalc = GetExcelContext()
                            action.Parameters("range") = If(ctxCalc.ContainsKey("selection") AndAlso ctxCalc("selection") IsNot Nothing, ctxCalc("selection").ToString(), "A1")
                        End If
                End Select
            End If

            Return action
        End Function

        ''' <summary>
        ''' 从 decision.parameters 中提取 operations 列表。
        ''' </summary>
        Private Function GetOperationsList(parameters As JObject) As List(Of String)
            Dim ops = New List(Of String)
            If parameters Is Nothing OrElse Not parameters.ContainsKey("operations") Then
                Return ops
            End If

            Dim opsToken = parameters("operations")
            If TypeOf opsToken Is JArray Then
                ops = CType(opsToken, JArray).Select(Function(t) t.ToString()).ToList()
            ElseIf opsToken IsNot Nothing Then
                ops.Add(opsToken.ToString())
            End If

            Return ops
        End Function

        ''' <summary>
        ''' 根据 operations 数组构建 format JSON 对象。
        ''' </summary>
        Private Function BuildFormatParameters(parameters As JObject) As String
            Dim format As New JObject()
            Dim ops = GetOperationsList(parameters)

            For Each op In ops
                Select Case op.ToLower()
                    Case "bold"
                        format("bold") = True
                    Case "italic"
                        format("italic") = True
                    Case "underline"
                        format("underline") = True
                    Case "center"
                        format("horizontalAlignment") = "center"
                    Case "left"
                        format("horizontalAlignment") = "left"
                    Case "right"
                        format("horizontalAlignment") = "right"
                    Case "header", "header_style"
                        format("bold") = True
                        format("bgColor") = "#4472C4"
                        format("color") = "#FFFFFF"
                        format("horizontalAlignment") = "center"
                End Select
            Next

            Return format.ToString()
        End Function

        ''' <summary>
        ''' 根据 skill 和参数生成动作描述
        ''' </summary>
        Private Function BuildSkillActionDescription(skill As String, parameters As JObject) As String
            Dim extra As String = ""
            If parameters IsNot Nothing Then
                If parameters.ContainsKey("range") Then
                    extra &= "，范围：" & parameters("range").ToString()
                End If
                If parameters.ContainsKey("operations") Then
                    extra &= "，操作：" & parameters("operations").ToString()
                End If
            End If

            Select Case skill.ToLower()
                Case "format_table"
                    Return "为Excel表格设置格式（表头、边框、列宽、数字格式等）" & extra
                Case "calculate"
                    Return "根据需求生成计算VBA代码并执行（求和、平均值、公式等）" & extra
                Case "create_chart"
                    Return "生成创建图表的VBA代码并执行" & extra
                Case "data_analysis"
                    Return "生成数据分析VBA代码并执行（统计、排序、筛选等）" & extra
                Case "pivot_table"
                    Return "生成创建数据透视表的VBA代码并执行" & extra
                Case Else
                    Return "执行Excel操作：" & skill & extra
            End Select
        End Function

        ''' <summary>
        ''' 生成最终回答
        ''' 基于实际执行结果和思考步骤生成有针对性的结论，不再使用固定模板。
        ''' </summary>
        Private Function GenerateFinalAnswer(response As ThinkingResponse) As String
            ' 有执行动作：根据实际执行结果生成结论
            If response.Actions IsNot Nothing AndAlso response.Actions.Count > 0 Then
                Dim action = response.Actions(0)
                Dim success = False
                If action.Parameters IsNot Nothing AndAlso action.Parameters.ContainsKey("success") Then
                    success = Convert.ToBoolean(action.Parameters("success"))
                End If

                If success Then
                    Return "已完成：" & action.Description & "。如有需要可通过撤销恢复。"
                Else
                    Return action.Description & " 请检查输入或尝试其他方式。"
                End If
            End If

            ' 无执行动作：基于最后一步思考给出总结性结论
            If response.ThinkingSteps IsNot Nothing AndAlso response.ThinkingSteps.Count > 0 Then
                Dim lastStep = response.ThinkingSteps(response.ThinkingSteps.Count - 1)
                If Not String.IsNullOrEmpty(lastStep.Detail) Then
                    Return lastStep.Detail
                End If
            End If

            Return "操作规划完成"
        End Function

#End Region

        ''' <summary>
        ''' 执行思考(调用LLM)
        ''' </summary>
        Private Async Function ExecuteThinking(
            prompt As String,
            hasImage As Boolean
        ) As Task(Of String)

            ' 构建多模态内容
            Dim contents = New List(Of MultimodalContent) From {
                MultimodalContent.CreateText(prompt)
            }

            ' 调用LLM
            Dim response = Await _llmEngine.ProcessContentAsync(contents)

            If response.Success Then
                Return response.Content
            Else
                Throw New Exception("LLM调用失败: " & response.ErrorMessage)
            End If
        End Function

        ''' <summary>
        ''' 解析思考响应
        ''' </summary>
        Private Function ParseThinkingResponse(
            llmResponse As String,
            mode As ThinkingMode
        ) As ThinkingResponse

            Dim result As New ThinkingResponse

            Try
                ' 提取 <thinking> 标签内容
                Dim thinkingMatch = System.Text.RegularExpressions.Regex.Match(
                    llmResponse,
                    "<thinking>([\s\S]*?)</thinking>",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase
                )

                If thinkingMatch.Success Then
                    Dim thinkingContent = thinkingMatch.Groups(1).Value.Trim()
                    result.ThinkingSteps = ParseThinkingSteps(thinkingContent, mode)
                End If

                ' 提取 <answer> 标签内容
                Dim answerMatch = System.Text.RegularExpressions.Regex.Match(
                    llmResponse,
                    "<answer>([\s\S]*?)</answer>",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase
                )

                If answerMatch.Success Then
                    result.FinalAnswer = answerMatch.Groups(1).Value.Trim()
                Else
                    ' 如果没有answer标签,使用整个响应
                    result.FinalAnswer = llmResponse
                End If

                ' 提取 <actions> 标签内容
                Dim actionsMatch = System.Text.RegularExpressions.Regex.Match(
                    llmResponse,
                    "<actions>([\s\S]*?)</actions>",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase
                )

                If actionsMatch.Success Then
                    Dim actionsJson = actionsMatch.Groups(1).Value.Trim()
                    result.Actions = ParseActions(actionsJson)
                End If

                ' 计算置信度
                result.Confidence = CalculateConfidence(result)

            Catch ex As Exception
                ' 解析失败,返回原始响应
                result.FinalAnswer = llmResponse
                result.Confidence = 0.5
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 解析思考步骤
        ''' </summary>
        Private Function ParseThinkingSteps(
            thinkingContent As String,
            mode As ThinkingMode
        ) As List(Of ThinkingStep)

            Dim steps = New List(Of ThinkingStep)

            ' 按行分割
            Dim lines = thinkingContent.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.RemoveEmptyEntries)

            Dim currentStep As ThinkingStep = Nothing
            Dim stepCounter As Integer = 0

            For Each line In lines
                ' 检测新步骤标记
                If line.Trim().StartsWith("**") OrElse line.Trim().StartsWith("步骤") OrElse line.Trim().StartsWith("-") Then
                    If currentStep IsNot Nothing Then
                        steps.Add(currentStep)
                    End If

                    stepCounter += 1
                    currentStep = New ThinkingStep With {
                        .StepId = stepCounter,
                        .Phase = "理解",
                        .StepAction = "解析步骤",
                        .Detail = line.Trim(),
                        .Status = "completed",
                        .Result = "",
                        .Timestamp = DateTime.Now
                    }
                ElseIf currentStep IsNot Nothing Then
                    ' 追加到当前步骤
                    currentStep.Detail += vbCrLf + line.Trim()
                End If
            Next

            ' 添加最后一个步骤
            If currentStep IsNot Nothing Then
                steps.Add(currentStep)
            End If

            Return steps
        End Function

        ''' <summary>
        ''' 根据思考模式获取步骤类型
        ''' </summary>
        Private Function GetStepTypeFromMode(mode As ThinkingMode) As String
            Select Case mode
                Case ThinkingMode.Quick
                    Return "quick_execution"
                Case ThinkingMode.Standard
                    Return "standard_planning"
                Case ThinkingMode.Deep
                    Return "deep_analysis"
                Case ThinkingMode.Expert
                    Return "expert_interaction"
                Case Else
                    Return "standard_planning"
            End Select
        End Function

        ''' <summary>
        ''' 解析动作列表
        ''' </summary>
        Private Function ParseActions(actionsJson As String) As List(Of ActionItem)
            Dim actions = New List(Of ActionItem)

            Try
                ' 尝试解析为JSON数组
                Dim jsonArray = JArray.Parse(actionsJson)

                For Each item In jsonArray
                    Dim typeToken = item("type")
                    Dim descToken = item("description")
                    Dim confToken = item("confidence")
                    Dim actionTypeStr As String = If(typeToken IsNot Nothing, typeToken.ToString(), "unknown")
                    Dim descStr As String = If(descToken IsNot Nothing, descToken.ToString(), "")
                    Dim confStr As String = If(confToken IsNot Nothing, confToken.ToString(), "0.5")
                    Dim confValue As Double = 0.5
                    Double.TryParse(confStr, confValue)
                    Dim action = New ActionItem With {
                        .ActionType = actionTypeStr,
                        .Description = descStr,
                        .Confidence = confValue
                    }

                    ' 解析参数
                    If item("parameters") IsNot Nothing Then
                        Dim paramsObj = CType(item("parameters"), JObject)
                        For Each prop In paramsObj.Properties()
                            action.Parameters(prop.Name) = prop.Value.ToString()
                        Next
                    End If

                    actions.Add(action)
                Next

            Catch ex As Exception
                ' JSON解析失败,返回空列表
            End Try

            Return actions
        End Function

        ''' <summary>
        ''' 计算置信度
        ''' </summary>
        Private Function CalculateConfidence(response As ThinkingResponse) As Double
            Dim confidence As Double = 0.5

            ' 有思考步骤增加置信度
            If response.ThinkingSteps.Count > 0 Then
                confidence += 0.1
            End If

            ' 有明确答案增加置信度
            If Not String.IsNullOrEmpty(response.FinalAnswer) Then
                confidence += 0.2
            End If

            ' 有高置信度动作增加
            If response.Actions.Any(Function(a) a.Confidence > 0.8) Then
                confidence += 0.1
            End If

            Return Math.Min(confidence, 1.0)
        End Function

        ''' <summary>
        ''' 自动 MCP 路由：用 McpCapabilityRegistry.Match() 匹配工具 → LLM 生成参数 → CallLocalTool 执行
        ''' 零维护 — 加新 MCP 工具只要注册 McpTool 带 keywords，AI 自动能驱动
        ''' </summary>
        Private Async Function TryAutoMcpRoute(userMessage As String) As Task(Of ThinkingResponse)
            Try
                ' 兜底确保所有内置工具已注册
                Mcp.McpCapabilityRegistry.RegisterAllBuiltInToolSets()
                Dim registry = Mcp.McpCapabilityRegistry.Instance

                ' 1. 关键词/n-gram 匹配
                Dim matches = registry.Match(userMessage, topN:=5)
                If matches Is Nothing OrElse matches.Count = 0 Then
                    System.Diagnostics.Debug.WriteLine("[AutoMcp] Match 返回 0 结果，回退思维链")
                    Return Nothing
                End If

                ' 🔥 详细诊断 — 打印所有 top5 分数来源
                System.Diagnostics.Debug.WriteLine($"[AutoMcp] 用户消息: ""{userMessage}""")
                For i = 0 To matches.Count - 1
                    Dim m = matches(i)
                    Dim typeTag = If(m.Capability.Source.EndsWith("-local"), m.Capability.Source, "external")
                    System.Diagnostics.Debug.WriteLine($"[AutoMcp] top{i+1}: score={m.Score,5:N0}  {m.Capability.Id,-35} [{typeTag}] 命中kw=[{String.Join(",", m.MatchedKeywords)}]")
                Next

                Dim top1 = matches(0)
                ' 置信度门槛：DisplayName 命中(≥100) 或 关键词命中(≥20) 就自动路由
                Dim hasDisplayNameHit = top1.Score >= 100
                Dim hasKeywordHit = top1.Score >= 20
                If Not (hasDisplayNameHit OrElse hasKeywordHit) Then
                    System.Diagnostics.Debug.WriteLine($"[AutoMcp] top1 score={top1.Score} 不够门槛(需要≥20)，回退思维链")
                    Return Nothing
                End If

                ' 🔥 额外安全：如果 top1 不是 computer 工具，但有 computer 工具挤进了 top3 且分数差距 < 30，
                ' 说明用户大概率想操作电脑但被 Excel 工具干扰了——优先 computer
                If Not top1.Capability.Id.StartsWith("computer_") Then
                    For i = 1 To Math.Min(2, matches.Count - 1)
                        Dim alt = matches(i)
                        If alt.Capability.Id.StartsWith("computer_") AndAlso (top1.Score - alt.Score) < 30 Then
                            System.Diagnostics.Debug.WriteLine($"[AutoMcp] ⚠️ top1={top1.Capability.Id} 但 {alt.Capability.Id} 分数接近({alt.Score})，切换到 computer")
                            top1 = alt
                            Exit For
                        End If
                    Next
                End If

                Dim toolDesc = top1.Capability
                Dim toolName = toolDesc.Id
                System.Diagnostics.Debug.WriteLine($"[AutoMcp] ✅ 最终选择: {toolName} (score={top1.Score}, kw=[{String.Join(",", top1.MatchedKeywords)}])")

                ' 2. 让 LLM 基于 InputSchema 生成 args JSON
                Dim argsPrompt = BuildMcpArgsPrompt(toolDesc, userMessage)
                Dim argsJson = Await GenerateMcpArgs(argsPrompt)
                If String.IsNullOrWhiteSpace(argsJson) Then argsJson = "{}"

                System.Diagnostics.Debug.WriteLine($"[AutoMcp] 生成参数: {argsJson}")

                ' 3. 执行（经接口解析，避免依赖具体 McpBridge/内部 McpManager）
                Dim mcpBridge = Bridges.BridgeRegistry.Instance.Resolve(Of Bridges.IMcpBridge)()
                If mcpBridge Is Nothing Then Return Nothing

                Dim resultJson = mcpBridge.CallLocalTool(toolName, argsJson)
                Dim resultObj = JObject.Parse(resultJson)
                Dim successToken = resultObj("success")
                Dim success = If(successToken IsNot Nothing, successToken.ToString() = "True", False)
                Dim resultToken = resultObj("result")
                Dim resultText = If(resultToken IsNot Nothing, resultToken.ToString(), "")
                Dim errorToken = resultObj("error")
                Dim errorText = If(errorToken IsNot Nothing, errorToken.ToString(), "")
                Dim isErrorToken = resultObj("isError")
                Dim isError = If(isErrorToken IsNot Nothing, isErrorToken.ToString() = "True", False)

                ' 4. 构建返回
                Dim answer As String
                If isError OrElse Not success Then
                    answer = $"执行 {toolName} 失败：{If(String.IsNullOrEmpty(errorText), resultText, errorText)}"
                Else
                    answer = BuildMcpSuccessMessage(toolDesc, resultText, userMessage)
                End If

                Dim resp As New ThinkingResponse With {
                    .FinalAnswer = answer,
                    .Confidence = top1.Confidence,
                    .ThinkingSteps = New List(Of ThinkingStep)()
                }
                resp.ThinkingSteps.Add(New ThinkingStep With {
                    .StepId = 1,
                    .Phase = "MCP工具匹配",
                    .Detail = $"自动匹配: {toolName} (命中: {String.Join(", ", top1.MatchedKeywords)})",
                    .Status = "completed"
                })
                resp.ThinkingSteps.Add(New ThinkingStep With {
                    .StepId = 2,
                    .Phase = "执行",
                    .Detail = $"调用参数: {argsJson}",
                    .Status = "completed"
                })
                resp.ThinkingSteps.Add(New ThinkingStep With {
                    .StepId = 3,
                    .Phase = "结果",
                    .Detail = answer,
                    .Status = "completed"
                })

                Return resp

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[AutoMcp] 自动路由失败，回退到思维链: {ex.Message}")
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 构建让 LLM 生成 MCP 工具参数的 prompt
        ''' </summary>
        Private Function BuildMcpArgsPrompt(toolDesc As Mcp.CapabilityDescriptor, userMessage As String) As String
            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("你是 MCP 工具参数生成器。根据工具 InputSchema 和用户需求，输出一个 JSON 对象作为工具参数。")
            sb.AppendLine()
            sb.AppendLine($"工具名: {toolDesc.Id}")
            sb.AppendLine($"工具描述: {toolDesc.Description}")
            sb.AppendLine()
            sb.AppendLine("InputSchema:")
            If toolDesc.InputSchema IsNot Nothing AndAlso toolDesc.InputSchema.Properties IsNot Nothing Then
                For Each kvp In toolDesc.InputSchema.Properties
                    Dim req = If(toolDesc.InputSchema.Required IsNot Nothing AndAlso toolDesc.InputSchema.Required.Contains(kvp.Key), "(必填)", "(可选)")
                    sb.AppendLine($"  - {kvp.Key} {req}: [{kvp.Value.Type}] {kvp.Value.Description}")
                Next
            Else
                sb.AppendLine("  (无参数)")
            End If
            sb.AppendLine()
            sb.AppendLine($"用户需求: {userMessage}")
            sb.AppendLine()
            sb.AppendLine("要求：只输出纯 JSON 对象（不要 markdown 代码块，不要文字解释）。无参数时输出 {}。")
            Return sb.ToString()
        End Function

        ''' <summary>
        ''' 调 LLM 生成 MCP 工具参数（JSON）
        ''' </summary>
        Private Async Function GenerateMcpArgs(prompt As String) As Task(Of String)
            Try
                Dim contents = New List(Of MultimodalContent) From {
                    MultimodalContent.CreateText(prompt)
                }
                Dim response = Await _llmEngine.ProcessContentAsync(contents)
                If Not response.Success OrElse String.IsNullOrWhiteSpace(response.Content) Then Return "{}"

                ' 清洗：LLM 可能输出 ```json ... ``` 或 带文字，提取纯 JSON
                Dim raw = response.Content.Trim()
                ' 剥 markdown 代码块
                If raw.StartsWith("```") Then
                    Dim lastFence = raw.LastIndexOf("```")
                    If lastFence > 0 Then
                        raw = raw.Substring(raw.IndexOf(vbCrLf) + 2, lastFence - raw.IndexOf(vbCrLf) - 2).Trim()
                    End If
                End If
                ' 找第一个 { 到最后一个 }
                Dim startBrace = raw.IndexOf("{"c)
                Dim endBrace = raw.LastIndexOf("}"c)
                If startBrace >= 0 AndAlso endBrace > startBrace Then
                    raw = raw.Substring(startBrace, endBrace - startBrace + 1)
                End If
                ' 验证是合法 JSON
                Try
                    JObject.Parse(raw)
                    Return raw
                Catch
                    Return "{}"
                End Try
            Catch
                Return "{}"
            End Try
        End Function

        ''' <summary>
        ''' 构建 MCP 工具执行成功的用户友好回复
        ''' </summary>
        Private Function BuildMcpSuccessMessage(toolDesc As Mcp.CapabilityDescriptor, resultText As String, userMessage As String) As String
            If String.IsNullOrWhiteSpace(resultText) Then
                Return $"✅ 已执行 {toolDesc.Id}"
            End If
            ' 对结果做截断（太长的 JSON 或日志）
            Dim displayResult = resultText
            If displayResult.Length > 500 Then
                displayResult = displayResult.Substring(0, 500) & "..."
            End If
            Return $"✅ {toolDesc.DisplayName} 执行成功：" & vbCrLf & displayResult
        End Function



    End Class

End Namespace