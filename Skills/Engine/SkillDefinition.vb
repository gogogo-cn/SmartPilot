Imports Newtonsoft.Json

Namespace Skills

    ''' <summary>
    ''' 统一步骤类型枚举（整合原 WorkflowStepType + CustomStepType）
    ''' 按领域分组：
    '''   通用控制流  0-19
    '''   代码执行    20-39
    '''   OCR/文档    40-59
    '''   Excel操作   60-79
    '''   AI/LLM      80-99
    '''   外部集成    100+
    ''' </summary>
    Public Enum WorkflowStepType
        ' ==================== 通用控制流 (0-19) ====================
        Condition = 0
        [Loop] = 1
        Wait = 2
        InvokeSubSkill = 3

        ' ==================== 代码执行 (20-39) ====================
        ExecuteCode = 20
        CallApi = 21

        ' ==================== AI/LLM (40-59) ====================
        LLMProcess = 40
        MultimodalProcess = 41
        ''' <summary>原 CustomStepType.CustomLLM，自定义 LLM 提示词调用</summary>
        CustomLLM = 42

        ' ==================== OCR/文档 (60-79) ====================
        ''' <summary>原 CustomStepType.UploadImage</summary>
        UploadImage = 60
        ''' <summary>原 CustomStepType.OCRRecognition</summary>
        OCRRecognition = 61
        ''' <summary>原 CustomStepType.ExtractFields</summary>
        ExtractFields = 62

        ' ==================== Excel操作 (80-99) ====================
        ''' <summary>原 CustomStepType.WriteToExcel，按映射表写入字段</summary>
        WriteToExcel = 80

        ' ==================== 外部集成 (100+) ====================
        ''' <summary>
        ''' 调用 MCP 协议外部工具
        ''' 通过已连接的 MCP Server 执行外部操作（文件系统、邮件、数据库等）
        ''' </summary>
        InvokeMcpTool = 100
    End Enum

    ''' <summary>
    ''' Skill 处理模式枚举
    ''' </summary>
    Public Enum SkillProcessingMode
        ''' <summary>
        ''' 传统 VBA 代码模式
        ''' </summary>
        VBACode
        ''' <summary>
        ''' 多模态 LLM 处理模式
        ''' </summary>
        MultimodalLLM
        ''' <summary>
        ''' 混合模式（根据输入自动选择）
        ''' </summary>
        Hybrid
        ''' <summary>
        ''' 本地 OCR 模式（使用 PaddleOCR/Windows OCR 等专业 OCR 引擎）
        ''' </summary>
        LocalOCR
        ''' <summary>
        ''' 【扩展6】Python 脚本执行模式（通过本地 Python 或远程 API 执行数据分析脚本）
        ''' </summary>
        PythonCode
        ''' <summary>
        ''' LLM 输出结构化 spec（JSON），前端 skillSpecEngine
        ''' 确定性渲染 VBA 后执行。LLM 负责规划（主题/样式/格式参数），模板负责代码正确性。
        ''' 此模式由前端 skill-pipeline 处理（executeSkillSpecMode），后端仅需透传该枚举值。
        ''' </summary>
        SkillSpec
    End Enum

    ''' <summary>
    ''' OCR 预处理配置
    ''' </summary>
    Public Class OcrPreprocessingConfig
        Public Property Deskew As Boolean = True
        Public Property Denoise As Boolean = True
        Public Property ContrastEnhancement As Boolean = True
    End Class

    ''' <summary>
    ''' OCR 后处理配置
    ''' </summary>
    Public Class OcrPostProcessingConfig
        Public Property UseLLM As Boolean = True
        Public Property LLMModel As String = ""
        Public Property ExtractionPrompt As String = ""
    End Class

    ''' <summary>
    ''' OCR 配置类（用于 LocalOCR 模式）
    ''' </summary>
    Public Class OcrConfig
        Public Property Engine As String = "PaddleOCR"
        Public Property FallbackEngine As String = "PaddleOCR"
        Public Property Preprocessing As OcrPreprocessingConfig = New OcrPreprocessingConfig()
        Public Property PostProcessing As OcrPostProcessingConfig = New OcrPostProcessingConfig()
    End Class

    ''' <summary>
    ''' 参数定义类
    ''' </summary>
    Public Class ParameterDefinition
        Public Property Name As String
        Public Property DataType As String
        Public Property Description As String
        Public Property IsRequired As Boolean = False
        Public Property DefaultValue As String
    End Class

    ''' <summary>
    ''' 统一工作流步骤定义类（v2：整合原 WorkflowStep + CustomWorkflowStep）
    ''' 设计原则：强类型属性优先，Config 字典作为逃生舱（用于领域特定步骤的扩展配置）
    ''' </summary>
    Public Class WorkflowStep
        Public Property StepId As String
        Public Property StepName As String
        Public Property StepType As WorkflowStepType
        Public Property Code As String
        Public Property ApiEndpoint As String
        Public Property ApiMethod As String = "POST"
        Public Property SubSkillName As String
        Public Property InputMapping As Dictionary(Of String, String)
        Public Property OutputMapping As Dictionary(Of String, String)
        Public Property Condition As String
        Public Property NextStepOnSuccess As String
        Public Property NextStepOnFailure As String
        Public Property RetryCount As Integer = 0
        Public Property Timeout As Integer = 30
        ''' <summary>
        ''' 【优先级2扩展】MCP 客户端名称（对应 mcp.json 中的 server key）
        ''' </summary>
        Public Property McpClientName As String
        ''' <summary>
        ''' 【优先级2扩展】MCP 工具名称（要调用的具体工具）
        ''' </summary>
        Public Property McpToolName As String
        ''' <summary>
        ''' 【优先级2扩展】MCP 工具调用参数（JSON 字符串，支持 {{变量}} 占位符）
        ''' </summary>
        Public Property McpArguments As String
        ''' <summary>
        ''' 【v2新增】通用配置字典（逃生舱），用于 OCRRecognition / WriteToExcel / CustomLLM 等领域特定步骤
        ''' 支持变量占位符 {{变量名}}，运行时会被变量存储替换
        ''' </summary>
        Public Property Config As Dictionary(Of String, Object)
        ''' <summary>
        ''' 【v2新增】最大循环迭代次数上限（仅 Loop 步骤生效，防止死循环）。默认 1000。
        ''' </summary>
        Public Property MaxLoopIterations As Integer = 1000

        Public Sub New()
            InputMapping = New Dictionary(Of String, String)
            OutputMapping = New Dictionary(Of String, String)
            Config = New Dictionary(Of String, Object)
        End Sub
    End Class

    ''' <summary>
    ''' 字段映射定义
    ''' </summary>
    Public Class FieldMapping
        Public Property SourceField As String
        Public Property TargetField As String
        Public Property TargetCell As String
        Public Property Format As String
    End Class
    
    ''' <summary>
    ''' Excel 字段映射配置
    ''' </summary>
    Public Class ExcelFieldMapping
        ''' <summary>
        ''' Excel 列字母（如 A, B, C）
        ''' </summary>
        Public Property Column As String

        ''' <summary>
        ''' 列标题
        ''' </summary>
        Public Property Header As String
    End Class

    ''' <summary>
    ''' Excel 配置（用于图片识别 Skill）
    ''' </summary>
    Public Class ExcelConfig
        ''' <summary>
        ''' 默认工作表名称
        ''' </summary>
        Public Property DefaultWorksheet As String

        ''' <summary>
        ''' 字段到 Excel 列的映射
        ''' </summary>
        Public Property FieldMapping As Dictionary(Of String, ExcelFieldMapping)

        ''' <summary>
        ''' 英文字段名到中文的映射
        ''' </summary>
        Public Property FieldNameMapping As Dictionary(Of String, String)

        ''' <summary>
        ''' 默认构造函数
        ''' </summary>
        Public Sub New()
            FieldMapping = New Dictionary(Of String, ExcelFieldMapping)()
            FieldNameMapping = New Dictionary(Of String, String)()
        End Sub
    End Class

    ''' <summary>
    ''' 单个声明式步骤（SkillCreator 默认产物 / STEPS_CONTRACT 后端镜像）。
    ''' 契约（前端 skill-domain.js STEPS_CONTRACT 为唯一事实源，此处为后端负载镜像）：
    '''   tool     (必填) MCP 工具名（direct 步）；形如 skill.Smart 的表单表示委托给另一 Skill
    '''   strategy (选填) direct（默认）或 code（显式 VBA）
    '''   params   (可选) 传给工具的参数对象（任意 JSON）
    '''   code     (仅 strategy 为 code 必填) 原始 VBA Sub，缺失即校验失败（不产生不可审计黑盒）
    '''   waitOn   (预留) 依赖上一步结果自动回填参数
    ''' </summary>
    Public Class DeclarativeStep
        ''' <summary>步骤显示名（缺省用 tool 名，供绿面板渲染）</summary>
        Public Property Name As String
        ''' <summary>步骤说明</summary>
        Public Property Description As String
        ''' <summary>MCP 工具名 / skill.Smart 形态的委托</summary>
        Public Property Tool As String
        ''' <summary>'direct'（默认）| 'code'</summary>
        Public Property Strategy As String = "direct"
        ''' <summary>传给工具的参数对象</summary>
        Public Property Params As Dictionary(Of String, Object) = New Dictionary(Of String, Object)()
        ''' <summary>仅 strategy='code' 必填：原始 VBA Sub</summary>
        Public Property Code As String = ""
        ''' <summary>预留：是否依赖上一步结果自动回填参数</summary>
        Public Property WaitOn As Boolean = False

        Public Sub New()
            Params = New Dictionary(Of String, Object)()
        End Sub
    End Class

    ''' <summary>
    ''' Skill 定义类
    ''' </summary>
    Public Class SkillDefinition
        ''' <summary>
        ''' Skill 唯一标识
        ''' </summary>
        Public Property Name As String

        ''' <summary>
        ''' 显示名称
        ''' </summary>
        Public Property DisplayName As String

        ''' <summary>
        ''' 功能描述
        ''' </summary>
        Public Property Description As String

        ''' <summary>
        ''' 匹配关键词数组
        ''' </summary>
        Public Property Keywords As String()

        ''' <summary>
        ''' 【v2新增】Skill 级工具白名单（前端 agent 模式生效）：声明后 LLM 仅可见白名单内的
        ''' MCP 工具，未声明（或空）则保持现状——全部业务 MCP 工具可见。
        ''' 支持精确工具名与尾缀通配符（如 "excel_format_range" / "excel_*"）。
        ''' 经 GetAllSkills camelCase 序列化为 'allowedTools' 透传前端。
        ''' </summary>
        Public Property AllowedTools As String()

        ''' <summary>
        ''' VBA 代码模板（传统模式使用）
        ''' </summary>
        Public Property TemplateCode As String

        ''' <summary>
        ''' 示例输入数组
        ''' </summary>
        Public Property Examples As String()

        ''' <summary>
        ''' 是否为通用 Skill（需要 AI 生成代码）
        ''' </summary>
        Public Property IsGeneric As Boolean = False

        ''' <summary>
        ''' 是否需要配置
        ''' </summary>
        Public Property NeedsConfig As Boolean = False

        ''' <summary>
        ''' 配置字段列表
        ''' </summary>
        Public Property ConfigFields As String()

        ''' <summary>
        ''' Skill 版本号（用于版本控制，兼容旧版整数版本）
        ''' 新代码请使用 SemanticVersion 属性
        ''' </summary>
        Public Property Version As Integer = 1

        ''' <summary>
        ''' 【v2新增】语义化版本号（SemVer，如 "1.3.0"）。
        ''' 规则：主版本号（不兼容API变更）.次版本号（向下兼容的功能性新增）.修订号（向下兼容的问题修正）
        ''' </summary>
        Public Property SemanticVersion As String = "1.0.0"

        ''' <summary>
        ''' 【v2新增】最低兼容插件版本（SemVer格式）。低于此版本的插件加载该Skill时将给出警告。
        ''' </summary>
        Public Property MinAppVersion As String = ""

        ''' <summary>
        ''' 是否为复合 Skill（包含多个工作流步骤）
        ''' </summary>
        Public Property IsComposite As Boolean = False

        ''' <summary>
        ''' 工作流步骤列表
        ''' </summary>
        Public Property WorkflowSteps As List(Of WorkflowStep)

        ''' <summary>
        ''' 【声明式步骤】SkillCreator 默认产物：有序工具序列（STEPS_CONTRACT 后端镜像）。
        ''' 前端 skill-pipeline 按 executionMode='steps'（或自检到 steps 非空）走声明式执行器；
        ''' 每步 = 真实 MCP 工具（strategy='direct'）或显式 VBA（strategy='code'）。
        ''' 经 GetAllSkills JSON 序列化为 camelCase 'steps' 透传给前端执行。
        ''' </summary>
        Public Property Steps As List(Of DeclarativeStep) = New List(Of DeclarativeStep)()

        ''' <summary>
        ''' 【执行模式】前端 skill-pipeline 的路由决策字段（skill-shared-util.js resolveExecutionMode 为消费方）。
        ''' 取值：steps（声明式工具序列）/ spec（LLM 规划 + VBA 渲染）/ agent（LLM 自主调白名单工具）/ fixed（默认固定流程）。
        ''' 序列化为 camelCase 'executionMode' 透传前端；未声明时前端按 steps 非空 → steps、否则按名称/processingMode 推导。
        ''' 必须有此属性：SkillJsonHelper 走反射遍历 SkillDefinition 属性，无属性则该字段在 JSON→md 往返中被静默丢弃。
        ''' </summary>
        Public Property ExecutionMode As String = ""

        ''' <summary>
        ''' 【v2新增】统一步骤视图：合并 WorkflowSteps（WorkflowEngine用） + CustomWorkflowSteps（旧格式字段，加载时已归一化为统一 WorkflowStep）
        ''' 遍历所有步骤时请优先使用此属性。写入时会写入 WorkflowSteps。
        ''' </summary>
        <JsonIgnore()>
        Public ReadOnly Property AllSteps As IEnumerable(Of WorkflowStep)
            Get
                Dim merged As New List(Of WorkflowStep)()
                If WorkflowSteps IsNot Nothing Then merged.AddRange(WorkflowSteps)
                If CustomWorkflowSteps IsNot Nothing AndAlso CustomWorkflowSteps.Count > 0 Then
                    merged.AddRange(CustomWorkflowSteps)
                End If
                Return merged
            End Get
        End Property

        ''' <summary>
        ''' 【v2新增】Skill依赖声明，启动时做可达性检查
        ''' </summary>
        Public Property Dependencies As SkillDependencyDeclaration = New SkillDependencyDeclaration()

        ''' <summary>
        ''' 【v2新增】预期执行性能SLA声明（仅供调度参考，非强制）
        ''' </summary>
        Public Property PerformanceSLA As SkillPerformanceSLA = New SkillPerformanceSLA()

        ''' <summary>
        ''' 输入参数定义列表
        ''' </summary>
        Public Property InputParameters As List(Of ParameterDefinition)

        ''' <summary>
        ''' 输出参数定义列表
        ''' </summary>
        Public Property OutputParameters As List(Of ParameterDefinition)

        ''' <summary>
        ''' 输入描述
        ''' </summary>
        Public Property InputDescription As String

        ''' <summary>
        ''' 输出描述
        ''' </summary>
        Public Property OutputDescription As String

        ''' <summary>
        ''' 参数定义列表（兼容属性）
        ''' </summary>
        Public Property Parameters As List(Of ParameterDefinition)
            Get
                Return InputParameters
            End Get
            Set(value As List(Of ParameterDefinition))
                InputParameters = value
            End Set
        End Property

        ''' <summary>
        ''' 处理模式（新增：VBA代码 / 多模态LLM / 混合）
        ''' </summary>
        Public Property ProcessingMode As SkillProcessingMode = SkillProcessingMode.VBACode

        ''' <summary>
        ''' 是否允许调用后端 MCP 工具确定性执行（数据处理家族：合并/拆分/关联等）。
        ''' 前端徽章据此显示"MCP 执行"而非"VBA 执行"。
        ''' </summary>
        Public Property AllowMcpTools As Boolean = False

        ''' <summary>
        ''' 多模态 LLM 系统提示词
        ''' </summary>
        Public Property LLMSystemPrompt As String

        ''' <summary>
        ''' 是否支持图片输入
        ''' </summary>
        Public Property SupportsImageInput As Boolean = False

        ''' <summary>
        ''' 是否支持文件输入
        ''' </summary>
        Public Property SupportsFileInput As Boolean = False

        ''' <summary>
        ''' 是否把 Skill 的输出自动导入 OCR 智能解析面板（如 GenerateOCRParseScript 生成的解析脚本）
        ''' </summary>
        Public Property AutoImportParseScript As Boolean = False

        ''' <summary>
        ''' 预期输出格式（如：json, text, markdown, vba）
        ''' </summary>
        Public Property OutputFormat As String = "text"
        
        ''' <summary>
        ''' 用户自定义Skill标记
        ''' </summary>
        Public Property IsCustom As Boolean = False

        ''' <summary>
        ''' 【同名合并防御】本定义对应的磁盘 .md 路径（内置来自 bin\Skills，自定义来自 AppData\CustomSkills\skills）
        ''' </summary>
        Public Property SourceFilePath As String

        ''' <summary>
        ''' 【同名合并防御】上面那个文件的最后写入时间。
        ''' 背景：用户点"快捷方式 / 启用开关 / 保存"会把内置技能**整份快照**提升到 AppData CustomSkills，
        ''' 快照连同当时的 Version 一起落盘，于是 SkillManager 合并时的 `custom.Version >= system.Version`
        ''' 恒为真，旧提示词会长期压住系统升级（实测 23 个内置技能全部被同版本快照遮蔽，
        ''' 改 Skills\*.md 完全不生效）。文件时间能区分"系统刚升级"与"用户刚编辑"：
        ''' 内置文件更新 → 行为字段保留内置版（见 SkillManager.Loading 的 MergeCustomSkill）。
        ''' </summary>
        Public Property SourceLastWriteTime As DateTime = DateTime.MinValue
        
        ''' <summary>
        ''' 创建时间
        ''' </summary>
        Public Property CreatedDate As DateTime = DateTime.Now
        
        ''' <summary>
        ''' 创建时间（别名）
        ''' </summary>
        Public Property CreatedAt As DateTime
            Get
                Return CreatedDate
            End Get
            Set(value As DateTime)
                CreatedDate = value
            End Set
        End Property
        
        ''' <summary>
        ''' 最后修改时间
        ''' </summary>
        Public Property LastModifiedDate As DateTime = DateTime.Now
        
        ''' <summary>
        ''' 最后修改时间（别名）
        ''' </summary>
        Public Property LastModified As DateTime
            Get
                Return LastModifiedDate
            End Get
            Set(value As DateTime)
                LastModifiedDate = value
            End Set
        End Property
        
        ''' <summary>
        ''' 创建者
        ''' </summary>
        Public Property CreatedBy As String = "User"
        
        ''' <summary>
        ''' 使用次数
        ''' </summary>
        Public Property UsageCount As Integer = 0
        
        ''' <summary>
        ''' 分类（用于技能分组）
        ''' </summary>
        Public Property Category As String = ""
        
        ''' <summary>
        ''' 标签列表
        ''' </summary>
        Public Property Tags As List(Of String) = New List(Of String)
        
        ''' <summary>
        ''' LLM用户提示词模板
        ''' </summary>
        Public Property LLMUserPromptTemplate As String
        
        ''' <summary>
        ''' 旧版自定义工作流步骤（v2 起统一存储为 WorkflowStep；加载时旧 CustomStepType 数据已归一化）
        ''' 保留属性名与 JSON/YAML 键名（customWorkflowSteps）以兼容历史数据与前端契约
        ''' </summary>
        Public Property CustomWorkflowSteps As List(Of WorkflowStep) = New List(Of WorkflowStep)()
        
        ''' <summary>
        ''' 字段映射配置
        ''' </summary>
        Public Property FieldMappings As List(Of FieldMapping) = New List(Of FieldMapping)()

        ''' <summary>
        ''' Excel 配置（用于图片识别 Skill）
        ''' </summary>
        Public Property ExcelConfig As ExcelConfig

        ''' <summary>
        ''' OCR 配置（用于 LocalOCR 模式）
        ''' </summary>
        Public Property OcrConfig As OcrConfig

        ''' <summary>
        ''' 【Skill 开关】是否启用该 Skill。
        ''' 默认 True（启用），设为 False 后该 Skill 不参与路由匹配、快捷入口、命令列表。
        ''' 支持两层覆盖：Markdown frontmatter 声明 isEnabled: false（内置停用）
        ''' + CustomSkillManager 用户保存的自定义版本（运行时开关）。
        ''' </summary>
        Public Property IsEnabled As Boolean = True

        ''' <summary>
        ''' 是否标记为快捷入口（显示在欢迎面板的快捷提示词区域）
        ''' </summary>
        Public Property IsQuickShortcut As Boolean = False

        ''' <summary>
        ''' 快捷入口图标（emoji 或图标名）
        ''' </summary>
        Public Property QuickShortcutIcon As String = "⚡"

        ''' <summary>
        ''' 快捷入口提示文本（点击后填入输入框的文字，默认使用 DisplayName）
        ''' </summary>
        Public Property QuickShortcutPrompt As String = ""

        ''' <summary>
        ''' 输出卡片后渲染的一组快捷操作按钮。
        ''' 任意 Skill 在 frontmatter 中声明 quickPrompts 列表即可自动解析并渲染，
        ''' 点击按钮直接执行对应 prompt（无需专用渲染逻辑）。
        ''' 序列化为 camelCase: quickPrompts
        ''' </summary>
        Public Property QuickPrompts As List(Of QuickPromptDefinition) = New List(Of QuickPromptDefinition)()

        ''' <summary>
        ''' SkillSpec 模式的参数声明列表。
        ''' 任意 Skill 在 frontmatter 中声明 specFields 即自动解析并传给前端：
        '''   1. skill 编辑器可查看/编辑这些字段（含 select 选项）
        '''   2. skill-spec-engine 的规划 prompt 以 specFields 为准（不再硬编码主题）
        ''' 序列化为 camelCase: specFields
        ''' </summary>
        Public Property SpecFields As List(Of SpecFieldDefinition) = New List(Of SpecFieldDefinition)()

        ''' <summary>
        ''' UI操作按钮配置（控制AI消息底部显示的按钮）
        ''' </summary>
        Public Property UIActions As UIActionConfig

        ' ==================== 【扩展1】RAG 知识库联动属性 ====================

        ''' <summary>
        ''' 是否启用 RAG 知识库增强（执行前自动检索相关知识注入 Prompt）
        ''' </summary>
        Public Property RagEnabled As Boolean = False

        ''' <summary>
        ''' RAG 目标知识库 ID（为空则使用 RagOrchestrator 默认知识库）
        ''' </summary>
        Public Property RagKbId As String = ""

        ''' <summary>
        ''' RAG 检索返回的 Top-K 文档数量（默认3）
        ''' </summary>
        Public Property RagTopK As Integer = 3

        ''' <summary>
        ''' RAG 检索分数阈值（低于此分数的结果被过滤，默认0.5）
        ''' </summary>
        Public Property RagScoreThreshold As Double = 0.5

        ''' <summary>
        ''' RAG 降级策略：true=无结果时降级到普通 LLM（不报错），false=无结果时返回"未找到相关知识"
        ''' </summary>
        Public Property RagFallbackToNormal As Boolean = True

        ' ==================== 【扩展2】管道编排属性 ====================

        ''' <summary>
        ''' 是否为管道型 Skill（包含多个串联的子 Skill 步骤）
        ''' </summary>
        Public Property IsPipeline As Boolean = False

        ''' <summary>
        ''' 管道步骤定义列表（每个步骤调用一个子 Skill，输出作为下一步输入）
        ''' </summary>
        Public Property PipelineSteps As List(Of PipelineStepDefinition) = New List(Of PipelineStepDefinition)()

        ' ==================== 【扩展3】事件触发配置 ====================

        ''' <summary>
        ''' 触发器配置（事件驱动/定时触发）
        ''' </summary>
        Public Property TriggerConfig As TriggerConfig

        ' ==================== 【扩展4】输出格式增强配置 ====================

        ''' <summary>
        ''' 图表配置（当 outputFormat="chart" 时使用）
        ''' </summary>
        Public Property ChartConfig As ChartConfig

        ''' <summary>
        ''' 数据透视表配置（当 outputFormat="pivotTable" 时使用）
        ''' </summary>
        Public Property PivotConfig As PivotTableConfig

        ' ==================== 【扩展5】安全级别 ====================

        ''' <summary>
        ''' 安全级别：Safe=直接执行 / NeedsConfirmation=需确认 / Critical=二次确认+审计日志
        ''' </summary>
        Public Property SecurityLevel As SkillSecurityLevel = SkillSecurityLevel.NeedsConfirmation

        ''' <summary>
        ''' 前端 Pipeline 路由目标（可选）。
        ''' 不设置或为空时走默认 skill-pipeline（VB后端执行）。
        ''' 设置为 "html" 则路由到 html-display-pipeline，"python" 路由到 python-pipeline，"text" 路由到 text-pipeline（agent-loop）。
        ''' 此属性仅用于前端 chat-orchestrator 路由决策，不影响 VB 后端执行逻辑。
        ''' </summary>
        Public Property Pipeline As String = ""

        ''' <summary>
        ''' 所需输入上下文声明（可选），用于前端路由时的上下文预检查。
        ''' 数组元素示例："selection"（需要选中区域）、"referenceRegions"（需要引用区域）、"attachments"（需要附件）。
        ''' </summary>
        Public Property InputContext As String() = New String() {}

        ''' <summary>
        ''' 默认构造函数
        ''' </summary>
        Public Sub New()
            Keywords = New String() {}
            Examples = New String() {}
            ConfigFields = New String() {}
            WorkflowSteps = New List(Of WorkflowStep)()
            InputParameters = New List(Of ParameterDefinition)()
            OutputParameters = New List(Of ParameterDefinition)()
            Version = 1
        End Sub

        ''' <summary>
        ''' 带参数的构造函数（简单 Skill）
        ''' </summary>
        Public Sub New(name As String, displayName As String, description As String,
                       keywords As String(), templateCode As String, examples As String())
            Me.Name = name
            Me.DisplayName = displayName
            Me.Description = description
            Me.Keywords = keywords
            Me.TemplateCode = templateCode
            Me.Examples = examples
            Me.IsGeneric = False
            Me.NeedsConfig = False
            Me.IsComposite = False
            Me.ConfigFields = New String() {}
            Me.WorkflowSteps = New List(Of WorkflowStep)()
            Me.InputParameters = New List(Of ParameterDefinition)()
            Me.OutputParameters = New List(Of ParameterDefinition)()
        End Sub
    End Class

    ''' <summary>
    ''' Skill 列表包装类（用于 JSON 反序列化）
    ''' </summary>
    Public Class SkillList
        ''' <summary>
        ''' Skill 列表全局版本号
        ''' </summary>
        Public Property Version As Integer = 1

        ''' <summary>
        ''' Skill 列表
        ''' </summary>
        Public Property Skills As List(Of SkillDefinition)

        Public Sub New()
            Version = 1
            Skills = New List(Of SkillDefinition)()
        End Sub
    End Class

    ''' <summary>
    ''' Skill 识别结果
    ''' </summary>
    Public Class SkillRecognitionResult
        ''' <summary>
        ''' 匹配到的 Skill
        ''' </summary>
        Public Property MatchedSkill As SkillDefinition

        ''' <summary>
        ''' 匹配置信度（0-1）
        ''' </summary>
        Public Property Confidence As Double

        ''' <summary>
        ''' 是否匹配成功
        ''' </summary>
        Public Property IsMatch As Boolean

        ''' <summary>
        ''' 匹配的详细信息
        ''' </summary>
        Public Property MatchDetails As String

        Public Sub New()
            Confidence = 0
            IsMatch = False
            MatchDetails = ""
        End Sub
    End Class

    ' ==================== 【扩展1-6】新增类型定义 ====================

    ''' <summary>
    ''' 【扩展5】Skill 安全级别枚举
    ''' </summary>
    Public Enum SkillSecurityLevel
        ''' <summary>安全操作，直接执行（如只读分析、格式化高亮）</summary>
        Safe
        ''' <summary>需用户确认后执行（默认：写入Excel、创建图表等）</summary>
        NeedsConfirmation
        ''' <summary>高危操作：二次确认 + 审计日志（MCP写操作、发送邮件、删除数据）</summary>
        Critical
    End Enum

    ''' <summary>
    ''' 【扩展2】管道步骤定义 - 用于 Pipeline 模式的子 Skill 调用配置
    ''' </summary>
    Public Class PipelineStepDefinition
        ''' <summary>步骤唯一标识</summary>
        Public Property StepId As String
        ''' <summary>步骤显示名称</summary>
        Public Property StepName As String
        ''' <summary>要调用的子 Skill 名称</summary>
        Public Property SkillName As String
        ''' <summary>输入字段映射（从上游输出到本步输入）{"本步输入名": "上游输出名"}</summary>
        Public Property InputMapping As Dictionary(Of String, String)
        ''' <summary>输出过滤（JSONPath 表达式，可选）</summary>
        Public Property OutputFilter As String
        ''' <summary>成功时下一步骤 ID</summary>
        Public Property NextStepOnSuccess As String
        ''' <summary>失败时下一步骤 ID（空则终止管道）</summary>
        Public Property NextStepOnFailure As String
        ''' <summary>重试次数</summary>
        Public Property RetryCount As Integer = 0
        ''' <summary>超时秒数</summary>
        Public Property TimeoutSeconds As Integer = 30
        ''' <summary>条件表达式（可选，满足才执行此步骤）</summary>
        Public Property Condition As String

        Public Sub New()
            InputMapping = New Dictionary(Of String, String)()
        End Sub
    End Class

    ''' <summary>
    ''' 【扩展3】触发器配置 - 支持事件驱动和定时触发
    ''' </summary>
    Public Class TriggerConfig
        ''' <summary>触发器类型</summary>
        Public Property TriggerType As TriggerTypeEnum = TriggerTypeEnum.Manual
        ''' <summary>事件类型（TriggerType=Event 时使用）</summary>
        Public Property EventType As String = ""
        ''' <summary>过滤条件（如 "targetColumn='B' &amp;&amp; targetSheet='订单表'"）</summary>
        Public Property FilterCondition As String = ""
        ''' <summary>防抖时间毫秒数（防止频繁触发，默认1000ms）</summary>
        Public Property DebounceMs As Integer = 1000
        ''' <summary>是否启用</summary>
        Public Property Enabled As Boolean = False
        ' --- 定时触发相关 ---
        ''' <summary>定时器间隔（秒）</summary>
        Public Property TimerIntervalSeconds As Integer = 0
        ''' <summary>定时器是否在启动时立即执行一次</summary>
        Public Property FireOnStart As Boolean = False
    End Class

    ''' <summary>
    ''' 触发器类型枚举
    ''' </summary>
    Public Enum TriggerTypeEnum
        ''' <summary>手动触发（默认）</summary>
        Manual
        ''' <summary>Excel 事件触发（Change/Open/Close/SelectionChange 等）</summary>
        [Event]
        ''' <summary>定时器触发</summary>
        Timer
    End Enum

    ''' <summary>
    ''' 【扩展4】图表配置 - 当 outputFormat="chart" 时使用
    ''' </summary>
    Public Class ChartConfig
        ''' <summary>图表类型：line/bar/pie/scatter/area/radar</summary>
        Public Property ChartType As String = "bar"
        ''' <summary>标题</summary>
        Public Property Title As String = ""
        ''' <summary>X 轴数据来源字段名</summary>
        Public Property XAxisField As String = ""
        ''' <summary>Y 轴数据来源字段名（可多个逗号分隔）</summary>
        Public Property YAxisFields As String = ""
        ''' <summary>目标位置（如 "F2" 或 "新建工作表"）</summary>
        Public Property TargetRange As String = ""
        ''' <summary>宽度（像素）</summary>
        Public Property Width As Integer = 400
        ''' <summary>高度（像素）</summary>
        Public Property Height As Integer = 250
        ''' <summary>是否显示图例</summary>
        Public Property ShowLegend As Boolean = True
    End Class

    ''' <summary>
    ''' 【扩展4】数据透视表配置 - 当 outputFormat="pivotTable" 时使用
    ''' </summary>
    Public Class PivotTableConfig
        ''' <summary>数据源范围</summary>
        Public Property SourceRange As String = ""
        ''' <summary>目标位置</summary>
        Public Property TargetRange As String = ""
        ''' <summary>行字段（列名列表）</summary>
        Public Property RowFields As List(Of String) = New List(Of String)()
        ''' <summary>列字段（列名列表）</summary>
        Public Property ColumnFields As List(Of String) = New List(Of String)()
        ''' <summary>数据字段（列名+聚合函数）</summary>
        Public Property DataFields As List(Of PivotDataField) = New List(Of PivotDataField)()
        ''' <summary>筛选器字段</summary>
        Public Property FilterFields As List(Of String) = New List(Of String)()
    End Class

    ''' <summary>
    ''' 数据透视表数据字段定义
    ''' </summary>
    Public Class PivotDataField
        Public Property FieldName As String
        Public Property Aggregation As String = "Sum" ' Sum/Average/Count/Max/Min
    End Class

    ''' <summary>
    ''' UI操作按钮配置
    ''' </summary>
    Public Class UIActionConfig
        ''' <summary>
        ''' 显示复制按钮
        ''' </summary>
        Public Property ShowCopy As Boolean = True

        ''' <summary>
        ''' 显示保存为MD文件按钮
        ''' </summary>
        Public Property ShowSaveMd As Boolean = False

        ''' <summary>
        ''' 显示重新生成按钮
        ''' </summary>
        Public Property ShowRegenerate As Boolean = True

        ''' <summary>
        ''' 自定义按钮列表（JSON格式）
        ''' </summary>
        Public Property CustomButtons As List(Of UICustomButton) = New List(Of UICustomButton)()
    End Class

    ''' <summary>
    ''' 自定义UI按钮
    ''' </summary>
    Public Class UICustomButton
        ''' <summary>
        ''' 按钮ID
        ''' </summary>
        Public Property Id As String

        ''' <summary>
        ''' 按钮显示文本
        ''' </summary>
        Public Property Label As String

        ''' <summary>
        ''' 按钮图标
        ''' </summary>
        Public Property Icon As String

        ''' <summary>
        ''' 点击动作（JavaScript代码或Bridge调用）
        ''' </summary>
        Public Property Action As String
    End Class

    ''' <summary>
    ''' 单条快捷操作定义。
    ''' frontmatter 示例：
    '''   quickPrompts:
    '''     - label: 商务蓝
    '''       prompt: 用商务蓝主题美化表格
    ''' </summary>
    Public Class QuickPromptDefinition
        ''' <summary>
        ''' 按钮显示文本
        ''' </summary>
        Public Property Label As String

        ''' <summary>
        ''' 点击后发送的提示词文本（将填入输入框并直接执行）
        ''' </summary>
        Public Property Prompt As String
    End Class

    ''' <summary>
    ''' 单条 spec 字段声明。
    ''' frontmatter 示例：
    '''   specFields:
    '''     - name: theme
    '''       label: 主题风格
    '''       type: select
    '''       options:
    '''         - business_blue
    '''         - tech_green
    '''       defaultValue: business_blue
    ''' </summary>
    Public Class SpecFieldDefinition
        ''' <summary>
        ''' 字段标识（如 theme、headerStyle、alignment）
        ''' </summary>
        Public Property Name As String

        ''' <summary>
        ''' 显示标签（如 主题风格）
        ''' </summary>
        Public Property Label As String

        ''' <summary>
        ''' 字段类型：select / boolean / string / number
        ''' </summary>
        Public Property Type As String = "string"

        ''' <summary>
        ''' select 类型的可选值列表
        ''' </summary>
        Public Property Options As List(Of String) = New List(Of String)()

        ''' <summary>
        ''' 默认值
        ''' </summary>
        Public Property DefaultValue As String
    End Class

    ''' <summary>
    ''' Skill依赖声明：启动加载时做可达性检查
    ''' </summary>
    Public Class SkillDependencyDeclaration
        ''' <summary>所需Bridge名称，如 ["ExcelDataBridge", "OcrServiceBridge"]</summary>
        Public Property Bridges As List(Of String) = New List(Of String)()
        ''' <summary>所需Excel工具名称，如 ["excel_read_range", "excel_format_range"]</summary>
        Public Property Tools As List(Of String) = New List(Of String)()
        ''' <summary>所需MCP Server key（对应mcp.json），如 ["filesystem"]</summary>
        Public Property McpServers As List(Of String) = New List(Of String)()
        ''' <summary>所依赖的其他Skill Name（IsPipeline模式时需声明）</summary>
        Public Property SubSkills As List(Of String) = New List(Of String)()
        ''' <summary>所需LLM模型ID，如 ["gpt-4o-mini"]；为空则不做检查</summary>
        Public Property LlmModels As List(Of String) = New List(Of String)()
    End Class

    ''' <summary>
    ''' Skill性能SLA声明（仅供调度和诊断参考，不作为硬性强制）
    ''' </summary>
    Public Class SkillPerformanceSLA
        ''' <summary>预期执行耗时（毫秒），-1 表示未评估</summary>
        Public Property ExpectedLatencyMs As Integer = -1
        ''' <summary>预期峰值内存增量（MB），-1 表示未评估</summary>
        Public Property MaxMemoryMB As Integer = -1
        ''' <summary>是否需要GPU（如本地大模型OCR/推理）</summary>
        Public Property RequiresGpu As Boolean = False
        ''' <summary>并发安全等级：None=不可重入, ReadOnly=可多并发读, Full=完全并发安全</summary>
        Public Property ConcurrencySafety As ConcurrencySafetyLevel = ConcurrencySafetyLevel.None
    End Class

    ''' <summary>并发安全等级</summary>
    Public Enum ConcurrencySafetyLevel
        ''' <summary>不可重入：全局串行执行（写操作默认）</summary>
        None = 0
        ''' <summary>只读安全：可多Skill并发读取，但不能与写者并发</summary>
        [ReadOnly] = 1
        ''' <summary>完全并发安全：无状态可任意并发</summary>
        Full = 2
    End Enum

End Namespace
