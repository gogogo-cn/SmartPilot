Option Strict On
Option Explicit On
Option Infer On
Imports System

Namespace Utils

    ''' <summary>
    ''' 配置常量定义类 - 集中管理所有配置键名称和默认值
    ''' </summary>
    Public Class ConfigConstants

        ' ========== LLM 相关配置 ==========
        Public Class LLM
            Public Const PROVIDER As String = "LLMProvider"
            Public Const MODEL As String = "LLMModel"
            Public Const API_KEY As String = "LLMApiKey"
            Public Const API_URL As String = "LLMApiUrl"
            Public Const TEMPERATURE As String = "LLMTemperature"
            Public Const MAX_TOKENS As String = "LLMMaxTokens"
            Public Const TIMEOUT As String = "LLMTimeout"
            Public Const HTTP_TIMEOUT As String = "LLMHttpTimeout"

            ' LLM API URLs
            Public Const API_URL_OPENAI As String = "LLMApiUrl_OpenAI"
            Public Const API_URL_AZURE As String = "LLMApiUrl_Azure"
            Public Const API_URL_QIANWEN As String = "LLMApiUrl_Qianwen"
            Public Const API_URL_ZHIPU As String = "LLMApiUrl_Zhipu"
            Public Const API_URL_WENXIN As String = "LLMApiUrl_Wenxin"
            Public Const API_URL_MOONSHOT As String = "LLMApiUrl_Moonshot"
            Public Const API_URL_VOLCENGINE As String = "LLMApiUrl_Volcengine"

            ' 默认值
            Public Const DEFAULT_PROVIDER As String = "openai"
            Public Const DEFAULT_MODEL As String = "gpt-4o"
            Public Const DEFAULT_TEMPERATURE As Double = 0.7
            Public Const DEFAULT_MAX_TOKENS As Integer = 4096
            Public Const DEFAULT_TIMEOUT As Integer = 120
            Public Const DEFAULT_HTTP_TIMEOUT As Integer = 120

            ' 重试配置
            Public Const DEFAULT_MAX_RETRIES As Integer = 2
            Public Const DEFAULT_RETRY_INITIAL_DELAY_MS As Integer = 1000

            ' 默认 API URLs
            Public Const DEFAULT_OPENAI_URL As String = "https://api.openai.com/v1/chat/completions"
            Public Const DEFAULT_AZURE_URL As String = ""
            Public Const DEFAULT_QIANWEN_URL As String = "https://dashscope-intl.aliyuncs.com/compatible-mode/v1/chat/completions"
            Public Const DEFAULT_ZHIPU_URL As String = "https://open.bigmodel.cn/api/paas/v4/chat/completions"
            Public Const DEFAULT_WENXIN_URL As String = "https://qianfan.baidubce.com/v2/chat/completions"
            Public Const DEFAULT_MOONSHOT_URL As String = "https://api.moonshot.cn/v1/chat/completions"
            Public Const DEFAULT_VOLCENGINE_URL As String = "https://ark.cn-beijing.volces.com/api/v3/chat/completions"

            ' 可用 LLM 提供商清单（单一事实源，与前端 PROVIDER_CONFIGS 一致，供后端暴露给 UI）
            Public Shared ReadOnly AVAILABLE_PROVIDERS As String() = New String() {
                "ollama", "openai", "azure", "wenxin", "qianwen", "zhipu", "xunfei", "moonshot", "deepseek", "volcengine", "custom"
            }
        End Class

        ' ========== OCR 相关配置 ==========
        Public Class OCR
            Public Const API_KEY As String = "BaiduOcrApiKey"
            Public Const SECRET_KEY As String = "BaiduOcrSecretKey"
            Public Const PROCESS_TIMEOUT As String = "OcrProcessTimeout"
            Public Const DEFAULT_PROCESS_TIMEOUT As Integer = 2000
        End Class

        ' ========== Ollama 相关配置 ==========
        Public Class Ollama
            Public Const API_URL As String = "OllamaApiUrl"
            Public Const TAGS_URL As String = "OllamaTagsUrl"
            Public Const DEFAULT_API_URL As String = "http://localhost:11434/api/chat"
            Public Const DEFAULT_TAGS_URL As String = "http://localhost:11434/api/tags"
            Public Const DEFAULT_MODEL As String = "qwen3-vl:4b"
        End Class

        ' ========== llama.cpp 相关配置（OpenAI 兼容本地服务，默认端口 8080） ==========
        Public Class LlamaCpp
            Public Const MODELS_URL As String = "LlamaCppModelsUrl"
            Public Const DEFAULT_MODELS_URL As String = "http://localhost:8080/v1/models"
        End Class

        Public Class RAG
            Public Const LOCAL_BGE_ENDPOINT As String = "RagLocalBgeEndpoint"
            Public Const QDRANT_ENDPOINT As String = "RagQdrantEndpoint"
            Public Const DEFAULT_LOCAL_BGE_ENDPOINT As String = "http://localhost:11434"
            Public Const DEFAULT_QDRANT_ENDPOINT As String = "http://localhost:6333"
        End Class

        ' ========== Translation 相关配置 ==========
        Public Class Translation
            Public Const TIMEOUT As String = "TranslationTimeout"
            Public Const APP_ID As String = "TranslationAppId"
            Public Const APP_KEY As String = "TranslationAppKey"
            Public Const DEFAULT_TIMEOUT As Integer = 30000
        End Class

        ' ========== Excel 相关配置 ==========
        Public Class Excel
            Public Const MAX_COLUMNS As String = "ExcelMaxColumns"
            Public Const MAX_SEARCH_ROWS As String = "ExcelMaxSearchRows"
            Public Const DEFAULT_MAX_COLUMNS As Integer = 100
            Public Const DEFAULT_MAX_SEARCH_ROWS As Integer = 1000
        End Class

        ' ========== Workflow 相关配置 ==========
        Public Class Workflow
            Public Const CHECKPOINT_INTERVAL As String = "WorkflowCheckpointInterval"
            Public Const HTTP_TIMEOUT As String = "WorkflowHttpTimeout"
            Public Const MAX_STEP_RETRIES As String = "WorkflowMaxStepRetries"
            Public Const SKILLS_FILE_PATH As String = "WorkflowSkillsFilePath"
            Public Const STATE_FILE_PATH As String = "WorkflowStateFilePath"
            Public Const VERSION As String = "WorkflowVersion"

            Public Const DEFAULT_CHECKPOINT_INTERVAL As Integer = 30
            Public Const DEFAULT_HTTP_TIMEOUT As Integer = 60
            Public Const DEFAULT_MAX_STEP_RETRIES As Integer = 3
            Public Const DEFAULT_VERSION As Integer = 8
        End Class

        ' ========== 数据类型检测配置 ==========
        Public Class DataType
            Public Const MIN_MATCH_FIELDS As String = "DataTypeMinMatchFields"
            Public Const ID_CARD_FIELDS As String = "DataTypeIDCardFieldNames"
            Public Const CONFIG_FILE As String = "DataTypeConfigFile"

            Public Const DEFAULT_MIN_MATCH_FIELDS As Integer = 2

            ' [已弃用 - 识别类Skill已移除]
            'Public Shared ReadOnly DEFAULT_ID_CARD_FIELDS As String() = New String() {
            '    "姓名", "性别", "民族", "出生日期", "出生",
            '    "住址", "公民身份号码", "身份证号码", "身份证号",
            '    "签发机关", "有效期限",
            '    "职位", "公司名称", "公司", "电话", "手机", "邮箱", "地址", "部门"
            '}
        End Class

        ' ========== UI 相关配置 ==========
        Public Class UI
            Public Const VIRTUAL_HOST_NAME As String = "UIVirtualHostName"
            Public Const SIDEBAR_WIDTH As String = "UISidebarWidth"
            Public Const DEFAULT_SIDEBAR_WIDTH As Integer = 420
            Public Const MIN_SIDEBAR_WIDTH As Integer = 28
            Public Const MAX_SIDEBAR_WIDTH As Integer = 1000
            Public Const DEFAULT_VIRTUAL_HOST As String = "app.excelai.local"
        End Class

        ' ========== Custom Skill 配置 ==========
        Public Class CustomSkill
            Public Const SYSTEM_PROMPT As String = "CustomSkillSystemPrompt"
            Public Const DEFAULT_SYSTEM_PROMPT As String = "你是一个Excel AI助手。"
            Public Const FORMAT_CURRENCY As String = "CustomSkillFormatCurrency"
            Public Const FORMAT_DATE As String = "CustomSkillFormatDate"
            Public Const FORMAT_NUMBER As String = "CustomSkillFormatNumber"

            Public Const DEFAULT_FORMAT_CURRENCY As String = "¥#,##0.00"
            Public Const DEFAULT_FORMAT_DATE As String = "yyyy-mm-dd"
            Public Const DEFAULT_FORMAT_NUMBER As String = "0.00"
        End Class

        ' ========== 图片识别配置 ==========
        Public Class ImageRecognition
            Public Const API_URL As String = "ImageRecognitionApiUrl"
            Public Const DEFAULT_API_URL As String = "https://api.moonshot.cn/v1/chat/completions"
        End Class

        ' ========== 路径配置 ==========
        Public Class Paths
            ' 基础路径
            Public Const APP_DATA_FOLDER As String = "ExcelAIVSTOAddIn"
            
            ' 配置文件路径
            Public Const CONFIG_FILE As String = "config.ini"
            Public Const SKILLS_FILE As String = "Skills.json"
            Public Const WORKFLOW_STATE_FILE As String = "WorkflowState.json"
            Public Const CUSTOM_SKILLS_FOLDER As String = "CustomSkills"
            
            ' Python 路径
            Public Const PYTHON_BASE_FOLDER As String = "Programs\Python"
            Public Const PYTHON_EXE As String = "python.exe"
            Public Shared ReadOnly PYTHON_VERSIONS As String() = New String() {"Python312", "Python311", "Python310", "Python39"}
            
            ' 日志文件
            Public Const DEBUG_LOG_FILE As String = "skill_debug.log"
        End Class

        Public Class MCP
            Public Const SERVER_PORT As String = "McpServerPort"
            Public Const UI_SERVER_PORT As String = "UiServerPort"
            Public Const DEFAULT_SERVER_PORT As Integer = 5173
            Public Const DEFAULT_UI_SERVER_PORT As Integer = 8088
        End Class

        ' ========== Python 按需下载配置 ==========
        Public Class PythonBootstrapper
            Public Const DOWNLOAD_URL_PREFIX As String = "https://www.python.org/ftp/python"
            Public Const EMBEDDED_PACKAGE As String = "python-{0}-embed-amd64.zip"
            Public Const MIN_PYTHON_VERSION As String = "3.9.0"
            Public Const RECOMMENDED_PYTHON_VERSION As String = "3.12.7"
            Public Const BOOTSTRAP_TIMEOUT_MS As Integer = 300000
            Public Const DOWNLOAD_RETRY_COUNT As Integer = 3
            Public Const DOWNLOAD_RETRY_DELAY_MS As Integer = 2000
            Public Const PYTHON_DIR_NAME As String = "PythonEmbed"
        End Class

        ' ========== 本地服务端口（与前端 app-endpoints.js PORTS 对齐，单一事实源） ==========
        Public Class LocalServices
            ''' <summary>VSTO 内置 HTTP 服务端口（番茄钟/任务规划面板 API）</summary>
            Public Const VSTO_HTTP_PORT As Integer = 9528
            ''' <summary>独立任务规划服务端口（Standalone Agent）</summary>
            Public Const STANDALONE_AGENT_PORT As Integer = 9529
            ''' <summary>独立 OCR 服务端口。只监听 127.0.0.1，Excel/Word/PPT 共用。</summary>
            Public Const OCR_HOST_PORT As Integer = 9530
        End Class

    End Class

End Namespace