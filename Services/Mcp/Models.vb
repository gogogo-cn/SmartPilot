Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json.Serialization

Namespace Services.Mcp

    '#Region "JSON-RPC 2.0 消息模型"

    ''' <summary>
    ''' JSON-RPC 2.0 请求消息
    ''' </summary>
    Public Class JsonRpcRequest
        <JsonProperty("jsonrpc")>
        Public Property JsonRpc As String = "2.0"

        <JsonProperty("id")>
        Public Property Id As Object

        <JsonProperty("method")>
        Public Property Method As String

        <JsonProperty("params")>
        Public Property Params As JToken
    End Class

    ''' <summary>
    ''' JSON-RPC 2.0 成功响应
    ''' </summary>
    Public Class JsonRpcResponse
        <JsonProperty("jsonrpc")>
        Public Property JsonRpc As String = "2.0"

        <JsonProperty("id")>
        Public Property Id As Object

        <JsonProperty("result")>
        Public Property Result As JToken
    End Class

    ''' <summary>
    ''' JSON-RPC 2.0 错误响应
    ''' </summary>
    Public Class JsonRpcError
        <JsonProperty("jsonrpc")>
        Public Property JsonRpc As String = "2.0"

        <JsonProperty("id")>
        Public Property Id As Object

        <JsonProperty("error")>
        Public Property [Error] As JsonRpcErrorDetail
    End Class

    ''' <summary>
    ''' JSON-RPC 错误详情
    ''' </summary>
    Public Class JsonRpcErrorDetail
        <JsonProperty("code")>
        Public Property Code As Integer

        <JsonProperty("message")>
        Public Property Message As String

        <JsonProperty("data")>
        Public Property Data As Object
    End Class

    ''' <summary>
    ''' JSON-RPC 通知（无 id 字段）
    ''' </summary>
    Public Class JsonRpcNotification
        <JsonProperty("jsonrpc")>
        Public Property JsonRpc As String = "2.0"

        <JsonProperty("method")>
        Public Property Method As String

        <JsonProperty("params")>
        Public Property Params As JToken
    End Class

    '#End Region

    '#Region "MCP 协议常量"

    ''' <summary>
    ''' MCP 协议版本和错误码
    ''' </summary>
    Public NotInheritable Class McpConstants
        Public Const PROTOCOL_VERSION As String = "2024-11-05"
        Public Const IMPLEMENTATION_NAME As String = "ExcelAIVSTOAddIn-MCP"
        Public Const IMPLEMENTATION_VERSION As String = "1.0.0"

        ' JSON-RPC 错误码
        Public Const PARSE_ERROR As Integer = -32700
        Public Const INVALID_REQUEST As Integer = -32600
        Public Const METHOD_NOT_FOUND As Integer = -32601
        Public Const INVALID_PARAMS As Integer = -32602
        Public Const INTERNAL_ERROR As Integer = -32603

        ' MCP 方法名
        Public Const METHOD_INITIALIZE As String = "initialize"
        Public Const METHOD_INITIALIZED As String = "notifications/initialized"
        Public Const METHOD_PING As String = "ping"
        Public Const METHOD_TOOLS_LIST As String = "tools/list"
        Public Const METHOD_TOOLS_CALL As String = "tools/call"
        Public Const METHOD_RESOURCES_LIST As String = "resources/list"
        Public Const METHOD_RESOURCES_READ As String = "resources/read"
        Public Const METHOD_RESOURCES_TEMPLATES_LIST As String = "resources/templates/list"
        Public Const METHOD_PROMPTS_LIST As String = "prompts/list"
        Public Const METHOD_PROMPTS_GET As String = "prompts/get"
        Public Const METHOD_CANCELLED As String = "notifications/cancelled"
        Public Const METHOD_PROGRESS As String = "notifications/progress"
    End Class

    '#End Region

    '#Region "MCP 数据模型"

    ''' <summary>
    ''' Server 能力声明
    ''' </summary>
    Public Class ServerCapabilities
        <JsonProperty("tools")>
        Public Property Tools As ToolsCapability

        <JsonProperty("resources")>
        Public Property Resources As ResourcesCapability

        <JsonProperty("prompts")>
        Public Property Prompts As PromptsCapability

        <JsonProperty("experimental")>
        Public Property Experimental As Dictionary(Of String, Object)

        Public Sub New()
            Tools = New ToolsCapability()
            Resources = New ResourcesCapability()
            Prompts = New PromptsCapability()
        End Sub
    End Class

    ''' <summary>
    ''' Tools 能力声明
    ''' </summary>
    Public Class ToolsCapability
        <JsonProperty("listChanged")>
        Public Property ListChanged As Boolean = False
    End Class

    ''' <summary>
    ''' Resources 能力声明
    ''' </summary>
    Public Class ResourcesCapability
        <JsonProperty("subscribe")>
        Public Property Subscribe As Boolean = False

        <JsonProperty("listChanged")>
        Public Property ListChanged As Boolean = False
    End Class

    ''' <summary>
    ''' Prompts 能力声明
    ''' </summary>
    Public Class PromptsCapability
        <JsonProperty("listChanged")>
        Public Property ListChanged As Boolean = False
    End Class

    ''' <summary>
    ''' Client 能力声明
    ''' </summary>
    Public Class ClientCapabilities
        <JsonProperty("roots")>
        Public Property Roots As RootsCapability

        Public Sub New()
            Roots = New RootsCapability()
        End Sub
    End Class

    ''' <summary>
    ''' Roots 能力
    ''' </summary>
    Public Class RootsCapability
        <JsonProperty("listChanged")>
        Public Property ListChanged As Boolean = False
    End Class

    ''' <summary>
    ''' Server 信息
    ''' </summary>
    Public Class ServerInfo
        <JsonProperty("name")>
        Public Property Name As String

        <JsonProperty("version")>
        Public Property Version As String
    End Class

    ''' <summary>
    ''' Client 信息
    ''' </summary>
    Public Class ClientInfo
        <JsonProperty("name")>
        Public Property Name As String

        <JsonProperty("version")>
        Public Property Version As String
    End Class

    '#End Region

    '#Region "Initialize 参数"

    ''' <summary>
    ''' Initialize 请求参数
    ''' </summary>
    Public Class InitializeParams
        <JsonProperty("protocolVersion")>
        Public Property ProtocolVersion As String

        <JsonProperty("capabilities")>
        Public Property Capabilities As ClientCapabilities

        <JsonProperty("clientInfo")>
        Public Property ClientInfo As ClientInfo
    End Class

    ''' <summary>
    ''' Initialize 响应结果
    ''' </summary>
    Public Class InitializeResult
        <JsonProperty("protocolVersion")>
        Public Property ProtocolVersion As String

        <JsonProperty("capabilities")>
        Public Property Capabilities As ServerCapabilities

        <JsonProperty("serverInfo")>
        Public Property ServerInfo As ServerInfo

        <JsonProperty("instructions")>
        Public Property Instructions As String
    End Class

    '#End Region

    '#Region "Tool 模型"

    ''' <summary>
    ''' MCP Tool 定义
    ''' </summary>
    Public Class McpTool
        <JsonProperty("name")>
        Public Property Name As String

        <JsonProperty("description")>
        Public Property Description As String

        <JsonProperty("inputSchema")>
        Public Property InputSchema As McpToolInputSchema

        Public Sub New()
            InputSchema = New McpToolInputSchema()
        End Sub

        Public Sub New(name As String, description As String, schema As McpToolInputSchema)
            Me.Name = name
            Me.Description = description
            Me.InputSchema = schema
        End Sub
    End Class

    ''' <summary>
    ''' Tool 输入参数 JSON Schema
    ''' </summary>
    Public Class McpToolInputSchema
        <JsonProperty("type")>
        Public Property Type As String = "object"

        <JsonProperty("properties")>
        Public Property Properties As Dictionary(Of String, McpToolProperty)

        <JsonProperty("required")>
        Public Property Required As List(Of String)

        Public Sub New()
            Properties = New Dictionary(Of String, McpToolProperty)()
            Required = New List(Of String)()
        End Sub
    End Class

    ''' <summary>
    ''' Tool 属性定义
    ''' </summary>
    Public Class McpToolProperty
        <JsonProperty("type")>
        Public Property Type As String

        <JsonProperty("description")>
        Public Property Description As String

        <JsonProperty("enum")>
        Public Property [Enum] As List(Of String)

        ''' <summary>
        ''' 无参构造函数（JSON 反序列化需要）
        ''' </summary>
        Public Sub New()
        End Sub

        Public Sub New(type As String, description As String)
            Me.Type = type
            Me.Description = description
        End Sub

        Public Sub New(type As String, description As String, enumValues As List(Of String))
            Me.Type = type
            Me.Description = description
            Me.[Enum] = enumValues
        End Sub
    End Class

    ''' <summary>
    ''' Tools/List 请求参数
    ''' </summary>
    Public Class ToolsListParams
        <JsonProperty("cursor")>
        Public Property Cursor As String
    End Class

    ''' <summary>
    ''' Tools/List 响应结果
    ''' </summary>
    Public Class ToolsListResult
        <JsonProperty("tools")>
        Public Property Tools As List(Of McpTool)

        <JsonProperty("nextCursor")>
        Public Property NextCursor As String

        Public Sub New()
            Tools = New List(Of McpTool)()
        End Sub

        Public Sub New(tools As List(Of McpTool))
            Me.Tools = tools
        End Sub
    End Class

    ''' <summary>
    ''' Tools/Call 请求参数
    ''' </summary>
    Public Class ToolsCallParams
        <JsonProperty("name")>
        Public Property Name As String

        <JsonProperty("arguments")>
        Public Property Arguments As Dictionary(Of String, Object)
    End Class

    ''' <summary>
    ''' Tools/Call 响应结果
    ''' </summary>
    Public Class ToolsCallResult
        <JsonProperty("content")>
        Public Property Content As List(Of McpContentItem)

        <JsonProperty("isError")>
        Public Property IsError As Boolean = False

        <JsonProperty("taskCompleted")>
        Public Property TaskCompleted As Boolean = True

        ''' <summary>
        ''' 结构化产物摘要（通用框架层字段）
        ''' 前端 agent-loop 自动提取后存在 turn._artifacts 数组，
        ''' 每轮注入 system prompt 让 LLM 知道自己已经创建了什么。
        ''' 典型内容：chart → {name, anchor, type, dataSource}
        '''           aggregate → {targetSheet, dimColumn, rowCount}
        '''           write_range → {sheet, range, rowCount}
        ''' 非看板专用——任何多步骤复合操作受益。
        ''' </summary>
        <JsonProperty("artifact")>
        Public Property Artifact As Dictionary(Of String, Object)

        Public Sub New()
            Content = New List(Of McpContentItem)()
            Artifact = New Dictionary(Of String, Object)()
        End Sub
    End Class

    ''' <summary>
    ''' MCP 内容项基类
    ''' </summary>
    <Newtonsoft.Json.JsonConverter(GetType(McpContentItemConverter))>
    Public MustInherit Class McpContentItem
        <JsonProperty("type")>
        Public MustOverride ReadOnly Property Type As String
    End Class

    ''' <summary>
    ''' 文本内容项
    ''' </summary>
    Public Class McpTextContent
        Inherits McpContentItem

        Public Overrides ReadOnly Property Type As String = "text"

        <JsonProperty("text")>
        Public Property Text As String
    End Class

    ''' <summary>
    ''' 图片内容项
    ''' </summary>
    Public Class McpImageContent
        Inherits McpContentItem

        Public Overrides ReadOnly Property Type As String = "image"

        <JsonProperty("data")>
        Public Property Data As String

        <JsonProperty("mimeType")>
        Public Property MimeType As String
    End Class

    ''' <summary>
    ''' 资源内容项
    ''' </summary>
    Public Class McpResourceContent
        Inherits McpContentItem

        Public Overrides ReadOnly Property Type As String = "resource"

        <JsonProperty("resource")>
        Public Property Resource As McpResourceContentDetail
    End Class

    ''' <summary>
    ''' 资源内容详情
    ''' </summary>
    Public Class McpResourceContentDetail
        <JsonProperty("uri")>
        Public Property Uri As String

        <JsonProperty("mimeType")>
        Public Property MimeType As String

        <JsonProperty("text")>
        Public Property Text As String

        <JsonProperty("blob")>
        Public Property Blob As String
    End Class

    ''' <summary>
    ''' MCP ContentItem 多态转换器
    ''' </summary>
    Public Class McpContentItemConverter
        Inherits JsonConverter

        Public Overrides Function CanConvert(objectType As Type) As Boolean
            Return GetType(McpContentItem).IsAssignableFrom(objectType)
        End Function

        Public Overrides Function ReadJson(reader As JsonReader, objectType As Type, existingValue As Object, serializer As JsonSerializer) As Object
            Dim jsonObject = JObject.Load(reader)
            Dim typeToken = jsonObject("type")
            Dim contentType = If(typeToken IsNot Nothing, typeToken.ToString(), Nothing)

            Select Case contentType
                Case "text"
                    Return jsonObject.ToObject(Of McpTextContent)()
                Case "image"
                    Return jsonObject.ToObject(Of McpImageContent)()
                Case "resource"
                    Return jsonObject.ToObject(Of McpResourceContent)()
                Case Else
                    Return jsonObject.ToObject(Of McpTextContent)()
            End Select
        End Function

        Public Overrides Sub WriteJson(writer As JsonWriter, value As Object, serializer As JsonSerializer)
            serializer.Serialize(writer, value)
        End Sub
    End Class

    '#End Region

    '#Region "Resource 模型"

    ''' <summary>
    ''' MCP Resource 定义
    ''' </summary>
    Public Class McpResource
        <JsonProperty("uri")>
        Public Property Uri As String

        <JsonProperty("name")>
        Public Property Name As String

        <JsonProperty("description")>
        Public Property Description As String

        <JsonProperty("mimeType")>
        Public Property MimeType As String
    End Class

    ''' <summary>
    ''' Resource 模板
    ''' </summary>
    Public Class McpResourceTemplate
        <JsonProperty("uriTemplate")>
        Public Property UriTemplate As String

        <JsonProperty("name")>
        Public Property Name As String

        <JsonProperty("description")>
        Public Property Description As String

        <JsonProperty("mimeType")>
        Public Property MimeType As String
    End Class

    ''' <summary>
    ''' Resources/List 响应结果
    ''' </summary>
    Public Class ResourcesListResult
        <JsonProperty("resources")>
        Public Property Resources As List(Of McpResource)

        <JsonProperty("nextCursor")>
        Public Property NextCursor As String

        Public Sub New()
            Resources = New List(Of McpResource)()
        End Sub
    End Class

    ''' <summary>
    ''' Resources/Read 请求参数
    ''' </summary>
    Public Class ResourcesReadParams
        <JsonProperty("uri")>
        Public Property Uri As String
    End Class

    ''' <summary>
    ''' Resources/Read 响应结果
    ''' </summary>
    Public Class ResourcesReadResult
        <JsonProperty("contents")>
        Public Property Contents As List(Of McpContentItem)

        Public Sub New()
            Contents = New List(Of McpContentItem)()
        End Sub
    End Class

    '#End Region

    '#Region "Prompt 模型"

    ''' <summary>
    ''' MCP Prompt 定义
    ''' </summary>
    Public Class McpPrompt
        <JsonProperty("name")>
        Public Property Name As String

        <JsonProperty("description")>
        Public Property Description As String

        <JsonProperty("arguments")>
        Public Property Arguments As List(Of McpPromptArgument)

        Public Sub New()
            Arguments = New List(Of McpPromptArgument)()
        End Sub
    End Class

    ''' <summary>
    ''' Prompt 参数
    ''' </summary>
    Public Class McpPromptArgument
        <JsonProperty("name")>
        Public Property Name As String

        <JsonProperty("description")>
        Public Property Description As String

        <JsonProperty("required")>
        Public Property Required As Boolean = False
    End Class

    ''' <summary>
    ''' Prompts/List 响应结果
    ''' </summary>
    Public Class PromptsListResult
        <JsonProperty("prompts")>
        Public Property Prompts As List(Of McpPrompt)

        <JsonProperty("nextCursor")>
        Public Property NextCursor As String

        Public Sub New()
            Prompts = New List(Of McpPrompt)()
        End Sub
    End Class

    ''' <summary>
    ''' 用于前端展示的工具简化信息
    ''' </summary>
    Public Class McpToolInfo
        <JsonProperty("name")>
        Public Property Name As String
        <JsonProperty("description")>
        Public Property Description As String
        <JsonProperty("source")>
        Public Property Source As String
        <JsonProperty("parameters")>
        Public Property Parameters As List(Of McpParameterInfo)
        <JsonProperty("enabled")>
        Public Property Enabled As Boolean = True

        ' 【修复·元数据丢失】以下字段 excel_tools.json 里都有（经 CapabilityDescriptor 已在内存中），
        ' 但 DTO 缺失 → 序列化时被丢弃 → 前端 buildMcpToolMeta 读不到 intentKeywords/displayName 等。
        ' 后果：agent-loop 的"要求覆盖度自检"因拿不到关键词而静默空转（实测 带intentKeywords=0）。
        <JsonProperty("displayName")>
        Public Property DisplayName As String
        <JsonProperty("intentKeywords")>
        Public Property IntentKeywords As List(Of String)
        <JsonProperty("operationType")>
        Public Property OperationType As String
        <JsonProperty("isWrite")>
        Public Property IsWrite As Boolean
        <JsonProperty("isDangerous")>
        Public Property IsDangerous As Boolean
        <JsonProperty("dangerSeverity")>
        Public Property DangerSeverity As String
        <JsonProperty("dangerMessage")>
        Public Property DangerMessage As String
        <JsonProperty("layer")>
        Public Property Layer As Integer

        Public Sub New()
            Parameters = New List(Of McpParameterInfo)()
        End Sub
    End Class

    ''' <summary>
    ''' 工具参数信息（前端展示用）
    ''' </summary>
    Public Class McpParameterInfo
        <JsonProperty("name")>
        Public Property Name As String
        <JsonProperty("type")>
        Public Property Type As String
        <JsonProperty("description")>
        Public Property Description As String
        <JsonProperty("required")>
        Public Property Required As Boolean
    End Class

    '#End Region

    '#Region "MCP Client 配置模型"

    ''' <summary>
    ''' MCP Client 连接配置
    ''' </summary>
    Public Class McpClientConfig
        <JsonProperty("name")>
        Public Property Name As String
        <JsonProperty("description")>
        Public Property Description As String
        <JsonProperty("transport")>
        Public Property Transport As String = "stdio"

        ' stdio 传输配置
        <JsonProperty("command")>
        Public Property Command As String
        <JsonProperty("args")>
        Public Property Args As List(Of String)

        ' HTTP 传输配置
        <JsonProperty("url")>
        Public Property Url As String
        <JsonProperty("headers")>
        Public Property Headers As Dictionary(Of String, String)

        <JsonProperty("enabled")>
        Public Property Enabled As Boolean = True

        ' 环境变量（用于传递认证凭证等，如 ms365 的 AZURE_CLIENT_ID）
        <JsonProperty("env")>
        Public Property Env As Dictionary(Of String, String)

        Public Sub New()
            Args = New List(Of String)()
            Headers = New Dictionary(Of String, String)()
            Env = New Dictionary(Of String, String)()
        End Sub
    End Class

    ''' <summary>
    ''' MCP 全局配置
    ''' </summary>
    Public Class McpGlobalConfig
        <JsonProperty("mcpServers")>
        Public Property McpServers As Dictionary(Of String, McpClientConfig)

        <JsonProperty("serverPort")>
        Public Property ServerPort As Integer = Utils.ConfigConstants.MCP.DEFAULT_SERVER_PORT

        <JsonProperty("serverEnabled")>
        Public Property ServerEnabled As Boolean = False

        <JsonProperty("autoStartClients")>
        Public Property AutoStartClients As Boolean = False

        Public Sub New()
            McpServers = New Dictionary(Of String, McpClientConfig)()
        End Sub
    End Class

    '#End Region

    '#Region "MCP 状态模型（前端展示用）"

    ''' <summary>
    ''' MCP 整体状态
    ''' </summary>
    Public Class McpStatusInfo
        <JsonProperty("serverRunning")>
        Public Property ServerRunning As Boolean
        <JsonProperty("serverPort")>
        Public Property ServerPort As Integer
        <JsonProperty("serverToolCount")>
        Public Property ServerToolCount As Integer
        <JsonProperty("clients")>
        Public Property Clients As List(Of McpClientStatusInfo)
        <JsonProperty("totalAvailableTools")>
        Public Property TotalAvailableTools As Integer

        Public Sub New()
            Clients = New List(Of McpClientStatusInfo)()
        End Sub
    End Class

    ''' <summary>
    ''' 单个 Client 连接状态
    ''' </summary>
    Public Class McpClientStatusInfo
        <JsonProperty("name")>
        Public Property Name As String
        <JsonProperty("description")>
        Public Property Description As String
        <JsonProperty("connected")>
        Public Property Connected As Boolean
        <JsonProperty("transport")>
        Public Property Transport As String
        <JsonProperty("toolCount")>
        Public Property ToolCount As Integer
        <JsonProperty("errorMessage")>
        Public Property ErrorMessage As String

        ' 编辑用：完整配置
        <JsonProperty("command")>
        Public Property Command As String
        <JsonProperty("args")>
        Public Property Args As List(Of String)
        <JsonProperty("url")>
        Public Property Url As String
        <JsonProperty("headers")>
        Public Property Headers As Dictionary(Of String, String)
        <JsonProperty("env")>
        Public Property Env As Dictionary(Of String, String)
        <JsonProperty("enabled")>
        Public Property Enabled As Boolean
    End Class

    '#End Region

End Namespace
