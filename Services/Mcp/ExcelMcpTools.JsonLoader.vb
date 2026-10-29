Option Strict On
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports Newtonsoft.Json
Imports ExcelAIVSTOAddIn.Utils

Namespace Services.Mcp

    Partial Public Class ExcelMcpTools

        '#Region "excel_tools.json 元数据加载（单一数据源）"

        ''' <summary>
        ''' excel_tools.json 配置模型（name/layer/description/parameters + 元数据）
        ''' </summary>
        Private Class ToolsConfigJson
            Public Property Version As String
            Public Property LastUpdated As String
            Public Property Tools As List(Of ToolJson)
        End Class

        Private Class ToolJson
            Public Property Name As String
            Public Property Layer As Integer
            Public Property Description As String
            Public Property Parameters As List(Of ToolParamJson)
            Public Property DisplayName As String
            Public Property IntentKeywords As List(Of String)
            Public Property OperationType As String
            Public Property IsWrite As Boolean
            Public Property IsDangerous As Boolean
            ''' <summary>危险等级 high|medium（可选；仅 isDangerous=true 时有意义）</summary>
            Public Property DangerSeverity As String
            ''' <summary>安全守卫确认文案（可选；前端弹窗用）</summary>
            Public Property DangerMessage As String
        End Class

        Private Class ToolParamJson
            Public Property Name As String
            Public Property Type As String
            Public Property Description As String
            Public Property Required As Boolean
        End Class

        Private Shared _jsonConfig As ToolsConfigJson
        Private Shared _jsonToolsCache As List(Of McpTool)
        Private Shared _jsonMetaCache As Dictionary(Of String, ToolMeta)
        Private Shared _jsonLoadAttempted As Boolean
        Private Shared _jsonLoadLock As New Object()

        ''' <summary>
        ''' 从 excel_tools.json 构建全部工具定义（缓存；加载失败返回 Nothing）
        ''' </summary>
        Private Shared Function TryLoadToolsFromJson() As List(Of McpTool)
            If _jsonToolsCache IsNot Nothing Then Return _jsonToolsCache
            Dim cfg = LoadToolsConfig()
            If cfg Is Nothing Then Return Nothing
            Dim result = New List(Of McpTool)()
            For Each tj In cfg.Tools
                result.Add(BuildMcpTool(tj))
            Next
            _jsonToolsCache = result
            Return result
        End Function

        ''' <summary>
        ''' 从 excel_tools.json 读取单个工具元数据（缓存；无该工具返回 Nothing）
        ''' </summary>
        Private Shared Function TryGetMetaFromJson(toolName As String) As ToolMeta
            If _jsonMetaCache IsNot Nothing AndAlso _jsonMetaCache.ContainsKey(toolName) Then
                Return _jsonMetaCache(toolName)
            End If
            Dim cfg = LoadToolsConfig()
            If cfg Is Nothing Then Return Nothing
            If _jsonMetaCache Is Nothing Then
                _jsonMetaCache = New Dictionary(Of String, ToolMeta)(StringComparer.OrdinalIgnoreCase)
            End If
            For Each tj In cfg.Tools
                If Not _jsonMetaCache.ContainsKey(tj.Name) Then
                    _jsonMetaCache(tj.Name) = BuildToolMeta(tj)
                End If
            Next
            If _jsonMetaCache.ContainsKey(toolName) Then Return _jsonMetaCache(toolName)
            Return Nothing
        End Function

        Private Shared Function LoadToolsConfig() As ToolsConfigJson
            If _jsonLoadAttempted Then Return _jsonConfig
            SyncLock _jsonLoadLock
                If _jsonLoadAttempted Then Return _jsonConfig
                _jsonLoadAttempted = True
                Try
                    Dim path As String = System.IO.Path.Combine(PathHelper.ConfigDirectory, "excel_tools.json")
                    If Not File.Exists(path) Then
                        Debug.WriteLine($"[ExcelMcpTools] excel_tools.json 不存在: {path}")
                        _jsonConfig = Nothing
                        Return Nothing
                    End If
                    Dim content = File.ReadAllText(path)
                    Dim cfg = JsonConvert.DeserializeObject(Of ToolsConfigJson)(content)
                    If cfg Is Nothing OrElse cfg.Tools Is Nothing OrElse cfg.Tools.Count = 0 Then
                        Debug.WriteLine("[ExcelMcpTools] excel_tools.json 无有效工具")
                        _jsonConfig = Nothing
                        Return Nothing
                    End If
                    _jsonConfig = cfg
                    Debug.WriteLine($"[ExcelMcpTools] ✅ 已从 excel_tools.json 加载 {cfg.Tools.Count} 个工具")
                Catch ex As Exception
                    Debug.WriteLine($"[ExcelMcpTools] excel_tools.json 加载失败: {ex.Message}")
                    _jsonConfig = Nothing
                End Try
                Return _jsonConfig
            End SyncLock
        End Function

        Private Shared Function BuildMcpTool(tj As ToolJson) As McpTool
            Dim schema = New McpToolInputSchema()
            If tj.Parameters IsNot Nothing Then
                For Each p In tj.Parameters
                    If String.IsNullOrEmpty(p.Name) Then Continue For
                    schema.Properties(p.Name) = New McpToolProperty(p.Type, p.Description)
                    If p.Required Then schema.Required.Add(p.Name)
                Next
            End If
            Return New McpTool(tj.Name, tj.Description, schema)
        End Function

        Private Shared Function BuildToolMeta(tj As ToolJson) As ToolMeta
            Dim cat As CapabilityCategory = CapabilityCategory.Utility
            If Not String.IsNullOrEmpty(tj.OperationType) Then
                If Not [Enum].TryParse(Of CapabilityCategory)(tj.OperationType, True, cat) Then
                    cat = CapabilityCategory.Utility
                End If
            End If
            Return New ToolMeta() With {
                .DisplayName = If(String.IsNullOrEmpty(tj.DisplayName), tj.Name, tj.DisplayName),
                .Category = cat,
                .Keywords = If(tj.IntentKeywords, New List(Of String)()),
                .IsWrite = tj.IsWrite,
                .IsDangerous = tj.IsDangerous,
                .DangerSeverity = tj.DangerSeverity,
                .DangerMessage = tj.DangerMessage
            }
        End Function

        '#End Region

    End Class

End Namespace
