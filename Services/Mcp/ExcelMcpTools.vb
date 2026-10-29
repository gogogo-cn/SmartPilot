Option Strict On
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Reflection
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ExcelAIVSTOAddIn.Utils

Namespace Services.Mcp

    '#Region "Excel MCP 工具注册表"

    ''' <summary>
    ''' Excel MCP 工具注册表 - 定义所有可暴露为 MCP Tool 的 Excel 操作
    ''' </summary>
    Partial Public Class ExcelMcpTools

        Private _bridge As Object

        ' ── Execute 方法反射缓存（启动时一次性扫描所有 Private Instance Execute* 方法）──
        Private Shared ReadOnly _executeMethodCache As New Dictionary(Of String, MethodInfo)(StringComparer.OrdinalIgnoreCase)
        Private Shared _cacheInitialized As Boolean = False

        ''' <summary>
        ''' 工具名 → Execute 方法名转换：excel_read_range → ExecuteReadRange
        ''' </summary>
        Private Shared Function ToolNameToMethodName(toolName As String) As String
            Dim shortName = toolName
            If shortName.StartsWith("excel_") Then shortName = shortName.Substring(6)
            Dim parts = shortName.Split(New Char() {"_"c}, StringSplitOptions.RemoveEmptyEntries)
            Dim sb As New System.Text.StringBuilder("Execute")
            For Each p In parts
                If p.Length > 0 Then
                    sb.Append(Char.ToUpperInvariant(p(0)))
                    If p.Length > 1 Then sb.Append(p.Substring(1))
                End If
            Next
            Return sb.ToString()
        End Function

        ''' <summary>
        ''' 首次调用时扫描类型，缓存所有 Private Instance Execute* 方法的 MethodInfo。
        ''' 加新工具只需写一个 Private Function ExecuteXxxYyy，ExecuteTool 自动发现。
        ''' </summary>
        Private Shared Sub EnsureMethodCache()
            If _cacheInitialized Then Return
            SyncLock _executeMethodCache
                If _cacheInitialized Then Return
                Dim t = GetType(ExcelMcpTools)
                Dim flags = BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.DeclaredOnly
                Dim dictType = GetType(Dictionary(Of String, Object))
                For Each mi In t.GetMethods(flags)
                    ' 只缓存真正的工具实现：1 个 Dictionary 参数 + 返回 ToolsCallResult
                    If mi.ReturnType Is GetType(ToolsCallResult) AndAlso
                       mi.GetParameters().Length = 1 AndAlso
                       mi.GetParameters()(0).ParameterType Is dictType Then
                        _executeMethodCache(mi.Name) = mi
                    End If
                Next
                _cacheInitialized = True
                Debug.WriteLine($"[ExcelMcpTools] ✅ Execute 方法缓存完成，共 {_executeMethodCache.Count} 个")
            End SyncLock
        End Sub

        ' ── 兼容旧 InvokePrivateExecute（已委托到统一缓存路由）──
        Private Function InvokePrivateExecute(methodName As String, arguments As Dictionary(Of String, Object)) As ToolsCallResult
            Return ExecuteByName(methodName, arguments)
        End Function

        ''' <summary>
        ''' 通过 MethodInfo 缓存直接调用指定 Execute 方法（所有 Private 工具实现的统一入口）
        ''' </summary>
        Private Function ExecuteByName(methodName As String, arguments As Dictionary(Of String, Object)) As ToolsCallResult
            EnsureMethodCache()
            Dim mi As MethodInfo = Nothing
            If _executeMethodCache.TryGetValue(methodName, mi) Then
                Try
                    Return DirectCast(mi.Invoke(Me, New Object() {arguments}), ToolsCallResult)
                Catch ex As TargetInvocationException
                    Dim errMsg As String = If(ex.InnerException IsNot Nothing, ex.InnerException.Message, ex.Message)
                    Return New ToolsCallResult() With {
                        .IsError = True,
                        .Content = New List(Of McpContentItem) From {New McpTextContent With {.Text = errMsg}}
                    }
                End Try
            End If
            Return New ToolsCallResult() With {
                .IsError = True,
                .Content = New List(Of McpContentItem) From {New McpTextContent With {.Text = "未找到执行方法: " & methodName}}
            }
        End Function

        ''' <summary>
        ''' 选区上下文：记录最近一次read_range的选区，写工具无参数调用时自动复用。
        ''' 生命周期（基于对话流程，非时间）：
        '''   1. excel_read_range 成功 → Update() 记录选区（新read覆盖旧read）
        '''   2. 写工具（clear_range/write_range等）无参数调用时 → TryGet() 使用，成功后 Clear() 消费
        '''   3. TryGet时检测到活动工作表已切换 → 自动失效返回False
        '''   4. 前端新会话开始时 → 显式调用 Reset() 清除
        ''' 没有TTL——上下文的生死完全由对话流程控制，不由时间控制。
        ''' </summary>
        Friend Class RangeContext
            Private Shared _worksheetName As String = ""
            Private Shared _rangeAddress As String = ""

            ''' <summary>
            ''' 更新上下文（由excel_read_range在成功读取后调用）
            ''' </summary>
            Public Shared Sub Update(sheetName As String, addr As String)
                _worksheetName = sheetName
                _rangeAddress = addr
                Debug.WriteLine($"[RangeContext] 更新: sheet={sheetName}, range={addr}")
            End Sub

            ''' <summary>
            ''' 消费并清除上下文（写工具成功执行后调用）
            ''' </summary>
            Public Shared Sub Clear()
                If Not String.IsNullOrEmpty(_rangeAddress) Then
                    Debug.WriteLine($"[RangeContext] 消费清除: sheet={_worksheetName}, range={_rangeAddress}")
                End If
                _worksheetName = ""
                _rangeAddress = ""
            End Sub

            ''' <summary>
            ''' 重置上下文（新会话开始时由前端显式调用）
            ''' </summary>
            Public Shared Sub Reset()
                _worksheetName = ""
                _rangeAddress = ""
                Debug.WriteLine($"[RangeContext] 重置（新会话）")
            End Sub

            ''' <summary>
            ''' 尝试获取有效上下文选区。
            ''' 返回True表示上下文有效，outAddr为A1格式地址（如A1:H6）。
            ''' 活动工作表已切换则自动失效。
            ''' </summary>
            Public Shared Function TryGet(currentSheetName As String, ByRef outAddr As String) As Boolean
                If String.IsNullOrEmpty(_worksheetName) OrElse String.IsNullOrEmpty(_rangeAddress) Then
                    outAddr = ""
                    Return False
                End If
                ' 工作表不匹配 → 上下文失效（用户切换了sheet）
                If Not String.Equals(_worksheetName, currentSheetName, StringComparison.OrdinalIgnoreCase) Then
                    Debug.WriteLine($"[RangeContext] 工作表已切换（{_worksheetName}→{currentSheetName}），上下文失效")
                    outAddr = ""
                    Return False
                End If
                outAddr = _rangeAddress
                Debug.WriteLine($"[RangeContext] 命中: range={outAddr}")
                Return True
            End Function
        End Class

        Public Sub New(bridge As Object)
            _bridge = bridge
        End Sub

        '#Region "工具定义"

        ''' <summary>
        ''' 获取所有 Excel MCP Tools 定义
        ''' </summary>
        ''' <Layer name="1">核心基础：读写/结构查询，每次任务几乎必用</Layer>
        ''' <Layer name="2">数据操作：格式/公式/编辑/排序/筛选，常用操作</Layer>
        ''' <Layer name="3">高级功能：VBA/图表/透视表/文件合并拆分/OCR/批处理，按需使用</Layer>
        Public Shared Function GetToolDefinitions() As List(Of McpTool)
            ' 单一数据源：优先从 excel_tools.json 加载（含 displayName/intentKeywords/operationType 等元数据）
            Dim jsonTools = TryLoadToolsFromJson()
            If jsonTools IsNot Nothing AndAlso jsonTools.Count > 0 Then
                Return jsonTools
            End If
            ' 兜底：配置文件缺失或加载失败时使用硬编码定义
            Return BuildHardcodedTools()
        End Function

        ''' <summary>
        ''' 硬编码工具定义（兜底：仅在 excel_tools.json 缺失或加载失败时触发）
        ''' 2026-09-04: 返回空列表，JSON 缺失视为配置错误（不再维护 69 个 tools.Add 硬编码）
        ''' </summary>
        Private Shared Function BuildHardcodedTools() As List(Of McpTool)
            Debug.WriteLine("[ExcelMcpTools] ⚠️ excel_tools.json 缺失或加载失败，硬编码兜底已废弃，返回空列表")
            Return New List(Of McpTool)()
        End Function

        ''' <summary>
        '''  将所有 MCP 工具注册到统一能力注册中心
        ''' 从此处开始，前端意图匹配不再需要单独的 KEYWORD_SCORE_BOOST 硬编码词表
        ''' </summary>
        Public Shared Sub RegisterAllToolsToRegistry()
            Dim tools = GetToolDefinitions()
            For Each tool In tools
                Dim meta = GetToolMetadata(tool.Name)
                McpCapabilityRegistry.Instance.RegisterMcpTool(
                    tool,
                    displayName:=meta.DisplayName,
                    category:=meta.Category,
                    keywords:=meta.Keywords,
                    isWriteOperation:=meta.IsWrite,
                    isDangerous:=meta.IsDangerous,
                    dangerSeverity:=meta.DangerSeverity,
                    dangerMessage:=meta.DangerMessage)
            Next
        End Sub

        ''' <summary>
        ''' 工具元数据映射表（每种工具一个条目）
        ''' </summary>
        Private Class ToolMeta
            Public Property DisplayName As String
            Public Property Category As CapabilityCategory
            Public Property Keywords As List(Of String)
            Public Property IsWrite As Boolean
            Public Property IsDangerous As Boolean
            Public Property DangerSeverity As String
            Public Property DangerMessage As String
        End Class

        Private Shared Function GetToolMetadata(toolName As String) As ToolMeta
            ' 单一数据源：优先从 excel_tools.json 读取元数据
            Dim jsonMeta = TryGetMetaFromJson(toolName)
            If jsonMeta IsNot Nothing Then
                Return jsonMeta
            End If
            ' 兜底：配置文件缺失或该工具无元数据时使用硬编码推导
            Return GetHardcodedToolMetadata(toolName)
        End Function

        Private Shared Function GetHardcodedToolMetadata(toolName As String) As ToolMeta
            ' 默认：根据工具名自动推导类型
            Dim defaultCat = CapabilityCategory.Utility
            Dim defaultWrite = False
            Dim defaultDanger = False

            If toolName.Contains("_read_") OrElse toolName.Contains("_get_") OrElse
               toolName.Contains("_list_") OrElse toolName.Contains("_search") OrElse
               toolName.Contains("_browse_") Then
                defaultCat = CapabilityCategory.Read
            ElseIf toolName.Contains("_aggregate") OrElse toolName.Contains("_date_group") OrElse
                    toolName.Contains("_remove_duplicates") OrElse toolName.Contains("_clean_data") OrElse
                    toolName.Contains("_join_tables") Then
                ' 纯计算/分析类工具：仅返回计算结果，不落盘（落地走 targetSheet/写表工具）
                defaultCat = CapabilityCategory.Compute
            ElseIf toolName.Contains("_write_") OrElse toolName.Contains("_set_") OrElse
                    toolName.Contains("_fill_") OrElse toolName.Contains("_add_") OrElse
                    toolName.Contains("_insert_") OrElse toolName.Contains("_delete_") OrElse
                    toolName.Contains("_remove_") OrElse toolName.Contains("_clear_") OrElse
                    toolName.Contains("_clean_") OrElse toolName.Contains("_merge_cells") OrElse
                    toolName.Contains("_unmerge_cells") Then
                defaultCat = CapabilityCategory.Write
                defaultWrite = True
            ElseIf toolName.Contains("_format_") OrElse toolName.Contains("_conditional_format") OrElse
                    toolName.Contains("_autofit") Then
                defaultCat = CapabilityCategory.Format
                defaultWrite = True
            ElseIf toolName.Contains("_pivot_") OrElse toolName.Contains("_sort_") OrElse
                    toolName.Contains("_filter") Then
                defaultCat = CapabilityCategory.Analyze
            ElseIf toolName.Contains("_split_") OrElse toolName.Contains("_merge_workbooks") OrElse
                    toolName.Contains("_merge_sheets") Then
                defaultCat = CapabilityCategory.Transform
                defaultWrite = True
            ElseIf toolName.Contains("_import_") Then
                defaultCat = CapabilityCategory.Import
                defaultWrite = True
            ElseIf toolName.Contains("_export_") Then
                defaultCat = CapabilityCategory.Export
                defaultWrite = True
            ElseIf toolName.Contains("_sheet") OrElse toolName.Contains("_workbook") Then
                defaultCat = CapabilityCategory.Manage
            ElseIf toolName.Contains("_chart") OrElse toolName.Contains("_display_html") Then
                defaultCat = CapabilityCategory.Visualize
                defaultWrite = True
            ElseIf toolName.Contains("_clipboard_") OrElse toolName.Contains("_screenshot") OrElse
                    toolName.Contains("_ocr_") Then
                defaultCat = CapabilityCategory.Utility
            End If

            ' 与 Config/excel_tools.json 危险集合对齐（JSON 为权威；此处仅 JSON 缺失时兜底）
            If toolName.Contains("_clear_sheet") OrElse toolName.Contains("_delete_sheet") OrElse
               toolName.Contains("_delete_row") OrElse toolName.Contains("_delete_column") OrElse
               toolName.Contains("_clear_range") OrElse toolName.Contains("_split_by_column") OrElse
               toolName.Contains("_merge_workbooks") OrElse toolName.Contains("_execute_vba") OrElse
               toolName.Contains("_run_python_script") OrElse toolName.Contains("_clean_data") OrElse
               toolName.Contains("_remove_duplicates") Then
                defaultDanger = True
            End If

            Dim meta As New ToolMeta() With {
                .DisplayName = toolName,
                .Category = defaultCat,
                .Keywords = New List(Of String)(),
                .IsWrite = defaultWrite,
                .IsDangerous = defaultDanger
            }

            ' ApplyExplicitToolMetadata(meta, toolName)  ' ← 2026-09-04 废弃，JSON 已是主数据源
            ' 原 69 个 Select Case 显式覆盖已废弃，Category/IsWrite/IsDangerous 由上面的自动推导逻辑覆盖

            Return meta
        End Function

        '#End Region

        '#Region "资源定义"

        ''' <summary>
        ''' 获取 Excel MCP 资源定义
        ''' </summary>
        Public Function GetResourceDefinitions() As List(Of McpResource)
            Dim resources = New List(Of McpResource)()

            Try
                Dim excelApp = Globals.ThisAddIn.Application
                If excelApp IsNot Nothing Then
                    Dim wb = excelApp.ActiveWorkbook
                    If wb IsNot Nothing Then
                        Dim sheetCount = wb.Worksheets.Count
                        For i As Integer = 1 To sheetCount
                            Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
                            Try
                                ws = DirectCast(wb.Worksheets(i), Microsoft.Office.Interop.Excel.Worksheet)
                                If ws.Visible <> Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible Then Continue For
                                Dim sheetName = ws.Name
                                resources.Add(New McpResource With {
                                    .Uri = $"workbook://sheets/{Uri.EscapeDataString(sheetName)}/range/used",
                                    .Name = $"工作表：{sheetName}",
                                    .Description = $"工作表 {sheetName} 的已使用范围数据",
                                    .MimeType = "application/json"
                                })
                            Finally
                                Utils.ComReleaseHelper.Release(ws)
                            End Try
                        Next
                        Utils.ComReleaseHelper.Release(wb)
                    End If
                End If
            Catch ex As Exception
                Debug.WriteLine($"获取资源定义异常: {ex.Message}")
            End Try

            ' 始终包含元数据资源
            resources.Insert(0, New McpResource With {
                .Uri = "workbook://metadata",
                .Name = "工作簿元数据",
                .Description = "当前工作簿的基本信息",
                .MimeType = "application/json"
            })

            Return resources
        End Function

        ''' <summary>
        ''' 获取资源模板定义
        ''' </summary>
        Public Function GetResourceTemplateDefinitions() As List(Of McpResourceTemplate)
            Return New List(Of McpResourceTemplate) From {
                New McpResourceTemplate With {
                    .UriTemplate = "workbook://sheets/{name}/range/{address}",
                    .Name = "工作表范围数据",
                    .Description = "读取指定工作表的指定范围数据",
                    .MimeType = "application/json"
                },
                New McpResourceTemplate With {
                    .UriTemplate = "workbook://sheets/{name}/structure",
                    .Name = "工作表结构",
                    .Description = "工作表的列结构信息",
                    .MimeType = "application/json"
                }
            }
        End Function

        '#End Region

        '#Region "工具执行"

        ''' <summary>
        ''' 加载意图识别配置（Config/intent_config.json）并返回原始 JSON 字符串。
        ''' 文件缺失或加载失败时返回空字符串。
        ''' </summary>
        Public Shared Function GetIntentConfigJson() As String
            Try
                Dim configDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config")
                Dim configPath = Path.Combine(configDir, "intent_config.json")
                If File.Exists(configPath) Then
                    Return File.ReadAllText(configPath)
                End If
                Debug.WriteLine("[ExcelMcpTools] GetIntentConfigJson: intent_config.json 不存在于 " & configPath)
            Catch ex As Exception
                Debug.WriteLine("[ExcelMcpTools] GetIntentConfigJson 异常: " & ex.Message)
            End Try
            Return ""
        End Function

        ''' <summary>
        ''' 执行 Excel MCP Tool 调用
        ''' </summary>
        ''' <param name="toolName">工具名称</param>
        ''' <param name="arguments">工具参数</param>
        ''' <returns>工具执行结果</returns>
        Public Function ExecuteTool(toolName As String, arguments As Dictionary(Of String, Object)) As ToolsCallResult
            ' ── 统一反射路由：工具名 → 方法名转换 → 缓存查找 MethodInfo → 调用 ──
            ' 加新工具只需写一个 Private Function ExecuteXxxYyy，ExecuteTool 自动发现（开闭原则）
            Dim methodName = ToolNameToMethodName(toolName)
            Debug.WriteLine($"[ExecuteTool] {toolName} → {methodName}")
            Return ExecuteByName(methodName, arguments)
        End Function

        '#End Region
    '#End Region
    End Class
End Namespace
