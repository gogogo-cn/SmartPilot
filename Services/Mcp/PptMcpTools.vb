Option Strict On
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ExcelAIVSTOAddIn.Services
Imports ExcelAIVSTOAddIn.Utils

Namespace Services.Mcp

    ''' <summary>
    ''' PPT MCP 工具注册表：AI 自由搓 PPT 的元基石工具集（与 Excel 工具完全分离）。
    ''' 工具定义单一数据源：Config/ppt_tools.json
    ''' 加新工具只需写一个 Private Function ExecuteXxxYyy，ExecuteTool 自动发现（开闭原则）。
    ''' 执行内核委托 PptGenerationService（PowerPoint COM 直出）与 PptDesignEngine（数据感知设计）。
    ''' </summary>
    Partial Public Class PptMcpTools

        Private Shared _executeMethodCache As New Dictionary(Of String, MethodInfo)(StringComparer.OrdinalIgnoreCase)
        Private Shared _cacheInitialized As Boolean = False
        Private _bridge As Object

        Public Sub New(bridge As Object)
            _bridge = bridge
        End Sub

        '#Region "工具定义"

        ''' <summary>
        ''' 获取所有 PPT MCP Tools 定义（单一数据源：Config/ppt_tools.json）
        ''' </summary>
        Public Shared Function GetToolDefinitions() As List(Of McpTool)
            Dim jsonTools = TryLoadToolsFromJson()
            If jsonTools IsNot Nothing AndAlso jsonTools.Count > 0 Then
                Return jsonTools
            End If
            ' 配置文件缺失视为配置错误（不维护硬编码工具表）
            Debug.WriteLine("[PptMcpTools] ⚠️ ppt_tools.json 缺失或加载失败，返回空列表")
            Return New List(Of McpTool)()
        End Function

        ''' <summary>
        ''' 将所有 PPT MCP 工具注册到统一能力注册中心
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
                    source:="ppt-local")
            Next
        End Sub

        ''' <summary>
        ''' 工具元数据（单一数据源：ppt_tools.json；缺失时按工具名自动推导兜底）
        ''' </summary>
        Private Class ToolMeta
            Public Property DisplayName As String
            Public Property Category As CapabilityCategory
            Public Property Keywords As List(Of String)
            Public Property IsWrite As Boolean
            Public Property IsDangerous As Boolean
        End Class

        Private Shared Function GetToolMetadata(toolName As String) As ToolMeta
            Dim jsonMeta = TryGetMetaFromJson(toolName)
            If jsonMeta IsNot Nothing Then
                Return jsonMeta
            End If
            Return GetHardcodedToolMetadata(toolName)
        End Function

        Private Shared Function GetHardcodedToolMetadata(toolName As String) As ToolMeta
            Dim defaultCat = CapabilityCategory.Utility
            Dim defaultWrite = False
            Dim defaultDangerous = False
            If toolName.Contains("_get_") OrElse toolName.Contains("_browse_") OrElse toolName.Contains("_list") Then
                defaultCat = CapabilityCategory.Read
            ElseIf toolName.Contains("_analyze_") Then
                defaultCat = CapabilityCategory.Compute
            ElseIf toolName.Contains("_generate") OrElse toolName.Contains("_import_") OrElse
                   toolName.Contains("_project_") OrElse toolName.Contains("_add_") OrElse
                   toolName.Contains("_update") OrElse toolName.Contains("_remove") OrElse
                   toolName.Contains("_build") Then
                defaultCat = CapabilityCategory.Write
                defaultWrite = True
                ' 与 ppt_tools.json 一致：删除类操作需危险确认
                If toolName.Contains("_remove") Then defaultDangerous = True
            End If
            Return New ToolMeta() With {
                .DisplayName = toolName,
                .Category = defaultCat,
                .Keywords = New List(Of String)(),
                .IsWrite = defaultWrite,
                .IsDangerous = defaultDangerous
            }
        End Function

        '#End Region

        '#Region "ppt_tools.json 元数据加载（单一数据源）"

        Private Class PptToolsConfigJson
            Public Property Version As String
            Public Property LastUpdated As String
            Public Property Tools As List(Of PptToolJson)
        End Class

        Private Class PptToolJson
            Public Property Name As String
            Public Property Layer As Integer
            Public Property Description As String
            Public Property Parameters As List(Of PptToolParamJson)
            Public Property DisplayName As String
            Public Property IntentKeywords As List(Of String)
            Public Property OperationType As String
            Public Property IsWrite As Boolean
            Public Property IsDangerous As Boolean
        End Class

        Private Class PptToolParamJson
            Public Property Name As String
            Public Property Type As String
            Public Property Description As String
            Public Property Required As Boolean
        End Class

        Private Shared _jsonConfig As PptToolsConfigJson
        Private Shared _jsonToolsCache As List(Of McpTool)
        Private Shared _jsonMetaCache As Dictionary(Of String, ToolMeta)
        Private Shared _jsonLoadAttempted As Boolean
        Private Shared _jsonLoadLock As New Object()

        ''' <summary>
        ''' 从 ppt_tools.json 构建全部工具定义（缓存；加载失败返回 Nothing）
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
        ''' 从 ppt_tools.json 读取单个工具元数据（缓存；无该工具返回 Nothing）
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

        Private Shared Function LoadToolsConfig() As PptToolsConfigJson
            If _jsonLoadAttempted Then Return _jsonConfig
            SyncLock _jsonLoadLock
                If _jsonLoadAttempted Then Return _jsonConfig
                _jsonLoadAttempted = True
                Try
                    Dim path As String = System.IO.Path.Combine(PathHelper.ConfigDirectory, "ppt_tools.json")
                    If Not File.Exists(path) Then
                        Debug.WriteLine($"[PptMcpTools] ppt_tools.json 不存在: {path}")
                        _jsonConfig = Nothing
                        Return Nothing
                    End If
                    Dim content = File.ReadAllText(path)
                    Dim cfg = JsonConvert.DeserializeObject(Of PptToolsConfigJson)(content)
                    If cfg Is Nothing OrElse cfg.Tools Is Nothing OrElse cfg.Tools.Count = 0 Then
                        Debug.WriteLine("[PptMcpTools] ppt_tools.json 无有效工具")
                        _jsonConfig = Nothing
                        Return Nothing
                    End If
                    _jsonConfig = cfg
                    Debug.WriteLine($"[PptMcpTools] ✅ 已从 ppt_tools.json 加载 {cfg.Tools.Count} 个工具")
                Catch ex As Exception
                    Debug.WriteLine($"[PptMcpTools] ppt_tools.json 加载失败: {ex.Message}")
                    _jsonConfig = Nothing
                End Try
                Return _jsonConfig
            End SyncLock
        End Function

        Private Shared Function BuildMcpTool(tj As PptToolJson) As McpTool
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

        Private Shared Function BuildToolMeta(tj As PptToolJson) As ToolMeta
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
                .IsDangerous = tj.IsDangerous
            }
        End Function

        '#End Region

        '#Region "工具执行"

        ''' <summary>
        ''' 执行 PPT MCP Tool 调用（统一反射路由：工具名 → Execute 方法名）
        ''' </summary>
        Public Function ExecuteTool(toolName As String, arguments As Dictionary(Of String, Object)) As ToolsCallResult
            Dim methodName = ToolNameToMethodName(toolName)
            Debug.WriteLine($"[PptMcpTools] {toolName} → {methodName}")
            Return ExecuteByName(methodName, arguments)
        End Function

        Private Shared Function ToolNameToMethodName(toolName As String) As String
            Dim shortName = toolName
            If shortName.StartsWith("ppt_") Then shortName = shortName.Substring(4)
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
        ''' 首次调用时扫描类型，缓存所有 Private Instance Execute* 方法的 MethodInfo
        ''' </summary>
        Private Shared Sub EnsureMethodCache()
            If _cacheInitialized Then Return
            SyncLock _executeMethodCache
                If _cacheInitialized Then Return
                Dim t = GetType(PptMcpTools)
                Dim flags = BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.DeclaredOnly
                Dim dictType = GetType(Dictionary(Of String, Object))
                For Each mi In t.GetMethods(flags)
                    If mi.ReturnType Is GetType(ToolsCallResult) AndAlso
                       mi.GetParameters().Length = 1 AndAlso
                       mi.GetParameters()(0).ParameterType Is dictType Then
                        _executeMethodCache(mi.Name) = mi
                    End If
                Next
                _cacheInitialized = True
                Debug.WriteLine($"[PptMcpTools] ✅ Execute 方法缓存完成，共 {_executeMethodCache.Count} 个")
            End SyncLock
        End Sub

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

        '#End Region

        '#Region "Execute 实现"

        ''' <summary>
        ''' ppt_get_templates：获取可用主题（配色+字体）与版式白名单
        ''' </summary>
        Private Function ExecuteGetTemplates(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim json = PptGenerationService.GetMergedTemplatesJson()
                If String.IsNullOrWhiteSpace(json) OrElse json = "{}" Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {
                        .Text = "未找到 PPT 主题配置：" & System.IO.Path.Combine(PathHelper.ConfigDirectory, "ppt_templates.json")
                    })
                    Return result
                End If
                result.Content.Add(New McpTextContent With {.Text = json})
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "获取 PPT 主题模板失败: " & ex.Message})
            End Try
            Return result
        End Function

        ''' <summary>
        ''' ppt_analyze_data：数据感知设计分析，返回推荐页面结构/洞察/画像/方向
        ''' </summary>
        Private Function ExecuteAnalyzeData(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim headers = ResolveStringList(GetArg(Of String)(args, "headers", ""))
                If headers.Count = 0 Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {
                        .Text = "缺少表头参数 headers（列名数组的 JSON 字符串）。请先用 excel_read_range 读取数据，或直接传入 headers 与 rows。"
                    })
                    Return result
                End If
                Dim rows = ResolveRows(GetArg(Of String)(args, "rows", ""))
                Dim reportTitle = GetArg(Of String)(args, "reportTitle", "数据分析报告")
                Dim requirement = GetArg(Of String)(args, "requirement", "")

                Dim columnHints As JObject = Nothing
                Dim columnHintsRaw = GetArg(Of String)(args, "columnHints", "")
                If Not String.IsNullOrWhiteSpace(columnHintsRaw) Then
                    Try
                        columnHints = JObject.Parse(columnHintsRaw)
                    Catch ex As Exception
                        Debug.WriteLine($"[PptMcpTools] columnHints 解析失败: {ex.Message}")
                        columnHints = Nothing
                    End Try
                End If

                Dim profile As PptDataProfile = PptDesignEngine.AnalyzeData(headers, rows)
                PptDesignEngine.ApplyColumnHints(profile, columnHints)
                Dim slides As JArray = PptDesignEngine.BuildDesign(profile, reportTitle, "Excel 选区", requirement)
                Dim insights As New JArray()
                For Each line As String In PptDesignEngine.GenerateInsights(profile)
                    insights.Add(line)
                Next

                Dim profileSummary As New JObject()
                profileSummary("rowCount") = profile.RowCount
                Dim numericNames As New JArray()
                For Each cp As PptColumnProfile In profile.NumericColumns
                    numericNames.Add(cp.Name)
                Next
                profileSummary("numericColumns") = numericNames
                profileSummary("timeSeriesColumn") = If(profile.TimeSeriesColumn Is Nothing, "", profile.TimeSeriesColumn.Name)
                profileSummary("mainCategoryColumn") = If(profile.MainCategoryColumn Is Nothing, "", profile.MainCategoryColumn.Name)

                Dim out As New JObject()
                out("success") = True
                out("slides") = slides
                out("insights") = insights
                out("profile") = profileSummary
                out("orientation") = PptDesignEngine.RecommendOrientation(profile)
                result.Content.Add(New McpTextContent With {.Text = out.ToString()})
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "数据分析失败: " & ex.Message})
            End Try
            Return result
        End Function

        ''' <summary>
        ''' ppt_generate：生成 PPT 演示文稿（.pptx 文件并自动打开）
        ''' </summary>
        Private Function ExecuteGenerate(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim out As New JObject()
                Dim keyList = New List(Of String)() From {
                    "reportTitle", "slides", "headers", "rows", "rangeAddress", "theme", "autoDesign", "requirement", "orientation", "applyLayout", "chartMode", "ratio"
                }
                For Each k In keyList
                    If args.ContainsKey(k) AndAlso args(k) IsNot Nothing Then
                        If k = "autoDesign" OrElse k = "applyLayout" Then
                            Dim b As Boolean = False
                            Boolean.TryParse(Convert.ToString(args(k)), b)
                            out(k) = b
                        Else
                            out(k) = JToken.FromObject(args(k))
                        End If
                    End If
                Next
                Dim argsJson = out.ToString(Formatting.None)
                Dim resp = PptGenerationService.Generate(argsJson)
                result.Content.Add(New McpTextContent With {.Text = resp})
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "PPT 生成失败: " & ex.Message})
            End Try
            Return result
        End Function

        ''' <summary>
        ''' ppt_import_theme：从已有 .pptx 提取配色字体，注册为自定义主题
        ''' </summary>
        Private Function ExecuteImportTheme(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim pptxPath = GetArg(Of String)(args, "pptxPath", "")
                If String.IsNullOrWhiteSpace(pptxPath) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {
                        .Text = "缺少参数 pptxPath（模板 .pptx 文件的绝对路径）"
                    })
                    Return result
                End If
                Dim resp = PptGenerationService.ImportThemeFromPptx(pptxPath)
                result.Content.Add(New McpTextContent With {.Text = resp})
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "主题导入失败: " & ex.Message})
            End Try
            Return result
        End Function

        ''' <summary>
        ''' ppt_create_theme：AI 根据需求创建新 PPT 主题（配色+字体）。
        ''' 主题 JSON 经 PptThemeDesignValidator 做 WCAG 对比度校验与自动微调，
        ''' 通过后复用 SaveCustomPptTheme 落盘并 InvalidateThemeCache 立即生效。
        ''' </summary>
        Private Function ExecuteCreateTheme(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim theme = GetArg(Of String)(args, "theme", "")
                If String.IsNullOrWhiteSpace(theme) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {
                        .Text = "缺少参数 theme（主题 JSON 字符串）：{""id"":""my_theme"",""name"":""我的主题"",""font"":""微软雅黑"",""colors"":{""primary"":[r,g,b],""secondary"":[r,g,b],""accent"":[r,g,b],""bg"":[r,g,b],""light"":[r,g,b],""text"":[r,g,b],""gray"":[r,g,b]}}。颜色也支持 ""#RRGGBB""。"
                    })
                    Return result
                End If

                Dim validation = PptThemeDesignValidator.ValidateAndFix(theme)
                Dim errors = TryCast(validation("errors"), JArray)
                If Not CBool(validation("success")) OrElse (errors IsNot Nothing AndAlso errors.Count > 0) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {
                        .Text = "主题配色未通过校验：" & Convert.ToString(validation("message")) &
                                "。请修正后重试，或改用内置主题（ppt_get_templates）。"
                    })
                    Return result
                End If

                Dim fixedJson = Convert.ToString(validation("fixedTheme"))
                Dim resp = PptGenerationService.SaveCustomPptTheme(fixedJson)

                ' 合并落盘结果与自动调整说明，返回给 LLM 以便知情
                Dim respObj As JObject = Nothing
                Try
                    respObj = JObject.Parse(resp)
                Catch
                End Try
                If respObj IsNot Nothing AndAlso CBool(respObj("success")) Then
                    Dim adjusted = TryCast(validation("adjusted"), JArray)
                    If adjusted IsNot Nothing AndAlso adjusted.Count > 0 Then
                        respObj("adjusted") = adjusted
                        respObj("message") = "主题已创建并生效。自动调整：" & String.Join("；", adjusted.Select(Function(x) Convert.ToString(x)))
                    Else
                        respObj("message") = "主题已创建并生效（配色校验通过，无调整）"
                    End If
                    result.Content.Add(New McpTextContent With {.Text = respObj.ToString(Formatting.None)})
                Else
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = Convert.ToString(resp)})
                End If
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "创建主题失败: " & ex.Message})
            End Try
            Return result
        End Function

        ''' <summary>
        ''' ppt_create_theme_from_image：从参考图提取品牌主色生成主题。
        ''' 参数：image（图片 URL 或本地路径，必填）、themeId/name（可选）、
        ''' styleDirection（可选，五选一，提取色覆盖 preset 配色、保留 preset 版式）、
        ''' font（可选）、designTokens（可选，JSON 字符串）。
        ''' 流程：提色 → 组装主题 JSON → ValidateAndFix（WCAG + 令牌）→ SaveCustomPptTheme 落盘。
        ''' </summary>
        Private Function ExecuteCreateThemeFromImage(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim image = GetArg(Of String)(args, "image", "")
                If String.IsNullOrWhiteSpace(image) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {
                        .Text = "缺少参数 image（参考图片的 URL 或本地路径）。可选参数：themeId/name（主题 id 与名称）、styleDirection（tech/academic/business/minimal/vivid）、font、designTokens（版式令牌 JSON 字符串）。"
                    })
                    Return result
                End If

                Dim palette = PptImagePaletteExtractor.Extract(image)
                If palette Is Nothing Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {
                        .Text = "图片无法读取：" & image & "。请确认路径/URL 有效且为常见图片格式（jpg/png/bmp/gif）。"
                    })
                    Return result
                End If

                Dim themeId = GetArg(Of String)(args, "themeId", "")
                Dim name = GetArg(Of String)(args, "name", "")
                Dim font = GetArg(Of String)(args, "font", "")
                Dim direction = GetArg(Of String)(args, "styleDirection", "")
                Dim tokensJson = GetArg(Of String)(args, "designTokens", "")

                Dim themeJson = PptImagePaletteExtractor.BuildThemeJson(palette, themeId, name, font, direction, tokensJson)
                Dim validation = PptThemeDesignValidator.ValidateAndFix(themeJson)
                Dim errors = TryCast(validation("errors"), JArray)
                If Not CBool(validation("success")) OrElse (errors IsNot Nothing AndAlso errors.Count > 0) Then
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {
                        .Text = "提取的配色未通过校验：" & Convert.ToString(validation("message")) &
                                "。建议调整图片或改用 ppt_create_theme 手工指定主题。"
                    })
                    Return result
                End If

                Dim fixedJson = Convert.ToString(validation("fixedTheme"))
                Dim resp = PptGenerationService.SaveCustomPptTheme(fixedJson)
                Dim respObj As JObject = Nothing
                Try
                    respObj = JObject.Parse(resp)
                Catch
                End Try
                If respObj IsNot Nothing AndAlso CBool(respObj("success")) Then
                    respObj("palette") = New JObject(
                        New JProperty("bg", New JArray(palette.Bg(0), palette.Bg(1), palette.Bg(2))),
                        New JProperty("primary", New JArray(palette.Primary(0), palette.Primary(1), palette.Primary(2))),
                        New JProperty("secondary", New JArray(palette.Secondary(0), palette.Secondary(1), palette.Secondary(2))),
                        New JProperty("accent", New JArray(palette.Accent(0), palette.Accent(1), palette.Accent(2))))
                    Dim adjusted = TryCast(validation("adjusted"), JArray)
                    If adjusted IsNot Nothing AndAlso adjusted.Count > 0 Then
                        respObj("adjusted") = adjusted
                        respObj("message") = "已从图片创建主题并生效。提取配色：" & String.Join(" / ",
                            "主色[" & palette.Primary(0) & "," & palette.Primary(1) & "," & palette.Primary(2) & "]",
                            "背景[" & palette.Bg(0) & "," & palette.Bg(1) & "," & palette.Bg(2) & "]") &
                            "。自动调整：" & String.Join("；", adjusted.Select(Function(x) Convert.ToString(x)))
                    Else
                        respObj("message") = "已从图片创建主题并生效（配色校验通过，无调整）"
                    End If
                    result.Content.Add(New McpTextContent With {.Text = respObj.ToString(Formatting.None)})
                Else
                    result.IsError = True
                    result.Content.Add(New McpTextContent With {.Text = Convert.ToString(resp)})
                End If
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "从图片创建主题失败: " & ex.Message})
            End Try
            Return result
        End Function

        ''' <summary>
        ''' ppt_browse_template：打开系统文件选择对话框，让用户选择 .pptx 模板文件
        ''' </summary>
        Private Function ExecuteBrowseTemplate(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim resp = PptGenerationService.BrowsePptFile()
                result.Content.Add(New McpTextContent With {.Text = resp})
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "打开模板选择对话框失败: " & ex.Message})
            End Try
            Return result
        End Function

        '#End Region

        '#Region "参数解析辅助"

        ''' <summary>
        ''' 安全获取参数值（与 ExcelMcpTools.GetArg 同契约）
        ''' </summary>
        Private Function GetArg(Of T)(args As Dictionary(Of String, Object), key As String, defaultValue As T) As T
            If args Is Nothing OrElse Not args.ContainsKey(key) Then
                Return defaultValue
            End If
            Try
                Dim value = args(key)
                If value Is Nothing Then Return defaultValue
                If GetType(T) = GetType(String) Then Return CType(CObj(value.ToString()), T)
                If GetType(T) = GetType(Boolean) Then
                    If TypeOf value Is Boolean Then Return CType(value, T)
                    Dim strVal = value.ToString().Trim()
                    If strVal.Equals("true", StringComparison.OrdinalIgnoreCase) Then Return CType(CObj(True), T)
                    If strVal.Equals("false", StringComparison.OrdinalIgnoreCase) Then Return CType(CObj(False), T)
                    Return defaultValue
                End If
                Return CType(value, T)
            Catch
                Return defaultValue
            End Try
        End Function

        ''' <summary>
        ''' 将参数解析为字符串列表（兼容 JSON 数组字符串 / JArray / List）
        ''' </summary>
        Private Shared Function ResolveStringList(value As Object) As List(Of String)
            Dim result = New List(Of String)()
            If value Is Nothing Then Return result
            If TypeOf value Is JArray Then
                For Each item In DirectCast(value, JArray)
                    result.Add(Convert.ToString(item))
                Next
                Return result
            End If
            If TypeOf value Is List(Of String) Then Return DirectCast(value, List(Of String))
            Dim s = Convert.ToString(value)
            If String.IsNullOrWhiteSpace(s) Then Return result
            Try
                Dim arr = JArray.Parse(s)
                For Each item In arr
                    result.Add(Convert.ToString(item))
                Next
                Return result
            Catch
                result.Add(s)
                Return result
            End Try
        End Function

        ''' <summary>
        ''' 将参数解析为数据行列表（兼容 JSON 数组字符串 / JArray）
        ''' </summary>
        Private Shared Function ResolveRows(value As Object) As List(Of Dictionary(Of String, Object))
            Dim result = New List(Of Dictionary(Of String, Object))()
            If value Is Nothing Then Return result
            Dim s = Convert.ToString(value)
            If String.IsNullOrWhiteSpace(s) Then Return result
            Try
                Dim arr = JArray.Parse(s)
                For Each item In arr
                    Dim obj = TryCast(item, JObject)
                    If obj IsNot Nothing Then
                        result.Add(obj.ToObject(Of Dictionary(Of String, Object))())
                    End If
                Next
                Return result
            Catch
                Return result
            End Try
        End Function

        '#End Region

    End Class

End Namespace
