Imports System.IO
Imports System.Reflection
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports ExcelAIVSTOAddIn.Context
Imports ExcelAIVSTOAddIn.Services
Imports ExcelAIVSTOAddIn.Services.Mcp
Imports ExcelAIVSTOAddIn.Bridges

Namespace Skills
    Public Class StepExecutionResult
        Public Property Success As Boolean
        Public Property ErrorMessage As String
        Public Property Outputs As Dictionary(Of String, Object)

        Public Sub New()
            Outputs = New Dictionary(Of String, Object)()
        End Sub
    End Class

    Public Class CustomSkillExecutor
        ''' <summary>
        ''' 合并/拆分/关联 MCP 执行结果：可读消息 + 结构化步骤数组。
        ''' Steps 会写入 result.Outputs("steps")，供前端 extractStepsPlan 渲染"步骤规划"清单。
        ''' </summary>
        Private Class MergeSplitterMcpResult
            Public Property Message As String
            Public Property Steps As List(Of String)

            Public Sub New(message As String, Optional steps As List(Of String) = Nothing)
                Me.Message = message
                Me.Steps = steps
            End Sub
        End Class

        Private _contextStore As ContextStore
        Private _llmEngine As MultimodalLLMEngine

        Public Sub New()
            _contextStore = ContextStore.GetInstance()
            _llmEngine = New MultimodalLLMEngine()
            InitializeLLM()
        End Sub

        Private Sub InitializeLLM()
            Try
                Dim config = Utils.ConfigurationManager.Instance
                Dim provider = ReadLLMConfig(config, Utils.ConfigConstants.LLM.PROVIDER, "LLM_PROVIDER", Utils.ConfigConstants.LLM.DEFAULT_PROVIDER)

                ' 根据提供商获取正确的 API URL 和模型
                Dim apiUrl As String = ""
                Dim model As String = ""
                If provider.ToLower() = "ollama" Then
                    apiUrl = config.GetSetting(Utils.ConfigConstants.Ollama.API_URL, Utils.ConfigConstants.Ollama.DEFAULT_API_URL)
                    model = ReadLLMConfig(config, Utils.ConfigConstants.LLM.MODEL, "LLM_MODEL", Utils.ConfigConstants.Ollama.DEFAULT_MODEL)
                Else
                    apiUrl = ReadLLMConfig(config, Utils.ConfigConstants.LLM.API_URL, "LLM_API_URL", "")
                    model = ReadLLMConfig(config, Utils.ConfigConstants.LLM.MODEL, "LLM_MODEL", Utils.ConfigConstants.LLM.DEFAULT_MODEL)
                End If

                Dim llmConfig = New LLMRequestConfig With {
                    .Provider = provider,
                    .Model = model,
                    .ApiKey = ReadLLMConfig(config, Utils.ConfigConstants.LLM.API_KEY, "LLM_API_KEY", ""),
                    .ApiUrl = apiUrl,
                    .SystemPrompt = config.GetSetting(Utils.ConfigConstants.CustomSkill.SYSTEM_PROMPT, Utils.ConfigConstants.CustomSkill.DEFAULT_SYSTEM_PROMPT)
                }
                _llmEngine.Initialize(llmConfig)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("初始化LLM失败: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' 读取 LLM 配置，兼容旧版下划线键名（LLM_PROVIDER 等）和新版驼峰键名（LLMProvider 等）
        ''' </summary>
        Private Function ReadLLMConfig(config As Utils.ConfigurationManager, key As String, legacyKey As String, defaultValue As String) As String
            Dim value = config.GetSetting(key, "")
            If String.IsNullOrEmpty(value) Then
                value = config.GetSetting(legacyKey, defaultValue)
            End If
            Return value
        End Function

        Public Async Function ExecuteCustomSkill(skill As SkillDefinition, inputs As Dictionary(Of String, Object)) As Task(Of WorkflowResult)
            Dim result As New WorkflowResult()
            result.Outputs = New Dictionary(Of String, Object)()

            Dim sw = System.Diagnostics.Stopwatch.StartNew()
            Dim skillName = If(skill?.Name, "Unknown")
            Dim processingMode = If(skill IsNot Nothing, skill.ProcessingMode.ToString(), "Unknown")

            Try
                If skill Is Nothing Then
                    result.Success = False
                    result.ErrorMessage = "Skill 定义为空"
                    Return result
                End If

                skill.UsageCount += 1

                ' 【扩展5】安全审批检查
                If Not Await SkillApprovalGate.Instance.CheckAndApproveAsync(skill) Then
                    result.Success = False
                    result.ErrorMessage = "用户拒绝了 Skill 执行请求: " & skill.DisplayName
                    result.ExecutionLog = New List(Of String) From {"[Approval] 用户拒绝执行"}
                    Return result
                End If

                If skill.ProcessingMode = SkillProcessingMode.MultimodalLLM AndAlso skill.LLMSystemPrompt IsNot Nothing Then
                    result = Await ExecuteLLMDrivenSkill(skill, inputs)
                ElseIf skill.ProcessingMode = SkillProcessingMode.Hybrid Then
                    ' 【优先级1扩展】Hybrid 混合模式：Phase 1 AI 分析 → Phase 2 VBA 执行
                    result = Await ExecuteHybridSkill(skill, inputs)
                ElseIf skill.IsPipeline AndAlso skill.PipelineSteps IsNot Nothing AndAlso skill.PipelineSteps.Count > 0 Then
                    ' 【扩展2】管道编排模式：多 Skill 串联执行
                    result = Await ExecutePipelineSkill(skill, inputs)
                ElseIf skill.ProcessingMode = SkillProcessingMode.PythonCode Then
                    ' 【扩展6】Python 脚本执行（架构预留）
                    result = Await ExecutePythonSkill(skill, inputs)
                ElseIf skill.CustomWorkflowSteps IsNot Nothing AndAlso skill.CustomWorkflowSteps.Count > 0 Then
                    result = Await ExecuteWorkflowSkill(skill, inputs)
                Else
                    result = ExecuteSimpleSkill(skill, inputs)
                End If

                ' 【自动导入 OCR 解析脚本】frontmatter 声明 autoImportParseScript: true 的 Skill
                ' （GenerateOCRParseScript）→ 生成的解析脚本直接写入 OCR 智能解析面板，用户不必再手动「上传 MD」。
                ' 此前该标记只有声明、没有消费者，Skill 承诺的"自动导入"从未发生。
                If skill.AutoImportParseScript AndAlso result IsNot Nothing AndAlso result.Success Then
                    Services.ParseScriptManager.AutoImportSkillParseScript(skill, result)
                End If

            Catch ex As Exception
                result.Success = False
                result.ErrorMessage = "执行自定义Skill失败: " & ex.Message
                System.Diagnostics.Debug.WriteLine("CustomSkillExecutor错误: " & ex.ToString())
            Finally
                sw.Stop()
                Dim status = If(result.Success, "成功", "失败")
                Dim logMsg = $"Skill [{skillName}] 执行完成 | 模式={processingMode} | 状态={status} | 耗时={sw.ElapsedMilliseconds}ms"
                If result.Success Then
                    Utils.AppLogger.LogInfo("CustomSkillExecutor", logMsg)
                Else
                    Utils.AppLogger.LogError("CustomSkillExecutor", $"{logMsg} | 错误={result.ErrorMessage}", Nothing)
                End If
                System.Diagnostics.Debug.WriteLine(logMsg)
            End Try

            Return result
        End Function

        Private Async Function ExecuteLLMDrivenSkill(skill As SkillDefinition, inputs As Dictionary(Of String, Object)) As Task(Of WorkflowResult)
            Dim result As New WorkflowResult()
            result.Outputs = New Dictionary(Of String, Object)()

            Try
                Dim prompt = BuildPrompt(skill, inputs)
                Dim contents = New List(Of MultimodalContent)()

                If inputs IsNot Nothing Then
                    Dim imagePath As String = ""
                    If inputs.ContainsKey("imagePath") Then
                        imagePath = inputs("imagePath")?.ToString()
                    End If
                    If imagePath IsNot Nothing AndAlso imagePath.Length > 0 AndAlso File.Exists(imagePath) Then
                        Dim imageBytes = File.ReadAllBytes(imagePath)
                        Dim base64 = Convert.ToBase64String(imageBytes)
                        contents.Add(MultimodalContent.CreateImage(base64, "image/jpeg"))
                    ElseIf inputs.ContainsKey("imageData") Then
                        contents.Add(MultimodalContent.CreateImage(inputs("imageData").ToString(), "image/jpeg"))
                    End If
                End If

                If inputs IsNot Nothing Then
                    If inputs.ContainsKey("text") Then
                        contents.Add(MultimodalContent.CreateText(inputs("text").ToString()))
                    ElseIf inputs.ContainsKey("prompt") Then
                        contents.Add(MultimodalContent.CreateText(inputs("prompt").ToString()))
                    End If
                End If
                If contents.Count = 0 Then
                    contents.Add(MultimodalContent.CreateText(prompt))
                End If

                Dim systemPrompt As String = prompt
                If skill.LLMSystemPrompt IsNot Nothing AndAlso skill.LLMSystemPrompt.Length > 0 Then
                    systemPrompt = skill.LLMSystemPrompt
                End If

                ' 【扩展1】RAG 知识库上下文增强
                Dim userQuestionForRAG As String = ""
                If inputs IsNot Nothing Then
                    If inputs.ContainsKey("text") Then userQuestionForRAG = inputs("text")?.ToString()
                    If inputs.ContainsKey("prompt") Then userQuestionForRAG = inputs("prompt")?.ToString()
                End If
                If String.IsNullOrWhiteSpace(userQuestionForRAG) Then userQuestionForRAG = prompt
                systemPrompt = Await EnrichWithRagContextAsync(skill, userQuestionForRAG, systemPrompt)

                Dim llmResult = Await _llmEngine.ProcessContentAsync(contents, systemPrompt)

                If llmResult.Success Then
                    result.Success = True
                    result.Outputs("content") = llmResult.Content

                    Try
                        Dim json = TryParseJson(llmResult.Content)
                        If json IsNot Nothing Then
                            For Each prop In json.Properties()
                                result.Outputs(prop.Name) = prop.Value?.ToObject(Of Object)()
                            Next
                        End If
                    Catch ex As Exception  ' 
                    End Try

                    If skill.FieldMappings IsNot Nothing AndAlso skill.FieldMappings.Count > 0 Then
                        WriteMappedFields(result, skill)
                    End If
                Else
                    result.Success = False
                    result.ErrorMessage = llmResult.ErrorMessage
                End If

            Catch ex As Exception
                result.Success = False
                result.ErrorMessage = ex.Message
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 【优先级1扩展】Hybrid 混合模式执行器
        ''' 两阶段流水线：Phase 1 LLM 分析/决策 → Phase 2 将分析结果注入 VBA 并执行
        ''' 适用场景：先 AI 理解数据再操作 Excel（异常检测高亮、智能格式化、分析后写入等）
        ''' </summary>
        Private Async Function ExecuteHybridSkill(skill As SkillDefinition, inputs As Dictionary(Of String, Object)) As Task(Of WorkflowResult)
            Dim result As New WorkflowResult()
            result.Outputs = New Dictionary(Of String, Object)()
            result.ExecutionLog = New List(Of String)()

            Try
                ' ========== Phase 1: LLM 智能分析 ==========
                result.ExecutionLog.Add("[Hybrid Phase 1] 开始 LLM 分析...")

                Dim analysisContents = New List(Of MultimodalContent)()

                ' 构建输入内容：支持文本 + 图片混合
                If inputs IsNot Nothing Then
                    ' 图片输入
                    Dim imagePath As String = ""
                    If inputs.ContainsKey("imagePath") Then
                        imagePath = inputs("imagePath")?.ToString()
                    End If
                    If imagePath IsNot Nothing AndAlso imagePath.Length > 0 AndAlso File.Exists(imagePath) Then
                        Dim imageBytes = File.ReadAllBytes(imagePath)
                        analysisContents.Add(MultimodalContent.CreateImage(Convert.ToBase64String(imageBytes), "image/jpeg"))
                    ElseIf inputs.ContainsKey("imageData") Then
                        analysisContents.Add(MultimodalContent.CreateImage(inputs("imageData").ToString(), "image/jpeg"))
                    End If

                    ' 文本输入
                    If inputs.ContainsKey("text") Then
                        analysisContents.Add(MultimodalContent.CreateText(inputs("text").ToString()))
                    ElseIf inputs.ContainsKey("prompt") Then
                        analysisContents.Add(MultimodalContent.CreateText(inputs("prompt").ToString()))
                    End If

                    ' 将当前 Excel 选区数据作为上下文传入 LLM（关键：让 AI "看到"数据）
                    If inputs.ContainsKey("excelData") Then
                        Dim excelDataText As String = inputs("excelData")?.ToString()
                        If excelDataText IsNot Nothing AndAlso excelDataText.Length > 0 Then
                            analysisContents.Add(MultimodalContent.CreateText(
                                vbCrLf & "[当前Excel数据]" & vbCrLf & excelDataText))
                        End If
                    End If
                End If

                If analysisContents.Count = 0 Then
                    analysisContents.Add(MultimodalContent.CreateText("请分析当前请求。"))
                End If

                ' 使用 Skill 定义的 System Prompt 进行分析
                Dim systemPrompt As String = skill.LLMSystemPrompt
                If String.IsNullOrWhiteSpace(systemPrompt) Then
                    systemPrompt = "你是 Excel 数据分析助手。分析用户需求并输出结构化的决策结果。"
                End If

                ' 【扩展1】RAG 知识库上下文增强（Hybrid 模式 Phase 1 也支持 RAG）
                Dim hybridUserQuestion As String = ""
                If inputs IsNot Nothing Then
                    If inputs.ContainsKey("text") Then hybridUserQuestion = inputs("text")?.ToString()
                    If inputs.ContainsKey("prompt") Then hybridUserQuestion = inputs("prompt")?.ToString()
                End If
                If String.IsNullOrWhiteSpace(hybridUserQuestion) Then hybridUserQuestion = "请分析当前请求。"
                systemPrompt = Await EnrichWithRagContextAsync(skill, hybridUserQuestion, systemPrompt)

                Dim llmResult = Await _llmEngine.ProcessContentAsync(analysisContents, systemPrompt)

                If Not llmResult.Success Then
                    result.Success = False
                    result.ErrorMessage = "Hybrid Phase 1 (LLM分析) 失败: " & llmResult.ErrorMessage
                    Return result
                End If

                Dim analysisOutput As String = llmResult.Content
                result.Outputs("analysis") = analysisOutput
                result.ExecutionLog.Add(String.Format("[Hybrid Phase 1] LLM 分析完成，输出长度: {0}", analysisOutput.Length))

                ' 尝试解析 LLM 输出为结构化 JSON（用于注入 VBA 变量）
                Dim analysisJson As Newtonsoft.Json.Linq.JObject = TryParseJson(analysisOutput)
                If analysisJson IsNot Nothing Then
                    For Each prop In analysisJson.Properties()
                        result.Outputs("llm_" & prop.Name) = prop.Value?.ToObject(Of Object)()
                    Next
                    result.ExecutionLog.Add(String.Format("[Hybrid Phase 1] 解析到 {0} 个结构化字段", analysisJson.Properties().Count()))
                End If

                ' ========== Phase 2: VBA 执行（将 AI 分析结果作为变量注入）==========
                result.ExecutionLog.Add("[Hybrid Phase 2] 开始 VBA 执行...")

                Dim vbaCode As String = ""
                If skill.TemplateCode IsNot Nothing AndAlso skill.TemplateCode.Length > 0 Then
                    vbaCode = skill.TemplateCode
                Else
                    ' 无模板代码时，仅返回分析结果
                    result.Success = True
                    result.Outputs("message") = String.Format("Skill '{0}' 分析完成（无VBA执行阶段）", skill.DisplayName)
                    Return result
                End If

                Dim executionResult As Object = Nothing
                Dim executionResultMessage As String = ""

                ' SmartDataAnalyzer 优先通过 VSTO Interop 直接操作 Excel，避免 VBA 宏依赖和模板注入问题
                If String.Equals(skill.Name, "SmartDataAnalyzer", StringComparison.OrdinalIgnoreCase) Then
                    result.ExecutionLog.Add("[Hybrid Phase 2] SmartDataAnalyzer 使用 VSTO Interop 直接操作 Excel")
                    Dim interopMessage = ExecuteSmartDataAnalyzerInterop(analysisJson)
                    executionResult = New With { .Message = interopMessage, .ExecutionMethod = "VSTOInterop" }
                    executionResultMessage = interopMessage
                ElseIf IsMergeSplitterFamily(skill.Name) Then
                    ' DataMerger/DataSplitter/TableJoiner 必须使用 MCP 工具完成，绝对不能执行有 "targetWs.Cells.Clear" 风险的 VBA 模板代码
                    result.ExecutionLog.Add($"[Hybrid Phase 2] {skill.DisplayName} 使用 MCP 工具执行（避免 VBA 清空数据的风险）")
                    Dim mcpResult = Await ExecuteMergeSplitterFamilyMcp(skill, inputs, analysisJson)
                    executionResult = New With {.Message = mcpResult.Message, .Steps = mcpResult.Steps, .ExecutionMethod = "MCPTools"}
                    executionResultMessage = mcpResult.Message
                    ' 把结构化步骤放进 Outputs，供前端 extractStepsPlan 渲染"步骤规划"清单
                    If mcpResult.Steps IsNot Nothing AndAlso mcpResult.Steps.Count > 0 Then
                        result.Outputs("steps") = mcpResult.Steps
                    End If
                Else
                    ' 将 LLM 分析结果中的关键字段替换进 VBA 代码的占位符
                    Dim finalVbaCode = InjectAnalysisToVBA(vbaCode, analysisOutput, analysisJson)

                    ' 在 VBA 沙箱中执行（用户主动触发的 Skill，跳过确认弹窗）
                    executionResult = Security.VBACodeSandbox.ExecuteCode(finalVbaCode, requireConfirmation:=False)

                    ' 解析执行结果中的 Message，判断是否真的成功
                    Try
                        executionResultMessage = CallByName(executionResult, "Message", CallType.Get).ToString()
                    Catch ex As Exception  ' 
                    End Try

                    result.ExecutionLog.Add("[Hybrid Phase 2] VBA 执行完成")
                End If

                result.Outputs("executionResult") = executionResult

                ' 如果有字段映射，写入 Excel
                If skill.FieldMappings IsNot Nothing AndAlso skill.FieldMappings.Count > 0 Then
                    WriteMappedFields(result, skill)
                    result.ExecutionLog.Add("[Hybrid Phase 2] 字段映射已写入 Excel")
                End If

                If executionResultMessage IsNot Nothing AndAlso executionResultMessage.StartsWith("✗") Then
                    result.Success = False
                    result.ErrorMessage = executionResultMessage.TrimStart("✗"c).Trim()
                    result.Outputs("message") = String.Format("Skill '{0}' 执行失败: {1}", skill.DisplayName, result.ErrorMessage)
                Else
                    result.Success = True
                    ' 把实际执行结果（含合并/拆分/关联的步骤规划）拼进 message，
                    ' 让"规划步骤"在技能卡片与 LLM 原生两条路径都可见，避免只显示一句"执行成功"。
                    If Not String.IsNullOrWhiteSpace(executionResultMessage) Then
                        result.Outputs("message") = String.Format("Skill '{0}' Hybrid 执行成功：{1}", skill.DisplayName, executionResultMessage)
                    Else
                        result.Outputs("message") = String.Format("Skill '{0}' Hybrid 执行成功", skill.DisplayName)
                    End If
                End If

            Catch ex As Exception
                result.Success = False
                result.ErrorMessage = "Hybrid Skill 执行失败: " & ex.Message
                result.ExecutionLog.Add("错误: " & ex.Message)
                System.Diagnostics.Debug.WriteLine("ExecuteHybridSkill 错误: " & ex.ToString())
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 【优先级1扩展】将 LLM 分析结果注入 VBA 代码
        ''' 支持三种注入方式：
        '''   1. {{ANALYSIS_RAW}} - 原始分析文本
        '''   2. {{ANALYSIS_JSON.path}} - JSON 字段值，支持嵌套路径（如 statistics.issueCount）
        '''   3. {{ANALYSIS_FIELD:key}} - 备用字段访问格式
        ''' </summary>
        Private Function InjectAnalysisToVBA(vbaCode As String, rawAnalysis As String, analysisJson As Newtonsoft.Json.Linq.JObject) As String
            Dim result = vbaCode

            ' 注入原始分析文本
            result = result.Replace("{{ANALYSIS_RAW}}", EscapeVBAString(rawAnalysis))

            ' 注入 JSON 结构化字段（支持嵌套路径）
            If analysisJson IsNot Nothing Then
                Dim jsonPlaceholderRegex As New System.Text.RegularExpressions.Regex("\{\{ANALYSIS_JSON\.([^{}]+)\}\}")
                result = jsonPlaceholderRegex.Replace(result, Function(m)
                    Dim path = m.Groups(1).Value
                    Dim fieldValue = GetJsonValueByPath(analysisJson, path)
                    Return EscapeVBAString(fieldValue)
                End Function)

                ' 也支持 {{ANALYSIS_FIELD:fieldName}} 格式（仅顶层）
                For Each prop In analysisJson.Properties()
                    Dim fieldValue As String = ""
                    If prop.Value IsNot Nothing Then
                        fieldValue = prop.Value.ToString()
                    End If

                    result = result.Replace(
                        String.Format("{{{{ANALYSIS_FIELD:{0}}}}}", prop.Name),
                        EscapeVBAString(fieldValue))
                Next
            End If

            Return result
        End Function

        ''' <summary>
        ''' 根据点分隔路径从 JObject 中取值（支持简单数组下标）
        ''' </summary>
        Private Function GetJsonValueByPath(json As Newtonsoft.Json.Linq.JObject, path As String) As String
            If json Is Nothing OrElse String.IsNullOrWhiteSpace(path) Then Return String.Empty

            Dim token As Newtonsoft.Json.Linq.JToken = json
            For Each part In path.Split("."c)
                If token Is Nothing Then Return String.Empty

                If TypeOf token Is Newtonsoft.Json.Linq.JObject Then
                    Dim child = TryCast(token, Newtonsoft.Json.Linq.JObject)(part)
                    token = child
                ElseIf TypeOf token Is Newtonsoft.Json.Linq.JArray Then
                    Dim index As Integer
                    If Integer.TryParse(part, index) Then
                        token = TryCast(token, Newtonsoft.Json.Linq.JArray)(index)
                    Else
                        Return String.Empty
                    End If
                Else
                    Return String.Empty
                End If
            Next

            If token Is Nothing Then Return String.Empty
            Return token.ToString()
        End Function

        ''' <summary>
        ''' SmartDataAnalyzer 的 VSTO 直接实现：根据 LLM 返回的 issues 对 Excel 选区进行高亮/格式化。
        ''' 避免依赖 VBA 宏注入和执行。
        ''' </summary>
        Private Function ExecuteSmartDataAnalyzerInterop(analysisJson As Newtonsoft.Json.Linq.JObject) As String
            Try
                Dim excelApp = Globals.ThisAddIn.GetExcelApp()
                If excelApp Is Nothing Then
                    Return "✗ 无法获取 Excel 应用"
                End If

                Dim ws = TryCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                If ws Is Nothing Then
                    Return "✗ 没有活动工作表"
                End If

                Dim rng = TryCast(excelApp.Selection, Microsoft.Office.Interop.Excel.Range)
                If rng Is Nothing OrElse rng.Cells.Count < 2 Then
                    Return "✗ 请先选择要分析的数据区域（至少2个单元格）"
                End If

                ' 重置选区格式
                rng.Interior.ColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexNone
                rng.Font.Bold = False
                rng.Font.Italic = False
                rng.Font.Strikethrough = False

                Dim issueCount As Integer = 0
                Dim summary As String = ""

                If analysisJson IsNot Nothing Then
                    Dim stats = analysisJson("statistics")
                    If stats IsNot Nothing AndAlso stats("issueCount") IsNot Nothing Then
                        issueCount = stats("issueCount").ToObject(Of Integer)()
                    End If

                    If analysisJson("summary") IsNot Nothing Then
                        summary = analysisJson("summary").ToString()
                    End If

                    Dim issues = analysisJson("issues")
                    If issues IsNot Nothing AndAlso TypeOf issues Is Newtonsoft.Json.Linq.JArray Then
                        For Each issue In CType(issues, Newtonsoft.Json.Linq.JArray)
                            Dim rowToken = issue("row")
                            Dim colToken = issue("column")
                            Dim actionToken = issue("action")
                            If rowToken Is Nothing OrElse colToken Is Nothing Then Continue For

                            Dim row As Integer
                            If Not Integer.TryParse(rowToken.ToString(), row) Then Continue For
                            Dim col As String = colToken.ToString()
                            If String.IsNullOrWhiteSpace(col) Then Continue For

                            Dim cell = ws.Range(col & row)
                            If cell Is Nothing Then Continue For

                            Dim action = If(actionToken IsNot Nothing, actionToken.ToString().ToLower(), "")
                            Select Case action
                                Case "highlight_red"
                                    cell.Interior.Color = System.Drawing.Color.Red.ToArgb()
                                Case "highlight_yellow"
                                    cell.Interior.Color = System.Drawing.Color.Yellow.ToArgb()
                                Case "highlight_orange"
                                    cell.Interior.Color = System.Drawing.Color.Orange.ToArgb()
                                Case "bold"
                                    cell.Font.Bold = True
                                Case "italic"
                                    cell.Font.Italic = True
                                Case "strikethrough"
                                    cell.Font.Strikethrough = True
                            End Select
                        Next
                    End If
                End If

                If issueCount > 0 Then
                    Return String.Format("✓ 分析完成：{0}，共发现 {1} 个问题，已自动标记。", summary, issueCount)
                Else
                    Return "✓ 未发现明显问题，数据质量良好！"
                End If

            Catch ex As Exception
                Return "✗ 执行失败: " & ex.Message
            End Try
        End Function

        ''' <summary>
        ''' DataMerger/DataSplitter/TableJoiner 的 MCP 工具执行：根据 skill 输入 + LLM 分析选择对应 MCP 工具。
        ''' 每个 Skill 只允许执行自己职责范围内的操作（白名单校验）。
        ''' 【安全保证】MCP 工具绝不清空用户选中的原始数据；结果写入新建的独立 Sheet / 文件。
        ''' </summary>
        Private Async Function ExecuteMergeSplitterFamilyMcp(skill As SkillDefinition,
                                                           inputs As Dictionary(Of String, Object),
                                                           analysisJson As Newtonsoft.Json.Linq.JObject) As Task(Of MergeSplitterMcpResult)
            ' BC42356 说明：本方法体全部为同步逻辑（MCP 工具同步执行 + 消息组装），没有异步等待点。
            ' 保留 Async 签名以兼容调用方 Await；其内部 Excel COM 操作必须留在当前线程，故不能包 Task.Run。
            Await Task.CompletedTask
            Try
                ' 1. 确定操作类型：优先从 LLM 分析解析，否则从 Skill 配置表单 inputs 读取
                Dim operationType As String = ""
                If analysisJson IsNot Nothing Then
                    Dim op = If(analysisJson("operationType"), analysisJson("operation"))
                    If op IsNot Nothing Then operationType = op.ToString().ToLower().Trim()
                End If
                If String.IsNullOrWhiteSpace(operationType) AndAlso inputs IsNot Nothing Then
                    If inputs.ContainsKey("operationType") Then operationType = inputs("operationType")?.ToString().ToLower().Trim()
                    If String.IsNullOrWhiteSpace(operationType) AndAlso inputs.ContainsKey("operation") Then operationType = inputs("operation")?.ToString().ToLower().Trim()
                End If

                ' 兜底：从默认参数 / 常见关键词推断
                If String.IsNullOrWhiteSpace(operationType) Then
                    Dim allText = ""
                    If analysisJson IsNot Nothing Then allText = analysisJson.ToString()
                    If inputs IsNot Nothing Then
                        For Each kv In inputs
                            allText &= " " & If(kv.Value, "").ToString()
                        Next
                    End If
                    If allText.Contains("按部门") OrElse allText.Contains("按列拆分") OrElse allText.Contains("分类拆分") OrElse allText.Contains("拆分到不同sheet") Then
                        operationType = "split_by_column"
                    ElseIf allText.Contains("每1000行") OrElse allText.Contains("行数拆分") OrElse allText.Contains("分批") Then
                        operationType = "split_by_rows"
                    ElseIf allText.Contains("合并文件") OrElse allText.Contains("合并工作簿") Then
                        operationType = "merge_files"
                    ElseIf allText.Contains("跨表") OrElse allText.Contains("JOIN") OrElse allText.Contains("关联") OrElse allText.Contains("匹配") Then
                        operationType = "join_tables"
                    Else
                        operationType = DefaultOperationForSkill(skill.Name) ' 默认：按当前 Skill 职责取合理操作
                    End If
                End If

                ' 统一为内部枚举名，兼容别名
                Select Case operationType
                    Case "merge_sheets", "mergeall", "merge_all_sheets", "合并所有sheet", "合并sheet", "合并工作表"
                        operationType = "merge_sheets"
                    Case "merge_files", "merge_workbooks", "合并文件", "合并工作簿", "汇总文件"
                        operationType = "merge_files"
                    Case "split_by_column", "splitbycolumn", "按列拆分", "按部门拆分", "按字段拆分"
                        operationType = "split_by_column"
                    Case "split_by_rows", "splitbyrows", "split_row_count", "按行数拆分", "每n行拆分"
                        operationType = "split_by_rows"
                    Case "join_tables", "join", "关联", "匹配", "跨表", "vlookup", "合并匹配"
                        operationType = "join_tables"
                    Case Else
                        operationType = DefaultOperationForSkill(skill.Name)
                End Select

                ' 2.5 按 skill.Name 限定允许的操作类型，防止跨 Skill 误调度
                Dim allowedOps As New List(Of String)()
                Select Case skill.Name.ToLower()
                    Case "datamerger" : allowedOps.AddRange({"merge_sheets", "merge_files"})
                    Case "datasplitter" : allowedOps.AddRange({"split_by_column", "split_by_rows"})
                    Case "tablejoiner" : allowedOps.Add("join_tables")
                End Select
                If allowedOps.Count > 0 AndAlso Not allowedOps.Contains(operationType) Then
                    Return New MergeSplitterMcpResult("✗ 操作类型(" & operationType & ")不属于 '" & skill.DisplayName & "' 的职责范围，请改用对应的合并/拆分/关联 Skill。")
                End If

                ' 2. 收集相关参数（优先 LLM 分析，其次 skill inputs）
                Dim args As New Dictionary(Of String, Object)()

                ' Helper：从 LLM/inputs 取参数，两者都没有就给默认值。
                ' LLM 分析结果可能把参数嵌套在 parameters 子对象里（如 { parameters: { column: "部门" } }），
                ' 因此顶层取不到时还要回退到 analysisJson("parameters")(key)，否则按列拆分等场景会误报"缺少参数"。
                Dim pick As Func(Of String, Object, Object) = Function(key, [default])
                                                                   If analysisJson IsNot Nothing Then
                                                                       If analysisJson(key) IsNot Nothing Then Return analysisJson(key).ToObject(Of Object)()
                                                                       Dim p = analysisJson("parameters")
                                                                       If p IsNot Nothing AndAlso p(key) IsNot Nothing Then Return p(key).ToObject(Of Object)()
                                                                   End If
                                                                   If inputs IsNot Nothing AndAlso inputs.ContainsKey(key) Then Return inputs(key)
                                                                   Return [default]
                                                               End Function

                ' 3. 分发到具体 MCP 工具（传入 SkillBridge.Instance，执行 merge/split 操作时不会用到 bridge，这里仅用于满足构造函数）
                Dim toolsInst = New ExcelMcpTools(SkillBridge.Instance)
                Select Case operationType
                    Case "merge_sheets"
                        args("targetSheet") = pick("targetSheet", "合并结果").ToString()
                        args("includeHeaders") = CBool(pick("includeHeaders", True))
                        args("copyFormatting") = CBool(pick("copyFormatting", False))
                        ' 用户明确说"合并所有/所有Sheet/全部表"时，强制合并全部可见表，
                        ' 忽略当前选中的工作表标签，避免"合并所有"只合并了选中的几张表（合并结果仅覆盖特定范围）。
                        args("mergeAll") = WantsMergeAllSheets(analysisJson, inputs)
                        Dim srcLst = pick("sourceSheets", Nothing)
                        If srcLst IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(srcLst.ToString()) Then
                            args("sourceSheets") = If(TypeOf srcLst Is String, srcLst.ToString(), JsonConvert.SerializeObject(srcLst))
                        End If
                        Dim res = toolsInst.ExecuteTool("excel_merge_sheets", args)
                        If res IsNot Nothing AndAlso res.IsError Then
                            Return New MergeSplitterMcpResult(ReadMcpMessage(res, "合并所有Sheet完成"))
                        End If
                        ' 把合并结果组装为可读消息（含源表清单、行数、执行步骤），
                        ' 让前端/LLM 能直接看到"规划步骤"，避免只返回裸 JSON 导致步骤不可见。
                        Dim mergeMsg = ComposeMergeMessage(res, "合并所有Sheet完成")
                        ' 同时提取结构化 steps 数组，供前端"步骤规划"清单渲染
                        Dim mergeSteps = ExtractStepsFromMcpResult(res)

                        ' 用户要求"合并并去重"时，合并完成后对结果表执行去重，
                        ' 避免"只合并不去重"却谎称已去重（原 excel_merge_sheets 无去重逻辑）。
                        If ShouldDeduplicateMerge(analysisJson, inputs) Then
                            Dim dedupArgs As New Dictionary(Of String, Object)()
                            dedupArgs("sheet") = args("targetSheet")
                            dedupArgs("hasHeaderRow") = CBool(pick("includeHeaders", True)).ToString().ToLower()
                            Dim dedupRes = toolsInst.ExecuteTool("excel_remove_duplicates", dedupArgs)
                            Dim dedupMsg = ReadMcpMessage(dedupRes, "合并结果去重完成")
                            Return New MergeSplitterMcpResult(mergeMsg & vbCrLf & dedupMsg, mergeSteps)
                        End If
                        Return New MergeSplitterMcpResult(mergeMsg, mergeSteps)

                    Case "merge_files"
                        Dim rawFiles = pick("filePaths", Nothing)
                        If rawFiles Is Nothing Then rawFiles = pick("files", Nothing)
                        If rawFiles Is Nothing OrElse String.IsNullOrWhiteSpace(rawFiles.ToString()) Then
                            Return New MergeSplitterMcpResult("✗ 合并文件模式需要提供文件路径列表(filePaths)。请选择要合并的文件后重试。")
                        End If
                        If TypeOf rawFiles Is String Then
                            args("filePaths") = rawFiles.ToString()
                        Else
                            args("filePaths") = JsonConvert.SerializeObject(rawFiles)
                        End If
                        args("targetPath") = pick("targetPath", "").ToString()
                        args("sheetPrefix") = pick("sheetPrefix", "").ToString()
                        Dim res2 = toolsInst.ExecuteTool("excel_merge_workbooks", args)
                        Return New MergeSplitterMcpResult(ReadMcpMessage(res2, "合并文件完成"))

                    Case "split_by_column"
                        Dim col = pick("keyColumn", pick("column", "")).ToString()
                        If String.IsNullOrWhiteSpace(col) Then
                            col = pick("splitColumn", pick("splitByColumn", "")).ToString()
                        End If
                        If String.IsNullOrWhiteSpace(col) Then
                            Return New MergeSplitterMcpResult("✗ 按列拆分模式需要指定拆分列（keyColumn 或 column 参数）。请输入要按哪一列拆分。")
                        End If
                        args("column") = col
                        args("outputMode") = pick("outputMode", pick("targetMode", "sheets")).ToString()
                        args("outputDir") = pick("targetFolder", pick("outputDir", "")).ToString()
                        args("namingPattern") = pick("namingPattern", "{original}_{value}").ToString()
                        Dim res3 = toolsInst.ExecuteTool("excel_split_by_column", args)
                        If res3 IsNot Nothing AndAlso res3.IsError Then
                            Return New MergeSplitterMcpResult(ReadMcpMessage(res3, "按列拆分完成"))
                        End If
                        ' 与合并对称：把拆分结果组装为可读消息（含拆分摘要、生成清单、执行步骤），
                        ' 并提取结构化 steps 数组，供前端"步骤规划"清单渲染，避免只返回裸 JSON。
                        Dim splitMsg = ComposeSplitMessage(res3, "按列拆分完成", "按 '" & col & "' ")
                        Dim splitSteps = ExtractStepsFromMcpResult(res3)
                        Return New MergeSplitterMcpResult(splitMsg, splitSteps)

                    Case "split_by_rows"
                        Dim rowsStr = pick("rowsPerFile", pick("rowsPerChunk", 1000)).ToString()
                        Dim rowsNum As Integer
                        If Not Integer.TryParse(rowsStr, rowsNum) OrElse rowsNum <= 0 Then rowsNum = 1000
                        args("rowsPerChunk") = rowsNum
                        args("outputMode") = pick("outputMode", pick("targetMode", "sheets")).ToString()
                        args("outputDir") = pick("targetFolder", pick("outputDir", "")).ToString()
                        args("keepHeader") = CBool(pick("keepHeader", True))
                        args("namingPattern") = pick("namingPattern", "{original}_{index}_{start}_{end}").ToString()
                        Dim res4 = toolsInst.ExecuteTool("excel_split_by_rows", args)
                        If res4 IsNot Nothing AndAlso res4.IsError Then
                            Return New MergeSplitterMcpResult(ReadMcpMessage(res4, "按行数拆分完成"))
                        End If
                        ' 与合并对称：把按行数拆分结果组装为可读消息，并提取结构化 steps 数组。
                        Dim splitMsg2 = ComposeSplitMessage(res4, "按行数拆分完成", "按每 " & rowsNum.ToString() & " 行 ")
                        Dim splitSteps2 = ExtractStepsFromMcpResult(res4)
                        Return New MergeSplitterMcpResult(splitMsg2, splitSteps2)

                    Case "join_tables"
                        Dim ls = pick("leftSheet", pick("sheet1", "")).ToString()
                        Dim rs = pick("rightSheet", pick("sheet2", "")).ToString()
                        Dim lk = pick("joinKey", pick("leftKeyColumn", pick("keyColumn", ""))).ToString()
                        If String.IsNullOrWhiteSpace(ls) OrElse String.IsNullOrWhiteSpace(rs) Then
                            Return New MergeSplitterMcpResult("✗ 跨表关联需要指定左表工作表(leftSheet)和右表工作表(rightSheet)。请填写要关联的两个Sheet名。")
                        End If
                        If String.IsNullOrWhiteSpace(lk) Then
                            Return New MergeSplitterMcpResult("✗ 跨表关联需要指定关键列(joinKey)。请输入按哪一列进行匹配。")
                        End If
                        args("leftSheet") = ls
                        args("rightSheet") = rs
                        args("leftKeyColumn") = lk
                        args("rightKeyColumn") = pick("rightKeyColumn", lk).ToString()
                        args("joinType") = pick("joinType", pick("matchMode", "inner")).ToString()
                        args("targetSheet") = pick("targetSheet", "关联结果").ToString()
                        Dim res5 = toolsInst.ExecuteTool("excel_join_tables", args)
                        Return New MergeSplitterMcpResult(ReadMcpMessage(res5, "跨表关联完成"))

                    Case Else
                        Return New MergeSplitterMcpResult("✗ 无法识别的操作类型: " & operationType)
                End Select

            Catch ex As Exception
                Return New MergeSplitterMcpResult("✗ " & skill.DisplayName & " MCP 执行失败: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' 判断是否为 数据合并/拆分/跨表关联 家庭 Skill（走 MCP 工具，禁止 VBA 模板）。
        ''' </summary>
        Private Shared Function IsMergeSplitterFamily(skillName As String) As Boolean
            Select Case skillName.ToLower()
                Case "datamerger", "datasplitter", "tablejoiner"
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        ''' <summary>
        ''' 当 LLM/表单未给出操作类型时，按当前 Skill 职责返回合理默认操作。
        ''' </summary>
        Private Shared Function DefaultOperationForSkill(skillName As String) As String
            Select Case skillName.ToLower()
                Case "datasplitter" : Return "split_by_rows"
                Case "tablejoiner" : Return "join_tables"
                Case Else : Return "merge_sheets"
            End Select
        End Function

        ''' <summary>
        ''' 读取 MCP 工具返回的信息，如果失败则返回错误消息，成功返回 message 字段或工具 JSON 摘要
        ''' </summary>
        Private Shared Function ReadMcpMessage(mcpResult As ToolsCallResult, defaultSuccessMsg As String) As String
            If mcpResult Is Nothing Then Return "✗ MCP 工具未返回结果"
            If mcpResult.IsError Then
                Dim errTxt = ""
                If mcpResult.Content IsNot Nothing Then
                    For Each c In mcpResult.Content
                        Try
                            Dim txtC = TryCast(c, McpTextContent)
                            If txtC IsNot Nothing Then errTxt &= txtC.Text & " "
                        Catch ex As Exception  ' 
                        End Try
                    Next
                End If
                Return "✗ " & If(String.IsNullOrWhiteSpace(errTxt), "未知错误", errTxt.Trim())
            End If
            If mcpResult.Content IsNot Nothing Then
                For Each c In mcpResult.Content
                    Try
                        Dim txtC = TryCast(c, McpTextContent)
                        If txtC Is Nothing Then Continue For
                        Dim txt = txtC.Text
                        If String.IsNullOrWhiteSpace(txt) Then Continue For
                        ' 如果是 JSON，尝试提取 message 字段
                        If txt.TrimStart().StartsWith("{") Then
                            Try
                                Dim j = Newtonsoft.Json.Linq.JObject.Parse(txt)
                                If j("message") IsNot Nothing Then
                                    Return "✓ " & j("message").ToString()
                                End If
                                If j("success") IsNot Nothing AndAlso j("success").ToObject(Of Boolean)() Then
                                    Return "✓ " & defaultSuccessMsg & "：" & txt
                                End If
                            Catch ex As Exception  ' 
                            End Try
                        End If
                        Return "✓ " & txt
                    Catch ex As Exception  ' 
                    End Try
                Next
            End If
            Return "✓ " & defaultSuccessMsg
        End Function

        ''' <summary>
        ''' 把合并结果 JSON 组装为可读消息（含源表清单、行数、执行步骤），
        ''' 让前端/LLM 能直接看到"规划步骤"，避免只返回裸 JSON 导致步骤不可见。
        ''' </summary>
        Private Shared Function ComposeMergeMessage(mcpResult As ToolsCallResult, defaultMsg As String) As String
            If mcpResult Is Nothing OrElse mcpResult.Content Is Nothing Then Return "✓ " & defaultMsg
            For Each c In mcpResult.Content
                Try
                    Dim txtC = TryCast(c, McpTextContent)
                    If txtC Is Nothing OrElse String.IsNullOrWhiteSpace(txtC.Text) Then Continue For
                    Dim txt = txtC.Text.Trim()
                    If Not txt.StartsWith("{") Then Return "✓ " & txt
                    Dim j = Newtonsoft.Json.Linq.JObject.Parse(txt)
                    Dim parts As New List(Of String)()
                    parts.Add("✓ " & defaultMsg)
                    If j("sourceSheets") IsNot Nothing AndAlso j("sourceSheets").Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                        Dim names = j("sourceSheets").Select(Function(t) t.ToString()).ToList()
                        If names.Count > 0 Then parts.Add("已合并 " & names.Count & " 张工作表：" & String.Join("、", names))
                    End If
                    If j("mergedRows") IsNot Nothing Then
                        parts.Add("合并结果共 " & j("mergedRows").ToString() & " 行（含表头）")
                    End If
                    If j("steps") IsNot Nothing AndAlso j("steps").Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                        Dim steps = j("steps").Select(Function(t) t.ToString()).ToList()
                        If steps.Count > 0 Then
                            parts.Add("执行步骤：")
                            For i = 0 To steps.Count - 1
                                parts.Add((i + 1).ToString() & ". " & steps(i))
                            Next
                        End If
                    End If
                    Return String.Join(vbCrLf, parts)
                Catch ex As Exception  ' 
                End Try
            Next
            Return "✓ " & defaultMsg
        End Function

        ''' <summary>
        ''' 把拆分结果 JSON 组装为可读消息（含拆分摘要、生成清单、执行步骤），
        ''' 与 ComposeMergeMessage 对称，让前端/LLM 直接看到"规划步骤"，
        ''' 避免拆分只返回裸 JSON 导致步骤不可见。
        ''' summaryPrefix 形如 "按 '部门' " 或 "按每 1000 行 "，用于拼装可读摘要行。
        ''' </summary>
        Private Shared Function ComposeSplitMessage(mcpResult As ToolsCallResult, defaultMsg As String, summaryPrefix As String) As String
            If mcpResult Is Nothing OrElse mcpResult.Content Is Nothing Then Return "✓ " & defaultMsg
            For Each c In mcpResult.Content
                Try
                    Dim txtC = TryCast(c, McpTextContent)
                    If txtC Is Nothing OrElse String.IsNullOrWhiteSpace(txtC.Text) Then Continue For
                    Dim txt = txtC.Text.Trim()
                    If Not txt.StartsWith("{") Then Return "✓ " & txt
                    Dim j = Newtonsoft.Json.Linq.JObject.Parse(txt)
                    Dim parts As New List(Of String)()
                    parts.Add("✓ " & defaultMsg)

                    ' 拆分摘要：{prefix}拆分为 {N} 组/份，生成 {M} 个工作表/文件：{names}
                    Dim countToken = If(j("groups"), j("chunks"))
                    If countToken IsNot Nothing Then
                        Dim unit = If(j("groups") IsNot Nothing, "组", "份")
                        Dim line = summaryPrefix & "拆分为 " & countToken.ToString() & " " & unit
                        If j("created") IsNot Nothing AndAlso j("created").Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                            Dim created = j("created").Select(Function(t) t.ToString()).ToList()
                            If created.Count > 0 Then
                                Dim isFile = created(0).IndexOf("\") >= 0 OrElse created(0).IndexOf("/") >= 0 OrElse created(0).IndexOf(":") >= 0
                                Dim kind = If(isFile, "个文件", "个工作表")
                                Dim shown = created.Take(8).Select(Function(s)
                                    If isFile Then
                                        Try : Return IO.Path.GetFileName(s) : Catch : Return s : End Try
                                    Else
                                        Return s
                                    End If
                                End Function).ToList()
                                line &= "，生成 " & created.Count.ToString() & kind & "：" & String.Join("、", shown)
                                If created.Count > shown.Count Then line &= " 等"
                            End If
                        End If
                        parts.Add(line)
                    End If

                    If j("steps") IsNot Nothing AndAlso j("steps").Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                        Dim steps = j("steps").Select(Function(t) t.ToString()).ToList()
                        If steps.Count > 0 Then
                            parts.Add("执行步骤：")
                            For i = 0 To steps.Count - 1
                                parts.Add((i + 1).ToString() & ". " & steps(i))
                            Next
                        End If
                    End If
                    Return String.Join(vbCrLf, parts)
                Catch ex As Exception  ' 
                End Try
            Next
            Return "✓ " & defaultMsg
        End Function

        ''' <summary>
        ''' 从合并 MCP 结果 JSON 中提取结构化 steps 数组，
        ''' 供前端 extractStepsPlan 渲染"步骤规划"清单（与 ComposeMergeMessage 的文本步骤同源）。
        ''' </summary>
        Private Shared Function ExtractStepsFromMcpResult(mcpResult As ToolsCallResult) As List(Of String)
            Dim steps As New List(Of String)()
            If mcpResult Is Nothing OrElse mcpResult.Content Is Nothing Then Return steps
            For Each c In mcpResult.Content
                Try
                    Dim txtC = TryCast(c, McpTextContent)
                    If txtC Is Nothing OrElse String.IsNullOrWhiteSpace(txtC.Text) Then Continue For
                    Dim txt = txtC.Text.Trim()
                    If Not txt.StartsWith("{") Then Continue For
                    Dim j = Newtonsoft.Json.Linq.JObject.Parse(txt)
                    If j("steps") IsNot Nothing AndAlso j("steps").Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                        For Each t In j("steps")
                            If t IsNot Nothing Then steps.Add(t.ToString())
                        Next
                    End If
                Catch ex As Exception  ' 
                End Try
            Next
            Return steps
        End Function

        ''' <summary>
        ''' 判断用户是否明确要求"合并所有Sheet"：
        ''' 优先取 LLM 分析输出的 mergeAll 字段，其次从用户原始文本关键词兜底。
        ''' 命中时后端向 excel_merge_sheets 传 mergeAll=true，强制合并全部可见表。
        ''' </summary>
        Private Shared Function WantsMergeAllSheets(analysisJson As Newtonsoft.Json.Linq.JObject, inputs As Dictionary(Of String, Object)) As Boolean
            If analysisJson IsNot Nothing Then
                Dim m = analysisJson("mergeAll")
                If m IsNot Nothing Then
                    Try
                        If m.ToObject(Of Boolean)() Then Return True
                    Catch ex As Exception  ' 
                    End Try
                End If
            End If
            Dim allText = ""
            If inputs IsNot Nothing Then
                For Each kv In inputs
                    allText &= " " & If(kv.Value, "").ToString()
                Next
            End If
            Return allText.Contains("合并所有") OrElse allText.Contains("所有sheet") OrElse allText.Contains("所有表") OrElse
                   allText.Contains("全部sheet") OrElse allText.Contains("全部表") OrElse allText.Contains("全部工作表") OrElse
                   allText.Contains("所有工作表") OrElse allText.Contains("全部合并") OrElse
                   allText.IndexOf("merge all", StringComparison.OrdinalIgnoreCase) >= 0
        End Function

        ''' <summary>
        ''' 判断合并是否需要串联去重：
        ''' 优先取 LLM 分析输出的 deduplicate 字段，其次从用户原始文本关键词兜底。
        ''' </summary>
        Private Shared Function ShouldDeduplicateMerge(analysisJson As Newtonsoft.Json.Linq.JObject, inputs As Dictionary(Of String, Object)) As Boolean
            If analysisJson IsNot Nothing Then
                Dim d = analysisJson("deduplicate")
                If d IsNot Nothing Then
                    Try
                        If d.ToObject(Of Boolean)() Then Return True
                    Catch ex As Exception  ' 
                    End Try
                End If
            End If
            Dim allText = ""
            If inputs IsNot Nothing Then
                For Each kv In inputs
                    allText &= " " & If(kv.Value, "").ToString()
                Next
            End If
            Return allText.Contains("去重") OrElse allText.Contains("去重复") OrElse
                   allText.IndexOf("dedup", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   allText.IndexOf("deduplicate", StringComparison.OrdinalIgnoreCase) >= 0
        End Function

        ''' <summary>
        ''' 转义字符串以安全嵌入 VBA 代码（处理引号和换行）
        ''' </summary>
        Private Function EscapeVBAString(value As String) As String
            If String.IsNullOrEmpty(value) Then Return """""" ' 空字符串
            ' 双引号转义为双引号（VBA 规范）
            Return """" & value.Replace("""", """""") & """"
        End Function

        ''' <summary>
        ''' 【扩展1】RAG 知识库上下文增强
        ''' 在 LLM 调用前，从知识库检索相关文档上下文并追加到 System Prompt
        ''' </summary>
        Private Async Function EnrichWithRagContextAsync(skill As SkillDefinition, userQuestion As String, originalPrompt As String) As Task(Of String)
            Try
                ' 检查 RAG 是否启用
                If Not skill.RagEnabled Then Return originalPrompt

                ' 初始化 RagOrchestrator（延迟初始化）
                Dim ragOrchestrator = Services.RAG.RagOrchestrator.Instance

                ' 调用 RAG 检索
                Dim ragResult = Await ragOrchestrator.RetrieveAndEnhanceAsync(
                    userQuestion,
                    If(String.IsNullOrWhiteSpace(skill.RagKbId), "", skill.RagKbId))

                If ragResult Is Nothing OrElse Not ragResult.HasContext Then
                    ' 无检索结果
                    If skill.RagFallbackToNormal Then
                        System.Diagnostics.Debug.WriteLine("[RAG] 无知识库结果，降级到普通模式")
                        Return originalPrompt
                    Else
                        Return originalPrompt & vbCrLf & vbCrLf &
                            "[知识库提示] 未找到相关知识，请基于通用知识回答。"
                    End If
                End If

                ' 构建增强后的 Prompt
                Dim enhancedPrompt As New System.Text.StringBuilder()
                enhancedPrompt.AppendLine(originalPrompt)
                enhancedPrompt.AppendLine()
                enhancedPrompt.AppendLine("========== 参考知识（来自企业知识库）=========")
                enhancedPrompt.AppendLine(ragResult.ContextText)

                If ragResult.SourceReferences IsNot Nothing AndAlso ragResult.SourceReferences.Count > 0 Then
                    enhancedPrompt.AppendLine()
                    enhancedPrompt.Append("【来源】: ")
                    Dim sources As New List(Of String)()
                    For Each srcRef In ragResult.SourceReferences.Take(5)
                        sources.Add(srcRef.DocTitle)
                    Next
                    enhancedPrompt.AppendLine(String.Join(", ", sources))
                End If

                enhancedPrompt.AppendLine("==================================================")
                System.Diagnostics.Debug.WriteLine(String.Format(
                    "[RAG] 增强 Prompt 成功，检索到 {0} 个文档片段",
                    ragResult.RetrievedChunks))

                Return enhancedPrompt.ToString()

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[RAG] 上下文增强失败: " & ex.Message)
                ' RAG 失败不阻塞主流程，返回原始 Prompt
                Return originalPrompt
            End Try
        End Function

        ''' <summary>
        ''' 【扩展2】管道编排执行器
        ''' 遍历 PipelineSteps，依次执行每个子 Skill，将上一步输出作为下一步输入
        ''' </summary>
        Private Async Function ExecutePipelineSkill(skill As SkillDefinition, inputs As Dictionary(Of String, Object)) As Task(Of WorkflowResult)
            Dim result As New WorkflowResult()
            result.Outputs = New Dictionary(Of String, Object)()
            result.ExecutionLog = New List(Of String)()

            Try
                result.ExecutionLog.Add("[Pipeline] 开始执行管道 Skill: " & skill.DisplayName)

                ' 当前管道数据流（在步骤间传递）
                Dim pipeData As Dictionary(Of String, Object) = New Dictionary(Of String, Object)()

                ' 初始化：将用户输入注入管道
                If inputs IsNot Nothing Then
                    For Each kvp In inputs
                        pipeData("input." & kvp.Key) = kvp.Value
                    Next
                End If

                ' 遍历每个管道步骤
                For Each pipeStep In skill.PipelineSteps
                    If pipeStep Is Nothing Then Continue For

                    result.ExecutionLog.Add(String.Format("[Pipeline] 执行步骤: {0} → Skill:{1}",
                        pipeStep.StepName, pipeStep.SkillName))

                    ' 构造当前步骤的输入参数
                    Dim stepInputs As New Dictionary(Of String, Object)()

                    ' 应用 InputMapping（从上游输出映射到本步输入）
                    If pipeStep.InputMapping IsNot Nothing Then
                        For Each mapping In pipeStep.InputMapping
                            Dim sourceKey = mapping.Value
                            If pipeData.ContainsKey(sourceKey) Then
                                stepInputs(mapping.Key) = pipeData(sourceKey)
                            ElseIf inputs IsNot Nothing AndAlso inputs.ContainsKey(mapping.Key) Then
                                stepInputs(mapping.Key) = inputs(mapping.Key)
                            Else
                                stepInputs(mapping.Key) = Nothing
                            End If
                        Next
                    End If

                    ' 条件检查（简化版）
                    If Not String.IsNullOrWhiteSpace(pipeStep.Condition) Then
                        Dim conditionMet = EvaluateSimpleCondition(pipeStep.Condition, pipeData)
                        If Not conditionMet Then
                            result.ExecutionLog.Add(String.Format(
                                "[Pipeline] 步骤 {0} 条件不满足，跳过", pipeStep.StepName))
                            Continue For
                        End If
                    End If

                    ' 获取目标子 Skill
                    Dim subSkill = SkillManager.Instance.GetSkill(pipeStep.SkillName)
                    If subSkill Is Nothing Then
                        result.ExecutionLog.Add(String.Format(
                            "[Pipeline] 错误: 子 Skill '{0}' 不存在", pipeStep.SkillName))
                        If String.IsNullOrEmpty(pipeStep.NextStepOnFailure) Then
                            result.Success = False
                            result.ErrorMessage = "管道步骤找不到子 Skill: " & pipeStep.SkillName
                            Return result
                        Else
                            Continue For
                        End If
                    End If

                    ' 执行子 Skill
                    Dim stepResult = Await ExecuteCustomSkill(subSkill, stepInputs)

                    If stepResult.Success Then
                        ' 将输出存入管道数据
                        If stepResult.Outputs IsNot Nothing Then
                            For Each kvp In stepResult.Outputs
                                pipeData(String.Format("step_{0}.{1}", pipeStep.StepId, kvp.Key)) = kvp.Value
                                ' 同时也用扁平 key 存储，方便后续映射
                                pipeData(kvp.Key) = kvp.Value
                            Next
                        End If
                        result.ExecutionLog.Add(String.Format(
                            "[Pipeline] 步骤 {0} 执行成功", pipeStep.StepName))
                    Else
                        result.ExecutionLog.Add(String.Format(
                            "[Pipeline] 步骤 {0} 执行失败: {1}", pipeStep.StepName, stepResult.ErrorMessage))

                        If String.IsNullOrEmpty(pipeStep.NextStepOnFailure) Then
                            result.Success = False
                            result.ErrorMessage = "管道在步骤 '" & pipeStep.StepName & "' 失败: " & stepResult.ErrorMessage
                            Return result
                        End If
                        ' 有失败回退步骤则继续
                    End If
                Next

                ' 收集最终结果
                result.Outputs = pipeData
                result.Success = True
                result.Outputs("message") = String.Format("管道 '{0}' 执行完成", skill.DisplayName)
                result.ExecutionLog.Add("[Pipeline] 管道执行完成")

            Catch ex As Exception
                result.Success = False
                result.ErrorMessage = "管道执行失败: " & ex.Message
                result.ExecutionLog.Add("错误: " & ex.Message)
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 【扩展2辅助】简单条件评估（支持基本比较操作）
        ''' </summary>
        Private Function EvaluateSimpleCondition(condition As String, data As Dictionary(Of String, Object)) As Boolean
            Try
                ' 简化版：检查变量是否存在且非空
                ' 格式: "variableName" 或 "variableName != null"
                Dim cond = condition.Trim()

                ' 包含 == 或 !=
                If cond.Contains("=") OrElse cond.Contains("<>") OrElse cond.Contains("!=") Then
                    Dim parts() = {"=", "<>", "!=", ">", "<", ">=", "<="}
                    For Each op In parts
                        If cond.Contains(op) Then
                            Dim pair = cond.Split({op}, 2, StringSplitOptions.None)
                            If pair.Length = 2 Then
                                Dim varName = pair(0).Trim()
                                Dim expectedValue = pair(1).Trim().Trim("'"c, """"c)
                                Dim actualValue As Object = Nothing
                                If data.ContainsKey(varName) Then actualValue = data(varName)
                                Dim actualStr = If(actualValue?.ToString(), "")
                                Select Case op
                                    Case "=", "=="
                                        Return actualStr = expectedValue
                                    Case "<>", "!="
                                        Return actualStr <> expectedValue
                                    Case ">"
                                        Return IsNumeric(actualStr) AndAlso IsNumeric(expectedValue) _
                                            AndAlso CDbl(actualStr) > CDbl(expectedValue)
                                    Case "<"
                                        Return IsNumeric(actualStr) AndAlso IsNumeric(expectedValue) _
                                            AndAlso CDbl(actualStr) < CDbl(expectedValue)
                                End Select
                            End If
                        End If
                    Next
                End If

                ' 默认：检查变量是否存在且非空
                Return data.ContainsKey(cond) AndAlso data(cond) IsNot Nothing _
                       AndAlso Not String.IsNullOrWhiteSpace(data(cond).ToString())
            Catch ex As Exception  ' 
                Return True ' 条件解析失败时默认通过
            End Try
        End Function

        ''' <summary>
        ''' 【扩展6】Python 脚本执行器（架构预留）
        ''' 通过 PythonExecutionService 执行 Python 脚本，支持本地/远程两种模式
        ''' </summary>
        Private Async Function ExecutePythonSkill(skill As SkillDefinition, inputs As Dictionary(Of String, Object)) As Task(Of WorkflowResult)
            Dim result As New WorkflowResult()
            result.Outputs = New Dictionary(Of String, Object)()
            result.ExecutionLog = New List(Of String)()

            Try
                result.ExecutionLog.Add("[Python] 开始执行 Python 脚本...")
                ' 诊断：把实际执行时的 Skill 名与 inputs 键名追加到 addin_path.log，
                ' 比 sel_debug.log 更可靠，可确认 dll 版本与参数传递。
                Try
                    Dim keysStr As String = If(inputs Is Nothing, "NOTHING", String.Join(",", inputs.Keys))
                    System.IO.File.AppendAllText(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "addin_path.log"), _
                        DateTime.Now.ToString("HH:mm:ss") & " | Skill=" & skill.Name & " | inputs.Keys=" & keysStr & vbCrLf)
                Catch
                End Try

                ' 【修复】所有 PythonCode Skill（不仅是 GeneratePython）都需要 Excel 数据。
                ' 若 inputs 里没数据，主动读取当前选区 / 活动表注入 excelData/selectionRange，
                ' 这样 SheetErrorChecker 等 TemplateCode Skill 才能拿到真实 rows。
                If inputs Is Nothing Then inputs = New Dictionary(Of String, Object)()
                EnsureExcelDataForPython(inputs)

                Dim scriptCode As String = ""

                ' 【GeneratePython 特殊处理】动态生成 Python 代码
                If String.Equals(skill.Name, "GeneratePython", StringComparison.OrdinalIgnoreCase) Then
                    result.ExecutionLog.Add("[Python] GeneratePython 模式：先调用 LLM 生成代码...")

                    ' 确保 inputs 中有 Excel 数据，没有则自动读取当前选区
                    If inputs Is Nothing Then inputs = New Dictionary(Of String, Object)()
                    EnsureExcelDataForPython(inputs)

                    scriptCode = Await GeneratePythonCodeAsync(skill, inputs)
                    If String.IsNullOrWhiteSpace(scriptCode) Then
                        result.Success = False
                        result.ErrorMessage = "LLM 生成 Python 代码失败"
                        Return result
                    End If
                    result.Outputs("generatedCode") = scriptCode
                    result.ExecutionLog.Add(String.Format("[Python] 代码生成完成，长度: {0}", scriptCode.Length))
                ElseIf skill.TemplateCode IsNot Nothing AndAlso skill.TemplateCode.Length > 0 Then
                    scriptCode = skill.TemplateCode
                Else
                    result.Success = False
                    result.ErrorMessage = "Python Skill 缺少脚本代码 (templateCode)"
                    Return result
                End If

                ' 构造输入数据（将 Excel 数据传递给 Python）
                Dim pythonInputs As New Dictionary(Of String, Object)()
                If inputs IsNot Nothing Then
                    For Each kvp In inputs
                        pythonInputs(kvp.Key) = kvp.Value
                    Next
                End If

                ' 提取引用区域和当前选区数据（优先使用传入的结构化数据，否则从 Excel 实时读取）
                Dim referenceData = ExtractReferenceData(inputs)
                Dim selectionData = ExtractSelectionData(inputs)

                ' 调用 Python 执行服务
                Dim pyService = Services.PythonExecutionService.Instance
                Dim pyResult = Await pyService.ExecuteScriptAsync(scriptCode, pythonInputs, True, Nothing, referenceData, selectionData)

                If Not pyResult.Success Then
                    result.Success = False
                    result.ErrorMessage = "Python 执行失败: " & pyResult.ErrorMessage
                    If Not String.IsNullOrWhiteSpace(pyResult.StandardError) Then
                        result.ErrorMessage &= vbCrLf & "stderr: " & pyResult.StandardError
                    End If
                    Return result
                End If

                ' 收集输出结果
                result.Outputs("message") = "Python 脚本执行成功"
                result.Outputs("exitCode") = pyResult.ExitCode
                result.Outputs("stdout") = pyResult.StandardOutput

                If pyResult.OutputData IsNot Nothing Then
                    result.Outputs("outputData") = pyResult.OutputData.ToString()
                    ' 将 JSON 字段展开到扁平 outputs 中
                    For Each prop In pyResult.OutputData.Properties()
                        result.Outputs("python_" & prop.Name) = prop.Value?.ToObject(Of Object)()
                    Next
                End If

                result.ExecutionLog.Add(String.Format("[Python] 执行完成，退出码: {0}", pyResult.ExitCode))
                result.Success = True

            Catch ex As Exception
                result.Success = False
                Dim fullMessage = BuildExceptionMessage(ex)
                result.ErrorMessage = "Python Skill 执行异常: " & fullMessage
                result.ExecutionLog.Add("[Python] 异常: " & fullMessage)
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 为 GeneratePython Skill 确保输入中包含 Excel 选区数据和引用区域数据
        ''' </summary>
        Private Sub EnsureExcelDataForPython(inputs As Dictionary(Of String, Object))
            If inputs Is Nothing Then inputs = New Dictionary(Of String, Object)()

            ' 读取当前选区
            If Not inputs.ContainsKey("excelData") Then
                Try
                    Dim reader As IExcelDataReader = BridgeRegistry.Instance.Resolve(Of IExcelDataReader)()
                    ' 【修复】ExcelDataBridge 是懒加载单例，若此前未被访问则 BridgeRegistry 中无实例，
                    ' 强制访问 Instance 触发构造函数里的 Register(Me)。
                    If reader Is Nothing Then
                        Dim unused = ExcelDataBridge.Instance
                        reader = BridgeRegistry.Instance.Resolve(Of IExcelDataReader)()
                    End If
                    Dim selectedDataJson As String = Nothing
                    If reader IsNot Nothing Then selectedDataJson = reader.GetSelectedRangeData()
                    Dim selectedDataObj = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(selectedDataJson)
                    ' 【修复】选区无效/为空（用户未框选或选区不含目标数据）时，回退读取整张活动工作表，
                    ' 保证"检查当前工作表"类 Skill 能覆盖所有行（含 #DIV/0! 等错误值）。
                    If selectedDataObj Is Nothing OrElse Not selectedDataObj.ContainsKey("success") OrElse Not CBool(selectedDataObj("success")) Then
                        If reader IsNot Nothing Then selectedDataJson = reader.GetUsedRangeData()
                        selectedDataObj = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(selectedDataJson)
                    End If
                    If selectedDataObj IsNot Nothing AndAlso selectedDataObj.ContainsKey("success") AndAlso CBool(selectedDataObj("success")) Then
                        inputs("excelData") = selectedDataJson
                        If selectedDataObj.ContainsKey("range") Then inputs("selectionRange") = selectedDataObj("range")
                        If selectedDataObj.ContainsKey("worksheetName") Then inputs("worksheetName") = selectedDataObj("worksheetName")
                    End If
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("[GeneratePython] 读取选区数据失败: " & ex.Message)
                End Try
            End If

            ' 读取引用区域（如果 ContextStore 已支持）
            If Not inputs.ContainsKey("referenceData") AndAlso Not inputs.ContainsKey("referenceRegions") Then
                Try
                    Dim store = ContextStore.GetInstance()
                    Dim refMethod = store.GetType().GetMethod("GetReferenceRegions")
                    If refMethod IsNot Nothing Then
                        Dim refs = refMethod.Invoke(store, Nothing)
                        If refs IsNot Nothing Then
                            Dim countProp = refs.GetType().GetProperty("Count")
                            If countProp IsNot Nothing AndAlso CInt(countProp.GetValue(refs)) > 0 Then
                                inputs("referenceData") = refs
                            End If
                        End If
                    End If
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("[GeneratePython] 读取引用区域失败: " & ex.Message)
                End Try
            End If
        End Sub

        ''' <summary>
        ''' 从 inputs 提取引用区域数据（支持 referenceData / referenceRegions / reference_regions 键）
        ''' </summary>
        Private Function ExtractReferenceData(inputs As Dictionary(Of String, Object)) As Object
            If inputs Is Nothing Then Return Nothing
            If inputs.ContainsKey("referenceData") Then Return inputs("referenceData")
            If inputs.ContainsKey("referenceRegions") Then Return inputs("referenceRegions")
            If inputs.ContainsKey("reference_regions") Then Return inputs("reference_regions")
            Return Nothing
        End Function

        ''' <summary>
        ''' 从 inputs 提取当前选区数据（支持 selectionData / selection / excelData 键）
        ''' </summary>
        Private Function ExtractSelectionData(inputs As Dictionary(Of String, Object)) As Object
            If inputs Is Nothing Then
                Try
                    System.IO.File.AppendAllText(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "sel_debug.log"), _
                        "[" & DateTime.Now.ToString("HH:mm:ss") & "] inputs=NOTHING" & vbCrLf)
                Catch
                End Try
                Return Nothing
            End If
            Try
                Dim keysStr As String = String.Join(",", inputs.Keys)
                System.IO.File.AppendAllText(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "sel_debug.log"), _
                    "[" & DateTime.Now.ToString("HH:mm:ss") & "] inputs.Keys=" & keysStr & vbCrLf)
            Catch
            End Try

            ' 【关键优先级】若 LLM 已传入区域地址，必须最优先按指定区域读取完整数据。
            ' 兼容多种键名：range / rangeAddress / selectionRange / address（LLM 生成参数名不固定）。
            ' 否则：前端在 selection.required=true 时会把"仅含表头、不含 rows"的轻量选区对象
            ' 以 selection/excelData 键注入，若被下方直接 Return，会导致 Python 端 rows 为空。
            Dim rangeAddressVal As String = Nothing
            For Each rk In New String() {"range", "rangeAddress", "selectionRange", "address"}
                If inputs.ContainsKey(rk) AndAlso inputs(rk) IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(inputs(rk).ToString()) Then
                    ' 【修复】Newtonsoft 把 JSON 字符串反序列化为 JValue，其 ToString() 会带双引号，
                    ' 必须 Trim 掉，否则 GetRangeData("$A$2:$F$9") 会因地址含引号而失败返回 null。
                    rangeAddressVal = inputs(rk).ToString().Trim(ChrW(34))
                    Exit For
                End If
            Next

            If rangeAddressVal IsNot Nothing Then
                Try
                    Dim reader As IExcelDataReader = BridgeRegistry.Instance.Resolve(Of IExcelDataReader)()
                    If reader IsNot Nothing Then
                        Dim rangeAddress As String = rangeAddressVal
                        If inputs.ContainsKey("sheet") Then
                            Dim sheetName As String = inputs("sheet").ToString().Trim(ChrW(34))
                            If Not String.IsNullOrWhiteSpace(sheetName) AndAlso Not rangeAddress.Contains("!"c) Then
                                rangeAddress = sheetName & "!" & rangeAddress
                            End If
                        End If
                        Dim rangeDataJson As String = reader.GetRangeData(rangeAddress)
                        Try
                            System.IO.File.AppendAllText(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "sel_debug.log"), _
                                "[" & DateTime.Now.ToString("HH:mm:ss") & "] range-branch addr=" & rangeAddress & vbCrLf & "json=" & If(rangeDataJson, "<null>") & vbCrLf)
                        Catch
                        End Try
                        If Not String.IsNullOrWhiteSpace(rangeDataJson) Then
                            Return JsonConvert.DeserializeObject(Of Object)(rangeDataJson)
                        End If
                    End If
                Catch ex As Exception  ' 
                    ' 失败则降级到下方选区/兜底读取
                End Try
            End If

            ' 其次：使用传入的结构化选区数据（前端可能注入，但可能不含 rows，故已让位给 range 分支）
            If inputs.ContainsKey("selectionData") Then Return inputs("selectionData")
            If inputs.ContainsKey("selection") Then Return inputs("selection")
            If inputs.ContainsKey("excelData") Then
                Try
                    Return JsonConvert.DeserializeObject(Of Object)(inputs("excelData").ToString())
                Catch ex As Exception  ' 
                    Return Nothing
                End Try
            End If

            ' 兜底：从 Excel 读取当前选区（经接口解析）
            Try
                Dim reader As IExcelDataReader = BridgeRegistry.Instance.Resolve(Of IExcelDataReader)()
                If reader Is Nothing Then Return Nothing
                Dim selJson As String = reader.GetSelectedRangeData()
                ' 【稳健兜底】若当前选区为空（未选中区域 / 选区在错误表），改读整张活动工作表，
                ' 彻底绕开"前端选区是否选中"的依赖，保证 SheetErrorChecker 一定能拿到数据。
                If String.IsNullOrWhiteSpace(selJson) Then
                    Try
                        selJson = reader.GetUsedRangeData()
                    Catch
                    End Try
                End If
                Try
                    System.IO.File.AppendAllText(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "sel_debug.log"), _
                        "[" & DateTime.Now.ToString("HH:mm:ss") & "] fallback-branch json=" & If(selJson, "<null>") & vbCrLf)
                Catch
                End Try
                Return JsonConvert.DeserializeObject(Of Object)(selJson)
            Catch ex As Exception  ' 
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 为 GeneratePython Skill 动态生成 Python 代码
        ''' </summary>
        Private Async Function GeneratePythonCodeAsync(skill As SkillDefinition, inputs As Dictionary(Of String, Object)) As Task(Of String)
            Try
                Dim userPrompt As String = ""
                If inputs IsNot Nothing Then
                    If inputs.ContainsKey("text") Then userPrompt = inputs("text")?.ToString()
                    If String.IsNullOrWhiteSpace(userPrompt) AndAlso inputs.ContainsKey("prompt") Then userPrompt = inputs("prompt")?.ToString()
                    If String.IsNullOrWhiteSpace(userPrompt) AndAlso inputs.ContainsKey("message") Then userPrompt = inputs("message")?.ToString()
                End If

                Dim contents = New List(Of MultimodalContent)()
                If Not String.IsNullOrWhiteSpace(userPrompt) Then
                    contents.Add(MultimodalContent.CreateText(userPrompt))
                End If

                ' 统一提取引用区域和当前选区数据，确保 LLM 能看到真实数据
                Dim referenceData = ExtractReferenceData(inputs)
                Dim selectionData = ExtractSelectionData(inputs)
                If referenceData IsNot Nothing Then
                    contents.Add(MultimodalContent.CreateText(vbCrLf & "[引用区域数据]" & vbCrLf & JsonConvert.SerializeObject(referenceData)))
                End If
                If selectionData IsNot Nothing Then
                    contents.Add(MultimodalContent.CreateText(vbCrLf & "[当前选区数据]" & vbCrLf & JsonConvert.SerializeObject(selectionData)))
                End If
                If contents.Count = 0 Then
                    contents.Add(MultimodalContent.CreateText("请根据当前上下文生成Python脚本。"))
                End If

                Dim systemPrompt = skill.LLMSystemPrompt
                If String.IsNullOrWhiteSpace(systemPrompt) Then
                    systemPrompt = "你是 Python 数据分析专家。根据用户需求生成完整可运行的 Python 脚本。"
                End If

                ' 无论 Skill 提示词如何，必须明确告知 LLM 数据源
                systemPrompt = systemPrompt & vbCrLf & vbCrLf &
                    "【强制规则 - 数据源】" & vbCrLf &
                    "1. 脚本运行环境中已经预加载了以下变量，可直接使用：" & vbCrLf &
                    "   - reference_data: 引用区域列表，每个元素包含 workbook/sheet/address/headers/rows" & vbCrLf &
                    "   - selection_data: 当前选区对象，包含 workbook/sheet/address/headers/rows" & vbCrLf &
                    "2. 严禁使用 pd.read_excel、pd.read_csv、open()、pathlib 等任何方式读取本地 Excel 文件。" & vbCrLf &
                    "3. 所有数据必须直接从 reference_data 或 selection_data 变量中获取。" & vbCrLf &
                    "4. 如果用户请求涉及多个表格对比/差异/匹配/关联，使用 reference_data[0]、reference_data[1] 等，禁止假设 Sheet1/Sheet2。" & vbCrLf &
                    "5. 表头/列名可能包含空格、括号、中文等字符，使用列名时必须用字符串形式访问，例如 df['A (ID)']、row['B (姓名)']；严禁把列名作为裸标识符使用，例如 df.A (ID)、{A (ID): value}。" & vbCrLf &
                    "6. 构造字典时，所有键必须加引号，例如 {key: value} 必须写成 {'key': value}。" & vbCrLf &
                    "7. 违反以上任何一条，脚本将执行失败。"

                ' 追加列名使用示例，强化 LLM 对特殊列名的处理
                contents.Add(MultimodalContent.CreateText(vbCrLf & "【列名使用示例 - 强制遵循】" & vbCrLf &
                    "headers 中的列名可能包含空格和括号，如 'A (ID)'、'B (姓名)'。" & vbCrLf &
                    "正确用法：" & vbCrLf &
                    "  df['A (ID)']" & vbCrLf &
                    "  df.loc[0, 'B (姓名)']" & vbCrLf &
                    "  row = {'A (ID)': 'xxx', 'B (姓名)': 'yyy'}" & vbCrLf &
                    "错误用法（会导致 SyntaxError）：" & vbCrLf &
                    "  df.A (ID)" & vbCrLf &
                    "  {A (ID): 'xxx'}"))

                Dim llmResult = Await _llmEngine.ProcessContentAsync(contents, systemPrompt)
                If Not llmResult.Success Then
                    Throw New InvalidOperationException("LLM 生成 Python 代码失败: " & If(llmResult.ErrorMessage, "未知错误"))
                End If

                Dim code = llmResult.Content

                ' 从 ```python 代码块中提取
                Dim pythonBlockMatch = System.Text.RegularExpressions.Regex.Match(code, "```python\s*([\s\S]*?)\s*```", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                If pythonBlockMatch.Success Then
                    code = pythonBlockMatch.Groups(1).Value.Trim()
                Else
                    ' 尝试从通用代码块中提取
                    Dim genericBlockMatch = System.Text.RegularExpressions.Regex.Match(code, "```\s*([\s\S]*?)\s*```")
                    If genericBlockMatch.Success Then
                        code = genericBlockMatch.Groups(1).Value.Trim()
                    End If
                End If

                Return code
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[GeneratePython] 生成代码异常: " & ex.Message)
                Throw
            End Try
        End Function

        Private Async Function ExecuteWorkflowSkill(skill As SkillDefinition, inputs As Dictionary(Of String, Object)) As Task(Of WorkflowResult)
            Dim result As New WorkflowResult()
            result.Outputs = New Dictionary(Of String, Object)()
            result.ExecutionLog = New List(Of String)()

            Try
                Dim variables As New Dictionary(Of String, Object)()

                If inputs IsNot Nothing Then
                    For Each kvp In inputs
                        variables("input." & kvp.Key) = kvp.Value
                    Next
                End If

                If skill.CustomWorkflowSteps IsNot Nothing Then
                    For Each wfStep In skill.CustomWorkflowSteps
                        If wfStep Is Nothing Then Continue For

                        Dim stepResult = Await ExecuteWorkflowStep(wfStep, variables, skill)

                        If Not stepResult.Success Then
                            result.Success = False
                            Dim stepName As String = "未知步骤"
                            If wfStep.StepName IsNot Nothing AndAlso wfStep.StepName.Length > 0 Then
                                stepName = wfStep.StepName
                            End If
                            result.ErrorMessage = "步骤 '" & stepName & "' 执行失败: " & stepResult.ErrorMessage
                            Return result
                        End If

                        If stepResult.Outputs IsNot Nothing Then
                            For Each kvp In stepResult.Outputs
                                variables(kvp.Key) = kvp.Value
                                If result.Outputs Is Nothing Then
                                    result.Outputs = New Dictionary(Of String, Object)()
                                End If
                                result.Outputs(kvp.Key) = kvp.Value
                            Next
                        End If

                        If result.ExecutionLog IsNot Nothing Then
                            Dim stepNameLog As String = "未知步骤"
                            If wfStep.StepName IsNot Nothing AndAlso wfStep.StepName.Length > 0 Then
                                stepNameLog = wfStep.StepName
                            End If
                            result.ExecutionLog.Add("步骤 '" & stepNameLog & "' 执行完成")
                        End If
                    Next
                End If

                result.Success = True

            Catch ex As Exception
                result.Success = False
                result.ErrorMessage = ex.Message
            End Try

            Return result
        End Function

        Private Async Function ExecuteWorkflowStep(wfStep As WorkflowStep, variables As Dictionary(Of String, Object), skill As SkillDefinition) As Task(Of StepExecutionResult)
            Dim result As New StepExecutionResult()
            result.Outputs = New Dictionary(Of String, Object)()

            Try
                If wfStep Is Nothing Then
                    result.Success = False
                    result.ErrorMessage = "步骤为空"
                    Return result
                End If

                Select Case wfStep.StepType
                    Case WorkflowStepType.OCRRecognition
                        Dim imagePath As String = ""
                        If variables IsNot Nothing AndAlso variables.ContainsKey("input.imagePath") Then
                            imagePath = variables("input.imagePath")?.ToString()
                        End If
                        If imagePath Is Nothing OrElse imagePath.Length = 0 Then
                            If wfStep.Config IsNot Nothing AndAlso wfStep.Config.ContainsKey("imagePath") Then
                                imagePath = wfStep.Config("imagePath")?.ToString()
                            End If
                        End If

                        If imagePath IsNot Nothing AndAlso imagePath.Length > 0 AndAlso File.Exists(imagePath) Then
                            result.Outputs("ocrText") = Await Services.OcrServiceManager.Instance.RecognizeText(imagePath)
                            result.Success = True
                        Else
                            result.ErrorMessage = "未找到图片文件"
                            result.Success = False
                        End If

                    Case WorkflowStepType.WriteToExcel
                        If skill IsNot Nothing AndAlso skill.FieldMappings IsNot Nothing AndAlso skill.FieldMappings.Count > 0 Then
                            Dim workflowResult As New WorkflowResult()
                            workflowResult.Outputs = New Dictionary(Of String, Object)()
                            WriteMappedFields(workflowResult, skill)
                        End If
                        result.Success = True

                    Case WorkflowStepType.CustomLLM
                        Dim prompt As String = ""
                        If wfStep.Config IsNot Nothing AndAlso wfStep.Config.ContainsKey("prompt") Then
                            prompt = wfStep.Config("prompt")?.ToString()
                        End If
                        If prompt IsNot Nothing AndAlso prompt.Length > 0 Then
                            Dim contents = New List(Of MultimodalContent)()
                            contents.Add(MultimodalContent.CreateText(prompt))
                            Dim llmResult = Await _llmEngine.ProcessContentAsync(contents)
                            If llmResult.Success Then
                                result.Outputs("llmResponse") = llmResult.Content
                                result.Success = True
                            Else
                                result.ErrorMessage = llmResult.ErrorMessage
                                result.Success = False
                            End If
                        End If

                    Case Else
                        result.Success = True
                End Select

            Catch ex As Exception
                result.Success = False
                result.ErrorMessage = ex.Message
            End Try

            Return result
        End Function

        Private Function ExecuteSimpleSkill(skill As SkillDefinition, inputs As Dictionary(Of String, Object)) As WorkflowResult
            Dim result As New WorkflowResult()
            result.Outputs = New Dictionary(Of String, Object)()

            Try
                ' 获取 VBA 代码
                Dim vbaCode As String = ""
                If skill.TemplateCode IsNot Nothing AndAlso skill.TemplateCode.Length > 0 Then
                    vbaCode = skill.TemplateCode
                ElseIf inputs IsNot Nothing AndAlso inputs.ContainsKey("code") Then
                    vbaCode = inputs("code").ToString()
                End If

                If String.IsNullOrWhiteSpace(vbaCode) Then
                    result.Success = False
                    result.ErrorMessage = "没有找到要执行的 VBA 代码"
                    Return result
                End If

                ' 使用 VBA 沙箱执行代码（全静态方法）
                ' 【修复】自定义 Skill 由用户主动触发，跳过交互式安全确认弹窗
                ' 安全检查仍会执行（Critical/Errors 仍会被阻止），仅跳过 MessageBox 确认步骤
                Dim executionResult = Security.VBACodeSandbox.ExecuteCode(vbaCode, requireConfirmation:=False)

                result.Success = True
                result.Outputs("message") = "Skill '" & skill.DisplayName & "' 执行成功"
                result.Outputs("executionResult") = executionResult

            Catch ex As Exception
                result.Success = False
                result.ErrorMessage = "执行 Skill 失败: " & ex.Message
                System.Diagnostics.Debug.WriteLine("ExecuteSimpleSkill 错误: " & ex.ToString())
            End Try

            Return result
        End Function

        Private Sub WriteMappedFields(result As WorkflowResult, skill As SkillDefinition)
            Try
                Dim excelApp = Globals.ThisAddIn.GetExcelApp()
                If excelApp.ActiveWorkbook Is Nothing Then
                    Return
                End If

                Dim sheet = TryCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                If sheet Is Nothing Then
                    Return
                End If

                Dim startRow = 2
                While startRow < 1000
                    Dim isEmpty = True
                    For Each mapping In skill.FieldMappings
                        If mapping.TargetCell IsNot Nothing AndAlso mapping.TargetCell.Length > 0 Then
                            Dim cell = sheet.Range(mapping.TargetCell & startRow)
                            If cell IsNot Nothing Then
                                Dim cellValue = cell.Value?.ToString()
                                If cellValue IsNot Nothing AndAlso cellValue.Length > 0 Then
                                    isEmpty = False
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                    If isEmpty Then
                        Exit While
                    End If
                    startRow += 1
                End While

                For Each mapping In skill.FieldMappings
                    If mapping.TargetCell IsNot Nothing AndAlso mapping.TargetCell.Length > 0 Then
                        If result.Outputs IsNot Nothing AndAlso result.Outputs.ContainsKey(mapping.SourceField) Then
                            Try
                                Dim value = result.Outputs(mapping.SourceField)?.ToString()
                                If value IsNot Nothing AndAlso value.Length > 0 Then
                                    Dim cell = sheet.Range(mapping.TargetCell & startRow)
                                    cell.Value = value

                                    If mapping.Format IsNot Nothing AndAlso mapping.Format.Length > 0 Then
                                        Select Case mapping.Format.ToLower()
                                            Case "currency", "货币"
                                                cell.NumberFormat = "¥#,##0.00"
                                            Case "date", "日期"
                                                cell.NumberFormat = "yyyy-mm-dd"
                                            Case "number", "数字"
                                                cell.NumberFormat = "0.00"
                                        End Select
                                    End If
                                End If
                            Catch ex As Exception
                                System.Diagnostics.Debug.WriteLine("写入单元格失败: " & ex.Message)
                            End Try
                        End If
                    End If
                Next

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("写入Excel失败: " & ex.Message)
            End Try
        End Sub

        Private Function BuildPrompt(skill As SkillDefinition, inputs As Dictionary(Of String, Object)) As String
            Dim prompt As String = ""

            If skill IsNot Nothing Then
                If skill.LLMSystemPrompt IsNot Nothing AndAlso skill.LLMSystemPrompt.Length > 0 Then
                    prompt = skill.LLMSystemPrompt & vbCrLf & vbCrLf
                End If

                If skill.LLMUserPromptTemplate IsNot Nothing AndAlso skill.LLMUserPromptTemplate.Length > 0 Then
                    Dim userPrompt As String = skill.LLMUserPromptTemplate
                    If inputs IsNot Nothing Then
                        For Each kvp In inputs
                            Dim placeholder As String = "{" & kvp.Key & "}"
                            Dim replacement As String = ""
                            If kvp.Value IsNot Nothing Then
                                replacement = kvp.Value.ToString()
                            End If
                            userPrompt = userPrompt.Replace(placeholder, replacement)
                        Next
                    End If
                    prompt &= userPrompt
                Else
                    prompt &= "请处理用户的输入。"
                End If
            End If

            Return prompt
        End Function

        Private Function TryParseJson(content As String) As Newtonsoft.Json.Linq.JObject
            If String.IsNullOrWhiteSpace(content) Then Return Nothing

            Dim trimmed = content.Trim()

            Try
                Return Newtonsoft.Json.Linq.JObject.Parse(trimmed)
            Catch ex As Exception  ' 
            End Try

            Dim extracted = ExtractJsonFromText(trimmed)
            If extracted IsNot Nothing Then
                Try
                    Return Newtonsoft.Json.Linq.JObject.Parse(extracted)
                Catch ex As Exception  ' 
                End Try
            End If

            Return Nothing
        End Function

        Private Function ExtractJsonFromText(text As String) As String
            Dim patterns As String() = {
                "```json\s*([\s\S]*?)\s*```",
                "```\s*([\s\S]*?)\s*```",
                "\{[\s\S]*\}"
            }

            For Each pattern In patterns
                Dim match = System.Text.RegularExpressions.Regex.Match(text, pattern)
                If match.Success Then
                    Dim candidate = match.Value
                    If match.Groups.Count > 1 AndAlso match.Groups(1).Success Then
                        candidate = match.Groups(1).Value
                    End If
                    If IsValidJsonObject(candidate) Then
                        Return candidate.Trim()
                    End If
                End If
            Next

            Return Nothing
        End Function

        Private Function IsValidJsonObject(text As String) As Boolean
            If String.IsNullOrWhiteSpace(text) Then Return False
            text = text.Trim()
            If Not text.StartsWith("{") OrElse Not text.EndsWith("}") Then Return False
            Try
                Newtonsoft.Json.Linq.JObject.Parse(text)
                Return True
            Catch ex As Exception  ' 
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 递归拼接异常及内部异常信息，方便定位根因
        ''' </summary>
        Private Function BuildExceptionMessage(ex As Exception) As String
            If ex Is Nothing Then Return ""
            Dim sb As New System.Text.StringBuilder()
            sb.Append(ex.Message)
            Dim inner = ex.InnerException
            While inner IsNot Nothing
                sb.Append(" -> ")
                sb.Append(inner.Message)
                inner = inner.InnerException
            End While
            Return sb.ToString()
        End Function
    End Class
End Namespace