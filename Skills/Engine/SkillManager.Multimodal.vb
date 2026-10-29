Imports System.IO
Imports System.Reflection
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Serialization
Imports System.Threading
Imports System.Threading.Tasks
Imports ExcelAIVSTOAddIn.Services
Imports ExcelAIVSTOAddIn.Bridges
Imports ExcelAIVSTOAddIn.Utils
Imports YamlDotNet.RepresentationModel

Namespace Skills

    ' SkillManager.Multimodal.vb — 多模态输入构建与结果处理
    ' （类级 XML 文档只在主文件 SkillManager.vb 保留一份，避免 BC42314 分部重复）
    Partial Public Class SkillManager

        ''' <summary>Skill 提示词中声明"可用 MCP 工具"插入点的占位符，运行时替换为实时工具清单。</summary>
        Private Const AVAILABLE_MCP_TOOLS_TOKEN As String = "{{AVAILABLE_MCP_TOOLS}}"

        Private Function ShouldUseMultimodal(inputs As Dictionary(Of String, Object)) As Boolean
            ' 检查输入中是否包含图片或文件数据
            If inputs Is Nothing Then Return False

            If inputs.ContainsKey("imageData") Then Return True
            If inputs.ContainsKey("imagePath") Then Return True
            If inputs.ContainsKey("fileData") Then Return True
            If inputs.ContainsKey("filePath") Then Return True

            Return False
        End Function

        ''' <summary>
        ''' 执行多模态 Skill（统一入口，路由到不同处理方式）
        ''' </summary>
        Private Async Function ExecuteMultimodalSkill(skill As SkillDefinition, inputs As Dictionary(Of String, Object), Optional progress As IProgress(Of String) = Nothing) As Task(Of WorkflowResult)
            Dim result As New WorkflowResult
            Dim stopwatch = System.Diagnostics.Stopwatch.StartNew()

            Try
                System.Diagnostics.Debug.WriteLine("执行多模态 Skill: " & skill.DisplayName)

                Dim contents As New List(Of MultimodalContent)()
                Dim userPrompt = BuildMultimodalContents(skill, inputs, contents)

                ' 【动态工具清单注入】Skill 提示词里可用 `{{AVAILABLE_MCP_TOOLS}}` 占位符声明
                ' "可用 MCP 工具"插入点，运行时替换为系统实时工具清单（经 McpBridge 动态构建，关联
                ' 全量内置工具）。替代在提示词里写死一份残缺工具列表——硬编码会让步骤 tool 名与真实
                ' 工具脱节/过期。不声明占位符的 Skill 走原样提示词，零影响。
                Dim systemPrompt As String = skill.LLMSystemPrompt
                If Not String.IsNullOrEmpty(systemPrompt) AndAlso systemPrompt.Contains(AVAILABLE_MCP_TOOLS_TOKEN) Then
                    Dim capText = McpBridge.Instance.BuildMcpCapabilitiesText()
                    systemPrompt = systemPrompt.Replace(AVAILABLE_MCP_TOOLS_TOKEN, capText)
                    System.Diagnostics.Debug.WriteLine($"[SkillManager] 已注入实时 MCP 工具清单到 Skill '{skill.Name}' 提示词")
                End If

                LLMEngine.Initialize(CreateLLMConfig(systemPrompt))

                Dim llmResult = Await LLMEngine.ProcessContentAsync(contents, userPrompt)

                stopwatch.Stop()
                ProcessMultimodalResult(llmResult, skill, result, stopwatch)

            Catch ex As Exception
                stopwatch.Stop()
                result.Success = False
                result.ErrorMessage = $"多模态处理失败: {ex.Message}"
                System.Diagnostics.Debug.WriteLine("执行多模态 Skill 出错: " & ex.Message)
            End Try

            Return result
        End Function

        Private Function BuildMultimodalContents(skill As SkillDefinition, inputs As Dictionary(Of String, Object),
                                                 contents As List(Of MultimodalContent)) As String
            Dim userPrompt As String = ""

            If inputs Is Nothing Then Return userPrompt

            ProcessImageDataInput(skill, inputs, contents)
            ProcessImagePathInput(skill, inputs, contents)
            ProcessFileInput(skill, inputs, contents)
            userPrompt = ExtractUserPrompt(inputs)

            If contents.Count = 0 AndAlso Not String.IsNullOrEmpty(userPrompt) Then
                contents.Add(MultimodalContent.CreateText(userPrompt))
            End If

            ' 【防御】若 inputs 里存在可提取的文本但未被识别为内容，仍保证存在至少一个文本项，
            ' 避免触发 MultimodalLLMEngine 的"contents 不能为空"误报。
            If contents.Count = 0 Then
                Dim anyText = inputs.Values _
                    .Where(Function(v) v IsNot Nothing) _
                    .Select(Function(v) v.ToString()) _
                    .FirstOrDefault(Function(s) Not String.IsNullOrWhiteSpace(s))
                If Not String.IsNullOrEmpty(anyText) Then
                    contents.Add(MultimodalContent.CreateText(anyText))
                End If
            End If

            Return userPrompt
        End Function

        Private Sub ProcessImageDataInput(skill As SkillDefinition, inputs As Dictionary(Of String, Object), contents As List(Of MultimodalContent))
            If Not inputs.ContainsKey("imageData") OrElse Not skill.SupportsImageInput Then Return

            Dim imageData = inputs("imageData").ToString()
            Dim mimeType As String = "image/png"
            If inputs.ContainsKey("mimeType") Then
                mimeType = inputs("mimeType").ToString()
            End If
            contents.Add(MultimodalContent.CreateImage(imageData, mimeType))
            System.Diagnostics.Debug.WriteLine("添加图片内容，长度: " & imageData.Length.ToString())
        End Sub

        Private Sub ProcessImagePathInput(skill As SkillDefinition, inputs As Dictionary(Of String, Object), contents As List(Of MultimodalContent))
            If Not inputs.ContainsKey("imagePath") OrElse Not skill.SupportsImageInput Then Return

            Dim imagePath = inputs("imagePath").ToString()
            If File.Exists(imagePath) Then
                Dim imageBytes = File.ReadAllBytes(imagePath)
                Dim base64Data = Convert.ToBase64String(imageBytes)
                Dim mimeType = GetMimeTypeFromExtension(Path.GetExtension(imagePath))
                contents.Add(MultimodalContent.CreateImage(base64Data, mimeType))
                System.Diagnostics.Debug.WriteLine("添加图片内容（从文件），路径: " & imagePath)
            End If
        End Sub

        Private Sub ProcessFileInput(skill As SkillDefinition, inputs As Dictionary(Of String, Object), contents As List(Of MultimodalContent))
            If Not inputs.ContainsKey("fileData") OrElse Not skill.SupportsFileInput Then Return

            Dim fileName As String = "file"
            If inputs.ContainsKey("fileName") Then
                fileName = inputs("fileName").ToString()
            End If
            Dim mimeType As String = "text/plain"
            If inputs.ContainsKey("mimeType") Then
                mimeType = inputs("mimeType").ToString()
            End If

            If IsImageMimeType(mimeType) AndAlso inputs.ContainsKey("fileBytes") Then
                Dim fileBytes = CType(inputs("fileBytes"), Byte())
                contents.Add(MultimodalContent.CreateImageFile(fileName, fileBytes, mimeType))
            Else
                Dim fileData = inputs("fileData").ToString()
                contents.Add(MultimodalContent.CreateFile(fileName, fileData, mimeType))
            End If
        End Sub

        Private Function ExtractUserPrompt(inputs As Dictionary(Of String, Object)) As String
            If inputs.ContainsKey("text") Then Return inputs("text").ToString()
            If inputs.ContainsKey("prompt") Then Return inputs("prompt").ToString()
            If inputs.ContainsKey("message") Then Return inputs("message").ToString()
            ' 【归一化】前端 buildSkillInputs 已将 LLM 不同键（contents/request/content/description/userPrompt）归一到 text；
            ' 此处再兜底，兼容未经前端链路直接进入后端的输入。
            If inputs.ContainsKey("contents") Then Return inputs("contents").ToString()
            If inputs.ContainsKey("request") Then Return inputs("request").ToString()
            If inputs.ContainsKey("content") Then Return inputs("content").ToString()
            If inputs.ContainsKey("userPrompt") Then Return inputs("userPrompt").ToString()
            If inputs.ContainsKey("description") Then Return inputs("description").ToString()
            Return ""
        End Function

        Private Sub ProcessMultimodalResult(llmResult As LLMResponse, skill As SkillDefinition,
                                            result As WorkflowResult, stopwatch As System.Diagnostics.Stopwatch)
            If llmResult.Success Then
                result.Success = True
                result.Outputs("content") = llmResult.Content
                result.Outputs("processingTime") = stopwatch.Elapsed.TotalSeconds
                result.Outputs("usage") = llmResult.Usage

                Select Case skill.OutputFormat.ToLower()
                    Case "json"
                        ' 【修复】禁止"假成功"，同时杜绝"假失败"：LLM 输出常被 ```json 代码块、
                        ' thinking 标签、正文混杂或尾部污染包裹，裸 JObject.Parse 会在此误判
                        ' "未返回结构化 JSON"。统一复用 SkillExecutionRouter 的容错提取器：
                        ' 只有确实提取不到任何合法 JSON 时才判失败（仍禁止假成功）。
                        Try
                            Dim jsonText = String.Empty
                            If Not String.IsNullOrEmpty(llmResult.Content) Then
                                jsonText = SkillExecutionRouter.ExtractJsonFromLLMResponse(llmResult.Content)
                            End If
                            If String.IsNullOrEmpty(jsonText) Then jsonText = llmResult.Content
                            Dim jsonObj = Newtonsoft.Json.Linq.JObject.Parse(jsonText)
                            For Each prop In jsonObj.Properties()
                                result.Outputs(prop.Name) = prop.Value.ToObject(Of Object)()
                            Next
                        Catch ex As Exception
                            ' 【修复】禁止"假成功"：Skill 声明输出 JSON 但 LLM 未返回合法 JSON 时，
                            ' 必须判定为失败，而不是空吞异常返回 Success=True，避免前端误报已生成。
                            result.Success = False
                            result.ErrorMessage =
                                String.Format("LLM 未按要求返回结构化 JSON（Skill: {0}）。原文：{1}",
                                              skill.Name,
                                              If(String.IsNullOrEmpty(llmResult.Content), "(空)", llmResult.Content))
                            Utils.ChatLogger.Instance.LogError(
                                "SKILL_JSON",
                                $"【校验失败】Skill '{skill.Name}' 的 JSON 输出解析失败: {ex.Message}")
                        End Try
                    Case "vba"
                        result.Outputs("code") = llmResult.Content
                End Select
            Else
                result.Success = False
                result.ErrorMessage = llmResult.ErrorMessage
            End If
        End Sub

        ''' <summary>
        ''' 执行本地 OCR Skill（使用 PaddleOCR/Windows OCR 等专业 OCR 引擎）
        ''' </summary>
        Private Async Function ExecuteLocalOCRSkill(skill As SkillDefinition, inputs As Dictionary(Of String, Object), Optional progress As IProgress(Of String) = Nothing) As Task(Of WorkflowResult)
            Dim result As New WorkflowResult
            Dim stopwatch = System.Diagnostics.Stopwatch.StartNew()
            Dim logger = Utils.ChatLogger.Instance

            Try
                logger.LogInfo("LOCAL_OCR", $"═══════════════════════════════════════════════════════")
                logger.LogInfo("LOCAL_OCR", $"【开始】执行本地 OCR Skill: {skill.DisplayName}")

                Dim ocrResult = Await PerformOcrRecognition(inputs, progress, logger)
                If Not ocrResult.Success Then
                    result.Success = False
                    result.ErrorMessage = ocrResult.ErrorMessage
                    Return result
                End If

                result.Outputs("rawText") = ocrResult.OcrText
                result.Outputs("ocrEngine") = ocrResult.OcrEngine

                Await ProcessOcrResultWithLLM(skill, ocrResult.OcrText, inputs, result, progress, logger)

                stopwatch.Stop()
                result.Success = True
                result.Outputs("processingTime") = stopwatch.Elapsed.TotalSeconds

                LogOcrResultSummary(result, stopwatch, logger)
                progress?.Report("步骤 4/4: 完成")

            Catch ex As Exception
                stopwatch.Stop()
                result.Success = False
                result.ErrorMessage = $"本地 OCR 处理失败: {ex.Message}"
                logger.LogError("LOCAL_OCR", $"  ❌ 执行本地 OCR Skill 出错: {ex.Message}", ex)
            End Try

            Return result
        End Function

    End Class

End Namespace
