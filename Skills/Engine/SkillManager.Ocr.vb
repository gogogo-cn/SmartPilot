Imports System.IO
Imports System.Reflection
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Serialization
Imports System.Threading
Imports System.Threading.Tasks
Imports ExcelAIVSTOAddIn.Services
Imports ExcelAIVSTOAddIn.Utils
Imports YamlDotNet.RepresentationModel

Namespace Skills

    ' SkillManager.Ocr.vb — 本地 OCR 执行
    ' （类级 XML 文档只在主文件 SkillManager.vb 保留一份，避免 BC42314 分部重复）
    Partial Public Class SkillManager

        Private Structure OcrRecognitionResult
            Public Property Success As Boolean
            Public Property OcrText As String
            Public Property OcrEngine As String
            Public Property ErrorMessage As String
        End Structure

        Private Async Function PerformOcrRecognition(inputs As Dictionary(Of String, Object), progress As IProgress(Of String), logger As Utils.ChatLogger) As Task(Of OcrRecognitionResult)
            Dim result As New OcrRecognitionResult With {.Success = False, .OcrEngine = "Unknown"}

            If inputs.ContainsKey("ocrText") Then
                result.OcrText = inputs("ocrText").ToString()
                If Not String.IsNullOrWhiteSpace(result.OcrText) Then
                    logger.LogInfo("LOCAL_OCR", $"【优化】使用前端传入的OCR文字，跳过本地OCR识别")
                    logger.LogInfo("LOCAL_OCR", $"  ├─ 传入文字长度: {result.OcrText.Length} 字符")
                    result.OcrEngine = "FrontendOCR"
                    progress?.Report("步骤 2/4: 使用前端传入的OCR文字...")
                    result.Success = True
                    Return result
                End If
            End If

            progress?.Report("步骤 1/4: 初始化 OCR 引擎...")
            Dim imageBase64 As String = ""
            Dim imagePath As String = ""

            If inputs.ContainsKey("imageData") Then
                imageBase64 = inputs("imageData").ToString()
                logger.LogInfo("LOCAL_OCR", $"  ├─ 输入类型: Base64 图片数据")
            ElseIf inputs.ContainsKey("imagePath") Then
                imagePath = inputs("imagePath").ToString()
                logger.LogInfo("LOCAL_OCR", $"  ├─ 输入类型: 图片文件路径: {imagePath}")
            Else
                result.ErrorMessage = "缺少图片输入（imageData 或 imagePath）"
                logger.LogError("LOCAL_OCR", $"  ❌ 错误: {result.ErrorMessage}")
                Return result
            End If

            progress?.Report("步骤 2/4: 使用 PaddleOCR 识别文字...")
            logger.LogInfo("LOCAL_OCR", $"【步骤 2/4】开始 OCR 文字识别")

            Dim ocrManager = OcrServiceManager.Instance
            result.OcrEngine = ocrManager.CurrentEngine.ToString()

            logger.LogInfo("LOCAL_OCR", $"  ├─ 当前引擎: {result.OcrEngine}")
            logger.LogInfo("LOCAL_OCR", $"  ├─ 引擎可用: {ocrManager.IsEngineAvailable(ocrManager.CurrentEngine)}")

            Try
                If Not String.IsNullOrEmpty(imageBase64) Then
                    result.OcrText = Await RecognizeFromBase64(imageBase64, ocrManager)
                ElseIf Not String.IsNullOrEmpty(imagePath) AndAlso File.Exists(imagePath) Then
                    result.OcrText = Await ocrManager.RecognizeText(imagePath)
                End If
            Catch ex As Exception
                result.ErrorMessage = $"OCR 识别失败: {ex.Message}"
                logger.LogError("LOCAL_OCR", $"  ❌ OCR 失败: {result.ErrorMessage}")
                Return result
            End Try

            If String.IsNullOrWhiteSpace(result.OcrText) Then
                result.ErrorMessage = "OCR 识别失败: 未能识别出文字（PaddleOCR 返回空结果）"
                logger.LogError("LOCAL_OCR", $"  ❌ OCR 失败: {result.ErrorMessage}")
                logger.LogError("LOCAL_OCR", $"     请检查: 1. PaddleOCR 是否正确安装 2. 图片是否清晰 3. OCR 服务是否启动")
                Return result
            End If

            logger.LogInfo("LOCAL_OCR", $"  ├─ OCR 引擎: {result.OcrEngine}")
            logger.LogInfo("LOCAL_OCR", $"  ├─ 识别文字数: {result.OcrText.Length} 字符")
            logger.LogInfo("LOCAL_OCR", $"  ├─ 识别结果前 200 字符: {result.OcrText.Substring(0, Math.Min(200, result.OcrText.Length))}...")
            result.Success = True

            Return result
        End Function

        Private Async Function RecognizeFromBase64(imageBase64 As String, ocrManager As OcrServiceManager) As Task(Of String)
            If imageBase64.Contains(",") Then
                imageBase64 = imageBase64.Substring(imageBase64.IndexOf(",") + 1)
            End If

            Dim tempPath = Path.GetTempFileName() & ".png"
            Try
                Dim imageBytes = Convert.FromBase64String(imageBase64)
                File.WriteAllBytes(tempPath, imageBytes)
                Return Await ocrManager.RecognizeText(tempPath)
            Finally
                If File.Exists(tempPath) Then File.Delete(tempPath)
            End Try
        End Function

        Private Async Function ProcessOcrResultWithLLM(skill As SkillDefinition, ocrText As String, inputs As Dictionary(Of String, Object),
                                                        result As WorkflowResult, progress As IProgress(Of String), logger As Utils.ChatLogger) As Task
            Dim forceUseLLM As Boolean = False
            If inputs.ContainsKey("useLLM") Then
                Dim llmVal = inputs("useLLM")
                If llmVal IsNot Nothing Then
                    Dim llmStr = llmVal.ToString().Trim()
                    forceUseLLM = llmStr.Equals("true", StringComparison.OrdinalIgnoreCase)
                End If
            End If

            If (skill.OcrConfig IsNot Nothing AndAlso skill.OcrConfig.PostProcessing.UseLLM) OrElse forceUseLLM Then
                progress?.Report("步骤 3/4: 使用 LLM 提取结构化数据...")
                logger.LogInfo("LOCAL_OCR", $"【步骤 3/4】LLM 后处理提取结构化数据")

                Dim extractionPrompt = BuildExtractionPrompt(skill, ocrText)
                Dim llmModel = GetLlmModelForOcr(skill)

                LLMEngine.Initialize(CreateLLMConfig(skill.LLMSystemPrompt, llmModel, 0.3, 2048))

                Dim llmResult = Await LLMEngine.ProcessTextAsync(extractionPrompt)

                If llmResult.Success Then
                    logger.LogInfo("LOCAL_OCR", $"  ├─ LLM 提取成功")
                    logger.LogInfo("LOCAL_OCR", $"  ├─ 提取结果: {llmResult.Content.Substring(0, Math.Min(200, llmResult.Content.Length))}...")
                    ParseLlmExtractionResult(llmResult.Content, result, logger)
                Else
                    logger.LogWarning("LOCAL_OCR", $"  ⚠️ LLM 提取失败: {llmResult.ErrorMessage}")
                    result.Outputs("extractedText") = ocrText
                End If
            Else
                progress?.Report("步骤 3/4: 跳过 LLM 后处理...")
                logger.LogInfo("LOCAL_OCR", $"【步骤 3/4】跳过 LLM 后处理，直接返回 OCR 结果")
                result.Outputs("extractedText") = ocrText
            End If
        End Function

        Private Function BuildExtractionPrompt(skill As SkillDefinition, ocrText As String) As String
            Dim extractionPrompt As String = ""
            If skill.OcrConfig IsNot Nothing AndAlso skill.OcrConfig.PostProcessing IsNot Nothing Then
                extractionPrompt = skill.OcrConfig.PostProcessing.ExtractionPrompt
            End If

            If String.IsNullOrEmpty(extractionPrompt) Then
                Return $"从以下 OCR 识别结果中提取关键信息，以 JSON 格式返回。识别结果：{vbCrLf}{vbCrLf}{ocrText}"
            Else
                Return $"{extractionPrompt}{vbCrLf}{vbCrLf}OCR 识别结果：{vbCrLf}{ocrText}"
            End If
        End Function

        Private Function GetLlmModelForOcr(skill As SkillDefinition) As String
            Dim llmModel As String = ""
            If skill.OcrConfig IsNot Nothing AndAlso skill.OcrConfig.PostProcessing IsNot Nothing Then
                llmModel = skill.OcrConfig.PostProcessing.LLMModel
            End If
            If String.IsNullOrEmpty(llmModel) OrElse llmModel.StartsWith("$") Then
                llmModel = GetConfigValue(Utils.ConfigConstants.LLM.MODEL, Utils.ConfigConstants.Ollama.DEFAULT_MODEL)
            End If
            Return llmModel
        End Function

        Private Sub ParseLlmExtractionResult(content As String, result As WorkflowResult, logger As Utils.ChatLogger)
            Dim cleanedContent = content.Trim()
            If cleanedContent.StartsWith("```json") Then
                cleanedContent = cleanedContent.Substring(7).Trim()
            End If
            If cleanedContent.StartsWith("```") Then
                cleanedContent = cleanedContent.Substring(3).Trim()
            End If
            If cleanedContent.EndsWith("```") Then
                cleanedContent = cleanedContent.Substring(0, cleanedContent.Length - 3).Trim()
            End If

            Try
                Dim jsonObj = Newtonsoft.Json.Linq.JObject.Parse(cleanedContent)
                result.Outputs("structuredData") = jsonObj.ToObject(Of Dictionary(Of String, Object))()

                For Each prop In jsonObj.Properties()
                    result.Outputs(prop.Name) = prop.Value.ToObject(Of Object)()
                Next

                logger.LogInfo("LOCAL_OCR", $"  ├─ JSON解析成功，字段: {String.Join(", ", jsonObj.Properties().Select(Function(p) p.Name))}")
            Catch ex As Exception
                logger.LogWarning("LOCAL_OCR", $"  ⚠️ JSON 解析失败，使用原始文本: {ex.Message}")
                logger.LogWarning("LOCAL_OCR", $"  ⚠️ 原始内容: {cleanedContent.Substring(0, Math.Min(200, cleanedContent.Length))}...")
                result.Outputs("extractedText") = cleanedContent
            End Try
        End Sub

        Private Sub LogOcrResultSummary(result As WorkflowResult, stopwatch As System.Diagnostics.Stopwatch, logger As Utils.ChatLogger)
            logger.LogInfo("LOCAL_OCR", $"  ├─ 返回 outputs 键: {String.Join(", ", result.Outputs.Keys)}")
            If result.Outputs.ContainsKey("rawText") Then
                logger.LogInfo("LOCAL_OCR", $"  ├─ rawText 长度: {result.Outputs("rawText").ToString().Length}")
            End If

            logger.LogInfo("LOCAL_OCR", $"【步骤 4/4】处理完成")
            logger.LogInfo("LOCAL_OCR", $"  ├─ 总耗时: {stopwatch.Elapsed.TotalSeconds:F2} 秒")
            logger.LogInfo("LOCAL_OCR", $"  └─ 成功: {result.Success}")
            logger.LogInfo("LOCAL_OCR", $"═══════════════════════════════════════════════════════")
        End Sub

    End Class

End Namespace
