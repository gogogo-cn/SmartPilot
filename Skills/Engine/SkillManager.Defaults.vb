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

    ' SkillManager.Defaults.vb — LLM 配置助手 / 默认与复合 Skill 创建 / 简单 Skill 执行
    ' （类级 XML 文档只在主文件 SkillManager.vb 保留一份，避免 BC42314 分部重复）
    Partial Public Class SkillManager

        ''' <summary>
        ''' 【修复】提取公共方法：创建 LLM 配置，避免重复代码
        ''' </summary>
        Private Function CreateLLMConfig(systemPrompt As String, Optional defaultModel As String = "", Optional temperature As Double? = Nothing, Optional maxTokens As Integer? = Nothing) As LLMRequestConfig
            Dim provider = ReadLLMConfig(Utils.ConfigConstants.LLM.PROVIDER, "LLM_PROVIDER", Utils.ConfigConstants.LLM.DEFAULT_PROVIDER)
            Dim apiUrl As String = ""
            Dim model As String = defaultModel

            If provider.ToLower() = "ollama" Then
                apiUrl = GetConfigValue(Utils.ConfigConstants.Ollama.API_URL, Utils.ConfigConstants.Ollama.DEFAULT_API_URL)
                If String.IsNullOrEmpty(model) Then
                    model = ReadLLMConfig(Utils.ConfigConstants.LLM.MODEL, "LLM_MODEL", Utils.ConfigConstants.Ollama.DEFAULT_MODEL)
                End If
            Else
                apiUrl = ReadLLMConfig(Utils.ConfigConstants.LLM.API_URL, "LLM_API_URL", "")
                If String.IsNullOrEmpty(model) Then
                    model = ReadLLMConfig(Utils.ConfigConstants.LLM.MODEL, "LLM_MODEL", Utils.ConfigConstants.LLM.DEFAULT_MODEL)
                End If
            End If

            Dim config As New LLMRequestConfig With {
                .Provider = provider,
                .Model = model,
                .ApiKey = ReadLLMConfig(Utils.ConfigConstants.LLM.API_KEY, "LLM_API_KEY", ""),
                .ApiUrl = apiUrl,
                .SystemPrompt = systemPrompt
            }

            If temperature.HasValue Then config.Temperature = temperature.Value
            If maxTokens.HasValue Then config.MaxTokens = maxTokens.Value

            Return config
        End Function

        ''' <summary>
        ''' 读取 LLM 配置，兼容旧版下划线键名（LLM_PROVIDER 等）和新版驼峰键名（LLMProvider 等）
        ''' </summary>
        Private Function ReadLLMConfig(key As String, legacyKey As String, defaultValue As String) As String
            Dim value = GetConfigValue(key, "")
            If String.IsNullOrEmpty(value) Then
                value = GetConfigValue(legacyKey, defaultValue)
            End If
            Return value
        End Function

        ''' <summary>
        ''' 从文件扩展名获取 MIME 类型
        ''' </summary>
        Private Function GetMimeTypeFromExtension(extension As String) As String
            Select Case extension.ToLower()
                Case ".jpg", ".jpeg"
                    Return "image/jpeg"
                Case ".png"
                    Return "image/png"
                Case ".gif"
                    Return "image/gif"
                Case ".bmp"
                    Return "image/bmp"
                Case ".webp"
                    Return "image/webp"
                Case Else
                    Return "image/png"
            End Select
        End Function

        ''' <summary>
        ''' 判断是否为图片 MIME 类型
        ''' </summary>
        Private Function IsImageMimeType(mimeType As String) As Boolean
            If String.IsNullOrEmpty(mimeType) Then Return False
            Dim imageTypes As String() = {"image/jpeg", "image/png", "image/gif", "image/webp", "image/bmp", "image/tiff"}
            Return imageTypes.Any(Function(t) mimeType.ToLower().StartsWith(t))
        End Function

        ''' <summary>
        ''' 获取配置值
        ''' </summary>
        Private Function GetConfigValue(key As String, defaultValue As String) As String
            Try
                Dim configManager = Utils.ConfigurationManager.Instance
                Return configManager.GetSetting(key, defaultValue)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"GetConfigValue 失败 [key={key}]: {ex.Message}")
                Return defaultValue
            End Try
        End Function

        ''' <summary>
        ''' 获取图片识别填入配置（配置驱动）
        ''' </summary>
        Private Function GetImageFillConfig(imageType As String) As ImageFillConfig
            Dim config As New ImageFillConfig With {
                .ImageType = imageType,
                .UseInvoiceMethod = False,
                .HasImagePath = False
            }

            config.SheetName = "数据录入"
            config.UseInvoiceMethod = False
            config.HasImagePath = False

            Return config
        End Function

        ''' <summary>
        ''' 执行简单 Skill
        ''' </summary>
        Private Function ExecuteSimpleSkill(skill As SkillDefinition, inputs As Dictionary(Of String, Object)) As WorkflowResult
            Dim result As New WorkflowResult

            Try
                If skill.IsGeneric Then
                    result.Success = True
                    result.Outputs("message") = $"Skill [{skill.DisplayName}] 已识别，等待 AI 生成响应"
                Else
                    result.Success = False
                    result.ErrorMessage = "非通用 Skill 缺少工作流定义"
                End If
            Catch ex As Exception
                result.Success = False
                result.ErrorMessage = ex.Message
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 创建复合 Skill
        ''' </summary>
        Public Function CreateCompositeSkill(name As String, displayName As String, description As String,
                                             keywords As String()) As SkillDefinition
            Dim skill As New SkillDefinition()
            skill.Name = name
            skill.DisplayName = displayName
            skill.Description = description
            skill.Keywords = keywords
            skill.IsComposite = True
            skill.IsGeneric = False
            skill.WorkflowSteps = New List(Of WorkflowStep)()
            skill.InputParameters = New List(Of ParameterDefinition)()
            skill.OutputParameters = New List(Of ParameterDefinition)()
            skill.Version = CurrentSkillVersion
            Return skill
        End Function

        Private Sub CreateDefaultSkills()
            _skills = New Dictionary(Of String, SkillDefinition)(StringComparer.OrdinalIgnoreCase)
            
            ' 优先从 Skills/Skills.json 加载完整的默认技能
            Dim appDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            Dim skillsPath = Path.Combine(appDir, "Skills", "Skills.json")
            
            If File.Exists(skillsPath) Then
                Try
                    Dim json = File.ReadAllText(skillsPath)
                    Dim skillList = JsonConvert.DeserializeObject(Of SkillList)(json)
                    
                    If skillList IsNot Nothing AndAlso skillList.Skills IsNot Nothing Then
                        For Each skill In skillList.Skills
                            _skills(skill.Name) = skill
                            System.Diagnostics.Debug.WriteLine("从 Skills/Skills.json 加载默认 Skill: " & skill.Name)
                        Next
                        System.Diagnostics.Debug.WriteLine(String.Format("成功从 Skills/Skills.json 加载 {0} 个默认 Skills", skillList.Skills.Count))
                        Return ' 成功加载，直接返回
                    End If
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("从 Skills/Skills.json 加载默认技能失败: " & ex.Message)
                End Try
            End If
            
            ' 【修复】如果 Skills.json 和 Markdown 都加载失败，创建最小化回退技能
            ' 确保系统始终有可用的基础技能，避免完全无技能可用的情况
            If _skills.Count = 0 Then
                System.Diagnostics.Debug.WriteLine("⚠️ [回退] Skills.json 和 Markdown 均未加载，创建最小化默认技能...")
                LogToFile("⚠️ [回退] 创建最小化默认技能")
                CreateFallbackSkills()
            End If
        End Sub

        ''' <summary>
        ''' 创建最小化回退技能（确保系统始终有可用技能）
        ''' </summary>
        Private Sub CreateFallbackSkills()
            Dim generateVBA As New SkillDefinition() With {
                .Name = "GenerateVBA",
                .DisplayName = "写 VBA 代码",
                .Description = "根据自然语言描述生成 Excel VBA 代码",
                .Keywords = New String() {"代码", "VBA", "宏", "编写", "生成"},
                .Examples = New String() {"写一个排序宏", "生成批量重命名工作表的代码"},
                .IsGeneric = True,
                .NeedsConfig = False,
                .ProcessingMode = SkillProcessingMode.MultimodalLLM,
                .SupportsImageInput = False,
                .SupportsFileInput = False,
                .OutputFormat = "vba",
                .IsCustom = False,
                .Tags = New List(Of String)(New String() {"VBA", "代码生成"}),
                .Version = CurrentSkillVersion,
                .CreatedDate = DateTime.Now,
                .LastModifiedDate = DateTime.Now,
                .CreatedBy = "Fallback",
                .UsageCount = 0
            }
            _skills(generateVBA.Name) = generateVBA

            Dim smartSummary As New SkillDefinition() With {
                .Name = "SmartSummary",
                .DisplayName = "智能总结",
                .Description = "对选中的 Excel 数据进行智能分析和总结",
                .Keywords = New String() {"总结", "分析", "统计", "汇总"},
                .Examples = New String() {"帮我总结一下这些数据", "分析表格内容"},
                .IsGeneric = True,
                .NeedsConfig = False,
                .ProcessingMode = SkillProcessingMode.MultimodalLLM,
                .SupportsImageInput = False,
                .SupportsFileInput = False,
                .OutputFormat = "text",
                .IsCustom = False,
                .Tags = New List(Of String)(New String() {"分析", "总结"}),
                .Version = CurrentSkillVersion,
                .CreatedDate = DateTime.Now,
                .LastModifiedDate = DateTime.Now,
                .CreatedBy = "Fallback",
                .UsageCount = 0
            }
            _skills(smartSummary.Name) = smartSummary

            System.Diagnostics.Debug.WriteLine($"✅ [回退] 已创建 {_skills.Count} 个最小化默认技能")
            LogToFile($"✅ [回退] 已创建 {_skills.Count} 个最小化默认技能")
        End Sub

    End Class

End Namespace
