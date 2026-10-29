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

    ' SkillManager.Persistence.vb — Skill 持久化（Save/Reset/CRUD）
    ' （类级 XML 文档只在主文件 SkillManager.vb 保留一份，避免 BC42314 分部重复）
    Partial Public Class SkillManager

        Public Sub SaveSkills()
            Try
                ' 保存 Skills 到 Markdown 文件(唯一存储格式)
                SaveSkillsToMarkdown()
                System.Diagnostics.Debug.WriteLine("Saved skills to Markdown files (version: " & CurrentSkillVersion.ToString() & ")")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Error saving skills: " & ex.Message)
            End Try
        End Sub
        
        ''' <summary>
        ''' 保存 Skills 到 Markdown 文件
        ''' 【修复】只保存非自定义的系统内置技能到 bin 目录
        ''' 自定义技能由 CustomSkillManager 负责保存到 AppData，不会被构建覆盖
        ''' </summary>
        Private Sub SaveSkillsToMarkdown()
            Try
                Dim mdDir = Utils.PathHelper.GetMarkdownSkillsDirectory()
                If Not Directory.Exists(mdDir) Then
                    Directory.CreateDirectory(mdDir)
                End If

                Dim savedCount As Integer = 0
                Dim skippedCount As Integer = 0
                For Each skill In _skills.Values
                    Try
                        ' 跳过自定义技能，由 CustomSkillManager 负责持久化
                        If skill.IsCustom Then
                            skippedCount += 1
                            Continue For
                        End If

                        Dim mdContent = ConvertSkillToMarkdown(skill)
                        Dim mdFilePath = Path.Combine(mdDir, skill.Name & ".md")
                        File.WriteAllText(mdFilePath, mdContent)
                        savedCount += 1
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine("Error saving skill to Markdown (" & skill.Name & "): " & ex.Message)
                    End Try
                Next
                
                System.Diagnostics.Debug.WriteLine(String.Format("Saved {0} system skills to Markdown files in: {1} (skipped {2} custom skills)", savedCount, mdDir, skippedCount))
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Error saving skills to Markdown: " & ex.Message)
            End Try
        End Sub
        
        ''' <summary>
        ''' 将 SkillDefinition 转换为 Markdown 格式
        ''' </summary>
        Private Function ConvertSkillToMarkdown(skill As SkillDefinition) As String
            Dim sb As New System.Text.StringBuilder()
            
            sb.AppendLine("---")
            AppendBasicFieldsToYaml(sb, skill)
            AppendShortcutConfigToYaml(sb, skill)
            AppendOcrConfigToYaml(sb, skill)
            AppendExcelConfigToYaml(sb, skill)
            sb.AppendLine("---")
            sb.AppendLine()
            
            AppendMarkdownBody(sb, skill)
            
            Return sb.ToString()
        End Function

        Private Sub AppendBasicFieldsToYaml(sb As System.Text.StringBuilder, skill As SkillDefinition)
            sb.AppendLine("name: " & skill.Name)
            sb.AppendLine("displayName: " & skill.DisplayName)
            sb.AppendLine("description: " & skill.Description)
            
            If skill.Keywords IsNot Nothing AndAlso skill.Keywords.Length > 0 Then
                sb.AppendLine("keywords:")
                For Each kw In skill.Keywords
                    sb.AppendLine("  - " & kw)
                Next
            End If
            
            If skill.Examples IsNot Nothing AndAlso skill.Examples.Length > 0 Then
                sb.AppendLine("examples:")
                For Each ex In skill.Examples
                    sb.AppendLine("  - " & ex)
                Next
            End If
            
            sb.AppendLine("isGeneric: " & skill.IsGeneric.ToString().ToLower())
            sb.AppendLine("needsConfig: " & skill.NeedsConfig.ToString().ToLower())
            
            If skill.ConfigFields IsNot Nothing AndAlso skill.ConfigFields.Length > 0 Then
                sb.AppendLine("configFields:")
                For Each cf In skill.ConfigFields
                    sb.AppendLine("  - " & cf)
                Next
            Else
                sb.AppendLine("configFields: []")
            End If
            
            sb.AppendLine("processingMode: " & skill.ProcessingMode.ToString())
            sb.AppendLine("supportsImageInput: " & skill.SupportsImageInput.ToString().ToLower())
            sb.AppendLine("supportsFileInput: " & skill.SupportsFileInput.ToString().ToLower())
            sb.AppendLine("outputFormat: " & skill.OutputFormat)
            sb.AppendLine("isCustom: " & skill.IsCustom.ToString().ToLower())

            ' Skill 开关持久化（只在 False 时写，True 是默认值）
            If Not skill.IsEnabled Then
                sb.AppendLine("isEnabled: false")
            End If

            ' Pipeline 路由声明
            If Not String.IsNullOrWhiteSpace(skill.Pipeline) Then
                sb.AppendLine("pipeline: " & skill.Pipeline)
            End If

            ' InputContext 声明
            If skill.InputContext IsNot Nothing AndAlso skill.InputContext.Length > 0 Then
                sb.AppendLine("inputContext:")
                For Each ctx In skill.InputContext
                    sb.AppendLine("  - " & ctx)
                Next
            End If

            ' Category
            If Not String.IsNullOrWhiteSpace(skill.Category) Then
                sb.AppendLine("category: " & skill.Category)
            End If

            If skill.Tags IsNot Nothing AndAlso skill.Tags.Count > 0 Then
                sb.AppendLine("tags:")
                For Each tag In skill.Tags
                    sb.AppendLine("  - " & tag)
                Next
            End If
            
            sb.AppendLine("version: " & skill.Version.ToString())
        End Sub

        Private Sub AppendShortcutConfigToYaml(sb As System.Text.StringBuilder, skill As SkillDefinition)
            sb.AppendLine("isQuickShortcut: " & skill.IsQuickShortcut.ToString().ToLower())
            If Not String.IsNullOrEmpty(skill.QuickShortcutIcon) AndAlso skill.QuickShortcutIcon <> "⚡" Then
                sb.AppendLine("quickShortcutIcon: " & skill.QuickShortcutIcon)
            End If
            If Not String.IsNullOrEmpty(skill.QuickShortcutPrompt) Then
                sb.AppendLine("quickShortcutPrompt: " & skill.QuickShortcutPrompt)
            End If
        End Sub

        Private Sub AppendOcrConfigToYaml(sb As System.Text.StringBuilder, skill As SkillDefinition)
            If skill.OcrConfig Is Nothing Then Return

            sb.AppendLine("ocrConfig:")
            sb.AppendLine("  engine: " & skill.OcrConfig.Engine)
            sb.AppendLine("  fallbackEngine: " & skill.OcrConfig.FallbackEngine)
            
            If skill.OcrConfig.Preprocessing IsNot Nothing Then
                sb.AppendLine("  preprocessing:")
                sb.AppendLine("    deskew: " & skill.OcrConfig.Preprocessing.Deskew.ToString().ToLower())
                sb.AppendLine("    denoise: " & skill.OcrConfig.Preprocessing.Denoise.ToString().ToLower())
                sb.AppendLine("    contrastEnhancement: " & skill.OcrConfig.Preprocessing.ContrastEnhancement.ToString().ToLower())
            End If
            
            If skill.OcrConfig.PostProcessing IsNot Nothing Then
                sb.AppendLine("  postprocessing:")
                sb.AppendLine("    useLLM: " & skill.OcrConfig.PostProcessing.UseLLM.ToString().ToLower())
                If Not String.IsNullOrEmpty(skill.OcrConfig.PostProcessing.LLMModel) Then
                    sb.AppendLine("    llmModel: " & skill.OcrConfig.PostProcessing.LLMModel)
                End If
                If Not String.IsNullOrEmpty(skill.OcrConfig.PostProcessing.ExtractionPrompt) Then
                    sb.AppendLine("    extractionPrompt: " & skill.OcrConfig.PostProcessing.ExtractionPrompt)
                End If
            End If
        End Sub

        Private Sub AppendExcelConfigToYaml(sb As System.Text.StringBuilder, skill As SkillDefinition)
            If skill.ExcelConfig Is Nothing Then Return

            sb.AppendLine("excelConfig:")
            sb.AppendLine("  defaultWorksheet: " & skill.ExcelConfig.DefaultWorksheet)
            
            If skill.ExcelConfig.FieldMapping IsNot Nothing AndAlso skill.ExcelConfig.FieldMapping.Count > 0 Then
                sb.AppendLine("  fieldMapping:")
                For Each kvp In skill.ExcelConfig.FieldMapping
                    sb.AppendLine("    " & kvp.Key & ":")
                    sb.AppendLine("      column: " & kvp.Value.Column)
                    sb.AppendLine("      header: " & kvp.Value.Header)
                Next
            End If
            
            If skill.ExcelConfig.FieldNameMapping IsNot Nothing AndAlso skill.ExcelConfig.FieldNameMapping.Count > 0 Then
                sb.AppendLine("  fieldNameMapping:")
                For Each kvp In skill.ExcelConfig.FieldNameMapping
                    sb.AppendLine("    " & kvp.Key & ": " & kvp.Value)
                Next
            End If
        End Sub

        Private Sub AppendMarkdownBody(sb As System.Text.StringBuilder, skill As SkillDefinition)
            sb.AppendLine("# " & skill.DisplayName)
            sb.AppendLine()
            
            If Not String.IsNullOrEmpty(skill.LLMSystemPrompt) Then
                sb.AppendLine("## System Prompt")
                sb.AppendLine()
                sb.AppendLine(skill.LLMSystemPrompt)
                sb.AppendLine()
            End If
            
            If Not String.IsNullOrEmpty(skill.TemplateCode) Then
                sb.AppendLine("## Template Code")
                sb.AppendLine()
                sb.AppendLine("```vba")
                sb.AppendLine(skill.TemplateCode)
                sb.AppendLine("```")
                sb.AppendLine()
            End If
        End Sub

        ''' <summary>
        ''' 还原 Skills 到默认状态(删除所有 Markdown 文件,下次启动时重新生成)
        ''' </summary>
        Public Function ResetSkillsToDefault() As Boolean
            Try
                System.Diagnostics.Debug.WriteLine("开始还原 Skills 到默认状态...")
        
                ' 删除所有 Markdown Skill 文件
                Dim mdDir = Utils.PathHelper.GetMarkdownSkillsDirectory()
                If Directory.Exists(mdDir) Then
                    Dim mdFiles = Directory.GetFiles(mdDir, "*.md")
                    For Each mdFile In mdFiles
                        File.Delete(mdFile)
                        System.Diagnostics.Debug.WriteLine("已删除 Markdown Skill: " & mdFile)
                    Next
                End If
        
                ' 清空当前内存中的 skills
                _skills.Clear()
                System.Diagnostics.Debug.WriteLine("已清空内存中的 skills")
        
                ' 重新加载默认 skills(从代码默认配置生成 Markdown 文件)
                LoadSkills()
                System.Diagnostics.Debug.WriteLine("已重新加载默认 skills,当前数量: " & _skills.Count.ToString())
        
                ' 保存到 Markdown 文件(重新生成)
                SaveSkills()
                System.Diagnostics.Debug.WriteLine("已重新生成 Markdown Skill 文件")
        
                Return True
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("还原 Skills 到默认状态失败: " & ex.Message)
                Return False
            End Try
        End Function

        Public Function GetSkill(name As String) As SkillDefinition
            If _skills.ContainsKey(name) Then
                Return _skills(name)
            End If
            Return Nothing
        End Function

        Public Function GetSpecificSkills() As List(Of SkillDefinition)
            Return _skills.Values.Where(Function(s) Not s.IsGeneric).ToList()
        End Function

        Public Sub AddSkill(skill As SkillDefinition)
            If skill Is Nothing Then Return
            _skills(skill.Name) = skill
            InvalidateCache()  ' 清除列表缓存，确保 GetAllSkills 返回最新数据

            ' 只保存真正的自定义技能到 CustomSkills 目录
            If skill.IsCustom Then
                Try
                    Dim customManager = CustomSkillManager.GetInstance()
                    customManager.SaveCustomSkill(skill)
                    System.Diagnostics.Debug.WriteLine($"Skill '{skill.Name}' 已保存到 CustomSkills 目录")
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"保存 Skill 到 CustomSkills 失败: {ex.Message}")
                End Try
            End If
        End Sub

        Public Function RemoveSkill(skillName As String) As Boolean
            If String.IsNullOrWhiteSpace(skillName) Then Return False

            ' 从内存中移除
            Dim result = _skills.Remove(skillName)
            If result Then InvalidateCache()  ' 清除列表缓存

            ' 同时从 CustomSkills 目录删除
            If result Then
                Try
                    Dim customManager = CustomSkillManager.GetInstance()
                    customManager.DeleteCustomSkill(skillName)
                    System.Diagnostics.Debug.WriteLine($"Skill '{skillName}' 已从 CustomSkills 目录删除")
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"从 CustomSkills 删除 Skill 失败: {ex.Message}")
                End Try
            End If

            Return result
        End Function

        Public Function GetCompositeSkills() As List(Of SkillDefinition)
            Return _skills.Values.Where(Function(s) s.IsComposite).ToList()
        End Function

    End Class

End Namespace
