Imports System.IO
Imports System.Collections.Generic
Imports System.Reflection
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports YamlDotNet.RepresentationModel
Imports YamlDotNet.Core

Namespace Skills
    ''' <summary>
    ''' 自定义Skill管理器（单例模式）
    ''' </summary>
    Public Class CustomSkillManager
        Private Shared _instance As CustomSkillManager
        Private _customSkillsPath As String
        Private _snippetsPath As String
        Private _customSkills As List(Of SkillDefinition)
        Private _codeSnippets As List(Of SkillDefinition)

        Public Shared Function GetInstance() As CustomSkillManager
            If _instance Is Nothing Then
                _instance = New CustomSkillManager()
            End If
            Return _instance
        End Function
        
        ''' <summary>
        ''' 重置单例实例（用于强制重新加载）
        ''' </summary>
        Public Shared Sub ResetInstance()
            _instance = Nothing
        End Sub

        Private Sub New()

            ' 使用 AppData 目录存储自定义 Skill
            _customSkillsPath = Path.Combine(Utils.PathHelper.CustomSkillsPath, "skills")
            _snippetsPath = Path.Combine(Utils.PathHelper.CustomSkillsPath, "snippets")
            
            Utils.AppLogger.LogInfo("CustomSkillManager", $"自定义Skill路径: {_customSkillsPath}")
            Utils.AppLogger.LogInfo("CustomSkillManager", $"代码片段路径: {_snippetsPath}")
            
            ' 确保根目录和子目录存在
            Utils.PathHelper.EnsureDirectoryExists(Utils.PathHelper.CustomSkillsPath)
            Utils.PathHelper.EnsureDirectoryExists(_customSkillsPath)
            Utils.PathHelper.EnsureDirectoryExists(_snippetsPath)
            
            ' 自动迁移旧目录中的文件到新的子目录结构
            MigrateOldFiles()
            
            LoadCustomSkills()
        End Sub

        ''' <summary>
        ''' 自动迁移旧目录中的文件到新的子目录结构
        ''' </summary>
        Private Sub MigrateOldFiles()
            Dim rootPath = Utils.PathHelper.CustomSkillsPath
            If Not Directory.Exists(rootPath) Then Return

            Dim movedCount = 0
            Dim mdFiles = Directory.GetFiles(rootPath, "*.md")

            For Each filePath In mdFiles
                Dim fileName = Path.GetFileName(filePath)
                Dim targetPath As String

                ' Snippet_ 开头的是代码片段，移到 snippets/ 目录
                If fileName.StartsWith("Snippet_", StringComparison.OrdinalIgnoreCase) Then
                    targetPath = Path.Combine(_snippetsPath, fileName)
                Else
                    ' 其他的移到 skills/ 目录
                    targetPath = Path.Combine(_customSkillsPath, fileName)
                End If

                ' 如果目标位置不存在则移动
                If Not File.Exists(targetPath) Then
                    Try
                        File.Move(filePath, targetPath)
                        movedCount += 1
                        Utils.AppLogger.LogInfo("CustomSkill", $"迁移: {fileName} -> {If(fileName.StartsWith("Snippet_"), "snippets", "skills")}/")
                    Catch ex As Exception
                        Utils.AppLogger.LogWarning("CustomSkill", $"迁移失败: {fileName}, {ex.Message}")
                    End Try
                End If
            Next

            If movedCount > 0 Then
                Utils.AppLogger.LogInfo("CustomSkill", $"迁移完成，共移动 {movedCount} 个文件")
            End If
        End Sub

        ''' <summary>
        ''' 加载所有自定义Skill
        ''' </summary>
        Private Sub LoadCustomSkills()
            _customSkills = New List(Of SkillDefinition)()
            _codeSnippets = New List(Of SkillDefinition)()

            If Not Directory.Exists(_customSkillsPath) Then
                Directory.CreateDirectory(_customSkillsPath)
            End If
            If Not Directory.Exists(_snippetsPath) Then
                Directory.CreateDirectory(_snippetsPath)
            End If

            ' 加载自定义技能（skills/ 目录）
            For Each skillFile In Directory.GetFiles(_customSkillsPath, "*.md")
                Try
                    Utils.AppLogger.LogDebug("CustomSkill", $"加载MD: {Path.GetFileName(skillFile)}")
                    Dim skill = ParseMarkdownSkillFile(skillFile)
                    ' 【同名合并防御】记录来源文件与写入时间：SkillManager 合并同名内置技能时，
                    ' 靠这个时间判断"是系统升级了还是用户刚编辑"（见 SkillManager.Loading 的合并逻辑）
                    If skill IsNot Nothing Then
                        skill.SourceFilePath = skillFile
                        Try
                            skill.SourceLastWriteTime = File.GetLastWriteTime(skillFile)
                        Catch exTime As Exception
                            Utils.AppLogger.LogWarning("CustomSkill", $"读取自定义 Skill 文件时间失败: {exTime.Message}")
                        End Try
                    End If
                    If skill IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(skill.Name) Then
                        skill.IsCustom = True
                        _customSkills.Add(skill)
                        Utils.AppLogger.LogInfo("CustomSkill", $"✓ 技能: {skill.Name}")
                    Else
                        Utils.AppLogger.LogWarning("CustomSkill", $"✗ 解析失败: {Path.GetFileName(skillFile)}")
                    End If
                Catch ex As Exception
                    Utils.AppLogger.LogError("CustomSkill", $"✗ 加载异常: {Path.GetFileName(skillFile)}", ex)
                End Try
            Next

            ' 加载代码片段（snippets/ 目录）
            For Each snippetFile In Directory.GetFiles(_snippetsPath, "*.md")
                Try
                    Utils.AppLogger.LogDebug("CustomSkill", $"加载代码片段: {Path.GetFileName(snippetFile)}")
                    Dim skill = ParseMarkdownSkillFile(snippetFile)
                    If skill IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(skill.Name) Then
                        skill.IsCustom = True
                        skill.IsGeneric = False
                        _codeSnippets.Add(skill)
                        Utils.AppLogger.LogInfo("CustomSkill", $"✓ 代码片段: {skill.Name}")
                    Else
                        Utils.AppLogger.LogWarning("CustomSkill", $"✗ 代码片段解析失败: {Path.GetFileName(snippetFile)}")
                    End If
                Catch ex As Exception
                    Utils.AppLogger.LogError("CustomSkill", $"✗ 代码片段加载异常: {Path.GetFileName(snippetFile)}", ex)
                End Try
            Next
        End Sub
        
        ''' <summary>
        ''' 解析 Markdown Skill 文件
        ''' </summary>
        Private Function ParseMarkdownSkillFile(filePath As String) As SkillDefinition
            Dim raw = File.ReadAllText(filePath, System.Text.Encoding.UTF8)
            ' 去除 UTF-8 BOM + 头尾空白（BOM 是 \uFEFF，很多编辑器保存带 BOM 的 UTF-8，导致 content.StartsWith("---")=False 进而 frontmatter 完全读不到）
            Dim content = If(raw IsNot Nothing AndAlso raw.Length > 0 AndAlso AscW(raw(0)) = &HFEFF, raw.Substring(1), raw).Trim()
            If String.IsNullOrWhiteSpace(content) Then Return Nothing

            ' 提取 YAML frontmatter（健壮版）：
            ' 1) content 必须以 --- 开头（Trim 后）
            ' 2) 第二个 --- 必须在**行首**独立一行，而不是任意位置出现（否则 YAML 值里有 "---" 就会提前截断）
            Dim yamlText As String = ""
            Dim frontmatterClosedAt As Integer = -1
            If content.StartsWith("---", StringComparison.Ordinal) Then
                ' 从索引 3 开始逐行寻找，找第 2 个独立的 "---" 行
                Dim scanPos = 3
                While scanPos < content.Length
                    ' 跳到下一行开头
                    Dim nextLineStart As Integer = -1
                    Dim crIdx = content.IndexOf(vbCr, scanPos)
                    Dim lfIdx = content.IndexOf(vbLf, scanPos)
                    If crIdx < 0 AndAlso lfIdx < 0 Then Exit While
                    If crIdx >= 0 AndAlso (lfIdx < 0 OrElse crIdx < lfIdx) Then
                        ' CRLF or CR only
                        nextLineStart = If(crIdx + 1 < content.Length AndAlso content(crIdx + 1) = CChar(vbLf), crIdx + 2, crIdx + 1)
                    Else
                        nextLineStart = lfIdx + 1
                    End If
                    If nextLineStart < 0 OrElse nextLineStart >= content.Length Then Exit While

                    ' 判断这一行是否就是独立的 "---"（允许前后空格）
                    Dim thisLineEnd = -1
                    Dim cr2 = content.IndexOf(vbCr, nextLineStart)
                    Dim lf2 = content.IndexOf(vbLf, nextLineStart)
                    If cr2 < 0 AndAlso lf2 < 0 Then thisLineEnd = content.Length
                    If thisLineEnd < 0 Then
                        If cr2 >= 0 AndAlso (lf2 < 0 OrElse cr2 < lf2) Then thisLineEnd = cr2 Else thisLineEnd = lf2
                    End If
                    Dim line = content.Substring(nextLineStart, thisLineEnd - nextLineStart).Trim()
                    If line = "---" Then
                        frontmatterClosedAt = nextLineStart  ' 第二个 --- 的行首位置
                        yamlText = content.Substring(3, nextLineStart - 3).Trim()
                        Exit While
                    End If
                    scanPos = nextLineStart
                End While
                ' 如果上面逐行没找到，fallback 到简单 IndexOf（兼容老格式）
                If String.IsNullOrWhiteSpace(yamlText) Then
                    Dim endIndex = content.IndexOf("---", 3, StringComparison.Ordinal)
                    If endIndex > 0 Then
                        yamlText = content.Substring(3, endIndex - 3).Trim()
                        frontmatterClosedAt = endIndex
                    End If
                End If
            End If

            If String.IsNullOrWhiteSpace(yamlText) Then
                Utils.AppLogger.LogWarning("CustomSkill", $"YAML frontmatter 为空/无法识别，跳过此文件: {Path.GetFileName(filePath)}")
                Return Nothing
            End If

            ' 使用 YamlDotNet 解析 YAML → JObject
            Dim yamlObj = ConvertYamlToJObject(yamlText)
            If yamlObj Is Nothing OrElse yamlObj.Count = 0 Then
                Utils.AppLogger.LogWarning("CustomSkill", "YAML解析结果为空对象，跳过此文件")
                Return Nothing
            End If

            ' 使用 SkillJsonHelper 反射解析所有字段（不区分大小写）
            Dim skill = SkillJsonHelper.ParseSkillFromJObject(yamlObj)
            If skill Is Nothing Then Return Nothing

            If String.IsNullOrWhiteSpace(skill.Name) OrElse skill.Name = "UnnamedSkill" Then
                Utils.AppLogger.LogWarning("CustomSkill", "解析后Name为空或UnnamedSkill，YAML可能解析不完整，跳过")
                Return Nothing
            End If

            ' === 提取 Markdown 正文 ===
            Dim markdownBody = ""
            Dim bodyStart = -1
            If frontmatterClosedAt > 0 Then
                ' 用新逻辑找到的第二个 --- 行首
                bodyStart = frontmatterClosedAt
            Else
                ' fallback：老逻辑（兼容）
                Dim frontmatterEnd = content.IndexOf("---", 3, StringComparison.Ordinal)
                If frontmatterEnd > 0 Then bodyStart = frontmatterEnd
            End If
            If bodyStart > 0 Then
                ' bodyStart 指向第二个 --- 所在行的行首或 --- 起点；要跳过这一行（跳过 "---" + 换行后的内容）
                Dim rest = content.Substring(bodyStart)
                Dim eolCr = rest.IndexOf(vbCr)
                Dim eolLf = rest.IndexOf(vbLf)
                Dim skipUntil As Integer = 3  ' 至少跳过 "---" 本身
                If eolCr >= 0 AndAlso (eolLf < 0 OrElse eolCr < eolLf) Then
                    skipUntil = If(eolCr + 2 <= rest.Length AndAlso rest(eolCr + 1) = CChar(vbLf), eolCr + 2, eolCr + 1)
                ElseIf eolLf >= 0 Then
                    skipUntil = eolLf + 1
                End If
                markdownBody = If(skipUntil < rest.Length, rest.Substring(skipUntil).Trim(), "")
            End If

            ' 提取 System Prompt 部分（优先从 Markdown 正文）
            Dim sysPrompt = ExtractMarkdownSection(markdownBody, "System Prompt")
            If Not String.IsNullOrEmpty(sysPrompt) Then
                skill.LLMSystemPrompt = sysPrompt
            End If

            ' 提取 User Prompt Template 部分
            Dim userPrompt = ExtractMarkdownSection(markdownBody, "User Prompt Template")
            If Not String.IsNullOrEmpty(userPrompt) Then
                skill.LLMUserPromptTemplate = userPrompt
            End If

            ' 提取 Template Code 部分（VBA代码）
            Dim templateCode = ExtractMarkdownCodeBlock(markdownBody, "Template Code")
            If Not String.IsNullOrEmpty(templateCode) Then
                skill.TemplateCode = templateCode
            End If

            ' 提取 Input Description 部分
            Dim inputDesc = ExtractMarkdownSection(markdownBody, "Input Description")
            If Not String.IsNullOrEmpty(inputDesc) Then
                skill.InputDescription = inputDesc
            End If

            ' 提取 Output Description 部分
            Dim outputDesc = ExtractMarkdownSection(markdownBody, "Output Description")
            If Not String.IsNullOrEmpty(outputDesc) Then
                skill.OutputDescription = outputDesc
            End If

            ' 确保默认值
            If String.IsNullOrWhiteSpace(skill.Name) Then skill.Name = "UnnamedSkill"
            If String.IsNullOrWhiteSpace(skill.DisplayName) Then skill.DisplayName = skill.Name
            If String.IsNullOrWhiteSpace(skill.OutputFormat) Then skill.OutputFormat = "text"
            If skill.Version <= 0 Then skill.Version = 1

            Utils.AppLogger.LogInfo("CustomSkill", $"解析完成 - Name:{skill.Name}, 字段数:{yamlObj.Count}")
            Return skill
        End Function
        
        ''' <summary>
        ''' 将字符串转换为 ProcessingMode
        ''' </summary>
        Private Function GetProcessingMode(modeStr As String) As SkillProcessingMode
            Select Case modeStr.ToLower()
                Case "vbatchcode", "vbacode"
                    Return SkillProcessingMode.VBACode
                Case "multimodalllm"
                    Return SkillProcessingMode.MultimodalLLM
                Case "hybrid"
                    Return SkillProcessingMode.Hybrid
                Case "localocr"
                    Return SkillProcessingMode.LocalOCR
                Case "pythoncode"
                    Return SkillProcessingMode.PythonCode
                Case Else
                    Return SkillProcessingMode.MultimodalLLM
            End Select
        End Function

        ''' <summary>
        ''' 安全解析布尔值（不区分大小写，兼容 Newtonsoft.Json JValue.ToString() 返回的小写格式）
        ''' </summary>
        Private Function SafeParseBool(strValue As String) As Boolean
            If String.IsNullOrEmpty(strValue) Then Return False
            Dim trimmed = strValue.Trim()
            If trimmed.Equals("true", StringComparison.OrdinalIgnoreCase) Then Return True
            If trimmed.Equals("false", StringComparison.OrdinalIgnoreCase) Then Return False
            Return False
        End Function

        ''' <summary>
        ''' 安全解析整数值
        ''' </summary>
        Private Function SafeParseInt(strValue As String) As Integer
            If String.IsNullOrEmpty(strValue) Then Return 1
            Dim result As Integer
            If Integer.TryParse(strValue.Trim(), result) Then Return result
            Return 1
        End Function
        
        ''' <summary>
        ''' 将 YAML 文本转换为 JObject
        ''' </summary>
        Private Function ConvertYamlToJObject(yamlText As String) As JObject
            Try
                If String.IsNullOrWhiteSpace(yamlText) Then Return Nothing

                Dim yamlStream = New YamlStream()
                yamlStream.Load(New System.IO.StringReader(yamlText))

                If yamlStream.Documents.Count = 0 Then 
                    Utils.AppLogger.LogWarning("CustomSkill", "YAML文档数为0")
                    Return Nothing
                End If

                Dim root = yamlStream.Documents(0).RootNode
                Utils.AppLogger.LogInfo("CustomSkill", $"RootNode类型:{root.GetType().Name}")
                
                Dim jsonObj = ConvertYamlNodeToJToken(root)
                Utils.AppLogger.LogInfo("CustomSkill", $"jsonObj类型:{jsonObj?.GetType().Name}")

                If TypeOf jsonObj Is JObject Then
                    Dim result = CType(jsonObj, JObject)
                    If result.Count = 0 Then
                        Utils.AppLogger.LogWarning("CustomSkill", "YAML解析结果为空JObject")
                    End If
                    Return result
                End If

                Utils.AppLogger.LogWarning("CustomSkill", $"YAML根节点不是MappingNode: {jsonObj?.GetType().Name}")
                Return Nothing
            Catch ex As Exception
                Utils.AppLogger.LogError("CustomSkill", $"YAML解析失败: {ex.GetType().Name} - {ex.Message}", ex)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 递归将 YamlDotNet 节点转换为 JToken
        ''' </summary>
        Private Function ConvertYamlNodeToJToken(node As YamlNode) As JToken
            If TypeOf node Is YamlMappingNode Then
                Dim mapping = CType(node, YamlMappingNode)
                Dim obj As New JObject()
                For Each kvp In mapping.Children
                    Dim key = kvp.Key.ToString()
                    obj(key) = ConvertYamlNodeToJToken(kvp.Value)
                Next
                Return obj
            ElseIf TypeOf node Is YamlSequenceNode Then
                Dim sequence = CType(node, YamlSequenceNode)
                Dim arr As New JArray()
                For Each item In sequence.Children
                    arr.Add(ConvertYamlNodeToJToken(item))
                Next
                Return arr
            ElseIf TypeOf node Is YamlScalarNode Then
                Dim scalar = CType(node, YamlScalarNode)
                Dim value = scalar.Value

                If value IsNot Nothing Then
                    Dim lowerValue = value.ToLower()
                    If lowerValue = "true" Then Return New JValue(True)
                    If lowerValue = "false" Then Return New JValue(False)
                    If lowerValue = "null" OrElse lowerValue = "~" OrElse lowerValue = "" Then Return JValue.CreateNull()

                    Dim intResult As Integer
                    If Integer.TryParse(value, intResult) Then Return New JValue(intResult)

                    Dim dblResult As Double
                    If Double.TryParse(value, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, dblResult) Then Return New JValue(dblResult)
                End If

                Return New JValue(value)
            End If

            Return JValue.CreateNull()
        End Function

        ''' <summary>
        ''' 从 Markdown 正文中提取指定二级标题下的文本内容
        ''' </summary>
        Private Function ExtractMarkdownSection(markdown As String, sectionTitle As String) As String
            If String.IsNullOrWhiteSpace(markdown) Then Return ""
            If String.IsNullOrWhiteSpace(sectionTitle) Then Return ""

            ' 按行扫描：找独立的一行以 "##" 开头（二级标题），标题文本（去掉首尾空白）等于 sectionTitle
            ' 这比旧版的 "## title + CRLF/LF 严格拼接" 更健壮，能处理：
            '   - 标题后尾随空格
            '   - ## 和 title 之间多个空格 / Tab
            '   - 行尾有 \r\n / \n / \r
            '   - sectionTitle 包含空格 (如 "Template Code", "User Prompt Template")
            Dim target = sectionTitle.Trim()
            Dim lines = markdown.Split(New String() {vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
            Dim headingStartLine As Integer = -1
            For i = 0 To lines.Length - 1
                Dim line = lines(i)
                If String.IsNullOrWhiteSpace(line) Then Continue For
                ' 必须以 ## 开头（不是 ### # 等其他级别）
                If Not line.StartsWith("##") Then Continue For
                ' 第 3 个字符必须是空格或 Tab（确认是二级标题，不是 "##Xyz"）
                If line.Length <= 2 Then Continue For
                Dim c3 = line(2)
                If c3 <> " "c AndAlso c3 <> ChrW(9) Then Continue For
                Dim titlePart = line.Substring(2).Trim()  ' 去掉 ## 以及之后的空格/Tab/行尾
                ' 末尾 "#"（如 "## Template Code ###" 带权重标记）也允许：先去尾 # 再 Trim
                While titlePart.Length > 0 AndAlso titlePart(titlePart.Length - 1) = "#"c
                    titlePart = titlePart.Substring(0, titlePart.Length - 1).TrimEnd()
                End While
                If String.Equals(titlePart, target, StringComparison.OrdinalIgnoreCase) Then
                    headingStartLine = i
                    Exit For
                End If
            Next
            If headingStartLine < 0 Then Return ""

            ' 收集从下一行直到下一个二级标题(##)之前的内容
            Dim sb As New Text.StringBuilder()
            For i = headingStartLine + 1 To lines.Length - 1
                Dim line = lines(i)
                ' 判断是否下一个二级标题：非空 + 以 ## 开头 + 第3字符是空格/Tab（不是 ###）
                If Not String.IsNullOrWhiteSpace(line) AndAlso
                    line.StartsWith("##") AndAlso
                    line.Length > 2 AndAlso
                    (line(2) = " "c OrElse line(2) = ChrW(9)) Then
                    Exit For
                End If
                sb.AppendLine(line)
            Next
            Return sb.ToString().Trim()
        End Function

        ''' <summary>
        ''' 从 Markdown 正文中提取指定二级标题下的代码块内容
        ''' </summary>
        Private Function ExtractMarkdownCodeBlock(markdown As String, sectionTitle As String) As String
            Dim sectionText = ExtractMarkdownSection(markdown, sectionTitle)
            If String.IsNullOrWhiteSpace(sectionText) Then Return ""

            Dim codeStart = sectionText.IndexOf("```")
            If codeStart < 0 Then Return sectionText.Trim()

            codeStart = codeStart + 3
            ' 跳过语言标识行（如 vba），兼容 CRLF 和 LF
            Dim langEndCrLf = sectionText.IndexOf(vbCrLf, codeStart)
            Dim langEndLf = sectionText.IndexOf(vbLf, codeStart)
            Dim langEnd As Integer = -1
            Dim newlineOffset As Integer = 1
            If langEndCrLf >= 0 AndAlso (langEndLf < 0 OrElse langEndCrLf <= langEndLf) Then
                langEnd = langEndCrLf
                newlineOffset = 2
            ElseIf langEndLf >= 0 Then
                langEnd = langEndLf
                newlineOffset = 1
            End If
            If langEnd > codeStart AndAlso langEnd - codeStart < 10 Then
                codeStart = langEnd + newlineOffset
            End If

            Dim codeEnd = sectionText.IndexOf("```", codeStart)
            If codeEnd < 0 Then Return sectionText.Substring(codeStart).Trim()

            Return sectionText.Substring(codeStart, codeEnd - codeStart).Trim()
        End Function

        ''' <summary>
        ''' 获取所有自定义Skill
        ''' </summary>
        Public Function GetAllCustomSkills() As List(Of SkillDefinition)
            Dim all = New List(Of SkillDefinition)(_customSkills)
            all.AddRange(_codeSnippets)
            Return all
        End Function

        ''' <summary>
        ''' 获取所有代码片段
        ''' </summary>
        Public Function GetAllCodeSnippets() As List(Of SkillDefinition)
            Return _codeSnippets.ToList()
        End Function
        
        ''' <summary>
        ''' 通过名称获取自定义Skill
        ''' （不区分大小写匹配，防止前端传 camelCase 名称找不到）
        ''' </summary>
        Public Function GetCustomSkill(name As String) As SkillDefinition
            Dim skill = _customSkills.FirstOrDefault(Function(s) String.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase))
            If skill Is Nothing Then
                skill = _codeSnippets.FirstOrDefault(Function(s) String.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase))
            End If
            Return skill
        End Function

        ''' <summary>
        ''' 判断是否是代码片段
        ''' </summary>
        Private Function IsCodeSnippet(name As String) As Boolean
            Return _codeSnippets.Any(Function(s) String.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase))
        End Function
        
        ''' <summary>
        ''' 保存自定义Skill
        ''' </summary>
        Public Sub SaveCustomSkill(skill As SkillDefinition)
            System.Diagnostics.Debug.WriteLine($"[CustomSkillManager.SaveCustomSkill] ⭐ ENTRY: name='{skill.Name}', IsEnabled={skill.IsEnabled}, IsQuickShortcut={skill.IsQuickShortcut}, IsCustom={skill.IsCustom}")
            If String.IsNullOrEmpty(skill.Name) Then
                Throw New ArgumentException("Skill名称不能为空")
            End If

            ' 判断是技能还是代码片段（有 TemplateCode 且无 LLMSystemPrompt = 代码片段）
            ' 代码片段存入 snippets/ 目录，技能存入 skills/ 目录（通过目录归类，无需名字前缀）
            Dim isSnippet = Not String.IsNullOrEmpty(skill.TemplateCode) AndAlso 
                            String.IsNullOrEmpty(skill.LLMSystemPrompt)
            Dim targetDir = If(isSnippet, _snippetsPath, _customSkillsPath)
            System.Diagnostics.Debug.WriteLine($"[CustomSkillManager.SaveCustomSkill] targetDir={targetDir}, isSnippet={isSnippet}")

            ' 确保目录存在
            Utils.PathHelper.EnsureDirectoryExists(targetDir)

            ' 更新或添加到对应列表
            Dim skillList = If(isSnippet, _codeSnippets, _customSkills)
            Dim existing = skillList.FirstOrDefault(Function(s) String.Equals(s.Name, skill.Name, StringComparison.OrdinalIgnoreCase))
            If existing IsNot Nothing Then
                skillList.Remove(existing)
                System.Diagnostics.Debug.WriteLine("[CustomSkillManager.SaveCustomSkill] 移除旧的内存副本")
            End If

            skill.IsCustom = True
            skill.LastModified = DateTime.Now
            skillList.Add(skill)

            ' 保存到对应目录
            Dim filePath = Path.Combine(targetDir, $"{skill.Name}.md")
            System.Diagnostics.Debug.WriteLine($"[CustomSkillManager.SaveCustomSkill] 准备写文件: {filePath}")
            SaveAsMarkdownFile(skill, filePath)
            System.Diagnostics.Debug.WriteLine($"[CustomSkillManager.SaveCustomSkill] ✅ SaveAsMarkdownFile 返回，文件存在={File.Exists(filePath)}")
            
            ' 验证文件内容里是否真的有 isEnabled
            If File.Exists(filePath) Then
                Dim verify = File.ReadAllText(filePath)
                Dim hasEnabled = verify.Contains("isEnabled")
                Dim enabledVal = If(hasEnabled, If(verify.Contains("isEnabled: false"), "false", "true"), "MISSING")
                System.Diagnostics.Debug.WriteLine($"[CustomSkillManager.SaveCustomSkill] 📄 文件验证: hasIsEnabled={hasEnabled}, value={enabledVal}")
            End If

            Utils.AppLogger.LogInfo("CustomSkill", $"已保存({If(isSnippet, "代码片段", "技能")}): {Path.GetFileName(filePath)}")
        End Sub
        
        ''' <summary>
        ''' 将Skill保存为Markdown文件
        ''' 【修复】使用反射自动写入所有简单类型字段，不遗漏 SkillDefinition 的任何属性
        ''' </summary>
        Private Sub SaveAsMarkdownFile(skill As SkillDefinition, filePath As String)
            Dim yamlBuilder As New System.Text.StringBuilder()
            yamlBuilder.AppendLine("---")

            ' 已手动处理的字段（数组/对象/特殊格式），避免重复
            Dim handledFields As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            handledFields.Add("Name")
            handledFields.Add("DisplayName")
            handledFields.Add("Description")
            handledFields.Add("Category")
            handledFields.Add("Keywords")
            handledFields.Add("Examples")
            handledFields.Add("Tags")
            handledFields.Add("InputContext")
            handledFields.Add("AllowedTools")
            handledFields.Add("UIActions")
            handledFields.Add("LLMSystemPrompt")
            handledFields.Add("TemplateCode")
            handledFields.Add("Pipeline")

            ' === 1. 手动写入核心字段 ===
            yamlBuilder.AppendLine($"name: {skill.Name}")
            yamlBuilder.AppendLine($"displayName: {skill.DisplayName}")
            yamlBuilder.AppendLine($"description: {EscapeYamlString(skill.Description)}")
            If Not String.IsNullOrEmpty(skill.Category) Then
                yamlBuilder.AppendLine($"category: {EscapeYamlString(skill.Category)}")
            End If

            ' === 2. 数组字段 ===
            If skill.Keywords IsNot Nothing AndAlso skill.Keywords.Length > 0 Then
                yamlBuilder.AppendLine("keywords:")
                For Each kw In skill.Keywords
                    yamlBuilder.AppendLine($"  - {EscapeYamlString(kw)}")
                Next
            End If

            If skill.Examples IsNot Nothing AndAlso skill.Examples.Length > 0 Then
                yamlBuilder.AppendLine("examples:")
                For Each ex In skill.Examples
                    yamlBuilder.AppendLine($"  - {EscapeYamlString(ex)}")
                Next
            End If

            If skill.Tags IsNot Nothing AndAlso skill.Tags.Count > 0 Then
                yamlBuilder.AppendLine("tags:")
                For Each tag In skill.Tags
                    yamlBuilder.AppendLine($"  - {EscapeYamlString(tag)}")
                Next
            End If

            If skill.InputContext IsNot Nothing AndAlso skill.InputContext.Length > 0 Then
                yamlBuilder.AppendLine("inputContext:")
                For Each ctx In skill.InputContext
                    yamlBuilder.AppendLine($"  - {EscapeYamlString(ctx)}")
                Next
            End If

            ' === 3. 布尔/数值/枚举字段（反射自动处理所有简单类型属性）===
            Dim simpleProps = GetType(SkillDefinition).GetProperties(
                System.Reflection.BindingFlags.Public Or
                System.Reflection.BindingFlags.Instance)

            For Each prop In simpleProps
                If handledFields.Contains(prop.Name) Then Continue For

                Dim propType = prop.PropertyType

                ' 处理可空类型
                Dim underlyingType = Nullable.GetUnderlyingType(propType)
                If underlyingType IsNot Nothing Then propType = underlyingType

                Dim value = prop.GetValue(skill)

                ' 跳过 null/Nothing
                If value Is Nothing Then Continue For

                Select Case propType
                    Case GetType(String)
                        Dim strVal = DirectCast(value, String)
                        If Not String.IsNullOrEmpty(strVal) Then
                            yamlBuilder.AppendLine($"{ToYamlKey(prop.Name)}: {EscapeYamlString(strVal)}")
                        End If

                    Case GetType(Boolean)
                        Dim boolVal = Convert.ToBoolean(value)
                        ' 跳过默认值 false（减少噪音），但关键开关字段必须始终显式写入
                        Dim forceWrite = boolVal = True OrElse prop.Name = "IsCustom" OrElse prop.Name = "IsGeneric" _
                           OrElse prop.Name = "IsComposite" OrElse prop.Name = "NeedsConfig" _
                           OrElse prop.Name = "SupportsImageInput" OrElse prop.Name = "SupportsFileInput" _
                           OrElse prop.Name = "IsQuickShortcut" OrElse prop.Name = "IsEnabled"
                        If prop.Name = "IsEnabled" Then
                            System.Diagnostics.Debug.WriteLine($"[SaveAsMarkdownFile] 🎯 IsEnabled REFLECT: raw={boolVal}, forceWrite={forceWrite}, ToYamlKey={ToYamlKey(prop.Name)}")
                        End If
                        If forceWrite Then
                            yamlBuilder.AppendLine($"{ToYamlKey(prop.Name)}: {boolVal.ToString().ToLower()}")
                            If prop.Name = "IsEnabled" Then
                                System.Diagnostics.Debug.WriteLine($"[SaveAsMarkdownFile] ✅ IsEnabled YAML LINE APPENDED: isEnabled={boolVal.ToString().ToLower()}")
                            End If
                        Else
                            If prop.Name = "IsEnabled" Then
                                System.Diagnostics.Debug.WriteLine("[SaveAsMarkdownFile] ❌ IsEnabled SKIPPED (forceWrite=False) - BUG!")
                            End If
                        End If

                    Case GetType(Integer), GetType(Long), GetType(Short)
                        Dim intVal = Convert.ToInt32(value)
                        If intVal <> 0 OrElse prop.Name = "Version" Then
                            yamlBuilder.AppendLine($"{ToYamlKey(prop.Name)}: {intVal}")
                        End If

                    Case GetType(Double), GetType(Decimal), GetType(Single)
                        Dim dblVal = Convert.ToDouble(value)
                        If dblVal <> 0 Then
                            yamlBuilder.AppendLine($"{ToYamlKey(prop.Name)}: {dblVal}")
                        End If

                    Case GetType(DateTime)
                        Dim dtVal = Convert.ToDateTime(value)
                        If dtVal <> DateTime.MinValue Then
                            yamlBuilder.AppendLine($"{ToYamlKey(prop.Name)}: {dtVal:yyyy-MM-ddTHH:mm:ss}")
                        End If

                    Case Else
                        ' 处理枚举类型
                        If propType.IsEnum Then
                            yamlBuilder.AppendLine($"{ToYamlKey(prop.Name)}: {value.ToString()}")
                        End If
                        ' 其他复杂类型（List, Dictionary 等）由下面的特定处理负责
                End Select
            Next

            ' === 4. 复杂对象的特定处理 ===
            If skill.UIActions IsNot Nothing Then
                yamlBuilder.AppendLine("UIActions:")
                yamlBuilder.AppendLine($"  ShowCopy: {skill.UIActions.ShowCopy.ToString().ToLower()}")
                yamlBuilder.AppendLine($"  ShowSaveMd: {skill.UIActions.ShowSaveMd.ToString().ToLower()}")
                yamlBuilder.AppendLine($"  ShowRegenerate: {skill.UIActions.ShowRegenerate.ToString().ToLower()}")
            End If

            ' Pipeline 路由声明
            If Not String.IsNullOrWhiteSpace(skill.Pipeline) Then
                yamlBuilder.AppendLine($"pipeline: {EscapeYamlString(skill.Pipeline)}")
            End If

            ' === 5. 复杂列表类型序列化（JSON-in-YAML 格式）===
            If skill.AllowedTools IsNot Nothing AndAlso skill.AllowedTools.Length > 0 Then
                yamlBuilder.AppendLine("allowedTools:")
                For Each at In skill.AllowedTools
                    yamlBuilder.AppendLine($"  - {EscapeYamlString(at)}")
                Next
            End If

            If skill.ConfigFields IsNot Nothing AndAlso skill.ConfigFields.Length > 0 Then
                yamlBuilder.AppendLine("configFields:")
                For Each cf In skill.ConfigFields
                    yamlBuilder.AppendLine($"  - {EscapeYamlString(cf)}")
                Next
            End If

            If skill.InputParameters IsNot Nothing AndAlso skill.InputParameters.Count > 0 Then
                yamlBuilder.AppendLine("inputParameters:")
                For Each ip In skill.InputParameters
                    yamlBuilder.AppendLine($"  - name: {EscapeYamlString(ip.Name)}")
                    yamlBuilder.AppendLine($"    dataType: {EscapeYamlString(ip.DataType)}")
                    yamlBuilder.AppendLine($"    description: {EscapeYamlString(ip.Description)}")
                    yamlBuilder.AppendLine($"    isRequired: {ip.IsRequired.ToString().ToLower()}")
                    If Not String.IsNullOrEmpty(ip.DefaultValue) Then
                        yamlBuilder.AppendLine($"    defaultValue: {EscapeYamlString(ip.DefaultValue)}")
                    End If
                Next
            End If

            If skill.OutputParameters IsNot Nothing AndAlso skill.OutputParameters.Count > 0 Then
                yamlBuilder.AppendLine("outputParameters:")
                For Each op In skill.OutputParameters
                    yamlBuilder.AppendLine($"  - name: {EscapeYamlString(op.Name)}")
                    yamlBuilder.AppendLine($"    dataType: {EscapeYamlString(op.DataType)}")
                    yamlBuilder.AppendLine($"    description: {EscapeYamlString(op.Description)}")
                Next
            End If

            If skill.WorkflowSteps IsNot Nothing AndAlso skill.WorkflowSteps.Count > 0 Then
                yamlBuilder.AppendLine("workflowSteps:")
                For Each ws In skill.WorkflowSteps
                    yamlBuilder.AppendLine($"  - stepId: {EscapeYamlString(ws.StepId)}")
                    yamlBuilder.AppendLine($"    stepName: {EscapeYamlString(ws.StepName)}")
                    yamlBuilder.AppendLine($"    stepType: {ws.StepType.ToString()}")
                    If Not String.IsNullOrEmpty(ws.Code) Then
                        yamlBuilder.AppendLine($"    code: {EscapeYamlString(ws.Code)}")
                    End If
                    If Not String.IsNullOrEmpty(ws.ApiEndpoint) Then
                        yamlBuilder.AppendLine($"    apiEndpoint: {EscapeYamlString(ws.ApiEndpoint)}")
                    End If
                Next
            End If

            If skill.CustomWorkflowSteps IsNot Nothing AndAlso skill.CustomWorkflowSteps.Count > 0 Then
                yamlBuilder.AppendLine("customWorkflowSteps:")
                For Each cws In skill.CustomWorkflowSteps
                    yamlBuilder.AppendLine($"  - stepId: {EscapeYamlString(cws.StepId)}")
                    yamlBuilder.AppendLine($"    stepName: {EscapeYamlString(cws.StepName)}")
                    yamlBuilder.AppendLine($"    stepType: {cws.StepType.ToString()}")
                Next
            End If

            If skill.FieldMappings IsNot Nothing AndAlso skill.FieldMappings.Count > 0 Then
                yamlBuilder.AppendLine("fieldMappings:")
                For Each fm In skill.FieldMappings
                    yamlBuilder.AppendLine($"  - sourceField: {EscapeYamlString(fm.SourceField)}")
                    yamlBuilder.AppendLine($"    targetField: {EscapeYamlString(fm.TargetField)}")
                    If Not String.IsNullOrEmpty(fm.TargetCell) Then
                        yamlBuilder.AppendLine($"    targetCell: {EscapeYamlString(fm.TargetCell)}")
                    End If
                    If Not String.IsNullOrEmpty(fm.Format) Then
                        yamlBuilder.AppendLine($"    format: {EscapeYamlString(fm.Format)}")
                    End If
                Next
            End If

            ' 序列化 specFields（SkillSpec 参数声明）
            If skill.SpecFields IsNot Nothing AndAlso skill.SpecFields.Count > 0 Then
                yamlBuilder.AppendLine("specFields:")
                For Each sf In skill.SpecFields
                    yamlBuilder.AppendLine($"  - name: {EscapeYamlString(sf.Name)}")
                    If Not String.IsNullOrWhiteSpace(sf.Label) Then
                        yamlBuilder.AppendLine($"    label: {EscapeYamlString(sf.Label)}")
                    End If
                    If Not String.IsNullOrWhiteSpace(sf.Type) Then
                        yamlBuilder.AppendLine($"    type: {EscapeYamlString(sf.Type)}")
                    End If
                    If sf.Options IsNot Nothing AndAlso sf.Options.Count > 0 Then
                        yamlBuilder.AppendLine("    options:")
                        For Each opt In sf.Options
                            yamlBuilder.AppendLine($"      - {EscapeYamlString(opt)}")
                        Next
                    End If
                    If Not String.IsNullOrWhiteSpace(sf.DefaultValue) Then
                        yamlBuilder.AppendLine($"    defaultValue: {EscapeYamlString(sf.DefaultValue)}")
                    End If
                Next
            End If

            ' 序列化 quickPrompts
            If skill.QuickPrompts IsNot Nothing AndAlso skill.QuickPrompts.Count > 0 Then
                yamlBuilder.AppendLine("quickPrompts:")
                For Each qp In skill.QuickPrompts
                    yamlBuilder.AppendLine($"  - label: {EscapeYamlString(qp.Label)}")
                    yamlBuilder.AppendLine($"    prompt: {EscapeYamlString(qp.Prompt)}")
                Next
            End If

            ' 序列化声明式步骤（steps，SkillCreator 默认产物）。
            ' params 以单行 JSON 标量输出，配合上面 ParseSkillFromJObject 的兼容解析，
            ' 保证含任意嵌套的参数对象无损往返（嵌套 YAML 手写不可靠）。
            If skill.Steps IsNot Nothing AndAlso skill.Steps.Count > 0 Then
                yamlBuilder.AppendLine("steps:")
                For Each ds In skill.Steps
                    yamlBuilder.AppendLine($"  - name: {EscapeYamlString(If(ds.Name, ""))}")
                    yamlBuilder.AppendLine($"    tool: {EscapeYamlString(If(ds.Tool, ""))}")
                    yamlBuilder.AppendLine($"    strategy: {EscapeYamlString(If(String.IsNullOrWhiteSpace(ds.Strategy), "direct", ds.Strategy))}")
                    If Not String.IsNullOrWhiteSpace(ds.Description) Then
                        yamlBuilder.AppendLine($"    description: {EscapeYamlString(ds.Description)}")
                    End If
                    If ds.Params IsNot Nothing AndAlso ds.Params.Count > 0 Then
                        yamlBuilder.AppendLine($"    params: {EscapeYamlString(Newtonsoft.Json.JsonConvert.SerializeObject(ds.Params))}")
                    End If
                    If Not String.IsNullOrWhiteSpace(ds.Code) Then
                        ' code 为 VBA，多行 → JSON 转义单行输出，保证换行无损保留
                        yamlBuilder.AppendLine($"    code: {EscapeYamlString(Newtonsoft.Json.JsonConvert.SerializeObject(ds.Code))}")
                    End If
                    If ds.WaitOn Then
                        yamlBuilder.AppendLine("    waitOn: true")
                    End If
                Next
            End If

            If skill.ExcelConfig IsNot Nothing Then
                Dim ecJson = Newtonsoft.Json.JsonConvert.SerializeObject(skill.ExcelConfig)
                yamlBuilder.AppendLine($"excelConfig: {EscapeYamlString(ecJson)}")
            End If

            If skill.OcrConfig IsNot Nothing Then
                Dim ocrJson = Newtonsoft.Json.JsonConvert.SerializeObject(skill.OcrConfig)
                yamlBuilder.AppendLine($"ocrConfig: {EscapeYamlString(ocrJson)}")
            End If

            ' === 5. 结束 YAML frontmatter ===
            yamlBuilder.AppendLine("---")
            yamlBuilder.AppendLine()

            ' === 6. Markdown 正文 ===
            yamlBuilder.AppendLine("# " & If(String.IsNullOrWhiteSpace(skill.DisplayName), skill.Name, skill.DisplayName))
            yamlBuilder.AppendLine()

            ' System Prompt 部分
            If Not String.IsNullOrWhiteSpace(skill.LLMSystemPrompt) Then
                yamlBuilder.AppendLine("## System Prompt")
                yamlBuilder.AppendLine()
                yamlBuilder.AppendLine(skill.LLMSystemPrompt)
                yamlBuilder.AppendLine()
            End If

            ' LLM User Prompt Template
            If Not String.IsNullOrWhiteSpace(skill.LLMUserPromptTemplate) Then
                yamlBuilder.AppendLine("## User Prompt Template")
                yamlBuilder.AppendLine()
                yamlBuilder.AppendLine(skill.LLMUserPromptTemplate)
                yamlBuilder.AppendLine()
            End If

            ' Template Code
            If Not String.IsNullOrWhiteSpace(skill.TemplateCode) Then
                yamlBuilder.AppendLine("## Template Code")
                yamlBuilder.AppendLine()
                yamlBuilder.AppendLine("```vba")
                yamlBuilder.AppendLine(skill.TemplateCode)
                yamlBuilder.AppendLine("```")
                yamlBuilder.AppendLine()
            End If

            ' Input Description
            If Not String.IsNullOrWhiteSpace(skill.InputDescription) Then
                yamlBuilder.AppendLine("## Input Description")
                yamlBuilder.AppendLine()
                yamlBuilder.AppendLine(skill.InputDescription)
                yamlBuilder.AppendLine()
            End If

            ' Output Description
            If Not String.IsNullOrWhiteSpace(skill.OutputDescription) Then
                yamlBuilder.AppendLine("## Output Description")
                yamlBuilder.AppendLine()
                yamlBuilder.AppendLine(skill.OutputDescription)
                yamlBuilder.AppendLine()
            End If

            System.IO.File.WriteAllText(filePath, yamlBuilder.ToString(), System.Text.Encoding.UTF8)
        End Sub

        ''' <summary>
        ''' 将 PascalCase 属性名转为 YAML 友好的 kebab-case 或保持原样
        ''' </summary>
        Private Shared Function ToYamlKey(propertyName As String) As String
            ' 保持 PascalCase，因为 YAML 支持
            Return propertyName.Substring(0, 1).ToLower() & propertyName.Substring(1)
        End Function

        ''' <summary>
        ''' 转义 YAML 字符串值（处理特殊字符）
        ''' </summary>
        Private Shared Function EscapeYamlString(value As String) As String
            If String.IsNullOrEmpty(value) Then Return ""
            If value.Contains(vbLf) OrElse value.Contains(vbCr) Then
                Dim folded = value.Replace(vbCrLf, " ").Replace(vbLf, " ").Replace(vbCr, " ")
                Return """" & folded.Replace("""", "\""") & """"
            End If
            Return """" & value.Replace("""", "\""") & """"
        End Function
        
        ''' <summary>
        ''' 删除自定义Skill
        ''' （不区分大小写匹配；同时扫 skills/ 与 snippets/ 两个目录，确保磁盘无残留）
        ''' </summary>
        Public Function DeleteCustomSkill(name As String) As Boolean
            Dim deletedFromMemory As Boolean = False
            Dim isSnippet As Boolean = False

            ' 从技能列表删除
            Dim skill = _customSkills.FirstOrDefault(Function(s) String.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase))
            If skill IsNot Nothing Then
                _customSkills.Remove(skill)
                deletedFromMemory = True
            Else
                ' 从代码片段列表删除
                skill = _codeSnippets.FirstOrDefault(Function(s) String.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase))
                If skill IsNot Nothing Then
                    _codeSnippets.Remove(skill)
                    isSnippet = True
                    deletedFromMemory = True
                End If
            End If

            Dim deletedFromDisk As Boolean = False
            Dim targetDir = If(isSnippet, _snippetsPath, _customSkillsPath)

            ' 无论内存中是否找到，都尝试从磁盘的两个目录删除文件（以防内存未成功加载但磁盘有残留）
            For Each dirPath In {_customSkillsPath, _snippetsPath}
                Try
                    Dim mdPath = Path.Combine(dirPath, $"{name}.md")
                    If File.Exists(mdPath) Then
                        File.Delete(mdPath)
                        deletedFromDisk = True
                        Utils.AppLogger.LogInfo("CustomSkill", $"已从磁盘删除MD文件: {mdPath}")
                    End If

                    Dim jsonPath = Path.Combine(dirPath, $"{name}.json")
                    If File.Exists(jsonPath) Then
                        File.Delete(jsonPath)
                        deletedFromDisk = True
                        Utils.AppLogger.LogInfo("CustomSkill", $"已从磁盘删除JSON文件: {jsonPath}")
                    End If
                Catch ex As Exception
                    Utils.AppLogger.LogError("CustomSkill", $"删除文件失败({dirPath})", ex)
                End Try
            Next

            If deletedFromMemory OrElse deletedFromDisk Then
                Utils.AppLogger.LogInfo("CustomSkill", $"删除完成: {name} (内存={deletedFromMemory}, 磁盘={deletedFromDisk})")
                Return True
            End If

            Utils.AppLogger.LogWarning("CustomSkill", $"未找到可删除的Skill: {name}")
            Return False
        End Function
        
        ''' <summary>
        ''' 复制Skill
        ''' </summary>
        Public Function DuplicateSkill(sourceName As String, newName As String) As SkillDefinition
            Dim source = GetCustomSkill(sourceName)
            If source Is Nothing Then
                Throw New ArgumentException($"找不到Skill: {sourceName}")
            End If

            Dim copy = Utils.JsonHelper.Deserialize(Of SkillDefinition)(Utils.JsonHelper.Serialize(source))
            copy.Name = newName
            copy.DisplayName = copy.DisplayName & " (副本)"
            copy.IsCustom = True
            copy.CreatedAt = DateTime.Now
            copy.LastModified = DateTime.Now
            copy.UsageCount = 0

            SaveCustomSkill(copy)
            Return copy
        End Function
        
        ''' <summary>
        ''' 从模板创建Skill
        ''' </summary>
        Public Function CreateFromTemplate(templateType As String, name As String) As SkillDefinition
            Dim skill As New SkillDefinition()
            skill.Name = name
            skill.IsCustom = True
            skill.Version = 1
            skill.CreatedAt = DateTime.Now
            skill.LastModified = DateTime.Now
            
            Select Case templateType.ToLower()
                Case "businesscard", "card"
                    skill.DisplayName = "名片识别"
                    skill.Description = "识别名片信息并导入Excel通讯录"
                    skill.Keywords = New String() {"名片", "通讯录", "联系人", "导入", "card", "contact"}
                    skill.ProcessingMode = SkillProcessingMode.MultimodalLLM
                    skill.SupportsImageInput = True
                    skill.SupportsFileInput = True
                    skill.InputDescription = "上传名片图片或PDF文件"
                    skill.OutputDescription = "提取的联系人信息"
                    skill.Parameters = New List(Of ParameterDefinition)()
                    skill.Examples = New String() {"识别这张名片", "导入名片信息"}
                    
                Case "invoice", "发票"
                    skill.DisplayName = "发票识别"
                    skill.Description = "识别发票信息并提取关键字段"
                    skill.Keywords = New String() {"发票", "报销", "识别", "invoice", "receipt"}
                    skill.ProcessingMode = SkillProcessingMode.MultimodalLLM
                    skill.SupportsImageInput = True
                    skill.SupportsFileInput = True
                    skill.InputDescription = "上传发票图片或PDF文件"
                    skill.OutputDescription = "提取的发票信息（发票号、金额、日期等）"
                    skill.Parameters = New List(Of ParameterDefinition)()
                    skill.Examples = New String() {"识别这张发票", "提取发票信息"}
                    
                Case "idcard", "身份证"
                    skill.DisplayName = "身份证识别"
                    skill.Description = "识别身份证信息"
                    skill.Keywords = New String() {"身份证", "证件", "识别", "idcard", "identity"}
                    skill.ProcessingMode = SkillProcessingMode.MultimodalLLM
                    skill.SupportsImageInput = True
                    skill.InputDescription = "上传身份证图片"
                    skill.OutputDescription = "提取的身份证信息"
                    skill.Parameters = New List(Of ParameterDefinition)()
                    skill.Examples = New String() {"识别身份证", "提取证件信息"}
                    
                Case Else
                    ' 通用模板
                    skill.DisplayName = name
                    skill.Description = "自定义Skill"
                    skill.Keywords = New String() {name}
                    skill.ProcessingMode = SkillProcessingMode.VBACode
                    skill.SupportsImageInput = False
                    skill.SupportsFileInput = False
                    skill.Parameters = New List(Of ParameterDefinition)()
                    skill.Examples = New String() {}
            End Select
            
            SaveCustomSkill(skill)
            Return skill
        End Function
        
        ''' <summary>
        ''' 重新加载所有自定义Skill
        ''' </summary>
        Public Sub Reload()
            _customSkills.Clear()
            _codeSnippets.Clear()
            LoadCustomSkills()
        End Sub
        
    End Class
End Namespace
