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

    ' SkillManager.Parsing.vb — LLM 响应 -> Skill 定义解析（function call / JSON 治理 / 命名映射）
    ' （类级 XML 文档只在主文件 SkillManager.vb 保留一份，避免 BC42314 分部重复）
    Partial Public Class SkillManager

        Private Function TryParseFunctionCallSkill(jsonText As String) As SkillDefinition
            Try
                Dim jObj = Newtonsoft.Json.Linq.JObject.Parse(jsonText)
                Dim argsObj = ExtractFunctionCallArguments(jObj)
                If argsObj Is Nothing Then Return Nothing

                Dim funcName = ExtractFunctionName(jObj)
                If String.IsNullOrWhiteSpace(funcName) Then Return Nothing

                Dim skill As New SkillDefinition()
                PopulateSkillFromFunctionCall(skill, argsObj, funcName)
                SetSkillCreatorDefaults(skill)

                Return skill
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[TryParseFunctionCallSkill] 解析失败: " & ex.Message)
                Return Nothing
            End Try
        End Function

        Private Function ExtractFunctionCallArguments(jObj As Newtonsoft.Json.Linq.JObject) As Newtonsoft.Json.Linq.JObject
            If jObj Is Nothing OrElse Not jObj.ContainsKey("function_call") Then Return Nothing

            Dim fc = jObj("function_call")
            If fc Is Nothing Then Return Nothing

            Dim args = fc("arguments")
            Return TryCast(args, Newtonsoft.Json.Linq.JObject)
        End Function

        Private Function ExtractFunctionName(jObj As Newtonsoft.Json.Linq.JObject) As String
            Dim fc = jObj("function_call")
            If fc Is Nothing Then Return Nothing
            Return fc("name")?.ToString()
        End Function

        Private Sub PopulateSkillFromFunctionCall(skill As SkillDefinition, argsObj As Newtonsoft.Json.Linq.JObject, funcName As String)
            skill.Name = argsObj("name")?.ToString()
            skill.DisplayName = argsObj("displayName")?.ToString()
            skill.Description = argsObj("description")?.ToString()
            skill.Category = argsObj("category")?.ToString()

            skill.Keywords = ParseJsonStringArray(argsObj, "keywords")
            skill.Examples = ParseJsonStringArray(argsObj, "examples")
            skill.Tags = ParseJsonStringList(argsObj, "tags")

            Dim supportsImage = ParseFlexibleBoolean(argsObj, "supportsImageInput")
            Dim supportsFile = ParseFlexibleBoolean(argsObj, "supportsFileInput")

            ConfigureSkillByFunctionType(skill, funcName, argsObj, supportsImage, supportsFile)
        End Sub

        Private Function ParseJsonStringArray(obj As Newtonsoft.Json.Linq.JObject, key As String) As String()
            Dim token = obj(key)
            If token Is Nothing OrElse token.Type <> Newtonsoft.Json.Linq.JTokenType.Array Then Return New String() {}
            Return token.Select(Function(t) t.ToString()).ToArray()
        End Function

        Private Function ParseJsonStringList(obj As Newtonsoft.Json.Linq.JObject, key As String) As List(Of String)
            Dim arr = ParseJsonStringArray(obj, key)
            Return If(arr Is Nothing, New List(Of String)(), arr.ToList())
        End Function

        Private Function ParseFlexibleBoolean(obj As Newtonsoft.Json.Linq.JObject, key As String) As Boolean
            Dim val = obj(key)
            If val Is Nothing Then Return False

            If val.Type = Newtonsoft.Json.Linq.JTokenType.Boolean Then
                Return val.ToObject(Of Boolean)()
            ElseIf val.Type = Newtonsoft.Json.Linq.JTokenType.Integer Then
                Return val.ToObject(Of Integer)() <> 0
            Else
                Dim strVal = val.ToString().Trim()
                Return strVal.Equals("true", StringComparison.OrdinalIgnoreCase)
            End If
        End Function

        ''' <summary>
        ''' 标准化 LLM 生成的 Skill JSON，修复常见类型问题：
        ''' 1. 数组字段（examples/keywords/tags）被错误生成为对象或字符串
        ''' 2. 整数字段（version/processingMode/supportsImageInput 等）被错误生成为字符串
        ''' 3. 布尔字段被错误生成为字符串
        ''' </summary>
        Private Function NormalizeJsonForSkill(jsonText As String) As String
            Try
                If String.IsNullOrWhiteSpace(jsonText) Then Return jsonText

                Dim jObj = Newtonsoft.Json.Linq.JObject.Parse(jsonText)
                Dim changed As Boolean = False

                ' 辅助函数：获取字段值（尝试 camelCase 和 PascalCase）
                Dim getField = Function(name As String) As Newtonsoft.Json.Linq.JToken
                                   Dim t = jObj(name)
                                   If t IsNot Nothing Then Return t
                                   Dim pascal = name.Substring(0, 1).ToUpper() & name.Substring(1)
                                   Return jObj(pascal)
                               End Function

                ' 辅助函数：设置字段值（尝试 camelCase 和 PascalCase，保持原始键名）
                Dim setField = Sub(name As String, value As Newtonsoft.Json.Linq.JToken)
                                   If jObj(name) IsNot Nothing Then
                                       jObj(name) = value
                                   Else
                                       Dim pascal = name.Substring(0, 1).ToUpper() & name.Substring(1)
                                       If jObj(pascal) IsNot Nothing Then
                                           jObj(pascal) = value
                                       Else
                                           jObj(name) = value
                                       End If
                                   End If
                               End Sub

                ' 1. 修复数组字段
                Dim arrayFields = {"examples", "keywords", "tags", "configFields", "inputContext"}
                For Each fieldName In arrayFields
                    Dim token = getField(fieldName)
                    If token Is Nothing Then Continue For

                    If token.Type = Newtonsoft.Json.Linq.JTokenType.Array Then Continue For

                    Dim newArray As New Newtonsoft.Json.Linq.JArray()
                    If token.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                        For Each prop In CType(token, Newtonsoft.Json.Linq.JObject).Properties()
                            Dim val = prop.Value?.ToString()
                            If Not String.IsNullOrWhiteSpace(val) Then newArray.Add(val)
                        Next
                        changed = True
                    ElseIf token.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                        Dim strVal = token.ToString()
                        If Not String.IsNullOrWhiteSpace(strVal) Then newArray.Add(strVal)
                        changed = True
                    ElseIf token.Type <> Newtonsoft.Json.Linq.JTokenType.Null Then
                        newArray.Add(token.ToString())
                        changed = True
                    Else
                        changed = True
                    End If

                    If changed Then setField(fieldName, newArray)
                Next

                ' 2. 修复整数字段
                Dim intFields = {"version", "processingMode"}
                For Each fieldName In intFields
                    Dim token = getField(fieldName)
                    If token Is Nothing Then Continue For

                    If token.Type = Newtonsoft.Json.Linq.JTokenType.Integer Then Continue For

                    Dim intVal As Integer
                    Dim strRep = token?.ToString()?.Trim()
                    If Integer.TryParse(strRep, intVal) Then
                        setField(fieldName, New Newtonsoft.Json.Linq.JValue(intVal))
                        changed = True
                    ElseIf Decimal.TryParse(strRep, Nothing) Then
                        Dim decVal = Decimal.Parse(strRep)
                        intVal = CInt(Math.Round(decVal))
                        setField(fieldName, New Newtonsoft.Json.Linq.JValue(intVal))
                        changed = True
                    End If
                Next

                ' 3. 修复布尔字段（字符串 "true"/"false" → 布尔）
                Dim boolFields = {"supportsImageInput", "supportsFileInput", "isGeneric", "needsConfig", "isCustom", "isComposite"}
                For Each fieldName In boolFields
                    Dim token = getField(fieldName)
                    If token Is Nothing Then Continue For

                    If token.Type = Newtonsoft.Json.Linq.JTokenType.Boolean Then Continue For

                    Dim strVal = token?.ToString()?.Trim()
                    If strVal IsNot Nothing Then
                        Dim boolVal As Boolean
                        If Boolean.TryParse(strVal, boolVal) Then
                            setField(fieldName, New Newtonsoft.Json.Linq.JValue(boolVal))
                            changed = True
                        End If
                    End If
                Next

                If changed Then
                    Return jObj.ToString(Newtonsoft.Json.Formatting.None)
                End If

                Return jsonText
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[NormalizeJsonForSkill] 标准化失败: " & ex.Message)
                Return jsonText
            End Try
        End Function

        Private Sub ConfigureSkillByFunctionType(skill As SkillDefinition, funcName As String,
                                                  argsObj As Newtonsoft.Json.Linq.JObject,
                                                  supportsImage As Boolean, supportsFile As Boolean)
            Select Case funcName.ToLower()
                Case "create_vba_skill"
                    skill.ProcessingMode = SkillProcessingMode.VBACode
                    skill.OutputFormat = "vba"
                    skill.TemplateCode = argsObj("templateCode")?.ToString()
                Case "create_llm_skill"
                    skill.ProcessingMode = SkillProcessingMode.MultimodalLLM
                    Dim outputFormat = argsObj("outputFormat")?.ToString()
                    skill.OutputFormat = If(String.IsNullOrWhiteSpace(outputFormat), "markdown", outputFormat)
                    skill.LLMSystemPrompt = argsObj("llmSystemPrompt")?.ToString()
                Case "create_hybrid_skill"
                    skill.ProcessingMode = SkillProcessingMode.Hybrid
                    Dim outputFormat = argsObj("outputFormat")?.ToString()
                    skill.OutputFormat = If(String.IsNullOrWhiteSpace(outputFormat), "vba", outputFormat)
                    skill.LLMSystemPrompt = argsObj("llmSystemPrompt")?.ToString()
                    skill.TemplateCode = argsObj("templateCode")?.ToString()
            End Select

            skill.SupportsImageInput = supportsImage
            skill.SupportsFileInput = supportsFile
        End Sub

        Private Sub SetSkillCreatorDefaults(skill As SkillDefinition)
            skill.IsGeneric = False
            skill.NeedsConfig = False
            skill.IsComposite = False
            skill.IsCustom = True
            skill.Version = 1
            skill.UsageCount = 0
            skill.CreatedBy = "SkillCreator"
            skill.CreatedDate = DateTime.Now
            skill.LastModifiedDate = DateTime.Now
        End Sub

        ''' <summary>
        ''' 从 LLM 响应文本中提取 JSON 对象字符串（增强版）
        ''' </summary>
        Private Function ExtractJsonFromLLMResponse(text As String) As String
            If String.IsNullOrWhiteSpace(text) Then Return Nothing

            text = text.Trim()

            ' 先清除 <thinking>...</thinking> 标签内容
            text = System.Text.RegularExpressions.Regex.Replace(text, "<thinking>[\s\S]*?</thinking>", "")
            ' 清除 <tool_call>...</tool_call> 等可能包裹的标签
            text = System.Text.RegularExpressions.Regex.Replace(text, "<[^>]+>[\s\S]*?</[^>]+>", "")
            text = text.Trim()

            ' 策略1: 尝试直接解析整个文本
            Try
                Dim testObj = Newtonsoft.Json.Linq.JObject.Parse(text)
                If testObj IsNot Nothing Then Return text
            Catch ex As Exception  ' 
            End Try

            ' 策略2: 匹配 ```json ... ``` 或 ``` ... ``` 代码块
            Dim codeBlockMatch = System.Text.RegularExpressions.Regex.Match(text, "```(?:json)?\s*([\s\S]*?)\s*```", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            If codeBlockMatch.Success AndAlso codeBlockMatch.Groups.Count > 1 Then
                Dim candidate = SanitizeJsonString(codeBlockMatch.Groups(1).Value.Trim())
                Dim valid = TryParseJObject(candidate)
                If valid Then Return candidate
            End If

            ' 策略3: 平衡括号匹配（找到完整的 { ... }）
            Dim braceMatch = ExtractBalancedBraces(text)
            If Not String.IsNullOrWhiteSpace(braceMatch) Then
                Dim candidate = SanitizeJsonString(braceMatch)
                Dim valid = TryParseJObject(candidate)
                If valid Then Return candidate
            End If

            ' 策略4: 扫描文本中所有可能的 { ... } 片段，取最大的可解析片段
            Dim allMatches = FindAllBraceCandidates(text)
            For Each candidate In allMatches
                Dim sanitized = SanitizeJsonString(candidate)
                If TryParseJObject(sanitized) Then
                    Return sanitized
                End If
            Next

            ' 策略5: 尝试提取包含 Skill 关键字段的 JSON
            Dim skillMatch = ExtractSkillJsonCandidate(text)
            If Not String.IsNullOrWhiteSpace(skillMatch) Then
                Return skillMatch
            End If

            ' 策略6: 暴力提取 — 找到第一个 { 和最后一个 }，截取中间全部内容
            Dim firstBrace = text.IndexOf("{"c)
            Dim lastBrace = text.LastIndexOf("}"c)
            If firstBrace >= 0 AndAlso lastBrace > firstBrace Then
                Dim bruteCandidate = text.Substring(firstBrace, lastBrace - firstBrace + 1)
                Dim sanitized = SanitizeJsonString(bruteCandidate)
                If TryParseJObject(sanitized) Then
                    Return sanitized
                End If
            End If

            ' 策略7: 逐字符扫描，对每个 { 位置尝试提取，找到第一个可解析的 JSON
            Dim idx As Integer = 0
            While idx < text.Length
                Dim bracePos = text.IndexOf("{"c, idx)
                If bracePos < 0 Then Exit While
                Dim candidate = ExtractBalancedBraces(text.Substring(bracePos))
                If Not String.IsNullOrWhiteSpace(candidate) Then
                    Dim sanitized = SanitizeJsonString(candidate)
                    If TryParseJObject(sanitized) Then
                        Return sanitized
                    End If
                    idx = bracePos + candidate.Length
                Else
                    idx = bracePos + 1
                End If
            End While

            Return Nothing
        End Function

        ''' <summary>
        ''' 清理 LLM 输出中常见的 JSON 格式污染
        ''' </summary>
        Private Function SanitizeJsonString(jsonText As String) As String
            If String.IsNullOrWhiteSpace(jsonText) Then Return jsonText

            Dim result = jsonText.Trim()

            ' 移除 BOM 和零宽字符，但保留正常换行/制表
            result = New String(result.Where(Function(c) Not Char.IsControl(c) OrElse c = vbCr Or c = vbLf Or c = vbTab).ToArray())

            ' 修复中文引号
            result = result.Replace(""""c, """"c).Replace(""""c, """"c)

            ' 修复多余的逗号：",}" → "}"
            result = System.Text.RegularExpressions.Regex.Replace(result, ",\s*}", "}")
            result = System.Text.RegularExpressions.Regex.Replace(result, ",\s*]", "]")

            ' 修复反引号字符串（LLM 有时用 JavaScript 反引号包裹代码，不是合法 JSON）
            ' 格式: "key": `code content`  →  "key": "code content"
            result = FixBacktickStrings(result)

            Return result
        End Function

        ''' <summary>
        ''' 将 JSON 中的反引号字符串（`...`）替换为合法的双引号字符串
        ''' LLM 有时会输出 JavaScript 风格的模板字符串，不是合法 JSON
        ''' </summary>
        Private Function FixBacktickStrings(jsonText As String) As String
            If String.IsNullOrWhiteSpace(jsonText) OrElse Not jsonText.Contains("`"c) Then Return jsonText

            ' 匹配反引号字符串：`...`（可跨行，非贪婪）
            Dim backtickRegex As New System.Text.RegularExpressions.Regex("`([\s\S]*?)`")
            Return backtickRegex.Replace(jsonText, Function(match As System.Text.RegularExpressions.Match)
                Dim inner = match.Groups(1).Value
                ' 转义内容中的双引号和反斜杠
                inner = inner.Replace("\", "\\")
                inner = inner.Replace("""", "\""")
                ' 转义换行符
                inner = inner.Replace(vbCrLf, "\n")
                inner = inner.Replace(vbCr, "\n")
                inner = inner.Replace(vbLf, "\n")
                ' 移除前后多余的空白
                inner = inner.Trim()
                Return """" & inner & """"
            End Function)
        End Function

        ''' <summary>
        ''' 扫描文本中所有可能的 JSON 对象候选（优先返回最长可解析的）
        ''' </summary>
        Private Function FindAllBraceCandidates(text As String) As List(Of String)
            Dim candidates As New List(Of String)()
            Dim idx As Integer = 0

            While idx < text.Length
                Dim braceStart = text.IndexOf("{"c, idx)
                If braceStart < 0 Then Exit While

                Dim balanced = ExtractBalancedBraces(text.Substring(braceStart))
                If Not String.IsNullOrWhiteSpace(balanced) Then
                    candidates.Add(balanced)
                    idx = braceStart + balanced.Length
                Else
                    idx = braceStart + 1
                End If
            End While

            Return candidates.OrderByDescending(Function(s) s.Length).ToList()
        End Function

        ''' <summary>
        ''' 尝试解析 JObject（不抛异常）
        ''' </summary>
        Private Function TryParseJObject(jsonText As String) As Boolean
            If String.IsNullOrWhiteSpace(jsonText) Then Return False
            Try
                Newtonsoft.Json.Linq.JObject.Parse(jsonText)
                Return True
            Catch ex As Exception  ' 
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 平衡括号提取完整 JSON 对象
        ''' </summary>
        Private Function ExtractBalancedBraces(text As String) As String
            Dim startIdx = text.IndexOf("{"c)
            If startIdx < 0 Then Return Nothing

            Dim depth As Integer = 0
            Dim inString As Boolean = False
            Dim escape As Boolean = False

            For i As Integer = startIdx To text.Length - 1
                Dim ch = text(i)

                If escape Then
                    escape = False
                    Continue For
                End If

                If ch = "\"c Then
                    escape = True
                    Continue For
                End If

                If ch = """"c Then
                    inString = Not inString
                    Continue For
                End If

                If Not inString Then
                    If ch = "{"c Then
                        depth += 1
                    ElseIf ch = "}"c Then
                        depth -= 1
                        If depth = 0 Then
                            Return text.Substring(startIdx, i - startIdx + 1)
                        End If
                    End If
                End If
            Next

            Return Nothing
        End Function

        ''' <summary>
        ''' 尝试从文本中找到包含 Skill 必需字段的 JSON 片段
        ''' </summary>
        Private Function ExtractSkillJsonCandidate(text As String) As String
            Dim requiredFields = {"name", "displayName", "description", "examples", "keywords", "category", "processingMode", "outputFormat"}

            ' 提取所有可能的 { ... } 片段
            Dim braceRegex As New System.Text.RegularExpressions.Regex("\{[\s\S]*?\}")
            Dim matches = braceRegex.Matches(text)

            For Each match As System.Text.RegularExpressions.Match In matches
                Dim candidate = match.Value.Trim()
                If Not TryParseJObject(candidate) Then Continue For

                Try
                    Dim jObj = Newtonsoft.Json.Linq.JObject.Parse(candidate)
                    If jObj Is Nothing Then Continue For

                    ' 检查是否包含至少 3 个 Skill 必需字段
                    Dim fieldCount As Integer = 0
                    For Each field In requiredFields
                        If jObj.ContainsKey(field) Then fieldCount += 1
                    Next

                    If fieldCount >= 3 Then
                        Return candidate
                    End If
                Catch ex As Exception  ' 
                End Try
            Next

            Return Nothing
        End Function

        ''' <summary>
        ''' 灵活反序列化 Skill：先标准化类型，再尝试直接反序列化和逐层解包
        ''' 【修复】使用 SkillJsonHelper 进行不区分大小写的反序列化
        ''' </summary>
        Private Function TryDeserializeSkillFlexibly(jsonText As String) As SkillDefinition
            Try
                Dim normalizedJson = NormalizeJsonForSkill(jsonText)

                ' 使用 SkillJsonHelper 进行不区分大小写的反序列化
                Dim skill = SkillJsonHelper.DeserializeSkill(normalizedJson)

                If skill IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(skill.Name) Then
                    System.Diagnostics.Debug.WriteLine("[SkillCreator] 通过 SkillJsonHelper 反序列化成功: " & skill.Name)
                    Return skill
                End If

                ' 反序列化成功但 name 为空 — 可能是嵌套结构，尝试逐层解包
                System.Diagnostics.Debug.WriteLine("[SkillCreator] SkillJsonHelper 返回的 Skill 名称为空，尝试解包嵌套结构")
                Try
                    Dim jObj = Newtonsoft.Json.Linq.JObject.Parse(normalizedJson)
                    Dim unwrapped = UnwrapNestedSkill(jObj)
                    If unwrapped IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(unwrapped.Name) Then
                        System.Diagnostics.Debug.WriteLine("[SkillCreator] 解包嵌套结构成功: " & unwrapped.Name)
                        Return unwrapped
                    End If

                    ' 尝试映射 LLM 使用的替代字段名 → 标准字段名
                    Dim mapped = TryMapAlternativeNames(jObj)
                    If mapped IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(mapped.Name) Then
                        System.Diagnostics.Debug.WriteLine("[SkillCreator] 映射替代字段名成功: " & mapped.Name)
                        Return mapped
                    End If
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("[SkillCreator] 解包嵌套结构异常: " & ex.Message)
                End Try

                System.Diagnostics.Debug.WriteLine("[SkillCreator] 所有反序列化方式均失败")
                Return Nothing
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[SkillCreator] TryDeserializeSkillFlexibly 异常: " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 尝试将 LLM 常用的替代字段名映射到 SkillDefinition 标准字段名
        ''' </summary>
        Private Function TryMapAlternativeNames(jObj As Newtonsoft.Json.Linq.JObject) As SkillDefinition
            Try
                ' 常见替代字段名映射表
                ' 格式: { PascalCase属性名, { 替代名称列表 } }
                Dim fieldMappings As New Dictionary(Of String, String()) From {
                    {"Name", {"skillName", "SkillName", "skill_name", "name_id", "name", "Name"}},
                    {"DisplayName", {"display_name", "displayName", "DisplayName", "Display_Name", "title", "Title"}},
                    {"Description", {"description", "Description", "desc", "Desc", "function_description", "FunctionDescription"}},
                    {"ProcessingMode", {"processing_mode", "processingMode", "ProcessingMode", "mode", "Mode", "type", "Type"}},
                    {"TemplateCode", {"template_code", "templateCode", "TemplateCode", "code", "Code", "vba_code", "VbaCode"}},
                    {"LLMSystemPrompt", {"llm_system_prompt", "llmSystemPrompt", "LlmSystemPrompt", "system_prompt", "SystemPrompt", "prompt", "Prompt"}},
                    {"Keywords", {"keywords", "Keywords", "keyword", "Keyword"}},
                    {"Examples", {"examples", "Examples", "example", "Example"}},
                    {"Category", {"category", "Category", "categories", "Categories"}},
                    {"Tags", {"tags", "Tags", "tag", "Tag"}},
                    {"OutputFormat", {"output_format", "outputFormat", "OutputFormat", "format", "Format"}},
                    {"IsGeneric", {"is_generic", "isGeneric", "IsGeneric", "generic", "Generic"}},
                    {"NeedsConfig", {"needs_config", "needsConfig", "NeedsConfig", "needsConfiguration", "NeedsConfiguration"}},
                    {"ConfigFields", {"config_fields", "configFields", "ConfigFields", "configuration", "Configuration", "inputParameters", "input_parameters", "InputParameters"}},
                    {"SupportsImageInput", {"supports_image_input", "supportsImageInput", "SupportsImageInput", "imageInput", "ImageInput"}},
                    {"SupportsFileInput", {"supports_file_input", "supportsFileInput", "SupportsFileInput", "fileInput", "FileInput"}},
                    {"FieldMappings", {"field_mappings", "fieldMappings", "FieldMappings", "mappings", "Mappings"}}
                }

                ' 构建标准 JSON 对象（使用 PascalCase 属性名匹配 SkillDefinition）
                Dim standardJson As New Dictionary(Of String, Object)()

                For Each kvp In fieldMappings
                    Dim standardKey = kvp.Key
                    Dim alternativeNames = kvp.Value

                    Dim foundValue As Newtonsoft.Json.Linq.JToken = Nothing
                    For Each altName In alternativeNames
                        If jObj.TryGetValue(altName, StringComparison.OrdinalIgnoreCase, foundValue) Then
                            Exit For
                        End If
                    Next

                    If foundValue IsNot Nothing Then
                        Select Case standardKey
                            Case "Name", "DisplayName", "Description", "TemplateCode",
                                 "LLMSystemPrompt", "Category", "OutputFormat"
                                standardJson(standardKey) = foundValue.ToString()
                            Case "ProcessingMode"
                                ' 转换 processingMode 为枚举整数值
                                Dim modeInt As Integer
                                If foundValue.Type = Newtonsoft.Json.Linq.JTokenType.Integer Then
                                    modeInt = CInt(foundValue)
                                ElseIf foundValue.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                                    Dim modeStr = foundValue.ToString().Trim()
                                    If Integer.TryParse(modeStr, modeInt) Then
                                        ' 成功解析为整数
                                    Else
                                        ' 尝试按枚举名称解析
                                        Dim parsedMode As SkillProcessingMode
                                        If [Enum].TryParse(modeStr, True, parsedMode) Then
                                            modeInt = CInt(parsedMode)
                                        Else
                                            modeInt = 0 ' 默认为 VBACode
                                        End If
                                    End If
                                Else
                                    modeInt = 0
                                End If
                                standardJson(standardKey) = modeInt
                            Case "IsGeneric", "NeedsConfig", "SupportsImageInput", "SupportsFileInput"
                                If foundValue.Type = Newtonsoft.Json.Linq.JTokenType.Boolean Then
                                    standardJson(standardKey) = CBool(foundValue)
                                Else
                                    Dim strVal = foundValue.ToString().ToLowerInvariant()
                                    standardJson(standardKey) = (strVal = "true" OrElse strVal = "yes")
                                End If
                            Case "Keywords", "Examples", "Tags", "ConfigFields", "FieldMappings"
                                If foundValue.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                                    standardJson(standardKey) = foundValue.ToObject(Of List(Of String))()
                                ElseIf foundValue.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                                    standardJson(standardKey) = New List(Of String) From {foundValue.ToString()}
                                Else
                                    standardJson(standardKey) = New List(Of String)()
                                End If
                        End Select
                    End If
                Next

                ' === 特殊处理：LLM 常用但结构不同的替代字段 ===

                ' 1. actions → TemplateCode：从 actions 数组的第一个元素提取 code 字段
                If Not standardJson.ContainsKey("TemplateCode") Then
                    Dim actionsToken As Newtonsoft.Json.Linq.JToken = Nothing
                    If jObj.TryGetValue("actions", StringComparison.OrdinalIgnoreCase, actionsToken) AndAlso
                       actionsToken.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                        Dim actionsArray = CType(actionsToken, Newtonsoft.Json.Linq.JArray)
                        If actionsArray.Count > 0 AndAlso actionsArray(0).Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                            Dim firstAction = CType(actionsArray(0), Newtonsoft.Json.Linq.JObject)
                            Dim codeToken As Newtonsoft.Json.Linq.JToken = Nothing
                            If firstAction.TryGetValue("code", StringComparison.OrdinalIgnoreCase, codeToken) Then
                                standardJson("TemplateCode") = codeToken.ToString()
                            End If
                        End If
                    End If
                End If

                ' 2. inputParameters → ConfigFields：从 inputParameters 数组提取各元素的 name 字段
                If Not standardJson.ContainsKey("ConfigFields") OrElse
                   (TypeOf standardJson("ConfigFields") Is List(Of String) AndAlso
                    DirectCast(standardJson("ConfigFields"), List(Of String)).Count = 0) Then
                    Dim inputParamsToken As Newtonsoft.Json.Linq.JToken = Nothing
                    If jObj.TryGetValue("inputParameters", StringComparison.OrdinalIgnoreCase, inputParamsToken) AndAlso
                       inputParamsToken.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                        Dim paramsArray = CType(inputParamsToken, Newtonsoft.Json.Linq.JArray)
                        Dim paramNames As New List(Of String)()
                        For Each param In paramsArray
                            If param.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                                Dim paramObj = CType(param, Newtonsoft.Json.Linq.JObject)
                                Dim nameToken As Newtonsoft.Json.Linq.JToken = Nothing
                                If paramObj.TryGetValue("name", StringComparison.OrdinalIgnoreCase, nameToken) Then
                                    paramNames.Add(nameToken.ToString())
                                End If
                            End If
                        Next
                        If paramNames.Count > 0 Then
                            standardJson("ConfigFields") = paramNames
                        End If
                    End If
                End If

                ' 3. 如果 Name 包含中文，尝试从 Name 或 Description 生成 PascalCase 英文 ID
                If standardJson.ContainsKey("Name") Then
                    Dim nameStr = standardJson("Name").ToString()
                    If System.Text.RegularExpressions.Regex.IsMatch(nameStr, "[\u4e00-\u9fff]") Then
                        ' 尝试从 Description 提取英文 ID，或生成一个通用 ID
                        Dim descStr = ""
                        If standardJson.ContainsKey("Description") Then
                            descStr = standardJson("Description").ToString()
                        End If
                        Dim generatedName = GeneratePascalCaseName(nameStr, descStr)
                        If Not String.IsNullOrWhiteSpace(generatedName) Then
                            standardJson("Name") = generatedName
                        End If
                    End If
                End If

                ' 如果连 Name 都没有，无法构建有效 Skill
                If Not standardJson.ContainsKey("Name") OrElse String.IsNullOrWhiteSpace(standardJson("Name").ToString()) Then
                    Return Nothing
                End If

                ' 从 standardJson 序列化再反序列化为 SkillDefinition
                ' 使用 SkillJsonHelper 进行不区分大小写的反序列化
                Dim mappedJsonStr = Newtonsoft.Json.JsonConvert.SerializeObject(standardJson)
                Dim mappedSkill = SkillJsonHelper.DeserializeSkill(mappedJsonStr)
                Return mappedSkill
            Catch ex As Exception  ' 
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 从中文名或描述中生成 PascalCase 英文 ID
        ''' 例如: "自动整理表格" → "AutoFormatTable"，"整理表格" → "FormatTable"
        ''' </summary>
        Private Function GeneratePascalCaseName(nameCn As String, description As String) As String
            Try
                ' 尝试从描述中提取关键词
                Dim desc = If(String.IsNullOrWhiteSpace(description), nameCn, description)

                ' 常见中文操作词 → 英文映射
                Dim actionMap As New Dictionary(Of String, String) From {
                    {"整理", "Format"}, {"清理", "Clean"}, {"删除", "Delete"}, {"移除", "Remove"},
                    {"提取", "Extract"}, {"导入", "Import"}, {"导出", "Export"}, {"生成", "Generate"},
                    {"创建", "Create"}, {"合并", "Merge"}, {"拆分", "Split"}, {"查找", "Find"},
                    {"替换", "Replace"}, {"转换", "Convert"}, {"分析", "Analyze"}, {"统计", "Statistic"},
                    {"汇总", "Summary"}, {"排序", "Sort"}, {"筛选", "Filter"}, {"格式化", "Format"},
                    {"美化", "Beautify"}, {"检查", "Check"}, {"验证", "Validate"}, {"对比", "Compare"},
                    {"发送", "Send"}, {"下载", "Download"}, {"上传", "Upload"}, {"备份", "Backup"},
                    {"恢复", "Restore"}, {"更新", "Update"}, {"同步", "Sync"}, {"填报", "Fill"},
                    {"采集", "Collect"}, {"识别", "Recognize"}, {"检测", "Detect"}, {"监控", "Monitor"}
                }

                ' 常见中文名词 → 英文映射
                Dim nounMap As New Dictionary(Of String, String) From {
                    {"表格", "Table"}, {"数据", "Data"}, {"文件", "File"}, {"图片", "Image"},
                    {"邮件", "Email"}, {"报表", "Report"}, {"图表", "Chart"}, {"工作簿", "Workbook"},
                    {"工作表", "Worksheet"}, {"单元格", "Cell"}, {"列", "Column"}, {"行", "Row"},
                    {"内容", "Content"}, {"格式", "Format"}, {"样式", "Style"}, {"颜色", "Color"},
                    {"字体", "Font"}, {"边框", "Border"}, {"公式", "Formula"}, {"函数", "Function"},
                    {"名称", "Name"}, {"日期", "Date"}, {"时间", "Time"}, {"数字", "Number"},
                    {"文本", "Text"}, {"值", "Value"}, {"区域", "Range"}, {"地址", "Address"},
                    {"链接", "Link"}, {"批注", "Comment"}, {"条件", "Condition"}, {"规则", "Rule"},
                    {"模板", "Template"}, {"配置", "Config"}, {"设置", "Setting"}, {"属性", "Property"},
                    {"信息", "Info"}, {"记录", "Record"}, {"日志", "Log"}, {"任务", "Task"},
                    {"项目", "Project"}, {"订单", "Order"}, {"客户", "Customer"}, {"产品", "Product"},
                    {"库存", "Inventory"}, {"财务", "Finance"}, {"预算", "Budget"}, {"发票", "Invoice"}
                }

                ' 从描述中匹配操作词和名词
                Dim actionWord As String = ""
                Dim nounWord As String = ""

                For Each kvp In actionMap
                    If desc.Contains(kvp.Key) Then
                        actionWord = kvp.Value
                        Exit For
                    End If
                Next

                For Each kvp In nounMap
                    If desc.Contains(kvp.Key) Then
                        nounWord = kvp.Value
                        Exit For
                    End If
                Next

                ' 如果找不到匹配，使用通用名称
                If String.IsNullOrWhiteSpace(actionWord) AndAlso String.IsNullOrWhiteSpace(nounWord) Then
                    Return "CustomSkill_" & DateTime.Now.ToString("yyyyMMddHHmmss")
                End If

                Return actionWord & nounWord
            Catch ex As Exception  ' 
                Return "CustomSkill_" & DateTime.Now.ToString("yyyyMMddHHmmss")
            End Try
        End Function

        ''' <summary>
        ''' 从原始文本中递归搜索包含 Skill 必需字段的 JSON
        ''' </summary>
        Private Function TryExtractSkillFromNestedJson(rawText As String) As SkillDefinition
            If String.IsNullOrWhiteSpace(rawText) Then Return Nothing

            ' 清除 thinking 标签
            Dim cleaned = System.Text.RegularExpressions.Regex.Replace(rawText, "<thinking>[\s\S]*?</thinking>", "")

            ' 提取所有 { ... } 平衡括号片段
            Dim candidates As New List(Of String)()
            Dim idx As Integer = 0
            While idx < cleaned.Length
                Dim braceStart = cleaned.IndexOf("{"c, idx)
                If braceStart < 0 Then Exit While

                Dim balanced = ExtractBalancedBraces(cleaned.Substring(braceStart))
                If Not String.IsNullOrWhiteSpace(balanced) Then
                    candidates.Add(balanced)
                    idx = braceStart + balanced.Length
                Else
                    idx = braceStart + 1
                End If
            End While

            ' 对每个候选进行标准化和反序列化
            For Each candidate In candidates
                Try
                    Dim normalized = NormalizeJsonForSkill(candidate)
                    ' 使用 SkillJsonHelper 进行不区分大小写的反序列化
                    Dim skill = SkillJsonHelper.DeserializeSkill(normalized)
                    If skill IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(skill.Name) Then
                        Return skill
                    End If

                    ' 尝试解包嵌套
                    Dim jObj = Newtonsoft.Json.Linq.JObject.Parse(normalized)
                    Dim unwrapped = UnwrapNestedSkill(jObj)
                    If unwrapped IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(unwrapped.Name) Then
                        Return unwrapped
                    End If
                Catch ex As Exception  ' 
                End Try
            Next

            Return Nothing
        End Function

        ''' <summary>
        ''' 逐层解包嵌套 JSON 对象，找到包含 Skill 必需字段的那一层
        ''' </summary>
        Private Function UnwrapNestedSkill(jObj As Newtonsoft.Json.Linq.JObject) As SkillDefinition
            If jObj Is Nothing Then Return Nothing

            ' 检查当前层是否已经是 Skill
            If jObj.ContainsKey("name") AndAlso Not String.IsNullOrWhiteSpace(jObj("name")?.ToString()) Then
                Try
                    Dim normalized = NormalizeJsonForSkill(jObj.ToString(Newtonsoft.Json.Formatting.None))
                    ' 使用 SkillJsonHelper 进行不区分大小写的反序列化
                    Dim skill = SkillJsonHelper.DeserializeSkill(normalized)
                    If skill IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(skill.Name) Then
                        Return skill
                    End If
                Catch ex As Exception  ' 
                End Try
            End If

            ' 递归查找子对象
            For Each prop In jObj.Properties()
                If prop.Value.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                    Dim childObj = CType(prop.Value, Newtonsoft.Json.Linq.JObject)
                    Dim result = UnwrapNestedSkill(childObj)
                    If result IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(result.Name) Then
                        Return result
                    End If
                End If
            Next

            Return Nothing
        End Function

        ' 【历史备注】判断是否应使用多模态处理（该方法定义在其他分部文件中）

    End Class

End Namespace
