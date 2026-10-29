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

    ''' <summary>
    ''' 不区分大小写的 Skill 反序列化辅助类
    ''' 解决 LLM 返回 camelCase JSON 与 .NET PascalCase 属性不匹配的问题
    ''' </summary>
    Public Module SkillJsonHelper

        ''' <summary>
        ''' 反序列化 JSON 为 SkillDefinition（不区分大小写）
        ''' </summary>
        Public Function DeserializeSkill(jsonText As String) As SkillDefinition
            Try
                If String.IsNullOrWhiteSpace(jsonText) Then Return Nothing

                Dim jObj = Newtonsoft.Json.Linq.JObject.Parse(jsonText)
                If jObj Is Nothing Then Return Nothing

                Dim skill As New SkillDefinition()

                ' 使用反射获取所有公开属性
                Dim props = GetType(SkillDefinition).GetProperties(
                    System.Reflection.BindingFlags.Public Or
                    System.Reflection.BindingFlags.Instance)

                For Each prop In props
                    Dim jsonValue = FindJsonValueCaseInsensitive(jObj, prop.Name)
                    If jsonValue Is Nothing Then Continue For

                    Try
                        SetPropertyValue(skill, prop, jsonValue)
                    Catch ex As Exception
                        ' 忽略单个属性的转换错误
                        System.Diagnostics.Debug.WriteLine($"[SkillJsonHelper] 设置属性 {prop.Name} 失败: {ex.Message}")
                    End Try
                Next

                Return skill
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[SkillJsonHelper] 反序列化失败: {ex.Message}")
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 在 JObject 中查找值（不区分大小写）
        ''' </summary>
        Private Function FindJsonValueCaseInsensitive(jObj As Newtonsoft.Json.Linq.JObject, propertyName As String) As Newtonsoft.Json.Linq.JToken
            ' 首先尝试精确匹配
            Dim val = jObj(propertyName)
            If val IsNot Nothing Then Return val

            ' 尝试不区分大小写匹配
            For Each prop In jObj.Properties()
                If String.Equals(prop.Name, propertyName, StringComparison.OrdinalIgnoreCase) Then
                    Return prop.Value
                End If
            Next

            ' 尝试常见别名
            Dim aliases = GetPropertyAliases(propertyName)
            For Each [alias] In aliases
                For Each prop In jObj.Properties()
                    If String.Equals(prop.Name, [alias], StringComparison.OrdinalIgnoreCase) Then
                        Return prop.Value
                    End If
                Next
            Next

            Return Nothing
        End Function

        ''' <summary>
        ''' 获取属性的常见别名（LLM 可能使用的字段名）
        ''' </summary>
        Private Function GetPropertyAliases(propertyName As String) As String()
            Select Case propertyName
                Case "Name"
                    Return {"skillName", "SkillName", "skill_name", "name_id"}
                Case "DisplayName"
                    Return {"display_name", "title", "Title"}
                Case "Description"
                    Return {"desc", "Desc", "function_description"}
                Case "ProcessingMode"
                    Return {"processing_mode", "mode", "Mode", "type", "Type"}
                Case "TemplateCode"
                    Return {"template_code", "code", "Code", "vba_code"}
                Case "LLMSystemPrompt"
                    Return {"llm_system_prompt", "system_prompt", "prompt", "Prompt"}
                Case "OutputFormat"
                    Return {"output_format", "format", "Format"}
                Case "IsGeneric"
                    Return {"is_generic", "generic", "Generic"}
                Case "NeedsConfig"
                    Return {"needs_config", "needsConfiguration"}
                Case "SupportsImageInput"
                    Return {"supports_image_input", "imageInput"}
                Case "SupportsFileInput"
                    Return {"supports_file_input", "fileInput"}
                Case "ConfigFields"
                    Return {"config_fields", "configuration", "inputParameters"}
                Case "Keywords"
                    Return {"keyword", "Keyword"}
                Case "Examples"
                    Return {"example", "Example"}
                Case "Tags"
                    Return {"tag", "Tag"}
                Case "Category"
                    Return {"categories", "Categories"}
                Case "FieldMappings"
                    Return {"field_mappings", "mappings"}
                Case "IsComposite"
                    Return {"is_composite", "composite"}
                Case "IsCustom"
                    Return {"is_custom", "custom"}
                Case "Version"
                    Return {"ver"}
                Case Else
                    Return New String() {}
            End Select
        End Function

        ''' <summary>
        ''' 设置属性值（处理类型转换）
        ''' </summary>
        Private Sub SetPropertyValue(target As Object, prop As System.Reflection.PropertyInfo, jsonValue As Newtonsoft.Json.Linq.JToken)
            Dim propType = prop.PropertyType

            ' 处理可空类型
            If Nullable.GetUnderlyingType(propType) IsNot Nothing Then
                propType = Nullable.GetUnderlyingType(propType)
            End If

            Select Case propType
                Case GetType(String)
                    prop.SetValue(target, jsonValue?.ToString())

                Case GetType(Integer)
                    If jsonValue.Type = Newtonsoft.Json.Linq.JTokenType.Integer Then
                        prop.SetValue(target, CInt(jsonValue))
                    ElseIf jsonValue.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                        Dim intVal As Integer
                        If Integer.TryParse(jsonValue.ToString(), intVal) Then
                            prop.SetValue(target, intVal)
                        End If
                    End If

                Case GetType(Boolean)
                    If jsonValue.Type = Newtonsoft.Json.Linq.JTokenType.Boolean Then
                        prop.SetValue(target, CBool(jsonValue))
                    ElseIf jsonValue.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                        Dim boolVal As Boolean
                        If Boolean.TryParse(jsonValue.ToString(), boolVal) Then
                            prop.SetValue(target, boolVal)
                        End If
                    ElseIf jsonValue.Type = Newtonsoft.Json.Linq.JTokenType.Integer Then
                        prop.SetValue(target, CInt(jsonValue) <> 0)
                    End If

                Case GetType(Double)
                    If jsonValue.Type = Newtonsoft.Json.Linq.JTokenType.Float Then
                        prop.SetValue(target, CDbl(jsonValue))
                    ElseIf jsonValue.Type = Newtonsoft.Json.Linq.JTokenType.Integer Then
                        prop.SetValue(target, CDbl(CInt(jsonValue)))
                    ElseIf jsonValue.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                        Dim dblVal As Double
                        If Double.TryParse(jsonValue.ToString(), dblVal) Then
                            prop.SetValue(target, dblVal)
                        End If
                    End If

                Case GetType(DateTime)
                    If jsonValue.Type = Newtonsoft.Json.Linq.JTokenType.Date Then
                        prop.SetValue(target, CDate(jsonValue))
                    ElseIf jsonValue.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                        Dim dtVal As DateTime
                        If DateTime.TryParse(jsonValue.ToString(), dtVal) Then
                            prop.SetValue(target, dtVal)
                        End If
                    End If

                Case GetType(SkillProcessingMode)
                    If jsonValue.Type = Newtonsoft.Json.Linq.JTokenType.Integer Then
                        prop.SetValue(target, CType(CInt(jsonValue), SkillProcessingMode))
                    ElseIf jsonValue.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                        Dim modeStr = jsonValue.ToString().Trim()
                        Dim modeInt As Integer
                        If Integer.TryParse(modeStr, modeInt) Then
                            prop.SetValue(target, CType(modeInt, SkillProcessingMode))
                        Else
                            Dim parsedMode As SkillProcessingMode
                            If [Enum].TryParse(modeStr, True, parsedMode) Then
                                prop.SetValue(target, parsedMode)
                            End If
                        End If
                    End If

                Case GetType(String())
                    If jsonValue.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                        Dim strArray = jsonValue.Select(Function(t) t?.ToString()).ToArray()
                        prop.SetValue(target, strArray)
                    End If

                Case Else
                    ' 通用枚举（SecurityLevel / ConcurrencySafety 等）：字符串或整数按名/值解析。
                    ' 原先只有 SkillProcessingMode 被特殊处理，securityLevel 等枚举字段在
                    ' JSON → SkillDefinition 时被静默丢弃。
                    If propType.IsEnum Then
                        If jsonValue.Type = Newtonsoft.Json.Linq.JTokenType.Integer Then
                            prop.SetValue(target, [Enum].ToObject(propType, CInt(jsonValue)))
                        Else
                            Dim enumVal As Object = Nothing
                            Try
                                enumVal = [Enum].Parse(propType, jsonValue.ToString().Trim(), True)
                            Catch
                                enumVal = Nothing
                            End Try
                            If enumVal IsNot Nothing Then prop.SetValue(target, enumVal)
                        End If
                        Exit Sub
                    End If

                    ' 对于复杂类型（List, Dictionary 等），使用 Newtonsoft.Json 的转换
                    If jsonValue.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                        If propType.IsGenericType Then
                            Dim itemType = propType.GetGenericArguments()(0)
                            Dim listType = GetType(List(Of)).MakeGenericType(itemType)
                            Dim list = Activator.CreateInstance(listType)
                            For Each item In CType(jsonValue, Newtonsoft.Json.Linq.JArray)
                                If item.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                                    Dim itemObj = item.ToObject(itemType)
                                    listType.GetMethod("Add").Invoke(list, New Object() {itemObj})
                                Else
                                    listType.GetMethod("Add").Invoke(list, New Object() {item?.ToString()})
                                End If
                            Next
                            prop.SetValue(target, list)
                        End If
                    End If
            End Select
        End Sub

        ''' <summary>
        ''' 从 JObject 解析 SkillDefinition（不区分大小写，反射自动处理所有简单类型字段）
        ''' 用于从 YAML frontmatter 或其他 JObject 来源加载 Skill
        ''' </summary>
        Public Function ParseSkillFromJObject(jObj As Newtonsoft.Json.Linq.JObject) As SkillDefinition
            If jObj Is Nothing Then Return Nothing

            Dim skill As New SkillDefinition()

            ' 已手动处理的复杂字段（数组、嵌套对象等）
            Dim handledFields As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            ' === 1. 数组字段（String()）===
            handledFields.Add("Keywords")
            If jObj("keywords") IsNot Nothing Then
                skill.Keywords = jObj("keywords").Select(Function(t) t?.ToString()).ToArray()
            End If

            handledFields.Add("Examples")
            If jObj("examples") IsNot Nothing Then
                skill.Examples = jObj("examples").Select(Function(t) t?.ToString()).ToArray()
            End If

            handledFields.Add("ConfigFields")
            If jObj("configFields") IsNot Nothing Then
                skill.ConfigFields = jObj("configFields").Select(Function(t) t?.ToString()).ToArray()
            End If

            handledFields.Add("InputContext")
            If jObj("inputContext") IsNot Nothing Then
                skill.InputContext = jObj("inputContext").Select(Function(t) t?.ToString()).ToArray()
            End If

            ' 工具白名单（agent 模式）。反射路径 SetPropertyValueFromJToken 不处理 String()，
            ' 缺失时 allowedTools 在 md 加载时会被静默丢弃，agent 模式退化成全量工具可见。
            handledFields.Add("AllowedTools")
            If jObj("allowedTools") IsNot Nothing Then
                skill.AllowedTools = jObj("allowedTools").Select(Function(t) t?.ToString()).ToArray()
            End If

            ' === 2. 列表字段 ===
            handledFields.Add("Tags")
            If jObj("tags") IsNot Nothing Then
                skill.Tags = New List(Of String)(jObj("tags").Select(Function(t) t?.ToString()))
            End If

            ' === 3. 复杂嵌套对象 ===
            handledFields.Add("UIActions")
            If jObj("UIActions") IsNot Nothing Then
                Dim uiObj = jObj("UIActions")
                skill.UIActions = New Skills.UIActionConfig()
                If uiObj("ShowCopy") IsNot Nothing Then
                    skill.UIActions.ShowCopy = Convert.ToBoolean(uiObj("ShowCopy").ToString())
                End If
                If uiObj("ShowSaveMd") IsNot Nothing Then
                    skill.UIActions.ShowSaveMd = Convert.ToBoolean(uiObj("ShowSaveMd").ToString())
                End If
                If uiObj("ShowRegenerate") IsNot Nothing Then
                    skill.UIActions.ShowRegenerate = Convert.ToBoolean(uiObj("ShowRegenerate").ToString())
                End If
            End If

            ' === 4. 反射处理所有简单类型属性 ===
            Dim simpleProps = GetType(SkillDefinition).GetProperties(
                System.Reflection.BindingFlags.Public Or
                System.Reflection.BindingFlags.Instance)

            For Each prop In simpleProps
                If handledFields.Contains(prop.Name) Then Continue For

                Dim yamlKey = prop.Name.Substring(0, 1).ToLower() & prop.Name.Substring(1)
                Dim jToken = FindJTokenCaseInsensitive(jObj, yamlKey)
                If jToken Is Nothing Then Continue For

                Try
                    SetPropertyValueFromJToken(skill, prop, jToken)
                Catch ex As Exception  '  忽略单个属性的转换错误
                End Try
            Next

            ' === 5. 处理复杂 List 类型（List(Of WorkflowStep) 等）===
            handledFields.Add("WorkflowSteps")
            If jObj("workflowSteps") IsNot Nothing Then
                Try
                    skill.WorkflowSteps = New List(Of WorkflowStep)()
                    For Each item In jObj("workflowSteps")
                        If item.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                            Dim ws = item.ToObject(Of WorkflowStep)()
                            skill.WorkflowSteps.Add(ws)
                        End If
                    Next
                Catch ex As Exception  ' 
                End Try
            End If

            ' === 声明式步骤（steps，SkillCreator 默认产物）===
            handledFields.Add("Steps")
            If jObj("steps") IsNot Nothing Then
                Try
                    skill.Steps = New List(Of DeclarativeStep)()
                    For Each item In jObj("steps")
                        If item.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                            Dim ds = item.ToObject(Of DeclarativeStep)()
                            ' params 兼容两种写法：嵌套 YAML 对象 或 单行 JSON 标量字符串
                            Dim paramsToken = CType(item, Newtonsoft.Json.Linq.JObject).Property("params")
                            If paramsToken IsNot Nothing AndAlso paramsToken.Value.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                                Dim raw = paramsToken.Value.ToString()
                                If Not String.IsNullOrWhiteSpace(raw) Then
                                    Try
                                        ds.Params = Newtonsoft.Json.Linq.JObject.Parse(raw).ToObject(Of Dictionary(Of String, Object))()
                                    Catch ex As Exception
                                        ds.Params = New Dictionary(Of String, Object)()
                                    End Try
                                End If
                            End If
                            skill.Steps.Add(ds)
                        End If
                    Next
                Catch ex As Exception  ' 
                End Try
            End If

            handledFields.Add("InputParameters")
            If jObj("inputParameters") IsNot Nothing Then
                Try
                    skill.InputParameters = New List(Of ParameterDefinition)()
                    For Each item In jObj("inputParameters")
                        If item.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                            Dim pd = item.ToObject(Of ParameterDefinition)()
                            skill.InputParameters.Add(pd)
                        End If
                    Next
                Catch ex As Exception  ' 
                End Try
            End If

            handledFields.Add("OutputParameters")
            If jObj("outputParameters") IsNot Nothing Then
                Try
                    skill.OutputParameters = New List(Of ParameterDefinition)()
                    For Each item In jObj("outputParameters")
                        If item.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                            Dim pd = item.ToObject(Of ParameterDefinition)()
                            skill.OutputParameters.Add(pd)
                        End If
                    Next
                Catch ex As Exception  ' 
                End Try
            End If

            handledFields.Add("CustomWorkflowSteps")
            If jObj("customWorkflowSteps") IsNot Nothing Then
                Try
                    skill.CustomWorkflowSteps = New List(Of WorkflowStep)()
                    For Each item In jObj("customWorkflowSteps")
                        If item.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                            ' 【v2 迁移】旧 CustomWorkflowStep/CustomStepType 数据归一化为统一 WorkflowStep：
                            ' 旧 stepType 名 CallAPI→CallApi、Conditional→Condition（其余同名）；
                            ' 旧数字序号按 v1 CustomStepType 声明序（0-7）映射为名字后反序列化。
                            Dim stepTypeToken = CType(item, Newtonsoft.Json.Linq.JObject).Property("stepType")
                            If stepTypeToken IsNot Nothing Then
                                Dim rawType As String = stepTypeToken.Value.ToString().Trim()
                                Dim normalizedType As String
                                Select Case rawType
                                    Case "0", "UploadImage" : normalizedType = "UploadImage"
                                    Case "1", "OCRRecognition" : normalizedType = "OCRRecognition"
                                    Case "2", "ExtractFields" : normalizedType = "ExtractFields"
                                    Case "3", "WriteToExcel" : normalizedType = "WriteToExcel"
                                    Case "4", "CallAPI" : normalizedType = "CallApi"
                                    Case "5", "CustomLLM" : normalizedType = "CustomLLM"
                                    Case "6", "Conditional" : normalizedType = "Condition"
                                    Case "7", "InvokeMcpTool" : normalizedType = "InvokeMcpTool"
                                    Case Else : normalizedType = rawType
                                End Select
                                stepTypeToken.Value = normalizedType
                            End If
                            Dim ws = item.ToObject(Of WorkflowStep)()
                            skill.CustomWorkflowSteps.Add(ws)
                        End If
                    Next
                Catch ex As Exception  '
                End Try
            End If

            handledFields.Add("FieldMappings")
            If jObj("fieldMappings") IsNot Nothing Then
                Try
                    skill.FieldMappings = New List(Of FieldMapping)()
                    For Each item In jObj("fieldMappings")
                        If item.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                            Dim fm = item.ToObject(Of FieldMapping)()
                            skill.FieldMappings.Add(fm)
                        End If
                    Next
                Catch ex As Exception  ' 
                End Try
            End If

            handledFields.Add("Parameters")
            If jObj("parameters") IsNot Nothing Then
                Try
                    skill.Parameters = New List(Of ParameterDefinition)()
                    For Each item In jObj("parameters")
                        If item.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                            Dim pd = item.ToObject(Of ParameterDefinition)()
                            skill.Parameters.Add(pd)
                        End If
                    Next
                Catch ex As Exception  ' 
                End Try
            End If

            handledFields.Add("ExcelConfig")
            If jObj("excelConfig") IsNot Nothing Then
                Try
                    Dim ecToken = jObj("excelConfig")
                    If ecToken.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                        ' 从 JSON 字符串反序列化
                        Dim ecJson = ecToken.ToString()
                        If Not String.IsNullOrWhiteSpace(ecJson) Then
                            skill.ExcelConfig = Newtonsoft.Json.JsonConvert.DeserializeObject(Of ExcelConfig)(ecJson)
                        End If
                    Else
                        skill.ExcelConfig = ecToken.ToObject(Of ExcelConfig)()
                    End If
                Catch ex As Exception  ' 
                End Try
            End If

            handledFields.Add("OcrConfig")
            If jObj("ocrConfig") IsNot Nothing Then
                Try
                    Dim ocrToken = jObj("ocrConfig")
                    If ocrToken.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                        Dim ocrJson = ocrToken.ToString()
                        If Not String.IsNullOrWhiteSpace(ocrJson) Then
                            skill.OcrConfig = Newtonsoft.Json.JsonConvert.DeserializeObject(Of OcrConfig)(ocrJson)
                        End If
                    Else
                        skill.OcrConfig = ocrToken.ToObject(Of OcrConfig)()
                    End If
                Catch ex As Exception  ' 
                End Try
            End If

            ' 手动解析 specFields（List(Of SpecFieldDefinition)）
            handledFields.Add("SpecFields")
            If jObj("specFields") IsNot Nothing Then
                Try
                    skill.SpecFields = New List(Of SpecFieldDefinition)()
                    For Each item In jObj("specFields")
                        If item.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                            skill.SpecFields.Add(item.ToObject(Of SpecFieldDefinition)())
                        End If
                    Next
                Catch ex As Exception  ' 
                End Try
            End If

            ' 手动解析 quickPrompts（List(Of QuickPromptDefinition)）
            handledFields.Add("QuickPrompts")
            If jObj("quickPrompts") IsNot Nothing Then
                Try
                    skill.QuickPrompts = New List(Of QuickPromptDefinition)()
                    For Each item In jObj("quickPrompts")
                        If item.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                            skill.QuickPrompts.Add(item.ToObject(Of QuickPromptDefinition)())
                        End If
                    Next
                Catch ex As Exception  ' 
                End Try
            End If

            ' === 6. 确保默认值 ===
            If String.IsNullOrWhiteSpace(skill.Name) Then skill.Name = "UnnamedSkill"
            If String.IsNullOrWhiteSpace(skill.DisplayName) Then skill.DisplayName = skill.Name
            If skill.Version <= 0 Then skill.Version = 1
            If skill.OutputFormat Is Nothing Then skill.OutputFormat = "text"

            Return skill
        End Function

        ''' <summary>
        ''' 在 JObject 中查找 JToken（不区分大小写）
        ''' </summary>
        Private Function FindJTokenCaseInsensitive(jObj As Newtonsoft.Json.Linq.JObject, key As String) As Newtonsoft.Json.Linq.JToken
            ' 精确匹配
            Dim val = jObj(key)
            If val IsNot Nothing Then Return val

            ' 不区分大小写匹配
            For Each prop In jObj.Properties()
                If String.Equals(prop.Name, key, StringComparison.OrdinalIgnoreCase) Then
                    Return prop.Value
                End If
            Next

            Return Nothing
        End Function

        ''' <summary>
        ''' 从 JToken 设置属性值（处理类型转换）
        ''' </summary>
        Private Sub SetPropertyValueFromJToken(target As Object, prop As System.Reflection.PropertyInfo, jToken As Newtonsoft.Json.Linq.JToken)
            Dim propType = prop.PropertyType

            ' 处理可空类型
            If Nullable.GetUnderlyingType(propType) IsNot Nothing Then
                propType = Nullable.GetUnderlyingType(propType)
            End If

            If jToken Is Nothing OrElse jToken.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Return

            Select Case propType
                Case GetType(String)
                    prop.SetValue(target, jToken.ToString())

                Case GetType(Integer)
                    If jToken.Type = Newtonsoft.Json.Linq.JTokenType.Integer Then
                        prop.SetValue(target, CInt(jToken))
                    Else
                        Dim intVal As Integer
                        If Integer.TryParse(jToken.ToString(), intVal) Then
                            prop.SetValue(target, intVal)
                        End If
                    End If

                Case GetType(Boolean)
                    If jToken.Type = Newtonsoft.Json.Linq.JTokenType.Boolean Then
                        prop.SetValue(target, CBool(jToken))
                    Else
                        prop.SetValue(target, Convert.ToBoolean(jToken.ToString()))
                    End If

                Case GetType(Double)
                    If jToken.Type = Newtonsoft.Json.Linq.JTokenType.Float Then
                        prop.SetValue(target, CDbl(jToken))
                    Else
                        Dim dblVal As Double
                        If Double.TryParse(jToken.ToString(), dblVal) Then
                            prop.SetValue(target, dblVal)
                        End If
                    End If

                Case GetType(DateTime)
                    If jToken.Type = Newtonsoft.Json.Linq.JTokenType.Date Then
                        prop.SetValue(target, CDate(jToken))
                    Else
                        Dim dtVal As DateTime
                        If DateTime.TryParse(jToken.ToString(), dtVal) Then
                            prop.SetValue(target, dtVal)
                        End If
                    End If

                Case Else
                    ' 处理枚举类型
                    If propType.IsEnum Then
                        If jToken.Type = Newtonsoft.Json.Linq.JTokenType.Integer Then
                            ' Integer → Enum：使用 Enum.ToObject
                            prop.SetValue(target, [Enum].ToObject(propType, CInt(jToken)))
                        Else
                            ' String → Enum：使用 Enum.Parse（不区分大小写）
                            Try
                                Dim parsed = [Enum].Parse(propType, jToken.ToString(), True)
                                prop.SetValue(target, parsed)
                            Catch ex As Exception  ' 
                            End Try
                        End If
                    End If
            End Select
        End Sub

        ''' <summary>
        ''' 精简序列化 SkillDefinition（仅包含核心字段）
        ''' 用于前端 data-skill-json 属性，避免序列化 50+ 字段导致 HTML 膨胀
        ''' 策略：手动构建字典，只输出有意义的字段
        ''' </summary>
        Public Function SerializeSkillCompact(skill As SkillDefinition) As String
            If skill Is Nothing Then Return "{}"

            Dim dict As New Dictionary(Of String, Object)()

            ' === 必填字段 ===
            dict("Name") = skill.Name
            dict("DisplayName") = skill.DisplayName
            If Not String.IsNullOrWhiteSpace(skill.Description) Then
                dict("Description") = skill.Description
            End If

            ' === 分类字段 ===
            If Not String.IsNullOrWhiteSpace(skill.Category) Then
                dict("Category") = skill.Category
            End If
            If skill.Tags IsNot Nothing AndAlso skill.Tags.Count > 0 Then
                dict("Tags") = skill.Tags
            End If

            ' === 核心配置 ===
            dict("ProcessingMode") = CInt(skill.ProcessingMode)
            If Not String.IsNullOrWhiteSpace(skill.TemplateCode) Then
                dict("TemplateCode") = skill.TemplateCode
            End If
            If Not String.IsNullOrWhiteSpace(skill.LLMSystemPrompt) Then
                dict("LLMSystemPrompt") = skill.LLMSystemPrompt
            End If
            If Not String.IsNullOrWhiteSpace(skill.LLMUserPromptTemplate) Then
                dict("LLMUserPromptTemplate") = skill.LLMUserPromptTemplate
            End If
            If Not String.IsNullOrWhiteSpace(skill.OutputFormat) AndAlso skill.OutputFormat <> "text" Then
                dict("OutputFormat") = skill.OutputFormat
            End If

            ' === 布尔/数值字段 ===
            If skill.IsGeneric Then dict("IsGeneric") = True
            If skill.NeedsConfig Then dict("NeedsConfig") = True
            If skill.IsComposite Then dict("IsComposite") = True
            If skill.SupportsImageInput Then dict("SupportsImageInput") = True
            If skill.SupportsFileInput Then dict("SupportsFileInput") = True
            If skill.AutoImportParseScript Then dict("AutoImportParseScript") = True
            If skill.IsCustom Then dict("IsCustom") = True
            If skill.Version <> 1 Then dict("Version") = skill.Version
            If skill.UsageCount > 0 Then dict("UsageCount") = skill.UsageCount

            ' === 声明式 steps 与路由关键字段 ===
            ' 【P0】这些字段原来是缺失的：SkillCreator 生成的技能经本函数序列化后 steps/allowedTools/
            ' pipeline/executionMode 全被剥掉，前端卡片保存回去时触发 SaveSkill 的
            ' NO_EXECUTABLE_IMPL 空壳校验（或退化成空实现）。此处必须与 SkillDefinition 保持同步。
            If skill.Steps IsNot Nothing AndAlso skill.Steps.Count > 0 Then
                dict("Steps") = skill.Steps
            End If
            If skill.AllowedTools IsNot Nothing AndAlso skill.AllowedTools.Length > 0 Then
                dict("AllowedTools") = skill.AllowedTools
            End If
            If Not String.IsNullOrWhiteSpace(skill.ExecutionMode) Then
                dict("ExecutionMode") = skill.ExecutionMode
            End If
            If Not String.IsNullOrWhiteSpace(skill.Pipeline) Then
                dict("Pipeline") = skill.Pipeline
            End If
            If skill.SecurityLevel <> SkillSecurityLevel.NeedsConfirmation Then
                dict("SecurityLevel") = skill.SecurityLevel.ToString()
            End If
            If skill.SpecFields IsNot Nothing AndAlso skill.SpecFields.Count > 0 Then
                dict("SpecFields") = skill.SpecFields
            End If
            If skill.QuickPrompts IsNot Nothing AndAlso skill.QuickPrompts.Count > 0 Then
                dict("QuickPrompts") = skill.QuickPrompts
            End If
            If skill.InputContext IsNot Nothing AndAlso skill.InputContext.Length > 0 Then
                dict("InputContext") = skill.InputContext
            End If

            ' === 数组字段（仅非空时输出）===
            If skill.Keywords IsNot Nothing AndAlso skill.Keywords.Length > 0 Then
                dict("Keywords") = skill.Keywords
            End If
            If skill.Examples IsNot Nothing AndAlso skill.Examples.Length > 0 Then
                dict("Examples") = skill.Examples
            End If
            If skill.ConfigFields IsNot Nothing AndAlso skill.ConfigFields.Length > 0 Then
                dict("ConfigFields") = skill.ConfigFields
            End If
            If skill.InputParameters IsNot Nothing AndAlso skill.InputParameters.Count > 0 Then
                dict("InputParameters") = skill.InputParameters
            End If
            If skill.OutputParameters IsNot Nothing AndAlso skill.OutputParameters.Count > 0 Then
                dict("OutputParameters") = skill.OutputParameters
            End If
            If skill.WorkflowSteps IsNot Nothing AndAlso skill.WorkflowSteps.Count > 0 Then
                dict("WorkflowSteps") = skill.WorkflowSteps
            End If
            If skill.CustomWorkflowSteps IsNot Nothing AndAlso skill.CustomWorkflowSteps.Count > 0 Then
                dict("CustomWorkflowSteps") = skill.CustomWorkflowSteps
            End If
            If skill.FieldMappings IsNot Nothing AndAlso skill.FieldMappings.Count > 0 Then
                dict("FieldMappings") = skill.FieldMappings
            End If

            ' === 元数据 ===
            dict("CreatedDate") = skill.CreatedDate.ToString("yyyy-MM-ddTHH:mm:ss")
            dict("LastModifiedDate") = skill.LastModifiedDate.ToString("yyyy-MM-ddTHH:mm:ss")
            If Not String.IsNullOrWhiteSpace(skill.CreatedBy) Then
                dict("CreatedBy") = skill.CreatedBy
            End If

            ' === 快捷入口 ===
            If skill.IsQuickShortcut Then
                dict("IsQuickShortcut") = True
                If Not String.IsNullOrWhiteSpace(skill.QuickShortcutIcon) Then
                    dict("QuickShortcutIcon") = skill.QuickShortcutIcon
                End If
                If Not String.IsNullOrWhiteSpace(skill.QuickShortcutPrompt) Then
                    dict("QuickShortcutPrompt") = skill.QuickShortcutPrompt
                End If
            End If

            ' === 可选配置 ===
            If skill.ExcelConfig IsNot Nothing Then
                dict("ExcelConfig") = skill.ExcelConfig
            End If
            If skill.OcrConfig IsNot Nothing Then
                dict("OcrConfig") = skill.OcrConfig
            End If

            ' 序列化为精简 JSON
            Dim settings As New Newtonsoft.Json.JsonSerializerSettings() With {
                .NullValueHandling = Newtonsoft.Json.NullValueHandling.Ignore
            }
            Return Newtonsoft.Json.JsonConvert.SerializeObject(dict, settings)
        End Function

    End Module

    ''' <summary>
    ''' 图片识别填入配置（配置驱动）
    ''' </summary>
    Public Class ImageFillConfig
        Public Property ImageType As String
        Public Property SheetName As String
        Public Property UseInvoiceMethod As Boolean
        Public Property HasImagePath As Boolean
    End Class

    ''' <summary>
    ''' 工作流持久化状态 - 用于支持长时间运行的工作流恢复
    ''' </summary>
    Public Class WorkflowPersistenceState
        Public Property WorkflowId As String
        Public Property SkillName As String
        Public Property CurrentStepId As String
        Public Property VariableStore As Dictionary(Of String, Object)
        Public Property ExecutionLog As List(Of String)
        Public Property StepExecutionCount As Dictionary(Of String, Integer)
        Public Property StartedAt As DateTime
        Public Property LastCheckpointAt As DateTime
        Public Property UserId As String
        Public Property Status As WorkflowStatus

        Public Enum WorkflowStatus
            InProgress
            Completed
            Failed
            Cancelled
        End Enum

        Public Sub New()
            VariableStore = New Dictionary(Of String, Object)()
            ExecutionLog = New List(Of String)()
            StepExecutionCount = New Dictionary(Of String, Integer)()
            StartedAt = DateTime.Now
            LastCheckpointAt = DateTime.Now
            Status = WorkflowStatus.InProgress
        End Sub
    End Class

    ''' <summary>
    ''' SkillManager - Skill 管理与执行核心类
    ''' 负责 Skill 的加载、缓存、热更新、工作流持久化
    ''' </summary>
    Partial Public Class SkillManager
        Implements IDisposable


        Private ReadOnly _persistenceFilePath As String
        Private ReadOnly _activeWorkflows As Dictionary(Of String, WorkflowPersistenceState)
        Private ReadOnly _persistenceLock As New Object()

        Private _skills As Dictionary(Of String, SkillDefinition)
        Private _workflowEngine As WorkflowEngine
        Private _llmEngine As MultimodalLLMEngine
        Private _configWatcher As Utils.ConfigFileWatcher
        Private _hotReloadEnabled As Boolean

        ' 热重载防抖（1 秒 trailing-edge，合并事件风暴并等待文件写完）
        Private _hotReloadDebounceTimer As System.Threading.Timer
        Private _hotReloadSyncLock As New Object()

        ' 缓存机制
        Private _skillsListCache As List(Of SkillDefinition)
        Private _skillsListCacheValid As Boolean = False
        Private _skillCategoriesCache As Dictionary(Of String, List(Of SkillDefinition))
        Private _skillCategoriesCacheValid As Boolean = False
        Private _cacheLock As New Object()

        ' 懒加载标记
        Private _workflowsLoaded As Boolean = False
        Private _llmInitialized As Boolean = False

        ' Skill执行队列，确保严格的顺序执行
        Private Shared _skillQueue As New Queue(Of Tuple(Of String, Dictionary(Of String, Object), IProgress(Of String), TaskCompletionSource(Of WorkflowResult)))
        Private Shared _isProcessingQueue As Boolean = False
        Private Shared _queueLock As New Object()

        Private Shared ReadOnly _lazyInstance As New Lazy(Of SkillManager)(Function() New SkillManager(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        Public Shared ReadOnly Property Instance As SkillManager
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        ''' <summary>
        ''' 配置热更新启用状态
        ''' </summary>
        Public Property HotReloadEnabled As Boolean
            Get
                Return _hotReloadEnabled
            End Get
            Set(value As Boolean)
                _hotReloadEnabled = value
                If value Then
                    StartConfigWatching()
                Else
                    StopConfigWatching()
                End If
            End Set
        End Property

        Private ReadOnly Property CurrentSkillVersion As Integer
            Get
                Return Integer.Parse(Utils.ConfigurationManager.Instance.GetSetting(
                    Utils.ConfigConstants.Workflow.VERSION,
                    Utils.ConfigConstants.Workflow.DEFAULT_VERSION.ToString()))
            End Get
        End Property



        Private Sub New()
            Try
                Utils.AppLogger.LogInfo("SkillManager", "初始化开始")
        
                _skills = New Dictionary(Of String, SkillDefinition)(StringComparer.OrdinalIgnoreCase)
                _persistenceFilePath = Utils.PathHelper.WorkflowStateFilePath
                _activeWorkflows = New Dictionary(Of String, WorkflowPersistenceState)(StringComparer.OrdinalIgnoreCase)
                _hotReloadEnabled = False
        
                Utils.AppLogger.LogInfo("SkillManager", "基础字段初始化完成")
        
                ' 只加载 Skills(轻量级操作)
                LoadSkills()
                Utils.AppLogger.LogInfo("SkillManager", $"Skills加载完成,共{_skills.Count}个")
        
                ' 其他组件采用懒加载
                Utils.AppLogger.LogInfo("SkillManager", "构造函数完成(懒加载模式)")
            Catch ex As Exception
                Utils.AppLogger.LogError("SkillManager", "初始化失败", ex)
                Throw
            End Try
        End Sub



        ''' <summary>
        ''' 获取 LLM 引擎（懒加载）
        ''' </summary>
        Private ReadOnly Property LLMEngine As MultimodalLLMEngine
            Get
                If _llmEngine Is Nothing Then
                    SyncLock _cacheLock
                        If _llmEngine Is Nothing Then
                            _llmEngine = MultimodalLLMEngine.SharedInstance
                            If Not _llmInitialized Then
                                InitializeLLMEngine()
                                _llmInitialized = True
                            End If
                        End If
                    End SyncLock
                End If
                Return _llmEngine
            End Get
        End Property

        ''' <summary>
        ''' 获取工作流引擎（懒加载）
        ''' </summary>
        Private ReadOnly Property WorkflowEngine As WorkflowEngine
            Get
                If _workflowEngine Is Nothing Then
                    SyncLock _cacheLock
                        If _workflowEngine Is Nothing Then
                            _workflowEngine = New WorkflowEngine(Me)
                            System.Diagnostics.Debug.WriteLine("WorkflowEngine 懒加载完成")
                        End If
                    End SyncLock
                End If
                Return _workflowEngine
            End Get
        End Property

        ''' <summary>
        ''' 确保配置监控器已初始化（懒加载）
        ''' </summary>
        Private Sub EnsureConfigWatcherInitialized()
            If _configWatcher Is Nothing Then
                SyncLock _cacheLock
                    If _configWatcher Is Nothing Then
                        _configWatcher = Utils.ConfigFileWatcher.Instance
                        AddHandler _configWatcher.ConfigFileChanged, AddressOf OnConfigFileChanged
                    End If
                End SyncLock
            End If
        End Sub

        ''' <summary>
        ''' 确保工作流已加载（懒加载）
        ''' </summary>
        Private Sub EnsureWorkflowsLoaded()
            If Not _workflowsLoaded Then
                SyncLock _persistenceLock
                    If Not _workflowsLoaded Then
                        LoadPersistedWorkflows()
                        _workflowsLoaded = True
                    End If
                End SyncLock
            End If
        End Sub


    End Class

End Namespace
