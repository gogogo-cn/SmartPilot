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

    ' SkillManager.YamlHelpers.vb — YAML / JSON 辅助方法
    ' （类级 XML 文档只在主文件 SkillManager.vb 保留一份，避免 BC42314 分部重复）
    Partial Public Class SkillManager

        ''' <summary>
        ''' 简单 YAML 转 JSON（处理基本缩进、列表、键值对）
        ''' </summary>
        Private Function ConvertYamlToJObject(yamlText As String) As Newtonsoft.Json.Linq.JObject
            Dim logPath = System.IO.Path.Combine(Utils.PathHelper.AppDataRoot, "yaml_debug.log")
            Try
                If String.IsNullOrWhiteSpace(yamlText) Then 
                    System.IO.File.AppendAllText(logPath, $"[YAML] 空文本{Environment.NewLine}")
                    Return New Newtonsoft.Json.Linq.JObject() 
                End If

                System.IO.File.AppendAllText(logPath, $"[YAML] 开始解析,长度:{yamlText.Length}{Environment.NewLine}")
                System.IO.File.AppendAllText(logPath, $"[YAML] 前100字符:{yamlText.Substring(0, Math.Min(100, yamlText.Length))}{Environment.NewLine}")

                Dim yamlStream = New YamlStream()
                yamlStream.Load(New System.IO.StringReader(yamlText))

                System.IO.File.AppendAllText(logPath, $"[YAML] Documents.Count:{yamlStream.Documents.Count}{Environment.NewLine}")

                If yamlStream.Documents.Count = 0 Then 
                    System.IO.File.AppendAllText(logPath, $"[YAML] 错误:Documents为空{Environment.NewLine}")
                    Return New Newtonsoft.Json.Linq.JObject() 
                End If

                Dim root = yamlStream.Documents(0).RootNode
                System.IO.File.AppendAllText(logPath, $"[YAML] RootNode类型:{root.GetType().Name}{Environment.NewLine}")

                Dim jsonObj = ConvertYamlNodeToJToken(root)
                System.IO.File.AppendAllText(logPath, $"[YAML] jsonObj类型:{jsonObj?.GetType().Name}{Environment.NewLine}")

                If TypeOf jsonObj Is Newtonsoft.Json.Linq.JObject Then
                    System.IO.File.AppendAllText(logPath, $"[YAML] 成功{Environment.NewLine}")
                    Return CType(jsonObj, Newtonsoft.Json.Linq.JObject)
                End If

                System.IO.File.AppendAllText(logPath, $"[YAML] 错误:jsonObj不是JObject{Environment.NewLine}")
                Return New Newtonsoft.Json.Linq.JObject()
            Catch ex As Exception
                System.IO.File.AppendAllText(logPath, $"[YAML] 解析失败:{ex.Message}{Environment.NewLine}{ex.StackTrace}{Environment.NewLine}")
                Return New Newtonsoft.Json.Linq.JObject()
            End Try
        End Function

        ''' <summary>
        ''' 递归将 YamlDotNet 节点转换为 Newtonsoft.Json JToken
        ''' </summary>
        Private Function ConvertYamlNodeToJToken(node As YamlDotNet.RepresentationModel.YamlNode) As Newtonsoft.Json.Linq.JToken
            If TypeOf node Is YamlMappingNode Then
                Dim mapping = CType(node, YamlMappingNode)
                Dim obj As New Newtonsoft.Json.Linq.JObject()
                For Each kvp In mapping.Children
                    Dim key = kvp.Key.ToString()
                    obj(key) = ConvertYamlNodeToJToken(kvp.Value)
                Next
                Return obj
            ElseIf TypeOf node Is YamlSequenceNode Then
                Dim sequence = CType(node, YamlSequenceNode)
                Dim arr As New Newtonsoft.Json.Linq.JArray()
                For Each item In sequence.Children
                    arr.Add(ConvertYamlNodeToJToken(item))
                Next
                Return arr
            ElseIf TypeOf node Is YamlScalarNode Then
                Dim scalar = CType(node, YamlScalarNode)
                Dim value = scalar.Value

                If value IsNot Nothing Then
                    Dim lowerValue = value.ToLower()
                    If lowerValue = "true" Then Return New Newtonsoft.Json.Linq.JValue(True)
                    If lowerValue = "false" Then Return New Newtonsoft.Json.Linq.JValue(False)
                    If lowerValue = "null" OrElse lowerValue = "~" OrElse lowerValue = "" Then Return Newtonsoft.Json.Linq.JValue.CreateNull()

                    Dim intResult As Integer
                    If Integer.TryParse(value, intResult) Then Return New Newtonsoft.Json.Linq.JValue(intResult)
                End If

                Return New Newtonsoft.Json.Linq.JValue(value)
            End If

            Return Newtonsoft.Json.Linq.JValue.CreateNull()
        End Function

        Private Function EscapeJson(value As String) As String
            If String.IsNullOrEmpty(value) Then Return ""
            Dim sb = New System.Text.StringBuilder()
            For Each ch As Char In value
                If ch = Chr(34) Then
                    sb.Append(""""c)
                ElseIf ch = Chr(92) Then
                    sb.Append("\"c)
                ElseIf ch = vbCr Then
                    sb.Append("\r")
                ElseIf ch = vbLf Then
                    sb.Append("\n")
                ElseIf ch = vbTab Then
                    sb.Append("\t")
                Else
                    sb.Append(ch)
                End If
            Next
            Return sb.ToString()
        End Function

        Private Function GetJsonString(obj As Newtonsoft.Json.Linq.JObject, key As String, Optional defaultValue As String = "") As String
            If obj.ContainsKey(key) Then
                Dim token = obj(key)
                If token IsNot Nothing Then Return token.ToString()
            End If
            Return defaultValue
        End Function

        Private Function GetJsonBool(obj As Newtonsoft.Json.Linq.JObject, key As String, Optional defaultValue As Boolean = False) As Boolean
            If obj.ContainsKey(key) Then
                Dim token = obj(key)
                If token IsNot Nothing Then
                    Dim strValue = token.ToString().Trim()
                    If strValue.Equals("true", StringComparison.OrdinalIgnoreCase) Then Return True
                    If strValue.Equals("false", StringComparison.OrdinalIgnoreCase) Then Return False
                End If
            End If
            Return defaultValue
        End Function

        Private Function GetJsonInt(obj As Newtonsoft.Json.Linq.JObject, key As String, Optional defaultValue As Integer = 0) As Integer
            If obj.ContainsKey(key) Then
                Dim token = obj(key)
                If token IsNot Nothing Then
                    Dim result As Integer
                    If Integer.TryParse(token.ToString(), result) Then Return result
                End If
            End If
            Return defaultValue
        End Function

        Private Function GetJsonStringArray(obj As Newtonsoft.Json.Linq.JObject, key As String) As String()
            If Not obj.ContainsKey(key) Then Return New String() {}
            Dim token = obj(key)
            If token Is Nothing Then Return New String() {}
            If token.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                Return token.Select(Function(t) t.ToString()).ToArray()
            End If
            Dim str = token.ToString()
            If String.IsNullOrEmpty(str) Then Return New String() {}
            Return New String() {str}
        End Function

        Private Function GetJsonStringList(obj As Newtonsoft.Json.Linq.JObject, key As String) As List(Of String)
            Dim arr = GetJsonStringArray(obj, key)
            If arr Is Nothing Then Return New List(Of String)()
            Return arr.ToList()
        End Function

        ''' <summary>
        ''' 解析 frontmatter 中的 quickPrompts 对象数组：
        '''   quickPrompts:
        '''     - label: 商务蓝
        '''       prompt: 用商务蓝主题美化表格
        ''' 仅接受 label+prompt 都非空的对象，其余忽略。
        ''' </summary>
        Private Function GetJsonQuickPrompts(obj As Newtonsoft.Json.Linq.JObject, key As String) As List(Of QuickPromptDefinition)
            Dim result = New List(Of QuickPromptDefinition)()
            If obj Is Nothing OrElse Not obj.ContainsKey(key) Then Return result
            Dim token = obj(key)
            If token Is Nothing OrElse token.Type <> Newtonsoft.Json.Linq.JTokenType.Array Then Return result

            For Each item As Newtonsoft.Json.Linq.JToken In token
                If item.Type <> Newtonsoft.Json.Linq.JTokenType.Object Then Continue For
                Dim itemObj = CType(item, Newtonsoft.Json.Linq.JObject)
                Dim label = GetJsonString(itemObj, "label", "")
                Dim prompt = GetJsonString(itemObj, "prompt", "")
                If Not String.IsNullOrWhiteSpace(label) AndAlso Not String.IsNullOrWhiteSpace(prompt) Then
                    result.Add(New QuickPromptDefinition() With {.Label = label.Trim(), .Prompt = prompt.Trim()})
                End If
            Next
            Return result
        End Function

        ''' <summary>
        ''' 解析 frontmatter 中的 uiActions 对象（键名大小写不敏感，兼容 md 的小写
        ''' showCopy/showSaveMd/showRegenerate 与 CustomSkillManager 写回的 PascalCase）：
        '''   uiActions:
        '''     showCopy: true
        '''     showSaveMd: true
        '''     showRegenerate: true
        ''' 缺省值沿用 UIActionConfig 默认（ShowCopy/ShowRegenerate=True，ShowSaveMd=False）。
        ''' 背景：此前 Markdown 加载路径从未解析 uiActions，导致 GenerateOCRParseScript
        ''' 这类技能到前端后 skill.uiActions=null，可读脚本卡（复制/保存/重新生成）永不渲染。
        ''' </summary>
        Private Function GetJsonUIActions(obj As Newtonsoft.Json.Linq.JObject, key As String) As UIActionConfig
            If obj Is Nothing Then Return Nothing
            Dim token As Newtonsoft.Json.Linq.JToken = Nothing
            For Each p As Newtonsoft.Json.Linq.JProperty In obj.Properties()
                If String.Equals(p.Name, key, StringComparison.OrdinalIgnoreCase) Then
                    token = p.Value
                    Exit For
                End If
            Next
            If token Is Nothing OrElse token.Type <> Newtonsoft.Json.Linq.JTokenType.Object Then Return Nothing
            Dim o = CType(token, Newtonsoft.Json.Linq.JObject)
            Dim cfg = New UIActionConfig()
            cfg.ShowCopy = ReadBoolInsensitive(o, "showCopy", cfg.ShowCopy)
            cfg.ShowSaveMd = ReadBoolInsensitive(o, "showSaveMd", cfg.ShowSaveMd)
            cfg.ShowRegenerate = ReadBoolInsensitive(o, "showRegenerate", cfg.ShowRegenerate)
            Return cfg
        End Function

        ''' <summary>大小写不敏感地读取布尔字段（"true"/"false"），缺失时返回默认值</summary>
        Private Function ReadBoolInsensitive(o As Newtonsoft.Json.Linq.JObject, key As String, dflt As Boolean) As Boolean
            If o Is Nothing Then Return dflt
            For Each p As Newtonsoft.Json.Linq.JProperty In o.Properties()
                If String.Equals(p.Name, key, StringComparison.OrdinalIgnoreCase) Then
                    Dim s = If(p.Value Is Nothing, "", p.Value.ToString().Trim())
                    If s.Equals("true", StringComparison.OrdinalIgnoreCase) Then Return True
                    If s.Equals("false", StringComparison.OrdinalIgnoreCase) Then Return False
                End If
            Next
            Return dflt
        End Function

        ''' <summary>
        ''' 解析 frontmatter 中的 specFields 对象数组：
        '''   specFields:
        '''     - name: theme
        '''       label: 主题风格
        '''       type: select
        '''       options:
        '''         - business_blue
        '''       defaultValue: business_blue
        ''' 仅接受 name 非空的对象，其余忽略。
        ''' </summary>
        Private Function GetJsonSpecFields(obj As Newtonsoft.Json.Linq.JObject, key As String) As List(Of SpecFieldDefinition)
            Dim result = New List(Of SpecFieldDefinition)()
            If obj Is Nothing OrElse Not obj.ContainsKey(key) Then Return result
            Dim token = obj(key)
            If token Is Nothing OrElse token.Type <> Newtonsoft.Json.Linq.JTokenType.Array Then Return result

            For Each item As Newtonsoft.Json.Linq.JToken In token
                If item.Type <> Newtonsoft.Json.Linq.JTokenType.Object Then Continue For
                Dim itemObj = CType(item, Newtonsoft.Json.Linq.JObject)
                Dim name = GetJsonString(itemObj, "name", "")
                If String.IsNullOrWhiteSpace(name) Then Continue For

                Dim options = New List(Of String)()
                If itemObj.ContainsKey("options") Then
                    Dim optToken = itemObj("options")
                    If optToken IsNot Nothing AndAlso optToken.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                        For Each opt In optToken
                            If opt.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                                options.Add(CStr(opt))
                            End If
                        Next
                    End If
                End If

                result.Add(New SpecFieldDefinition() With {
                    .Name = name.Trim(),
                    .Label = GetJsonString(itemObj, "label", name.Trim()),
                    .Type = GetJsonString(itemObj, "type", "string"),
                    .Options = options,
                    .DefaultValue = GetJsonString(itemObj, "defaultValue", "")
                })
            Next
            Return result
        End Function

        ''' <summary>
        ''' 从 Markdown 正文中提取指定二级标题下的文本内容
        ''' </summary>
        Private Function ExtractMarkdownSection(markdown As String, sectionTitle As String) As String
            If String.IsNullOrWhiteSpace(markdown) Then Return ""
            Dim pattern = "## " & sectionTitle & vbCrLf
            Dim idx = markdown.IndexOf(pattern, StringComparison.OrdinalIgnoreCase)
            If idx < 0 Then
                pattern = "## " & sectionTitle & vbLf
                idx = markdown.IndexOf(pattern, StringComparison.OrdinalIgnoreCase)
            End If
            If idx < 0 Then Return ""

            Dim start = idx + pattern.Length
            ' 查找下一个二级标题时，跳过三级及以上标题（###, #### 等）
            Dim nextHeading = FindNextH2(markdown, start)
            Dim sectionText As String
            If nextHeading > 0 Then
                sectionText = markdown.Substring(start, nextHeading - start).Trim()
            Else
                sectionText = markdown.Substring(start).Trim()
            End If
            Return sectionText
        End Function

        ''' <summary>
        ''' 查找下一个二级标题（## 开头，且 ## 后不是 #），跳过 ###、#### 等
        ''' </summary>
        Private Function FindNextH2(markdown As String, start As Integer) As Integer
            Dim pos = start
            While pos < markdown.Length
                Dim found = markdown.IndexOf("## ", pos)
                If found < 0 Then Return -1
                ' 检查 ## 后面是否不是另一个 #（即排除 ###, #### 等）
                If found + 3 < markdown.Length AndAlso markdown(found + 3) <> "#"c Then
                    ' 还要确保这一行以 ## 开头（前面是换行或开头）
                    Dim lineStart = markdown.LastIndexOf(vbLf, found)
                    If lineStart < 0 Then lineStart = 0 Else lineStart = lineStart + 1
                    Dim prefix = markdown.Substring(lineStart, found - lineStart).Trim()
                    If prefix.Length = 0 Then
                        Return found
                    End If
                End If
                pos = found + 3
            End While
            Return -1
        End Function

        ''' <summary>
        ''' 从 Markdown 正文中提取指定二级标题下的代码块内容
        ''' </summary>
        Private Function ExtractMarkdownCodeBlock(markdown As String, sectionTitle As String) As String
            Dim sectionText = ExtractMarkdownSection(markdown, sectionTitle)
            If String.IsNullOrWhiteSpace(sectionText) Then Return ""

            ' 查找 ```vba 或 ``` 代码块
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
        ''' 重新加载所有 Skills（支持热更新）
        ''' </summary>
        Public Function ReloadSkills() As Boolean
            Try
                System.Diagnostics.Debug.WriteLine("重新加载所有 Skills...")
                LoadSkills()
                Return True
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("重新加载 Skills 失败: " & ex.Message)
                Return False
            End Try
        End Function

    End Class

End Namespace
