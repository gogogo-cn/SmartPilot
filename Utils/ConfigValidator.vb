Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Utils

    ''' <summary>
    ''' 配置验证结果
    ''' </summary>
    Public Class ConfigValidationResult
        Public Property IsValid As Boolean
        Public Property Errors As List(Of String)
        Public Property Warnings As List(Of String)

        Public Sub New()
            IsValid = True
            Errors = New List(Of String)()
            Warnings = New List(Of String)()
        End Sub

        Public Sub AddError(message As String)
            IsValid = False
            Errors.Add(message)
        End Sub

        Public Sub AddWarning(message As String)
            Warnings.Add(message)
        End Sub
    End Class

    ''' <summary>
    ''' 配置验证器 - 验证配置文件的正确性
    ''' </summary>
    Public Class ConfigValidator

        ''' <summary>
        ''' 验证数据类型配置文件
        ''' </summary>
        Public Shared Function ValidateDataTypeConfig(filePath As String) As ConfigValidationResult
            Dim result = New ConfigValidationResult()

            Try
                If Not File.Exists(filePath) Then
                    result.AddError(String.Format("配置文件不存在: {0}", filePath))
                    Return result
                End If

                Dim json = File.ReadAllText(filePath)
                Dim configObj As JObject

                Try
                    configObj = JObject.Parse(json)
                Catch ex As Exception
                    result.AddError(String.Format("JSON 格式无效: {0}", ex.Message))
                    Return result
                End Try

                ' 验证必需字段
                If configObj("version") Is Nothing Then
                    result.AddError("缺少必需字段: version")
                ElseIf Not Integer.TryParse(configObj("version").ToString(), New Integer) Then
                    result.AddError("version 必须是整数")
                End If

                If configObj("dataTypes") Is Nothing Then
                    result.AddError("缺少必需字段: dataTypes")
                    Return result
                End If

                Dim dataTypes = configObj("dataTypes")
                If Not dataTypes.HasValues Then
                    result.AddWarning("dataTypes 数组为空")
                End If

                Dim index = 0
                For Each dataType In dataTypes
                    ValidateDataTypeItem(dataType, index, result)
                    index += 1
                Next

            Catch ex As Exception
                result.AddError(String.Format("验证过程发生异常: {0}", ex.Message))
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 验证单个数据类型配置项
        ''' </summary>
        Private Shared Sub ValidateDataTypeItem(dataType As JToken, index As Integer, result As ConfigValidationResult)
            Dim prefix = String.Format("dataTypes[{0}]", index)

            ' 验证必需字段
            If dataType("typeName") Is Nothing OrElse String.IsNullOrWhiteSpace(dataType("typeName").ToString()) Then
                result.AddError(String.Format("{0}: 缺少必需字段 typeName", prefix))
            End If

            If dataType("displayName") Is Nothing OrElse String.IsNullOrWhiteSpace(dataType("displayName").ToString()) Then
                result.AddError(String.Format("{0}: 缺少必需字段 displayName", prefix))
            End If

            If dataType("defaultSheetName") Is Nothing Then
                result.AddWarning(String.Format("{0}: 缺少字段 defaultSheetName，将使用空值", prefix))
            End If

            ' [已弃用 - 识别类Skill已移除]
            'If dataType("fieldMappings") IsNot Nothing Then
            '    If Not TypeOf dataType("fieldMappings") Is JObject Then
            '        result.AddError(String.Format("{0}: fieldMappings 必须是对象类型", prefix))
            '    End If
            'End If

            ' 验证 detectionFields
            If dataType("detectionFields") IsNot Nothing Then
                If Not TypeOf dataType("detectionFields") Is JArray Then
                    result.AddError(String.Format("{0}: detectionFields 必须是数组类型", prefix))
                End If
            End If

            ' 验证 excelColumns
            If dataType("excelColumns") IsNot Nothing Then
                ValidateExcelColumns(dataType("excelColumns"), prefix, result)
            End If
        End Sub

        ''' <summary>
        ''' 验证 Excel 列配置
        ''' </summary>
        Private Shared Sub ValidateExcelColumns(columns As JToken, prefix As String, result As ConfigValidationResult)
            If Not TypeOf columns Is JArray Then
                result.AddError(String.Format("{0}: excelColumns 必须是数组类型", prefix))
                Return
            End If

            Dim colIndex = 0
            For Each col In columns
                Dim colPrefix = String.Format("{0}.excelColumns[{1}]", prefix, colIndex)

                If col("field") Is Nothing OrElse String.IsNullOrWhiteSpace(col("field").ToString()) Then
                    result.AddError(String.Format("{0}: 缺少必需字段 field", colPrefix))
                End If

                If col("header") Is Nothing OrElse String.IsNullOrWhiteSpace(col("header").ToString()) Then
                    result.AddError(String.Format("{0}: 缺少必需字段 header", colPrefix))
                End If

                If col("width") IsNot Nothing Then
                    Dim width As Integer
                    If Not Integer.TryParse(col("width").ToString(), width) OrElse width <= 0 Then
                        result.AddError(String.Format("{0}: width 必须是正整数", colPrefix))
                    End If
                End If

                colIndex += 1
            Next
        End Sub

        ''' <summary>
        ''' 验证 Skills 配置文件
        ''' </summary>
        Public Shared Function ValidateSkillsConfig(filePath As String) As ConfigValidationResult
            Dim result = New ConfigValidationResult()

            Try
                If Not File.Exists(filePath) Then
                    result.AddError(String.Format("配置文件不存在: {0}", filePath))
                    Return result
                End If

                Dim json = File.ReadAllText(filePath)
                Dim configObj As JObject

                Try
                    configObj = JObject.Parse(json)
                Catch ex As Exception
                    result.AddError(String.Format("JSON 格式无效: {0}", ex.Message))
                    Return result
                End Try

                ' 验证必需字段
                If configObj("version") Is Nothing Then
                    result.AddError("缺少必需字段: version")
                End If

                If configObj("skills") Is Nothing Then
                    result.AddError("缺少必需字段: skills")
                    Return result
                End If

                Dim skills = configObj("skills")
                If Not skills.HasValues Then
                    result.AddWarning("skills 数组为空")
                End If

                Dim index = 0
                For Each skill In skills
                    ValidateSkillItem(skill, index, result)
                    index += 1
                Next

            Catch ex As Exception
                result.AddError(String.Format("验证过程发生异常: {0}", ex.Message))
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 验证单个 Skill 配置项
        ''' </summary>
        Private Shared Sub ValidateSkillItem(skill As JToken, index As Integer, result As ConfigValidationResult)
            Dim prefix = String.Format("skills[{0}]", index)

            ' 验证必需字段
            If skill("name") Is Nothing OrElse String.IsNullOrWhiteSpace(skill("name").ToString()) Then
                result.AddError(String.Format("{0}: 缺少必需字段 name", prefix))
            End If

            If skill("displayName") Is Nothing OrElse String.IsNullOrWhiteSpace(skill("displayName").ToString()) Then
                result.AddError(String.Format("{0}: 缺少必需字段 displayName", prefix))
            End If

            If skill("description") Is Nothing OrElse String.IsNullOrWhiteSpace(skill("description").ToString()) Then
                result.AddWarning(String.Format("{0}: 缺少字段 description", prefix))
            End If

            ' 验证 processingMode
            If skill("processingMode") IsNot Nothing Then
                Dim validModes = New String() {"Traditional", "MultimodalLLM", "Hybrid"}
                Dim mode = skill("processingMode").ToString()
                If Not validModes.Contains(mode) Then
                    result.AddError(String.Format("{0}: processingMode 必须是 Traditional、MultimodalLLM 或 Hybrid 之一", prefix))
                End If
            End If

            ' 验证 outputFormat
            If skill("outputFormat") IsNot Nothing Then
                Dim validFormats = New String() {"text", "json", "vba", "xml", "html"}
                Dim format = skill("outputFormat").ToString().ToLower()
                If Not validFormats.Contains(format) Then
                    result.AddWarning(String.Format("{0}: outputFormat '{1}' 不是标准格式", prefix, format))
                End If
            End If

            ' 验证数组字段
            ValidateStringArrayField(skill, "keywords", prefix, result)
            ValidateStringArrayField(skill, "examples", prefix, result)
            ValidateStringArrayField(skill, "configFields", prefix, result)
            ValidateStringArrayField(skill, "tags", prefix, result)
        End Sub

        ''' <summary>
        ''' 验证字符串数组字段
        ''' </summary>
        Private Shared Sub ValidateStringArrayField(obj As JToken, fieldName As String, prefix As String, result As ConfigValidationResult)
            If obj(fieldName) IsNot Nothing Then
                If Not TypeOf obj(fieldName) Is JArray Then
                    result.AddError(String.Format("{0}: {1} 必须是数组类型", prefix, fieldName))
                End If
            End If
        End Sub

        ''' <summary>
        ''' 验证应用程序配置
        ''' </summary>
        Public Shared Function ValidateAppConfig(configManager As ConfigurationManager) As ConfigValidationResult
            Dim result = New ConfigValidationResult()

            Try
                ' 验证必需的 LLM 配置
                Dim provider = configManager.GetSetting(ConfigConstants.LLM.PROVIDER, ConfigConstants.LLM.DEFAULT_PROVIDER)
                If String.IsNullOrWhiteSpace(provider) Then
                    result.AddError(String.Format("缺少必需配置: {0}", ConfigConstants.LLM.PROVIDER))
                End If

                Dim apiKey = configManager.GetSetting(ConfigConstants.LLM.API_KEY, "")
                If String.IsNullOrWhiteSpace(apiKey) Then
                    result.AddWarning(String.Format("未配置 API Key: {0}", ConfigConstants.LLM.API_KEY))
                End If

                ' 验证数值配置
                ValidateNumericConfig(configManager, ConfigConstants.LLM.TEMPERATURE, 0.0, 2.0, result)
                ValidateNumericConfig(configManager, ConfigConstants.LLM.MAX_TOKENS, 1, 100000, result)
                ValidateNumericConfig(configManager, ConfigConstants.Workflow.HTTP_TIMEOUT, 1, 3600, result)

            Catch ex As Exception
                result.AddError(String.Format("验证过程发生异常: {0}", ex.Message))
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 验证数值配置范围
        ''' </summary>
        Private Shared Sub ValidateNumericConfig(configManager As ConfigurationManager, key As String, minValue As Double, maxValue As Double, result As ConfigValidationResult)
            Dim valueStr = configManager.GetSetting(key, "")
            If Not String.IsNullOrWhiteSpace(valueStr) Then
                Dim value As Double
                If Double.TryParse(valueStr, value) Then
                    If value < minValue OrElse value > maxValue Then
                        result.AddWarning(String.Format("配置 {0} 的值 {1} 不在推荐范围 [{2}, {3}] 内", key, value, minValue, maxValue))
                    End If
                Else
                    result.AddError(String.Format("配置 {0} 的值 '{1}' 不是有效的数字", key, valueStr))
                End If
            End If
        End Sub

    End Class

End Namespace
