Imports System.IO
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Utils

    ''' <summary>
    ''' JSON 序列化帮助类 - 提供统一的 JSON 序列化和反序列化方法
    ''' 消除代码重复，提供错误处理和日志记录
    ''' </summary>
    Public Class JsonHelper

        ' 默认的 JSON 序列化设置
        Private Shared ReadOnly _defaultSettings As JsonSerializerSettings
        
        ' 紧凑格式的 JSON 序列化设置（用于网络传输）
        Private Shared ReadOnly _compactSettings As JsonSerializerSettings
        
        ' 共享构造函数 - 在类首次被引用时执行
        Shared Sub New()
            ' 初始化 _defaultSettings
            _defaultSettings = New JsonSerializerSettings()
            _defaultSettings.NullValueHandling = NullValueHandling.Ignore
            _defaultSettings.ReferenceLoopHandling = ReferenceLoopHandling.Ignore
            _defaultSettings.Formatting = Formatting.Indented
            ' JSON字段名不区分大小写 - 通过反射设置，兼容旧版VB.NET编译器
            Try
                Dim prop = _defaultSettings.GetType().GetProperty("PropertyNameCaseInsensitive")
                If prop IsNot Nothing Then
                    prop.SetValue(_defaultSettings, True)
                End If
            Catch ex As Exception  '  忽略，旧版本不支持
            End Try
            
            ' 初始化 _compactSettings
            _compactSettings = New JsonSerializerSettings()
            _compactSettings.NullValueHandling = NullValueHandling.Ignore
            _compactSettings.ReferenceLoopHandling = ReferenceLoopHandling.Ignore
            _compactSettings.Formatting = Formatting.None
            ' JSON字段名不区分大小写 - 通过反射设置
            Try
                Dim prop = _compactSettings.GetType().GetProperty("PropertyNameCaseInsensitive")
                If prop IsNot Nothing Then
                    prop.SetValue(_compactSettings, True)
                End If
            Catch ex As Exception  '  忽略
            End Try
        End Sub

        ''' <summary>
        ''' 将对象序列化为 JSON 字符串
        ''' </summary>
        ''' <typeparam name="T">对象类型</typeparam>
        ''' <param name="obj">要序列化的对象</param>
        ''' <param name="compact">是否使用紧凑格式（无缩进）</param>
        ''' <returns>JSON 字符串</returns>
        Public Shared Function Serialize(Of T)(obj As T, Optional compact As Boolean = False) As String
            If obj Is Nothing Then
                Return String.Empty
            End If

            Try
                Dim settings = If(compact, _compactSettings, _defaultSettings)
                Return JsonConvert.SerializeObject(obj, settings)
            Catch ex As JsonException
                System.Diagnostics.Debug.WriteLine($"JSON 序列化失败: {ex.Message}")
                Throw New InvalidOperationException($"序列化类型 {GetType(T).Name} 失败", ex)
            End Try
        End Function

        ''' <summary>
        ''' 将对象序列化为 JSON 字符串（非泛型版本）
        ''' </summary>
        ''' <param name="obj">要序列化的对象</param>
        ''' <param name="compact">是否使用紧凑格式</param>
        ''' <returns>JSON 字符串</returns>
        Public Shared Function Serialize(obj As Object, Optional compact As Boolean = False) As String
            If obj Is Nothing Then
                Return String.Empty
            End If

            Try
                Dim settings = If(compact, _compactSettings, _defaultSettings)
                Return JsonConvert.SerializeObject(obj, settings)
            Catch ex As JsonException
                System.Diagnostics.Debug.WriteLine($"JSON 序列化失败: {ex.Message}")
                Throw New InvalidOperationException($"序列化对象失败", ex)
            End Try
        End Function

        ''' <summary>
        ''' 将 JSON 字符串反序列化为对象
        ''' </summary>
        ''' <typeparam name="T">目标类型</typeparam>
        ''' <param name="json">JSON 字符串</param>
        ''' <returns>反序列化的对象</returns>
        Public Shared Function Deserialize(Of T)(json As String) As T
            If String.IsNullOrWhiteSpace(json) Then
                Return Nothing
            End If

            Try
                Return JsonConvert.DeserializeObject(Of T)(json, _defaultSettings)
            Catch ex As JsonException
                System.Diagnostics.Debug.WriteLine($"JSON 反序列化失败: {ex.Message}")
                Throw New InvalidOperationException($"反序列化到类型 {GetType(T).Name} 失败", ex)
            End Try
        End Function

        ''' <summary>
        ''' 将 JSON 字符串反序列化为对象（非泛型版本）
        ''' </summary>
        ''' <param name="json">JSON 字符串</param>
        ''' <returns>反序列化的对象</returns>
        Public Shared Function Deserialize(json As String) As Object
            If String.IsNullOrWhiteSpace(json) Then
                Return Nothing
            End If

            Try
                Return JsonConvert.DeserializeObject(json, _defaultSettings)
            Catch ex As JsonException
                System.Diagnostics.Debug.WriteLine($"JSON 反序列化失败: {ex.Message}")
                Throw New InvalidOperationException("反序列化对象失败", ex)
            End Try
        End Function

        ''' <summary>
        ''' 尝试反序列化，失败时返回默认值
        ''' </summary>
        ''' <typeparam name="T">目标类型</typeparam>
        ''' <param name="json">JSON 字符串</param>
        ''' <param name="defaultValue">默认值</param>
        ''' <returns>反序列化的对象或默认值</returns>
        Public Shared Function TryDeserialize(Of T)(json As String, Optional defaultValue As T = Nothing) As T
            If String.IsNullOrWhiteSpace(json) Then
                Return defaultValue
            End If

            Try
                Return JsonConvert.DeserializeObject(Of T)(json, _defaultSettings)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"JSON 尝试反序列化失败: {ex.Message}")
                Return defaultValue
            End Try
        End Function

        ''' <summary>
        ''' 从文件加载并反序列化 JSON
        ''' </summary>
        ''' <typeparam name="T">目标类型</typeparam>
        ''' <param name="filePath">文件路径</param>
        ''' <returns>反序列化的对象</returns>
        Public Shared Function LoadFromFile(Of T)(filePath As String) As T
            If Not File.Exists(filePath) Then
                Throw New FileNotFoundException($"JSON 文件不存在: {filePath}")
            End If

            Try
                Dim json = File.ReadAllText(filePath)
                Return Deserialize(Of T)(json)
            Catch ex As IOException
                System.Diagnostics.Debug.WriteLine($"读取 JSON 文件失败: {ex.Message}")
                Throw
            End Try
        End Function

        ''' <summary>
        ''' 尝试从文件加载并反序列化 JSON，失败时返回默认值
        ''' </summary>
        ''' <typeparam name="T">目标类型</typeparam>
        ''' <param name="filePath">文件路径</param>
        ''' <param name="defaultValue">默认值</param>
        ''' <returns>反序列化的对象或默认值</returns>
        Public Shared Function TryLoadFromFile(Of T)(filePath As String, Optional defaultValue As T = Nothing) As T
            If Not File.Exists(filePath) Then
                Return defaultValue
            End If

            Try
                Dim json = File.ReadAllText(filePath)
                Return TryDeserialize(Of T)(json, defaultValue)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"尝试读取 JSON 文件失败: {ex.Message}")
                Return defaultValue
            End Try
        End Function

        ''' <summary>
        ''' 将对象序列化并保存到文件
        ''' </summary>
        ''' <typeparam name="T">对象类型</typeparam>
        ''' <param name="obj">要保存的对象</param>
        ''' <param name="filePath">文件路径</param>
        Public Shared Sub SaveToFile(Of T)(obj As T, filePath As String)
            If obj Is Nothing Then
                Throw New ArgumentNullException(NameOf(obj))
            End If

            Try
                PathHelper.EnsureDirectoryExists(filePath)
                Dim json = Serialize(obj)
                File.WriteAllText(filePath, json)
            Catch ex As IOException
                System.Diagnostics.Debug.WriteLine($"保存 JSON 文件失败: {ex.Message}")
                Throw
            End Try
        End Sub

        ''' <summary>
        ''' 尝试将对象序列化并保存到文件
        ''' </summary>
        ''' <typeparam name="T">对象类型</typeparam>
        ''' <param name="obj">要保存的对象</param>
        ''' <param name="filePath">文件路径</param>
        ''' <returns>是否成功</returns>
        Public Shared Function TrySaveToFile(Of T)(obj As T, filePath As String) As Boolean
            If obj Is Nothing Then
                Return False
            End If

            Try
                SaveToFile(obj, filePath)
                Return True
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"尝试保存 JSON 文件失败: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 验证字符串是否为有效的 JSON
        ''' </summary>
        ''' <param name="json">要验证的字符串</param>
        ''' <returns>是否为有效的 JSON</returns>
        Public Shared Function IsValidJson(json As String) As Boolean
            If String.IsNullOrWhiteSpace(json) Then
                Return False
            End If

            Try
                JToken.Parse(json)
                Return True
            Catch ex As Exception  ' 
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 将 JSON 字符串格式化为易读的格式
        ''' </summary>
        ''' <param name="json">JSON 字符串</param>
        ''' <returns>格式化后的 JSON 字符串</returns>
        Public Shared Function FormatJson(json As String) As String
            If String.IsNullOrWhiteSpace(json) Then
                Return String.Empty
            End If

            Try
                Dim token = JToken.Parse(json)
                Return token.ToString(Formatting.Indented)
            Catch ex As JsonException
                Return json
            End Try
        End Function

        ''' <summary>
        ''' 解析 JSON 字符串为 JObject（兼容旧 API 名称）
        ''' </summary>
        Public Shared Function ParseJson(json As String) As JObject
            If String.IsNullOrWhiteSpace(json) Then
                Return Nothing
            End If

            Try
                Return JObject.Parse(json)
            Catch ex As JsonException
                System.Diagnostics.Debug.WriteLine($"JSON 解析失败: {ex.Message}")
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 兼容旧的 SerializeToJson API
        ''' </summary>
        Public Shared Function SerializeToJson(obj As Object) As String
            Return Serialize(obj)
        End Function

        ''' <summary>
        ''' 尝试获取 JObject 中的键，返回是否存在（兼容旧测试用例）
        ''' </summary>
        Public Shared Function TryGetValue(jObj As JObject, key As String, defaultValue As String) As Boolean
            If jObj Is Nothing Then Return False
            Dim token = jObj.SelectToken(key)
            Return token IsNot Nothing
        End Function

        ''' <summary>
        ''' 获取 JObject 中的数组值，若不存在返回空 JArray（兼容旧测试用例）
        ''' </summary>
        Public Shared Function GetArray(jObj As JObject, key As String) As JArray
            If jObj Is Nothing Then Return New JArray()
            Dim token = jObj.SelectToken(key)
            If token Is Nothing Then Return New JArray()
            If TypeOf token Is JArray Then
                Return CType(token, JArray)
            Else
                Dim arr As New JArray()
                arr.Add(token)
                Return arr
            End If
        End Function

        ''' <summary>
        ''' 合并两个 JSON 对象
        ''' </summary>
        ''' <param name="original">原始 JSON</param>
        ''' <param name="update">更新的 JSON</param>
        ''' <returns>合并后的 JSON 字符串</returns>
        Public Shared Function MergeJson(original As String, update As String) As String
            If String.IsNullOrWhiteSpace(original) Then
                Return update
            End If

            If String.IsNullOrWhiteSpace(update) Then
                Return original
            End If

            Try
                Dim originalObj = JObject.Parse(original)
                Dim updateObj = JObject.Parse(update)
                originalObj.Merge(updateObj, New JsonMergeSettings With {
                    .MergeArrayHandling = MergeArrayHandling.Union
                })
                Return originalObj.ToString(Formatting.Indented)
            Catch ex As JsonException
                System.Diagnostics.Debug.WriteLine($"合并 JSON 失败: {ex.Message}")
                Throw
            End Try
        End Function

    End Class

End Namespace
