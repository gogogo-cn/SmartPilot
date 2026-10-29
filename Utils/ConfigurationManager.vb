Option Strict On

Imports System.IO
Imports System.Collections.Generic
Imports System.Threading
Imports System.Text

Namespace Utils

    ''' <summary>
    ''' 配置管理类 - 管理应用程序配置和API密钥
    ''' 使用 DPAPI 加密敏感数据
    ''' 优化：批量保存、延迟写入、解密缓存
    ''' </summary>
    Public Class ConfigurationManager
        Private _configFilePath As String
        Private _config As Dictionary(Of String, String)

        ' 需要加密的敏感配置项关键字（含实际使用的键名变体）
        ' 【修复】补充大写下划线格式键名：LLM/OCR相关配置
        Private ReadOnly _sensitiveKeys As HashSet(Of String) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            "LLMApiKey", "BaiduOcrApiKey", "BaiduOcrSecretKey",
            "AzureApiKey", "OpenAIApiKey", "QianwenApiKey",
            "ZhipuApiKey", "WenxinApiKey", "MoonshotApiKey",
            "ApiKey", "SecretKey", "Password", "Token",
            "LLM_API_KEY", "LLM_PROVIDER", "LLM_MODEL", "LLM_API_URL",
            "OCR_API_KEY", "OCR_SECRET_KEY", "OCR_CLOUD_API_KEY", "OCR_CLOUD_SECRET_KEY",
            "LICENSE_SESSION_TOKEN", "LICENSE_ACTIVATION"
        }

        ' 延迟保存相关
        Private _pendingSave As Boolean = False
        Private _saveTimer As Timer
        Private _saveDelayMs As Integer = 5000 ' 5秒延迟
        Private _saveLock As New Object()

        ' 解密缓存（带大小限制和过期机制）
        Private _decryptedCache As New Dictionary(Of String, String)()
        Private _cacheLock As New Object()
        Private Const MAX_CACHE_ENTRIES As Integer = 50
        Private _cacheAccessTime As New Dictionary(Of String, DateTime)()

        Private Shared ReadOnly _lazyInstance As New Lazy(Of ConfigurationManager)(Function() New ConfigurationManager(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        Public Shared ReadOnly Property Instance As ConfigurationManager
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        Private Sub New()
            _configFilePath = PathHelper.ConfigFilePath
            _config = New Dictionary(Of String, String)
            PathHelper.EnsureDirectoryExists(_configFilePath)
            LoadConfig()
        End Sub

        ''' <summary>
        ''' 获取设置值（自动解密敏感数据，使用缓存）
        ''' </summary>
        Public Function GetSetting(key As String, Optional defaultValue As String = "") As String
            SyncLock _cacheLock
                If _decryptedCache.ContainsKey(key) Then
                    _cacheAccessTime(key) = DateTime.Now
                    Return _decryptedCache(key)
                End If
            End SyncLock

            Dim value As String = defaultValue
            If _config.ContainsKey(key) Then
                value = _config(key)
            End If

            If Not String.IsNullOrEmpty(value) AndAlso IsSensitiveKey(key) Then
                If EncryptionHelper.IsEncrypted(value) Then
                    Try
                        value = EncryptionHelper.Decrypt(value)
                        SyncLock _cacheLock
                            If _decryptedCache.Count >= MAX_CACHE_ENTRIES Then
                                Dim oldestKey = _cacheAccessTime.OrderBy(Function(kvp) kvp.Value).FirstOrDefault().Key
                                If Not String.IsNullOrEmpty(oldestKey) Then
                                    _decryptedCache.Remove(oldestKey)
                                    _cacheAccessTime.Remove(oldestKey)
                                End If
                            End If
                            _decryptedCache(key) = value
                            _cacheAccessTime(key) = DateTime.Now
                        End SyncLock
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine($"解密配置项: {key} 失败: {ex.Message}")
                        Return defaultValue
                    End Try
                End If
            End If

            Return value
        End Function

        ''' <summary>
        ''' 设置值（自动加密敏感数据，延迟保存）
        ''' </summary>
        Public Sub SetSetting(key As String, value As String)
            Dim originalValue As String = Nothing
            If _config.ContainsKey(key) Then
                originalValue = _config(key)
            End If

            If String.Equals(originalValue, value, StringComparison.Ordinal) Then
                Return
            End If

            If String.IsNullOrEmpty(value) Then
                _config(key) = value
                SyncLock _cacheLock
                    _decryptedCache.Remove(key)
                    _cacheAccessTime.Remove(key)
                End SyncLock
            ElseIf IsSensitiveKey(key) Then
                Try
                    _config(key) = EncryptionHelper.Encrypt(value)
                    SyncLock _cacheLock
                        If _decryptedCache.Count >= MAX_CACHE_ENTRIES Then
                            Dim oldestKey = _cacheAccessTime.OrderBy(Function(k) k.Value).FirstOrDefault().Key
                            If Not String.IsNullOrEmpty(oldestKey) Then
                                _decryptedCache.Remove(oldestKey)
                                _cacheAccessTime.Remove(oldestKey)
                            End If
                        End If
                        _decryptedCache(key) = value
                        _cacheAccessTime(key) = DateTime.Now
                    End SyncLock
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"加密配置项: {key} 失败: {ex.Message}")
                    Throw
                End Try
            Else
                _config(key) = value
            End If

            ScheduleSave()
        End Sub

        ''' <summary>
        ''' 批量设置（减少IO操作）
        ''' </summary>
        Public Sub SetSettings(settings As Dictionary(Of String, String))
            If settings Is Nothing OrElse settings.Count = 0 Then
                Return
            End If

            Using PerformanceMonitor.Instance.StartOperation("ConfigurationManager.SetSettings")
                For Each kvp In settings
                    Dim value = kvp.Value
                    If String.IsNullOrEmpty(value) Then
                        _config(kvp.Key) = value
                        SyncLock _cacheLock
                            _decryptedCache.Remove(kvp.Key)
                            _cacheAccessTime.Remove(kvp.Key)
                        End SyncLock
                    ElseIf IsSensitiveKey(kvp.Key) Then
                        Try
                            _config(kvp.Key) = EncryptionHelper.Encrypt(value)
                            SyncLock _cacheLock
                                If _decryptedCache.Count >= MAX_CACHE_ENTRIES Then
                                    Dim oldestKey = _cacheAccessTime.OrderBy(Function(k) k.Value).FirstOrDefault().Key
                                    If Not String.IsNullOrEmpty(oldestKey) Then
                                        _decryptedCache.Remove(oldestKey)
                                        _cacheAccessTime.Remove(oldestKey)
                                    End If
                                End If
                                _decryptedCache(kvp.Key) = value
                                _cacheAccessTime(kvp.Key) = DateTime.Now
                            End SyncLock
                        Catch ex As Exception
                            System.Diagnostics.Debug.WriteLine($"加密配置项: {kvp.Key} 失败: {ex.Message}")
                        End Try
                    Else
                        _config(kvp.Key) = value
                    End If
                Next

                ' 只保存一次
                ScheduleSave()
            End Using
        End Sub

        ''' <summary>
        ''' 调度延迟保存
        ''' </summary>
        Private Sub ScheduleSave()
            SyncLock _saveLock
                _pendingSave = True

                If _saveTimer Is Nothing Then
                    _saveTimer = New Timer(Sub() FlushSave(), Nothing, _saveDelayMs, Timeout.Infinite)
                Else
                    _saveTimer.Change(_saveDelayMs, Timeout.Infinite)
                End If
            End SyncLock
        End Sub

        ''' <summary>
        ''' 立即保存（供外部调用）
        ''' </summary>
        Public Sub FlushSave()
            SyncLock _saveLock
                If Not _pendingSave Then
                    Return
                End If
                _pendingSave = False
            End SyncLock

            SaveConfig()
        End Sub

        ''' <summary>
        ''' 检查是否是敏感配置项
        ''' </summary>
        Private Function IsSensitiveKey(key As String) As Boolean
            If String.IsNullOrEmpty(key) Then Return False

            ' 检查完整匹配
            If _sensitiveKeys.Contains(key) Then Return True

            ' 检查部分匹配（例如包含 ApiKey、Secret 等关键字）
            Dim keyLower = key.ToLowerInvariant()
            Return keyLower.Contains("apikey") OrElse
                   keyLower.Contains("secret") OrElse
                   keyLower.Contains("password") OrElse
                   keyLower.Contains("token") OrElse
                   keyLower.Contains("privatekey")
        End Function

        ''' <summary>
        ''' 获取原始值（不解密）- 用于内部调试
        ''' </summary>
        Public Function GetRawSetting(key As String) As String
            If _config.ContainsKey(key) Then
                Return _config(key)
            End If
            Return String.Empty
        End Function

        ''' <summary>
        ''' 获取加密的API密钥（向后兼容）
        ''' </summary>
        <Obsolete("请直接使用 GetSetting，已自动处理加密")>
        Public Function GetEncryptedSetting(key As String) As String
            Return GetSetting(key)
        End Function

        ''' <summary>
        ''' 设置加密的API密钥（向后兼容）
        ''' </summary>
        <Obsolete("请直接使用 SetSetting，已自动处理加密")>
        Public Sub SetEncryptedSetting(key As String, value As String)
            SetSetting(key, value)
        End Sub

        ''' <summary>
        ''' 批量保存设置（兼容方法）
        ''' </summary>
        Public Sub SaveSettings(settings As Dictionary(Of String, Object))
            For Each kvp In settings
                SetSetting(kvp.Key, kvp.Value?.ToString())
            Next
        End Sub

        ''' <summary>
        ''' 加载配置（支持Base64编码值和标准值）
        ''' </summary>
        Private Sub LoadConfig()
            Try
                If Not File.Exists(_configFilePath) Then
                    Return
                End If

                Dim lines = File.ReadAllLines(_configFilePath)
                For Each line In lines
                    If String.IsNullOrWhiteSpace(line) OrElse line.StartsWith(";") Then
                        Continue For
                    End If

                    Dim separatorIndex = line.IndexOf("=")
                    If separatorIndex > 0 Then
                        Dim key = line.Substring(0, separatorIndex).Trim()
                        Dim value = line.Substring(separatorIndex + 1).Trim()

                        ' 检测Base64编码标记并解码
                        If value.StartsWith("|B64|") Then
                            Try
                                Dim base64Part = value.Substring(5) ' 移除|B64|前缀
                                value = Encoding.UTF8.GetString(Convert.FromBase64String(base64Part))
                            Catch ex As Exception
                                System.Diagnostics.Debug.WriteLine($"解码 Base64 配置值失败: [{key}]: {ex.Message}")
                                ' 解码失败则保留原始值
                            End Try
                        End If

                        _config(key) = value
                    End If
                Next
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("加载配置失败: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' 保存配置（INI 格式，对包含特殊字符的值进行转义）
        ''' </summary>
        Private Sub SaveConfig()
            Try
                PathHelper.EnsureDirectoryExists(_configFilePath)

                Using writer As New StreamWriter(_configFilePath, False)
                    writer.WriteLine("; Excel AI VSTO Add-in 配置文件")
                    writer.WriteLine("; 生成时间: " & DateTime.Now.ToString())
                    writer.WriteLine()

                    For Each kvp In _config
                        ' 对包含换行符或等号的值进行Base64编码保护
                        Dim value = kvp.Value
                        If Not String.IsNullOrEmpty(value) AndAlso
                           (value.Contains(vbCrLf) OrElse value.Contains(vbLf) OrElse value.Contains("="c)) Then
                            ' 使用特殊前缀标记 Base64 编码值
                            value = "|B64|" & Convert.ToBase64String(Encoding.UTF8.GetBytes(value))
                        End If
                        writer.WriteLine(String.Format("{0}={1}", kvp.Key, value))
                    Next
                End Using
            Catch ex As IOException
                System.Diagnostics.Debug.WriteLine($"保存配置失败(IO错误): {ex.Message}")
            Catch ex As UnauthorizedAccessException
                System.Diagnostics.Debug.WriteLine($"保存配置失败(权限错误): {ex.Message}")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"保存配置失败: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' 获取所有配置（敏感数据自动解密，安全返回给调用方）
        ''' </summary>
        Public Function GetAllSettings() As Dictionary(Of String, String)
            Dim result As New Dictionary(Of String, String)()
            For Each kvp In _config
                If IsSensitiveKey(kvp.Key) Then
                    ' 敏感数据通过 GetSetting 自动解密
                    result(kvp.Key) = GetSetting(kvp.Key)
                Else
                    result(kvp.Key) = kvp.Value
                End If
            Next
            Return result
        End Function

        ''' <summary>
        ''' 删除配置项（延迟保存，与其他写入操作保持一致）
        ''' </summary>
        Public Sub RemoveSetting(key As String)
            If _config.ContainsKey(key) Then
                _config.Remove(key)
                ' 同时清除解密缓存
                SyncLock _cacheLock
                    _decryptedCache.Remove(key)
                End SyncLock
                ScheduleSave()
            End If
        End Sub

        ''' <summary>
        ''' 清除所有配置（延迟保存）
        ''' </summary>
        Public Sub ClearAllSettings()
            _config.Clear()
            ClearDecryptedCache()
            ScheduleSave()
        End Sub

        ''' <summary>
        ''' 【修复】手动清除解密缓存，减少明文密钥在内存中的驻留时间
        ''' 建议在敏感操作完成后调用
        ''' </summary>
        Public Sub ClearDecryptedCache()
            SyncLock _cacheLock
                _decryptedCache.Clear()
                _cacheAccessTime.Clear()
            End SyncLock
        End Sub

        ''' <summary>
        ''' 重新加载配置文件
        ''' </summary>
        Public Sub Reload()
            ClearDecryptedCache()
            _config.Clear()
            LoadConfig()
        End Sub

        ''' <summary>
        ''' 获取配置的JSON字符串表示
        ''' </summary>
        Public Function GetConfigJson() As String
            Try
                Dim allSettings = GetAllSettings()
                Return Newtonsoft.Json.JsonConvert.SerializeObject(allSettings)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("获取配置 JSON 失败: " & ex.Message)
                Return "{}"
            End Try
        End Function

    End Class

End Namespace