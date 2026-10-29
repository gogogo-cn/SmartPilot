Option Strict On
Option Explicit On
Option Infer On
Imports System.Collections.Concurrent
Imports System.Security.Cryptography
Imports System.Text
Imports System.Linq

Namespace Utils

    ''' <summary>
    ''' 请求缓存 - 用于缓存OCR和LLM请求结果
    ''' </summary>
    Public Class RequestCache

        Private _cache As ConcurrentDictionary(Of String, CacheEntry)
        Private _maxSize As Integer
        Private _defaultTtlMinutes As Integer
        Private _isEnabled As Boolean
        Private _lock As New Object()

        ''' <summary>
        ''' 默认构造函数 - 无参构造用于 DI 容器注入（使用默认配置：size=100, TTL=30min, enabled=true）
        ''' </summary>
        Public Sub New()
            Me.New(maxSize:=100, defaultTtlMinutes:=30, isEnabled:=True)
        End Sub

        Public Sub New(Optional maxSize As Integer = 100, Optional defaultTtlMinutes As Integer = 30, Optional isEnabled As Boolean = True)
            _cache = New ConcurrentDictionary(Of String, CacheEntry)()
            _maxSize = maxSize
            _defaultTtlMinutes = defaultTtlMinutes
            _isEnabled = isEnabled
        End Sub

        ''' <summary>
        ''' 获取或生成缓存键
        ''' </summary>
        Public Shared Function GenerateKey(prefix As String, ParamArray args() As Object) As String
            Dim keyBuilder = New StringBuilder()
            keyBuilder.Append(prefix)
            keyBuilder.Append(":")

            For Each arg In args
                If arg IsNot Nothing Then
                    keyBuilder.Append(arg.ToString())
                End If
                keyBuilder.Append("|")
            Next

            ' 使用SHA256哈希缩短键长度
            Using sha256 As SHA256 = SHA256.Create()
                Dim hash = sha256.ComputeHash(Encoding.UTF8.GetBytes(keyBuilder.ToString()))
                Dim hashString = BitConverter.ToString(hash).Replace("-", "").ToLower()
                Return $"{prefix}:{hashString.Substring(0, 16)}"
            End Using
        End Function

        ''' <summary>
        ''' 获取缓存值
        ''' </summary>
        Public Function TryGetValue(key As String) As Tuple(Of Boolean, Object)
            If Not _isEnabled Then
                Return New Tuple(Of Boolean, Object)(False, Nothing)
            End If

            Dim entry As CacheEntry = Nothing
            If _cache.TryGetValue(key, entry) Then
                ' 检查是否过期
                If entry.ExpiresAt > DateTime.Now Then
                    Return New Tuple(Of Boolean, Object)(True, entry.Value)
                Else
                    ' 过期,移除
                    Dim removedEntry As CacheEntry = Nothing
                    _cache.TryRemove(key, removedEntry)
                End If
            End If

            Return New Tuple(Of Boolean, Object)(False, Nothing)
        End Function

        ''' <summary>
        ''' 设置缓存值
        ''' </summary>
        Public Sub SetValue(key As String, value As Object, Optional ttlMinutes As Integer = -1)
            If Not _isEnabled Then
                Return
            End If

            ' 清理过期条目
            CleanupExpired()

            ' 检查缓存大小
            SyncLock _lock
                If _cache.Count >= _maxSize Then
                    RemoveOldest()
                End If
            End SyncLock

            Dim ttl = If(ttlMinutes > 0, ttlMinutes, _defaultTtlMinutes)
            Dim entry = New CacheEntry With {
                .Value = value,
                .CreatedAt = DateTime.Now,
                .ExpiresAt = DateTime.Now.AddMinutes(ttl)
            }

            _cache.AddOrUpdate(key, entry, Function(k, old) entry)
        End Sub

        ''' <summary>
        ''' 清除所有缓存
        ''' </summary>
        Public Sub Clear()
            _cache.Clear()
        End Sub

        ''' <summary>
        ''' 清除过期缓存
        ''' </summary>
        Public Sub CleanupExpired()
            Dim now = DateTime.Now
            Dim expiredKeys = _cache.Where(Function(kvp) kvp.Value.ExpiresAt <= now).Select(Function(kvp) kvp.Key).ToList()

            For Each key In expiredKeys
                Dim removedEntry As CacheEntry = Nothing
                _cache.TryRemove(key, removedEntry)
            Next
        End Sub

        ''' <summary>
        ''' 移除最旧的条目
        ''' </summary>
        Private Sub RemoveOldest()
            Dim oldestKey = _cache.OrderBy(Function(kvp) kvp.Value.CreatedAt).FirstOrDefault().Key
            If oldestKey IsNot Nothing Then
                Dim removedEntry As CacheEntry = Nothing
                _cache.TryRemove(oldestKey, removedEntry)
            End If
        End Sub

        ''' <summary>
        ''' 获取缓存统计信息
        ''' </summary>
        Public Function GetStats() As Dictionary(Of String, Object)
            Dim now = DateTime.Now
            Dim validCount = _cache.Where(Function(kvp) kvp.Value.ExpiresAt > now).Count()
            Dim expiredCount = _cache.Count - validCount

            Return New Dictionary(Of String, Object) From {
                {"total", _cache.Count},
                {"valid", validCount},
                {"expired", expiredCount},
                {"maxSize", _maxSize},
                {"isEnabled", _isEnabled}
            }
        End Function

        ''' <summary>
        ''' 启用/禁用缓存
        ''' </summary>
        Public Sub SetEnabled(enabled As Boolean)
            _isEnabled = enabled
            If Not enabled Then
                Clear()
            End If
        End Sub

        ''' <summary>
        ''' 更新缓存配置
        ''' </summary>
        Public Sub UpdateConfig(Optional maxSize As Integer = -1, Optional defaultTtlMinutes As Integer = -1, Optional isEnabled As Nullable(Of Boolean) = Nothing)
            If maxSize > 0 Then
                _maxSize = maxSize
            End If
            If defaultTtlMinutes > 0 Then
                _defaultTtlMinutes = defaultTtlMinutes
            End If
            If isEnabled.HasValue Then
                SetEnabled(isEnabled.Value)
            End If
        End Sub

        Private Class CacheEntry
            Public Property Value As Object
            Public Property CreatedAt As DateTime
            Public Property ExpiresAt As DateTime
        End Class

    End Class

End Namespace