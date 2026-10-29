Option Strict On
Imports System.IO
Imports System.Threading
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services

    ''' <summary>
    ''' 单条聊天消息记录
    ''' </summary>
    Public Class ChatMessage
        Public Property Role As String          ' user | assistant | system
        Public Property Content As String
        Public Property Timestamp As DateTime
        Public Property Metadata As Dictionary(Of String, Object)

        Public Sub New()
            Metadata = New Dictionary(Of String, Object)()
            Timestamp = DateTime.Now
        End Sub

        Public Sub New(role As String, content As String)
            Me.New()
            Me.Role = role
            Me.Content = content
        End Sub
    End Class

    ''' <summary>
    ''' 用户偏好设置（跨会话持久化）
    ''' </summary>
    Public Class UserPreferences
        Public Property DefaultThinkingMode As String = "AUTO"
        Public Property PreferredLanguage As String = "zh-CN"
        Public Property AutoSaveIntervalSeconds As Integer = 30
        Public Property MaxHistoryMessages As Integer = 100
        Public Property EnableContextAwareness As Boolean = True
        Public Property CustomPromptPrefix As String = ""
        Public Property FavoriteSkillNames As List(Of String)

        Public Sub New()
            FavoriteSkillNames = New List(Of String)()
        End Sub
    End Class

    ''' <summary>
    ''' 会话状态快照（用于持久化）
    ''' </summary>
    Public Class ConversationSnapshot
        Public Property ConversationId As String
        Public Property WorkbookName As String
        Public Property WorksheetName As String
        Public Property Messages As List(Of ChatMessage)
        Public Property CreatedAt As DateTime
        Public Property LastModifiedAt As DateTime
        Public Property TurnCount As Integer
        Public Property IsActive As Boolean
        Public Property Tags As List(Of String)

        Public Sub New()
            Messages = New List(Of ChatMessage)()
            Tags = New List(Of String)()
            CreatedAt = DateTime.Now
            LastModifiedAt = DateTime.Now
            IsActive = True
        End Sub
    End Class

    ''' <summary>
    ''' 会话记忆持久化服务
    ''' 职责：
    '''   1. 内存中维护当前会话的消息历史
    '''   2. 自动持久化到磁盘（本地临时目录）
    '''   3. 支持会话恢复、用户偏好跨轮次保留
    '''   4. 提供上下文摘要压缩，避免提示词过长
    ''' </summary>
    Public Class ConversationMemoryService
        Implements IDisposable

        Private _memoryStorePath As String
        Private _preferencesPath As String
        Private _currentConversationId As String
        Private _messages As List(Of ChatMessage)
        Private _preferences As UserPreferences
        Private _autoSaveTimer As Timer
        Private _isDisposed As Boolean
        Private _syncLock As New Object()

        Private Shared ReadOnly _lazyInstance As New Lazy(Of ConversationMemoryService)(
            Function() New ConversationMemoryService(),
            LazyThreadSafetyMode.ExecutionAndPublication)

        Public Shared ReadOnly Property Instance As ConversationMemoryService
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        ''' <summary>
        ''' 当前会话ID
        ''' </summary>
        Public ReadOnly Property CurrentConversationId As String
            Get
                Return _currentConversationId
            End Get
        End Property

        ''' <summary>
        ''' 用户偏好设置
        ''' </summary>
        Public ReadOnly Property Preferences As UserPreferences
            Get
                Return _preferences
            End Get
        End Property

        Private Sub New()
            Try
                Dim rootDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ExcelAIVSTOAddIn")
                If Not Directory.Exists(rootDir) Then Directory.CreateDirectory(rootDir)

                Dim memoryDir = Path.Combine(rootDir, "ConversationMemory")
                If Not Directory.Exists(memoryDir) Then Directory.CreateDirectory(memoryDir)

                _memoryStorePath = Path.Combine(memoryDir, "conversations.json")
                _preferencesPath = Path.Combine(rootDir, "user_preferences.json")

                _messages = New List(Of ChatMessage)()
                _currentConversationId = GenerateConversationId()

                LoadPreferences()
                LoadRecentConversation()

                _autoSaveTimer = New Timer(AddressOf AutoSaveCallback, Nothing,
                    TimeSpan.FromSeconds(_preferences.AutoSaveIntervalSeconds),
                    TimeSpan.FromSeconds(_preferences.AutoSaveIntervalSeconds))

                System.Diagnostics.Debug.WriteLine($"[ConversationMemoryService] 初始化完成, 会话ID: {_currentConversationId}")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[ConversationMemoryService] 初始化失败: {ex.Message}")
                _messages = If(_messages, New List(Of ChatMessage)())
                _preferences = If(_preferences, New UserPreferences())
                _currentConversationId = If(_currentConversationId, GenerateConversationId())
            End Try
        End Sub

        ''' <summary>
        ''' 生成新的会话ID
        ''' </summary>
        Private Shared Function GenerateConversationId() As String
            Return "conv_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & "_" & Guid.NewGuid().ToString("N").Substring(0, 6)
        End Function

        ''' <summary>
        ''' 加载用户偏好设置
        ''' </summary>
        Private Sub LoadPreferences()
            Try
                If File.Exists(_preferencesPath) Then
                    Dim json = File.ReadAllText(_preferencesPath)
                    If Not String.IsNullOrWhiteSpace(json) Then
                        _preferences = JsonConvert.DeserializeObject(Of UserPreferences)(json)
                    End If
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[ConversationMemoryService] 加载偏好失败: {ex.Message}")
            End Try
            _preferences = If(_preferences, New UserPreferences())
        End Sub

        ''' <summary>
        ''' 保存用户偏好设置
        ''' </summary>
        Public Sub SavePreferences()
            Try
                SyncLock _syncLock
                    Dim json = JsonConvert.SerializeObject(_preferences, Formatting.Indented)
                    File.WriteAllText(_preferencesPath, json)
                    System.Diagnostics.Debug.WriteLine("[ConversationMemoryService] 用户偏好已保存")
                End SyncLock
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[ConversationMemoryService] 保存偏好失败: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' 加载最近的会话（恢复上下文）
        ''' </summary>
        Private Sub LoadRecentConversation()
            Try
                If Not File.Exists(_memoryStorePath) Then Return
                Dim json = File.ReadAllText(_memoryStorePath)
                If String.IsNullOrWhiteSpace(json) Then Return

                Dim allConversations = JsonConvert.DeserializeObject(Of List(Of ConversationSnapshot))(json)
                If allConversations Is Nothing OrElse allConversations.Count = 0 Then Return

                Dim recent = allConversations.OrderByDescending(Function(c) c.LastModifiedAt).FirstOrDefault()
                If recent IsNot Nothing AndAlso recent.IsActive Then
                    _messages = If(recent.Messages, New List(Of ChatMessage)())
                    _currentConversationId = recent.ConversationId
                    System.Diagnostics.Debug.WriteLine($"[ConversationMemoryService] 恢复会话: {_currentConversationId}, 消息数: {_messages.Count}")
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[ConversationMemoryService] 加载会话失败: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' 自动保存回调
        ''' </summary>
        Private Sub AutoSaveCallback(state As Object)
            Try
                PersistCurrentConversation()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[ConversationMemoryService] 自动保存失败: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' 持久化当前会话到磁盘
        ''' </summary>
        Public Sub PersistCurrentConversation()
            Try
                SyncLock _syncLock
                    Dim snapshot = New ConversationSnapshot With {
                        .ConversationId = _currentConversationId,
                        .Messages = New List(Of ChatMessage)(_messages),
                        .LastModifiedAt = DateTime.Now,
                        .TurnCount = _messages.Count \ 2,
                        .IsActive = True
                    }

                    Dim allConversations As New List(Of ConversationSnapshot)()
                    If File.Exists(_memoryStorePath) Then
                        Try
                            Dim existingJson = File.ReadAllText(_memoryStorePath)
                            If Not String.IsNullOrWhiteSpace(existingJson) Then
                                allConversations = JsonConvert.DeserializeObject(Of List(Of ConversationSnapshot))(existingJson)
                            End If
                        Catch ex As Exception  ' 
                        End Try
                    End If
                    allConversations = If(allConversations, New List(Of ConversationSnapshot)())

                    Dim existingIndex = allConversations.FindIndex(Function(c) c.ConversationId = _currentConversationId)
                    If existingIndex >= 0 Then
                        allConversations(existingIndex) = snapshot
                    Else
                        allConversations.Add(snapshot)
                    End If

                    If allConversations.Count > 20 Then
                        allConversations = allConversations.
                            OrderByDescending(Function(c) c.LastModifiedAt).
                            Take(20).ToList()
                    End If

                    Dim json = JsonConvert.SerializeObject(allConversations, Formatting.Indented)
                    File.WriteAllText(_memoryStorePath, json)
                End SyncLock
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[ConversationMemoryService] 持久化失败: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' 添加消息到记忆
        ''' </summary>
        Public Sub AddMessage(role As String, content As String, Optional metadata As Dictionary(Of String, Object) = Nothing)
            If String.IsNullOrWhiteSpace(role) Then Return
            If String.IsNullOrWhiteSpace(content) Then Return

            SyncLock _syncLock
                Dim msg As New ChatMessage(role, content)
                If metadata IsNot Nothing Then
                    For Each kv In metadata
                        msg.Metadata(kv.Key) = kv.Value
                    Next
                End If
                _messages.Add(msg)

                While _messages.Count > _preferences.MaxHistoryMessages
                    _messages.RemoveAt(0)
                End While
            End SyncLock
        End Sub

        ''' <summary>
        ''' 获取上下文摘要（用于提示词注入，自动控制长度）
        ''' </summary>
        ''' <param name="maxTokens">最大预估token数</param>
        ''' <returns>格式化的上下文历史字符串</returns>
        Public Function GetContextSummary(Optional maxTokens As Integer = 2000) As String
            SyncLock _syncLock
                If _messages.Count = 0 Then Return ""

                Dim sb As New System.Text.StringBuilder()
                Dim currentTokens As Integer = 0
                Dim startIdx As Integer = Math.Max(0, _messages.Count - 30)

                For i As Integer = startIdx To _messages.Count - 1
                    Dim msg = _messages(i)
                    Dim line = $"[{msg.Timestamp:HH:mm:ss}] {msg.Role}: {msg.Content}"
                    Dim estimatedTokens = line.Length \ 4 + 1

                    If currentTokens + estimatedTokens > maxTokens Then
                        If sb.Length > 0 Then
                            sb.Insert(0, $"[...已截断 {_messages.Count - i} 条历史记录...]" & vbCrLf)
                        End If
                        Exit For
                    End If

                    sb.AppendLine(line)
                    currentTokens += estimatedTokens
                Next

                Return sb.ToString().Trim()
            End SyncLock
        End Function

        ''' <summary>
        ''' 获取最近的消息列表（用于前端传递）
        ''' </summary>
        Public Function GetRecentMessages(Optional count As Integer = 20) As List(Of ChatMessage)
            SyncLock _syncLock
                Dim result As New List(Of ChatMessage)()
                Dim startIdx = Math.Max(0, _messages.Count - count)
                For i As Integer = startIdx To _messages.Count - 1
                    result.Add(_messages(i))
                Next
                Return result
            End SyncLock
        End Function

        ''' <summary>
        ''' 开始新会话（保留偏好但清空历史消息）
        ''' </summary>
        Public Function StartNewConversation() As String
            SyncLock _syncLock
                PersistCurrentConversation()

                _currentConversationId = GenerateConversationId()
                _messages.Clear()

                System.Diagnostics.Debug.WriteLine($"[ConversationMemoryService] 新会话开始: {_currentConversationId}")
                Return _currentConversationId
            End SyncLock
        End Function

        ''' <summary>
        ''' 清除当前会话历史
        ''' </summary>
        Public Sub ClearCurrentConversation()
            SyncLock _syncLock
                _messages.Clear()
            End SyncLock
        End Sub

        ''' <summary>
        ''' 更新用户偏好（单个键值）
        ''' </summary>
        Public Sub UpdatePreference(key As String, value As Object)
            If String.IsNullOrWhiteSpace(key) Then Return

            Select Case key.ToLower()
                Case "defaultthinkingmode"
                    _preferences.DefaultThinkingMode = If(value?.ToString(), "AUTO")
                Case "preferredlanguage"
                    _preferences.PreferredLanguage = If(value?.ToString(), "zh-CN")
                Case "maxhistorymessages"
                    Dim n As Integer
                    If Integer.TryParse(value?.ToString(), n) AndAlso n > 0 Then
                        _preferences.MaxHistoryMessages = n
                    End If
                Case "enablecontextawareness"
                    Dim b As Boolean
                    If Boolean.TryParse(value?.ToString(), b) Then
                        _preferences.EnableContextAwareness = b
                    End If
                Case "custompromptprefix"
                    _preferences.CustomPromptPrefix = If(value?.ToString(), "")
            End Select

            SavePreferences()
        End Sub

        ''' <summary>
        ''' 添加收藏的Skill
        ''' </summary>
        Public Sub ToggleFavoriteSkill(skillName As String)
            If String.IsNullOrWhiteSpace(skillName) Then Return

            SyncLock _syncLock
                If _preferences.FavoriteSkillNames.Contains(skillName) Then
                    _preferences.FavoriteSkillNames.Remove(skillName)
                Else
                    _preferences.FavoriteSkillNames.Add(skillName)
                End If
            End SyncLock
            SavePreferences()
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If _isDisposed Then Return
            _isDisposed = True

            Try
                If _autoSaveTimer IsNot Nothing Then
                    _autoSaveTimer.Dispose()
                    _autoSaveTimer = Nothing
                End If
                PersistCurrentConversation()
                SavePreferences()
            Catch ex As Exception  ' 
            End Try
        End Sub

    End Class

End Namespace