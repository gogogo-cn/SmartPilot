Imports System.IO

Namespace Utils

    ''' <summary>
    ''' 配置文件变更事件参数
    ''' </summary>
    Public Class ConfigFileChangedEventArgs
        Inherits EventArgs

        Public Property FilePath As String
        Public Property ConfigType As ConfigType
        Public Property ChangeTime As DateTime

        Public Sub New(filePath As String, configType As ConfigType)
            Me.FilePath = filePath
            Me.ConfigType = configType
            Me.ChangeTime = DateTime.Now
        End Sub
    End Class

    ''' <summary>
    ''' 配置类型枚举
    ''' </summary>
    Public Enum ConfigType
        DataType
        Skills
        AppSettings
        CustomSkills
        CustomSkillSnippets
    End Enum

    ''' <summary>
    ''' 配置文件监控器 - 支持配置热更新
    ''' </summary>
    Public Class ConfigFileWatcher
        Implements IDisposable

        Private Shared ReadOnly _lazyInstance As New Lazy(Of ConfigFileWatcher)(
            Function() New ConfigFileWatcher(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        Public Shared ReadOnly Property Instance As ConfigFileWatcher
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        Private _watchers As Dictionary(Of ConfigType, FileSystemWatcher)
        Private _lastReadTimes As Dictionary(Of String, DateTime)
        Private _debounceInterval As TimeSpan
        Private _isDisposed As Boolean

        Public Event ConfigFileChanged As EventHandler(Of ConfigFileChangedEventArgs)

        Private Sub New(Optional debounceSeconds As Double = 2.0)
            _watchers = New Dictionary(Of ConfigType, FileSystemWatcher)()
            _lastReadTimes = New Dictionary(Of String, DateTime)()
            _debounceInterval = TimeSpan.FromSeconds(debounceSeconds)
            _isDisposed = False
        End Sub

        ''' <summary>
        ''' 开始监控配置文件
        ''' </summary>
        Public Sub StartWatching(configType As ConfigType, filePath As String)
            If _isDisposed Then Throw New ObjectDisposedException("ConfigFileWatcher")
            If Not File.Exists(filePath) Then
                System.Diagnostics.Debug.WriteLine(String.Format("无法监控不存在的文件: {0}", filePath))
                Return
            End If

            ' 如果已经在监控该类型，先停止
            StopWatching(configType)

            Dim directory = Path.GetDirectoryName(filePath)
            Dim fileName = Path.GetFileName(filePath)

            Dim watcher = New FileSystemWatcher(directory, fileName) With {
                .NotifyFilter = NotifyFilters.LastWrite Or NotifyFilters.Size Or NotifyFilters.CreationTime,
                .EnableRaisingEvents = True
            }

            AddHandler watcher.Changed, AddressOf OnConfigFileChanged
            AddHandler watcher.Created, AddressOf OnConfigFileChanged
            AddHandler watcher.Renamed, AddressOf OnConfigFileChanged

            _watchers(configType) = watcher
            _lastReadTimes(filePath) = DateTime.MinValue

            System.Diagnostics.Debug.WriteLine(String.Format("开始监控配置文件: {0} [{1}]", filePath, configType.ToString()))
        End Sub
        
        ''' <summary>
        ''' 开始监控目录中的所有匹配文件
        ''' </summary>
        Public Sub StartWatchingDirectory(configType As ConfigType, directoryPath As String, searchPattern As String)
            If _isDisposed Then Throw New ObjectDisposedException("ConfigFileWatcher")
            If Not Directory.Exists(directoryPath) Then
                System.Diagnostics.Debug.WriteLine(String.Format("无法监控不存在的目录: {0}", directoryPath))
                Return
            End If

            ' 如果已经在监控该类型，先停止
            StopWatching(configType)

            Dim watcher = New FileSystemWatcher(directoryPath, searchPattern) With {
                .NotifyFilter = NotifyFilters.LastWrite Or NotifyFilters.Size Or NotifyFilters.CreationTime Or NotifyFilters.FileName,
                .EnableRaisingEvents = True,
                .IncludeSubdirectories = False
            }

            AddHandler watcher.Changed, AddressOf OnConfigFileChanged
            AddHandler watcher.Created, AddressOf OnConfigFileChanged
            AddHandler watcher.Deleted, AddressOf OnConfigFileChanged
            AddHandler watcher.Renamed, AddressOf OnConfigFileChanged

            _watchers(configType) = watcher
            _lastReadTimes(directoryPath) = DateTime.MinValue

            System.Diagnostics.Debug.WriteLine(String.Format("开始监控目录: {0}\{1} [{2}]", directoryPath, searchPattern, configType.ToString()))
        End Sub

        ''' <summary>
        ''' 停止监控指定类型的配置文件
        ''' </summary>
        Public Sub StopWatching(configType As ConfigType)
            If _watchers.ContainsKey(configType) Then
                Dim watcher = _watchers(configType)
                watcher.EnableRaisingEvents = False

                RemoveHandler watcher.Changed, AddressOf OnConfigFileChanged
                RemoveHandler watcher.Created, AddressOf OnConfigFileChanged
                RemoveHandler watcher.Renamed, AddressOf OnConfigFileChanged
                RemoveHandler watcher.Deleted, AddressOf OnConfigFileChanged

                watcher.Dispose()
                _watchers.Remove(configType)

                System.Diagnostics.Debug.WriteLine(String.Format("停止监控配置类型: {0}", configType.ToString()))
            End If
        End Sub

        ''' <summary>
        ''' 配置文件变更处理
        ''' </summary>
        Private Sub OnConfigFileChanged(sender As Object, e As FileSystemEventArgs)
            Try
                Dim filePath = e.FullPath

                ' 防抖处理：避免短时间内多次触发
                If _lastReadTimes.ContainsKey(filePath) Then
                    Dim lastRead = _lastReadTimes(filePath)
                    If DateTime.Now - lastRead < _debounceInterval Then
                        Return
                    End If
                End If

                _lastReadTimes(filePath) = DateTime.Now

                ' 找到对应的配置类型
                Dim configType = GetConfigTypeFromPath(filePath)

                System.Diagnostics.Debug.WriteLine(String.Format("配置文件变更 detected: {0} [{1}]", filePath, configType.ToString()))

                ' 触发事件
                RaiseEvent ConfigFileChanged(Me, New ConfigFileChangedEventArgs(filePath, configType))

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("处理配置文件变更事件失败: {0}", ex.Message))
            End Try
        End Sub

        ''' <summary>
        ''' 根据文件路径获取配置类型
        ''' </summary>
        Private Function GetConfigTypeFromPath(filePath As String) As ConfigType
            Dim fileName = Path.GetFileName(filePath).ToLower()
            Dim extension = Path.GetExtension(filePath).ToLower()

            ' 如果是 .md 文件，属于 Skills
            If extension = ".md" Then
                Return ConfigType.Skills
            End If

            Select Case fileName
                Case "datatypeconfigs.json"
                    Return ConfigType.DataType
                Case "defaultskills.json", "skills.json"
                    Return ConfigType.Skills
                Case "config.ini"
                    Return ConfigType.AppSettings
                Case Else
                    If fileName.Contains("custom") Then
                        Return ConfigType.CustomSkills
                    End If
                    Return ConfigType.AppSettings
            End Select
        End Function

        ''' <summary>
        ''' 停止所有监控
        ''' </summary>
        Public Sub StopAllWatching()
            For Each configType In _watchers.Keys.ToList()
                StopWatching(configType)
            Next
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            If _isDisposed Then Return

            StopAllWatching()
            _watchers.Clear()
            _lastReadTimes.Clear()

            _isDisposed = True
        End Sub
    End Class

End Namespace
