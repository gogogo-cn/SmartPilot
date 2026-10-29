Imports System
Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Net
Imports System.Net.Http
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Text.RegularExpressions

Namespace Services

    ''' <summary>
    ''' Python 按需下载引导服务
    '''
    ''' 替代原有的 PythonEmbed 捆绑方案：
    '''   1. 优先检测系统 Python（where/python.org 安装/Microsoft Store）
    '''   2. 未找到时，从 python.org 按需下载嵌入式 Python（~15MB 核心）
    '''   3. 下载到 %AppData%/ExcelAIVSTOAddIn/PythonEmbed/
    '''   4. 通过 IProgress 回调报告下载进度，供前端 UI 展示
    ''' </summary>
    Public Class PythonBootstrapper
        Implements IDisposable

        Private Shared ReadOnly _lazyInstance As New Lazy(Of PythonBootstrapper)(
            Function() New PythonBootstrapper(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        Public Shared ReadOnly Property Instance As PythonBootstrapper
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        Private Shared ReadOnly _httpClient As HttpClient
        Private _disposed As Boolean = False

        ' 缓存：已找到的 Python 路径
        Private _cachedPythonPath As String = Nothing
        Private _cacheLock As New Object()

        ' 下载状态
        Private _isDownloading As Boolean = False
        Private _downloadLock As New Object()

        ' 事件：供 UI 层订阅进度
        Public Event OnDownloadProgress As Action(Of String, Integer)  ' message, percent
        Public Event OnDownloadCompleted As Action(Of String)          ' pythonPath
        Public Event OnDownloadFailed As Action(Of String)             ' errorMessage

        Shared Sub New()
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            _httpClient = New HttpClient()
            _httpClient.Timeout = TimeSpan.FromMilliseconds(Utils.ConfigConstants.PythonBootstrapper.BOOTSTRAP_TIMEOUT_MS)
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "ExcelAIVSTOAddIn-PythonBootstrapper/1.0")
        End Sub

        ''' <summary>
        ''' 获取嵌入式 Python 的安装目录
        ''' </summary>
        Public Shared ReadOnly Property EmbeddedPythonDir As String
            Get
                Return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    Utils.ConfigConstants.Paths.APP_DATA_FOLDER,
                    Utils.ConfigConstants.PythonBootstrapper.PYTHON_DIR_NAME)
            End Get
        End Property

        ''' <summary>
        ''' 获取嵌入式 Python 的 exe 路径
        ''' </summary>
        Public Shared ReadOnly Property EmbeddedPythonExe As String
            Get
                Return Path.Combine(EmbeddedPythonDir, "python.exe")
            End Get
        End Property

        ''' <summary>
        ''' 检测系统是否已安装 Python
        ''' </summary>
        ''' <returns>可用的 Python 路径，或 Nothing</returns>
        Public Function DetectSystemPython() As String
            SyncLock _cacheLock
                If Not String.IsNullOrEmpty(_cachedPythonPath) Then
                    Return _cachedPythonPath
                End If
            End SyncLock

            Dim foundPath As String = Nothing

            ' 1. 尝试系统 PATH 中的 python3.exe
            foundPath = TryFindPython("python3.exe")
            If Not String.IsNullOrEmpty(foundPath) Then
                CacheAndReturn(foundPath)
                Return foundPath
            End If

            ' 2. 尝试系统 PATH 中的 python.exe
            foundPath = TryFindPython("python.exe")
            If Not String.IsNullOrEmpty(foundPath) Then
                CacheAndReturn(foundPath)
                Return foundPath
            End If

            ' 3. 尝试 py launcher (Python 3.7+ 自带)
            foundPath = TryFindPythonViaPyLauncher()
            If Not String.IsNullOrEmpty(foundPath) Then
                CacheAndReturn(foundPath)
                Return foundPath
            End If

            ' 4. 尝试 Microsoft Store 安装路径
            foundPath = TryFindPythonInMicrosoftStore()
            If Not String.IsNullOrEmpty(foundPath) Then
                CacheAndReturn(foundPath)
                Return foundPath
            End If

            ' 5. 尝试注册表
            foundPath = Utils.PythonLocator.FindFromRegistry()
            If Not String.IsNullOrEmpty(foundPath) Then
                CacheAndReturn(foundPath)
                Return foundPath
            End If

            ' 6. 尝试标准安装路径
            Dim localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
            Dim standardPaths As String() = {
                Path.Combine(localAppData, "Programs\Python\Python312\python.exe"),
                Path.Combine(localAppData, "Programs\Python\Python311\python.exe"),
                Path.Combine(localAppData, "Programs\Python\Python310\python.exe"),
                Path.Combine(localAppData, "Programs\Python\Python39\python.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Python312\python.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Python311\python.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Python310\python.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Python39\python.exe")
            }
            For Each path In standardPaths
                If File.Exists(path) Then
                    foundPath = path
                    CacheAndReturn(foundPath)
                    Return foundPath
                End If
            Next

            ' 7. 尝试已下载的嵌入式 Python
            If File.Exists(EmbeddedPythonExe) Then
                foundPath = EmbeddedPythonExe
                CacheAndReturn(foundPath)
                Return foundPath
            End If

            Return Nothing
        End Function

        ''' <summary>
        ''' 确保 Python 可用（检测 → 按需下载）
        ''' </summary>
        ''' <param name="progress">进度回调（可选）</param>
        ''' <returns>可用的 Python 路径</returns>
        Public Async Function EnsurePythonAsync(Optional progress As IProgress(Of PythonBootstrapProgress) = Nothing) As Task(Of String)
            ' 1. 先尝试系统 Python
            Dim systemPath = DetectSystemPython()
            If Not String.IsNullOrEmpty(systemPath) Then
                ReportProgress(progress, "已检测到系统 Python", 100, True, systemPath)
                Return systemPath
            End If

            ' 2. 系统未找到，尝试按需下载
            ReportProgress(progress, "未检测到 Python，正在准备下载...", 0, False, Nothing)

            Dim downloadResult = Await DownloadEmbeddedPythonAsync(progress).ConfigureAwait(False)
            If downloadResult Then
                CacheAndReturn(EmbeddedPythonExe)
                ReportProgress(progress, "Python 安装完成", 100, True, EmbeddedPythonExe)
                Return EmbeddedPythonExe
            End If

            ReportProgress(progress, "Python 下载失败，请手动安装 Python", 0, False, Nothing)
            Return Nothing
        End Function

        ''' <summary>
        ''' 下载嵌入式 Python（~15MB 核心包）
        ''' </summary>
        Private Async Function DownloadEmbeddedPythonAsync(Optional progress As IProgress(Of PythonBootstrapProgress) = Nothing) As Task(Of Boolean)
            SyncLock _downloadLock
                If _isDownloading Then Return False
                _isDownloading = True
            End SyncLock

            Try
                Dim version = Utils.ConfigConstants.PythonBootstrapper.RECOMMENDED_PYTHON_VERSION
                Dim packageName = String.Format(Utils.ConfigConstants.PythonBootstrapper.EMBEDDED_PACKAGE, version)
                Dim downloadUrl = $"{Utils.ConfigConstants.PythonBootstrapper.DOWNLOAD_URL_PREFIX}/{version}/{packageName}"

                System.Diagnostics.Debug.WriteLine($"[PythonBootstrapper] 开始下载: {downloadUrl}")

                ReportProgress(progress, $"正在从 python.org 下载 Python {version}...", 5, False, Nothing)

                ' 创建临时下载目录
                Dim tempDir = Path.Combine(Path.GetTempPath(), "SmartPilot_PythonBootstrap_" & Guid.NewGuid().ToString("N"))
                Directory.CreateDirectory(tempDir)

                Dim zipPath = Path.Combine(tempDir, packageName)
                Dim extractDir = EmbeddedPythonDir

                ' 带重试的下载
                Dim retryCount = Utils.ConfigConstants.PythonBootstrapper.DOWNLOAD_RETRY_COUNT
                Dim success = False
                Dim lastException As Exception = Nothing
                Dim retryIndex = 0
                Do While retryIndex < retryCount AndAlso Not success
                    Try
                        If retryIndex > 0 Then
                            Await Task.Delay(Utils.ConfigConstants.PythonBootstrapper.DOWNLOAD_RETRY_DELAY_MS).ConfigureAwait(False)
                        End If

                        ReportProgress(progress, $"正在下载 Python ({retryIndex + 1}/{retryCount})...", 10 + (retryIndex * 20), False, Nothing)

                        Using response = Await _httpClient.GetAsync(downloadUrl, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(False)
                            response.EnsureSuccessStatusCode()

                            Dim totalBytes = If(response.Content.Headers.ContentLength.HasValue, response.Content.Headers.ContentLength.Value, -1L)
                            Using stream = Await response.Content.ReadAsStreamAsync().ConfigureAwait(False)
                                Using fileStream = New FileStream(zipPath, FileMode.Create, FileAccess.Write, FileShare.None, 8192, useAsync:=True)
                                    Dim buffer(8191) As Byte
                                    Dim bytesRead As Integer = 0
                                    Dim totalRead As Long = 0
                                    Dim lastReportPercent As Integer = 10

                                    Do
                                        bytesRead = Await stream.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(False)
                                        If bytesRead = 0 Then Exit Do

                                        Await fileStream.WriteAsync(buffer, 0, bytesRead).ConfigureAwait(False)
                                        totalRead += bytesRead

                                        If totalBytes > 0 Then
                                            Dim currentPercent = CInt(10 + (totalRead * 60 / totalBytes))
                                            If currentPercent > lastReportPercent Then
                                                lastReportPercent = currentPercent
                                                ReportProgress(progress, $"下载中... {totalRead / 1024 / 1024}MB / {totalBytes / 1024 / 1024}MB", currentPercent, False, Nothing)
                                            End If
                                        End If
                                    Loop
                                End Using
                            End Using
                        End Using

                        success = True
                    Catch ex As Exception
                        lastException = ex
                        System.Diagnostics.Debug.WriteLine($"[PythonBootstrapper] 下载尝试 {retryIndex + 1} 失败: {ex.Message}")
                    End Try
                    retryIndex += 1
                Loop

                If Not success Then Return False

                ' 解压
                ReportProgress(progress, "正在解压...", 75, False, Nothing)

                If Directory.Exists(extractDir) Then
                    Directory.Delete(extractDir, recursive:=True)
                End If
                Directory.CreateDirectory(extractDir)

                ZipFile.ExtractToDirectory(zipPath, extractDir)

                ' 配置嵌入 Python（修改 python39._pth 以启用 site-packages 和 pip）
                ReportProgress(progress, "正在配置 Python 环境...", 85, False, Nothing)
                ConfigureEmbeddedPython(extractDir, version)

                ' 清理临时文件
                ReportProgress(progress, "清理临时文件...", 95, False, Nothing)
                Try
                    Directory.Delete(tempDir, recursive:=True)
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"[PythonBootstrapper] 清理临时文件失败: {ex.Message}")
                End Try

                Return True
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[PythonBootstrapper] 下载失败: {ex.Message}")
                RaiseEvent OnDownloadFailed(ex.Message)
                Return False
            Finally
                SyncLock _downloadLock
                    _isDownloading = False
                End SyncLock
            End Try
        End Function

        ''' <summary>
        ''' 配置嵌入式 Python 以支持 pip 和 site-packages
        ''' 嵌入式 Python 默认不加载 site-packages，需要修改 ._pth 文件
        ''' </summary>
        Private Sub ConfigureEmbeddedPython(extractDir As String, version As String)
            Try
                Dim sitePackagesDir = Path.Combine(extractDir, "Lib", "site-packages")

                ' 查找 ._pth 文件（例如 python39._pth）
                Dim pthFiles = Directory.GetFiles(extractDir, "python*._pth")
                If pthFiles.Length = 0 Then
                    ' 创建 site-packages 目录
                    Directory.CreateDirectory(sitePackagesDir)
                    Return
                End If

                Dim pthPath = pthFiles(0)
                Dim pthContent = File.ReadAllText(pthPath)

                ' 取消注释 import site（启用 site-packages）
                pthContent = pthContent.Replace("#import site", "import site")

                ' 确保 Lib\site-packages 在路径中
                Dim libSiteLine = "Lib\site-packages"
                If Not pthContent.Contains(libSiteLine) Then
                    pthContent = pthContent.TrimEnd() & vbCrLf & libSiteLine
                End If

                File.WriteAllText(pthPath, pthContent)

                ' 创建 site-packages 目录
                If Not Directory.Exists(sitePackagesDir) Then
                    Directory.CreateDirectory(sitePackagesDir)
                End If

                System.Diagnostics.Debug.WriteLine("[PythonBootstrapper] 嵌入式 Python 配置完成")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[PythonBootstrapper] 配置嵌入式 Python 失败: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' 检查 Python 是否已就绪（系统已安装或已下载）
        ''' </summary>
        Public Function IsPythonReady() As Boolean
            Dim path = DetectSystemPython()
            Return Not String.IsNullOrEmpty(path)
        End Function

        ''' <summary>
        ''' 清除缓存，强制重新检测
        ''' </summary>
        Public Sub ClearCache()
            SyncLock _cacheLock
                _cachedPythonPath = Nothing
            End SyncLock
        End Sub

#Region "检测辅助方法"

        Private Function TryFindPython(exeName As String) As String
            Try
                Using proc As New Process()
                    proc.StartInfo.FileName = "where"
                    proc.StartInfo.Arguments = exeName
                    proc.StartInfo.UseShellExecute = False
                    proc.StartInfo.RedirectStandardOutput = True
                    proc.StartInfo.RedirectStandardError = True
                    proc.StartInfo.CreateNoWindow = True
                    proc.Start()
                    Dim output = proc.StandardOutput.ReadToEnd().Trim()
                    proc.WaitForExit(3000)

                    If proc.ExitCode = 0 AndAlso Not String.IsNullOrEmpty(output) Then
                        Dim lines = output.Split({Environment.NewLine, vbLf, vbCr}, StringSplitOptions.RemoveEmptyEntries)
                        For Each line In lines
                            Dim trimmed = line.Trim()
                            If File.Exists(trimmed) AndAlso IsValidPythonPath(trimmed) Then
                                Return trimmed
                            End If
                        Next
                    End If
                End Using
            Catch ex As Exception  ' 
            End Try
            Return Nothing
        End Function

        Private Function TryFindPythonViaPyLauncher() As String
            Try
                Using proc As New Process()
                    proc.StartInfo.FileName = "py"
                    proc.StartInfo.Arguments = "-3 -c ""import sys; print(sys.executable)"""
                    proc.StartInfo.UseShellExecute = False
                    proc.StartInfo.RedirectStandardOutput = True
                    proc.StartInfo.RedirectStandardError = True
                    proc.StartInfo.CreateNoWindow = True
                    proc.Start()
                    Dim output = proc.StandardOutput.ReadToEnd().Trim()
                    proc.WaitForExit(3000)

                    If proc.ExitCode = 0 AndAlso Not String.IsNullOrEmpty(output) AndAlso File.Exists(output) Then
                        Return output
                    End If
                End Using
            Catch ex As Exception  ' 
            End Try
            Return Nothing
        End Function

        Private Function TryFindPythonInMicrosoftStore() As String
            Try
                Dim localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
                Dim storeBase = Path.Combine(localAppData, "Microsoft\WindowsApps")

                Dim candidates As String() = {
                    Path.Combine(storeBase, "python3.exe"),
                    Path.Combine(storeBase, "python.exe")
                }
                For Each path In candidates
                    If File.Exists(path) AndAlso IsValidPythonPath(path) Then
                        Return path
                    End If
                Next
            Catch ex As Exception  ' 
            End Try
            Return Nothing
        End Function

        Private Function IsValidPythonPath(path As String) As Boolean
            If String.IsNullOrEmpty(path) OrElse Not File.Exists(path) Then Return False
            Try
                Using proc As New Process()
                    proc.StartInfo.FileName = path
                    proc.StartInfo.Arguments = "--version"
                    proc.StartInfo.UseShellExecute = False
                    proc.StartInfo.RedirectStandardOutput = True
                    proc.StartInfo.RedirectStandardError = True
                    proc.StartInfo.CreateNoWindow = True
                    proc.Start()
                    proc.WaitForExit(5000)
                    Return proc.ExitCode = 0
                End Using
            Catch ex As Exception  ' 
                Return False
            End Try
        End Function

        Private Sub CacheAndReturn(path As String)
            SyncLock _cacheLock
                _cachedPythonPath = path
            End SyncLock
        End Sub

        Private Sub ReportProgress(progress As IProgress(Of PythonBootstrapProgress), message As String, percent As Integer, isComplete As Boolean, pythonPath As String)
            If progress IsNot Nothing Then
                progress.Report(New PythonBootstrapProgress With {
                    .Message = message,
                    .PercentComplete = percent,
                    .IsComplete = isComplete,
                    .PythonPath = pythonPath
                })
            End If

            RaiseEvent OnDownloadProgress(message, percent)

            If isComplete AndAlso Not String.IsNullOrEmpty(pythonPath) Then
                RaiseEvent OnDownloadCompleted(pythonPath)
            End If

            System.Diagnostics.Debug.WriteLine($"[PythonBootstrapper] {percent}%: {message}")
        End Sub

#End Region

#Region "IDisposable"

        Public Sub Dispose() Implements IDisposable.Dispose
            If Not _disposed Then
                _httpClient?.Dispose()
                _disposed = True
            End If
            GC.SuppressFinalize(Me)
        End Sub

#End Region

    End Class

    ''' <summary>
    ''' Python 下载进度报告
    ''' </summary>
    Public Class PythonBootstrapProgress
        Public Property Message As String = ""
        Public Property PercentComplete As Integer = 0
        Public Property IsComplete As Boolean = False
        Public Property PythonPath As String = Nothing
    End Class

End Namespace