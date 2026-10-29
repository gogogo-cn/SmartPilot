Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Net.Http
Imports System.Threading.Tasks
Imports Newtonsoft.Json

Namespace Services.RAG

    ''' <summary>
    ''' Qdrant 自动安装管理器
    ''' 负责下载、安装、启动 Qdrant 向量数据库
    ''' </summary>
    Public Class QdrantInstaller

        Private Const QDRANT_VERSION As String = "v1.13.2"
        Private Const QDRANT_DOWNLOAD_URL As String = "https://github.com/qdrant/qdrant/releases/download/{0}/qdrant-x86_64-pc-windows-msvc.zip"
        ' 国内镜像源（按优先级排列）
        Private QDRANT_MIRROR_URLS As String() = {
            "https://ghfast.top/https://github.com/qdrant/qdrant/releases/download/{0}/qdrant-x86_64-pc-windows-msvc.zip",
            "https://github.moeyy.xyz/https://github.com/qdrant/qdrant/releases/download/{0}/qdrant-x86_64-pc-windows-msvc.zip",
            "https://gh-proxy.com/https://github.com/qdrant/qdrant/releases/download/{0}/qdrant-x86_64-pc-windows-msvc.zip",
            "https://mirror.ghproxy.com/https://github.com/qdrant/qdrant/releases/download/{0}/qdrant-x86_64-pc-windows-msvc.zip"
        }
        Private Const QDRANT_EXE_NAME As String = "qdrant.exe"

        Private ReadOnly _installDir As String
        Private ReadOnly _qdrantExePath As String
        Private _qdrantProcess As Process

        Public Sub New(Optional installDir As String = Nothing)
            If String.IsNullOrEmpty(installDir) Then
                ' 统一安装到用户本地应用数据目录，插件版和 Web 版共享
                Dim localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
                _installDir = Path.Combine(localAppData, "SmartPilot", "Qdrant")
            Else
                _installDir = installDir
            End If

            _qdrantExePath = Path.Combine(_installDir, QDRANT_EXE_NAME)
        End Sub

        ''' <summary>
        ''' 检查 Qdrant 是否已安装
        ''' </summary>
        Public Function IsInstalled() As Boolean
            Return File.Exists(_qdrantExePath)
        End Function

        ''' <summary>
        ''' 检查 Qdrant 服务是否正在运行
        ''' </summary>
        Public Function IsRunning() As Boolean
            Try
                Dim processes = Process.GetProcessesByName("qdrant")
                Return processes.Length > 0
            Catch ex As Exception  ' 
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 获取安装状态信息
        ''' </summary>
        Public Function GetStatus() As Dictionary(Of String, Object)
            Dim status As New Dictionary(Of String, Object) From {
                {"installed", IsInstalled()},
                {"running", IsRunning()},
                {"installDir", _installDir},
                {"exePath", _qdrantExePath}
            }
            Return status
        End Function

        ''' <summary>
        ''' 下载并安装 Qdrant（自动尝试镜像源）
        ''' </summary>
        ''' <param name="progressCallback">进度回调 (message, percentage)</param>
        Public Async Function InstallAsync(Optional progressCallback As Action(Of String, Integer) = Nothing) As Task(Of Boolean)
            Try
                ' 创建安装目录
                If Not Directory.Exists(_installDir) Then
                    Directory.CreateDirectory(_installDir)
                    progressCallback?.Invoke("创建安装目录...", 10)
                End If

                Dim tempZipPath = Path.Combine(Path.GetTempPath(), $"qdrant-{QDRANT_VERSION}.zip")

                ' 先尝试主 URL，失败则依次尝试镜像
                Dim allUrls As New List(Of String) From {String.Format(QDRANT_DOWNLOAD_URL, QDRANT_VERSION)}
                For Each mirror In QDRANT_MIRROR_URLS
                    allUrls.Add(String.Format(mirror, QDRANT_VERSION))
                Next

                Dim downloadSuccess = False
                For Each downloadUrl In allUrls
                    Try
                        progressCallback?.Invoke($"正在下载 Qdrant {QDRANT_VERSION}...", 20)

                        Using httpClient = New HttpClient()
                            httpClient.Timeout = TimeSpan.FromMinutes(5) ' 5分钟超时

                            Using response = Await httpClient.GetAsync(downloadUrl, HttpCompletionOption.ResponseHeadersRead)
                                response.EnsureSuccessStatusCode()

                                Dim totalBytes = response.Content.Headers.ContentLength.GetValueOrDefault(-1)
                                Dim downloadedBytes = 0L

                                Using contentStream = Await response.Content.ReadAsStreamAsync()
                                    Using fileStream = New FileStream(tempZipPath, FileMode.Create, FileAccess.Write, FileShare.None, 8192, True)
                                        Dim buffer = New Byte(8191) {}
                                        Dim bytesRead As Integer

                                        Do
                                            bytesRead = Await contentStream.ReadAsync(buffer, 0, buffer.Length)
                                            If bytesRead = 0 Then Exit Do

                                            Await fileStream.WriteAsync(buffer, 0, bytesRead)
                                            downloadedBytes += bytesRead

                                            If totalBytes > 0 Then
                                                Dim percentage = CInt(20 + (downloadedBytes * 60 / totalBytes)) ' 20-80%
                                                progressCallback?.Invoke($"下载中 {CInt(downloadedBytes / 1024 / 1024)}MB / {CInt(totalBytes / 1024 / 1024)}MB", percentage)
                                            End If
                                        Loop
                                    End Using
                                End Using
                            End Using
                        End Using

                        downloadSuccess = True
                        Exit For
                    Catch ex As Exception
                        Debug.WriteLine($"[QdrantInstaller] 下载失败 ({downloadUrl}): {ex.Message}")
                        ' 继续尝试下一个 URL
                    End Try
                Next

                If Not downloadSuccess Then
                    progressCallback?.Invoke("所有下载源均失败，请手动下载安装", -1)
                    Return False
                End If

                progressCallback?.Invoke("下载完成，正在解压...", 80)

                ' 解压
                ZipFile.ExtractToDirectory(tempZipPath, _installDir)

                ' 删除临时文件
                If File.Exists(tempZipPath) Then
                    File.Delete(tempZipPath)
                End If

                progressCallback?.Invoke("安装完成", 100)

                ' 验证安装
                If IsInstalled() Then
                    Debug.WriteLine($"[QdrantInstaller] Qdrant 安装成功: {_qdrantExePath}")
                    Return True
                Else
                    Debug.WriteLine("[QdrantInstaller] 安装验证失败: qdrant.exe 未找到")
                    Return False
                End If

            Catch ex As Exception
                Debug.WriteLine($"[QdrantInstaller] 安装失败: {ex.Message}")
                progressCallback?.Invoke($"安装失败: {ex.Message}", -1)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 从本地文件安装 Qdrant
        ''' </summary>
        ''' <param name="zipFilePath">Qdrant ZIP 文件路径</param>
        ''' <param name="progressCallback">进度回调</param>
        Public Function InstallFromLocalFile(zipFilePath As String, Optional progressCallback As Action(Of String, Integer) = Nothing) As Boolean
            Try
                If Not File.Exists(zipFilePath) Then
                    progressCallback?.Invoke("文件不存在: " & zipFilePath, -1)
                    Return False
                End If

                ' 创建安装目录
                If Not Directory.Exists(_installDir) Then
                    Directory.CreateDirectory(_installDir)
                End If

                progressCallback?.Invoke("正在解压...", 50)

                ' 解压
                ZipFile.ExtractToDirectory(zipFilePath, _installDir)

                progressCallback?.Invoke("安装完成", 100)

                ' 验证安装
                If IsInstalled() Then
                    Debug.WriteLine($"[QdrantInstaller] Qdrant 本地安装成功: {_qdrantExePath}")
                    Return True
                Else
                    Debug.WriteLine("[QdrantInstaller] 本地安装验证失败: qdrant.exe 未找到")
                    Return False
                End If

            Catch ex As Exception
                Debug.WriteLine($"[QdrantInstaller] 本地安装失败: {ex.Message}")
                progressCallback?.Invoke($"安装失败: {ex.Message}", -1)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 启动 Qdrant 服务
        ''' </summary>
        Public Function StartService() As Boolean
            Try
                If Not IsInstalled() Then
                    Debug.WriteLine("[QdrantInstaller] Qdrant 未安装")
                    Return False
                End If

                If IsRunning() Then
                    Debug.WriteLine("[QdrantInstaller] Qdrant 已在运行")
                    Return True
                End If

                ' 必须显式指定 WorkingDirectory：Qdrant 的 storage 默认放在工作目录下。
                ' 不指定则继承 Excel 进程工作目录（可能是 Office 安装目录/System32），
                ' 导致 storage 位置随启动方式漂移——用户会看到「重启后向量数据像丢了」
                Dim startInfo = New ProcessStartInfo With {
                    .FileName = _qdrantExePath,
                    .Arguments = "--uri http://0.0.0.0:6333",
                    .UseShellExecute = False,
                    .CreateNoWindow = True,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True,
                    .WorkingDirectory = _installDir
                }

                _qdrantProcess = Process.Start(startInfo)
                Debug.WriteLine($"[QdrantInstaller] Qdrant 已启动, PID: {_qdrantProcess.Id}")

                ' 等待服务启动
                Threading.Thread.Sleep(2000)

                Return IsRunning()

            Catch ex As Exception
                Debug.WriteLine($"[QdrantInstaller] 启动失败: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 停止 Qdrant 服务
        ''' </summary>
        Public Function StopService() As Boolean
            Try
                Dim processes = Process.GetProcessesByName("qdrant")
                For Each proc In processes
                    proc.Kill()
                    proc.WaitForExit(5000)
                Next

                Debug.WriteLine("[QdrantInstaller] Qdrant 已停止")
                Return True

            Catch ex As Exception
                Debug.WriteLine($"[QdrantInstaller] 停止失败: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 卸载 Qdrant
        ''' </summary>
        Public Function Uninstall() As Boolean
            Try
                ' 先停止服务
                StopService()

                ' 删除安装目录
                If Directory.Exists(_installDir) Then
                    Directory.Delete(_installDir, True)
                End If

                Debug.WriteLine("[QdrantInstaller] Qdrant 已卸载")
                Return True

            Catch ex As Exception
                Debug.WriteLine($"[QdrantInstaller] 卸载失败: {ex.Message}")
                Return False
            End Try
        End Function

    End Class

End Namespace
