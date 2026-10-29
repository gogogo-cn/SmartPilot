Imports System.IO

Namespace Utils

    ''' <summary>
    ''' 系统 Python 定位。注册表、PATH 与常见安装路径只保留这一份实现。
    ''' </summary>
    Public Class PythonLocator

        ''' <summary>
        ''' 常见安装路径、注册表、PATH 依次查找，返回第一个可用的 python.exe。
        ''' </summary>
        Public Shared Function ResolvePreferredPath() As String
            Try
                Dim pythonPath = PathHelper.FindFirstExistingPath(PathHelper.GetPythonSearchPaths())
                If Not String.IsNullOrEmpty(pythonPath) Then
                    Return pythonPath
                End If

                Dim regPath = FindFromRegistry()
                If Not String.IsNullOrEmpty(regPath) Then
                    Return regPath
                End If

                Return FindFirstOnPath()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("获取Python路径失败: " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 从 HKLM / HKCU 的 PythonCore 注册表项查找版本最高的 python.exe。
        ''' </summary>
        Public Shared Function FindFromRegistry() As String
            Try
                Dim registryRoots As String() = {
                    "HKEY_LOCAL_MACHINE\SOFTWARE\Python\PythonCore",
                    "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Python\PythonCore",
                    "HKEY_CURRENT_USER\SOFTWARE\Python\PythonCore"
                }

                For Each root As String In registryRoots
                    Try
                        Dim keyBase = root.Replace("HKEY_LOCAL_MACHINE\", "").Replace("HKEY_CURRENT_USER\", "")
                        Dim hive = If(root.StartsWith("HKEY_LOCAL"),
                                      Microsoft.Win32.Registry.LocalMachine,
                                      Microsoft.Win32.Registry.CurrentUser)
                        Using key As Microsoft.Win32.RegistryKey = hive.OpenSubKey(keyBase)
                            If key Is Nothing Then Continue For
                            Dim subKeyNames As String() = key.GetSubKeyNames()
                            Dim bestVersion As String = ""
                            For Each versionName As String In subKeyNames
                                If versionName.StartsWith("3.") AndAlso
                                   String.Compare(versionName, bestVersion, StringComparison.OrdinalIgnoreCase) > 0 Then
                                    bestVersion = versionName
                                End If
                            Next

                            If String.IsNullOrEmpty(bestVersion) AndAlso subKeyNames.Length > 0 Then
                                bestVersion = subKeyNames(subKeyNames.Length - 1)
                            End If

                            If String.IsNullOrEmpty(bestVersion) Then Continue For
                            Using versionKey As Microsoft.Win32.RegistryKey = key.OpenSubKey(bestVersion & "\InstallPath")
                                If versionKey Is Nothing Then Continue For
                                Dim installPath As String = versionKey.GetValue("")?.ToString()
                                If String.IsNullOrEmpty(installPath) Then Continue For
                                Dim exePath = Path.Combine(installPath, "python.exe")
                                If File.Exists(exePath) Then
                                    System.Diagnostics.Debug.WriteLine($"[Python] 注册表找到: {exePath}")
                                    Return exePath
                                End If
                            End Using
                        End Using
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine($"[Python] 注册表子项查找失败: {ex.Message}")
                    End Try
                Next
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[Python] 注册表查找失败: {ex.Message}")
            End Try

            Return Nothing
        End Function

        ''' <summary>返回 where python.exe 的第一行，不校验文件是否存在。</summary>
        Public Shared Function FindFirstOnPath() As String
            Dim all = FindAllOnPath()
            If all.Count = 0 Then Return Nothing
            Return all(0)
        End Function

        ''' <summary>返回 where python.exe 列出的全部路径。</summary>
        Public Shared Function FindAllOnPath() As List(Of String)
            Dim found As New List(Of String)
            Try
                Using process As New System.Diagnostics.Process()
                    process.StartInfo.FileName = "where"
                    process.StartInfo.Arguments = "python.exe"
                    process.StartInfo.RedirectStandardOutput = True
                    process.StartInfo.UseShellExecute = False
                    process.StartInfo.CreateNoWindow = True
                    process.Start()
                    Dim output = process.StandardOutput.ReadToEnd()
                    process.WaitForExit(3000)

                    If process.ExitCode = 0 AndAlso Not String.IsNullOrWhiteSpace(output) Then
                        For Each line In output.Split({Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
                            Dim candidate = line.Trim()
                            If candidate.Length > 0 Then found.Add(candidate)
                        Next
                    End If
                End Using
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[Python] PATH查找失败: {ex.Message}")
            End Try
            Return found
        End Function

    End Class

End Namespace
