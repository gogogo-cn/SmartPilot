Imports System.IO

Namespace Utils

    ''' <summary>
    ''' 路径帮助类 - 集中管理应用程序所有路径操作
    ''' 消除硬编码路径，提供统一的路径获取方法
    ''' </summary>
    Public Class PathHelper

        Private Shared ReadOnly _appDataPath As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        Private Shared ReadOnly _programFilesPath As String = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
        Private Shared ReadOnly _programFilesX86Path As String = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)

        ''' <summary>
        ''' 获取应用程序数据根目录
        ''' </summary>
        Public Shared ReadOnly Property AppDataRoot As String
            Get
                Return Path.Combine(_appDataPath, ConfigConstants.Paths.APP_DATA_FOLDER)
            End Get
        End Property

        ''' <summary>
        ''' 获取配置文件完整路径
        ''' </summary>
        Public Shared ReadOnly Property ConfigFilePath As String
            Get
                Return Path.Combine(AppDataRoot, ConfigConstants.Paths.CONFIG_FILE)
            End Get
        End Property

        ''' <summary>
        ''' 获取 Config 目录路径（统一配置目录）
        ''' 优先级：运行时目录 Config/ → 开发源码目录 Config/
        ''' </summary>
        Public Shared ReadOnly Property ConfigDirectory As String
            Get
                Dim assemblyDir = PluginInstallPath

                ' 运行时目录 Config/（bin/Debug/Config/ 或发布目录 Config/）
                Dim runtimeConfig = Path.Combine(assemblyDir, "Config")
                If Directory.Exists(runtimeConfig) Then
                    Return runtimeConfig
                End If

                ' 开发模式：回退到源码目录 Config/
                Dim devConfig = Path.Combine(assemblyDir, "..", "..", "Config")
                devConfig = Path.GetFullPath(devConfig)
                If Directory.Exists(devConfig) Then
                    Return devConfig
                End If

                ' 最终回退：返回运行时目录（即使不存在，让调用方处理）
                Return runtimeConfig
            End Get
        End Property

        ''' <summary>
        ''' 获取 Skills 配置文件路径
        ''' </summary>
        Public Shared ReadOnly Property SkillsFilePath As String
            Get
                Return Path.Combine(AppDataRoot, ConfigConstants.Paths.SKILLS_FILE)
            End Get
        End Property

        ''' <summary>
        ''' 获取工作流状态文件路径
        ''' </summary>
        Public Shared ReadOnly Property WorkflowStateFilePath As String
            Get
                Return Path.Combine(AppDataRoot, ConfigConstants.Paths.WORKFLOW_STATE_FILE)
            End Get
        End Property

        ''' <summary>
        ''' 获取自定义 Skill 目录路径
        ''' </summary>
        Public Shared ReadOnly Property CustomSkillsPath As String
            Get
                Return Path.Combine(AppDataRoot, ConfigConstants.Paths.CUSTOM_SKILLS_FOLDER)
            End Get
        End Property
        
        ''' <summary>
        ''' 获取插件安装目录（用于加载内置的 MD Skill 文件）
        ''' </summary>
        Public Shared ReadOnly Property PluginInstallPath As String
            Get
                ' VSTO ClickOnce部署时，CodeBase指向临时目录，需要获取实际的安装目录
                Dim assemblyCodeBase = System.Reflection.Assembly.GetExecutingAssembly().CodeBase
                Dim assemblyPath = New Uri(assemblyCodeBase).LocalPath
                Return Path.GetDirectoryName(assemblyPath)
            End Get
        End Property

        ''' <summary>
        ''' 获取调试日志文件路径
        ''' </summary>
        Public Shared ReadOnly Property DebugLogFilePath As String
            Get
                Return Path.Combine(AppDataRoot, ConfigConstants.Paths.DEBUG_LOG_FILE)
            End Get
        End Property

        ''' <summary>
        ''' 获取内置 Markdown Skill 文件目录路径（仅用于系统Skill）
        ''' </summary>
        Public Shared Function GetMarkdownSkillsDirectory() As String
            Dim codeBase = New Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath
            Dim appDir = Path.GetDirectoryName(codeBase)
            
            ' 优先使用编译输出目录
            Dim projectPath = Path.Combine(appDir, "Skills")
            If Directory.Exists(projectPath) Then
                Return projectPath
            End If
            
            ' 开发时使用源代码目录
            Dim devPath = Path.Combine(appDir, "..", "..", "Skills")
            If Directory.Exists(devPath) Then
                Return devPath
            End If
            
            Return projectPath
        End Function

        ''' <summary>
        ''' 获取 Python 可能的安装路径列表
        ''' 优先级：系统 Python → 配置 Python → 按需下载的嵌入式 Python → PATH 兜底
        ''' </summary>
        Public Shared Function GetPythonSearchPaths() As String()
            Dim paths As New List(Of String)

            ' 【优先级1】系统 Python（通过 Bootstrapper 检测，已缓存）
            Dim systemPython = Services.PythonBootstrapper.Instance.DetectSystemPython()
            If Not String.IsNullOrEmpty(systemPython) AndAlso File.Exists(systemPython) Then
                paths.Add(systemPython)
            End If

            ' 【优先级2】从配置中读取的 Python 路径
            Dim customPythonPath = ConfigurationManager.Instance.GetSetting("PythonPath", String.Empty)
            If Not String.IsNullOrEmpty(customPythonPath) AndAlso File.Exists(customPythonPath) AndAlso Not paths.Contains(customPythonPath) Then
                paths.Add(customPythonPath)
            End If

            ' 【优先级3】按需下载的嵌入式 Python（%AppData%/ExcelAIVSTOAddIn/PythonEmbed/）
            Dim bootstrapperExe = Services.PythonBootstrapper.EmbeddedPythonExe
            If Not String.IsNullOrEmpty(bootstrapperExe) AndAlso File.Exists(bootstrapperExe) AndAlso Not paths.Contains(bootstrapperExe) Then
                paths.Add(bootstrapperExe)
            End If

            ' 【优先级4】标准安装路径
            For Each version In ConfigConstants.Paths.PYTHON_VERSIONS
                Dim stdPath = Path.Combine(_appDataPath, ConfigConstants.Paths.PYTHON_BASE_FOLDER, version, ConfigConstants.Paths.PYTHON_EXE)
                If File.Exists(stdPath) AndAlso Not paths.Contains(stdPath) Then
                    paths.Add(stdPath)
                End If
            Next

            ' 【优先级5】系统环境变量中的 Python（兜底）
            paths.Add("python.exe")
            paths.Add("python3.exe")

            Return paths.ToArray()
        End Function

        ''' <summary>
        ''' 获取嵌入式 Python 路径（按需下载位置，位于 %AppData%/ExcelAIVSTOAddIn/PythonEmbed/）
        ''' </summary>
        Public Shared Function GetEmbeddedPythonPath() As String
            Dim bootstrapperExe = Services.PythonBootstrapper.EmbeddedPythonExe
            If Not String.IsNullOrEmpty(bootstrapperExe) AndAlso File.Exists(bootstrapperExe) Then
                Return bootstrapperExe
            End If

            ' 兼容旧路径：插件安装目录下的 PythonEmbed
            Dim legacyCandidates As String() = {
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PythonEmbed", "python.exe"),
                Path.Combine(Path.GetDirectoryName(GetType(PathHelper).Assembly.Location), "PythonEmbed", "python.exe")
            }
            For Each path In legacyCandidates
                If File.Exists(path) Then
                    Return path
                End If
            Next
            Return Nothing
        End Function

        ''' <summary>
        ''' 查找第一个存在的路径
        ''' </summary>
        ''' <param name="paths">候选路径数组</param>
        ''' <returns>第一个存在的路径，或空字符串</returns>
        Public Shared Function FindFirstExistingPath(paths As String()) As String
            If paths Is Nothing Then Return String.Empty

            For Each path In paths
                If Not String.IsNullOrEmpty(path) AndAlso File.Exists(path) Then
                    Return path
                End If
            Next

            Return String.Empty
        End Function

        ''' <summary>
        ''' 确保目录存在，不存在则创建
        ''' </summary>
        ''' <param name="filePath">文件或目录路径</param>
        ''' <returns>是否成功</returns>
        Public Shared Function EnsureDirectoryExists(filePath As String) As Boolean
            Try
                Dim directory = System.IO.Path.GetDirectoryName(filePath)
                If Not String.IsNullOrEmpty(directory) AndAlso Not System.IO.Directory.Exists(directory) Then
                    System.IO.Directory.CreateDirectory(directory)
                End If
                Return True
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"创建目录失败: {filePath}, 错误: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 获取安全的文件名（移除非法字符）
        ''' </summary>
        ''' <param name="fileName">原始文件名</param>
        ''' <returns>安全的文件名</returns>
        Public Shared Function GetSafeFileName(fileName As String) As String
            If String.IsNullOrEmpty(fileName) Then
                Return String.Empty
            End If

            Dim invalidChars = Path.GetInvalidFileNameChars()
            For Each c In invalidChars
                fileName = fileName.Replace(c, "_"c)
            Next

            Return fileName
        End Function

        ''' <summary>
        ''' 获取安全的完整路径
        ''' </summary>
        ''' <param name="basePath">基础路径</param>
        ''' <param name="relativePath">相对路径</param>
        ''' <returns>安全的完整路径</returns>
        Public Shared Function GetSafePath(basePath As String, relativePath As String) As String
            If String.IsNullOrEmpty(basePath) Then
                Throw New ArgumentNullException(NameOf(basePath))
            End If

            If String.IsNullOrEmpty(relativePath) Then
                Return basePath
            End If

            ' 规范化路径，防止目录遍历攻击
            Dim fullPath = Path.GetFullPath(Path.Combine(basePath, relativePath))
            Dim baseFullPath = Path.GetFullPath(basePath)

            ' 确保结果路径在基础路径之下
            If Not fullPath.StartsWith(baseFullPath, StringComparison.OrdinalIgnoreCase) Then
                Throw New System.Security.SecurityException("路径遍历攻击检测：尝试访问基础路径之外的目录")
            End If

            Return fullPath
        End Function

        ''' <summary>
        ''' 输出路径里"编造的用户名/占位符"判定表。
        ''' 例：C:\Users\Yourusername\Desktop\output.csv、C:\path\to\<username>\x.csv、用户名、xxx、...
        ''' </summary>
        Private Shared ReadOnly PlaceholderPathRegex As New System.Text.RegularExpressions.Regex(
            "(?i)(" &
            "your[\s_\-]?user\w*|yourname|your[\s_\-]?name|some[\s_\-]?user|test[\s_\-]?user|" &
            "user[\s_\-]?name|<[^<>/]{1,24}>|\{[^{}/]{1,24}\}|%username%|" &
            "用户名|用户目录|示例用户|placeholder|dummy|" &
            "your[\s_\-]?path|path[\\/]+to[\\/]|example\.(com|org)|" &
            "xxx+|\.\.\.|…" &
            ")",
            System.Text.RegularExpressions.RegexOptions.Compiled)

        ''' <summary>
        ''' 净化"模型/用户给出的**文件**输出路径"，避免把文件写进编造出来的目录。
        ''' 动机（实测）：用户说"将当前工作表导出为 CSV 并保存到桌面"，模型把用户名当占位符，
        ''' 传入 C:\Users\Yourusername\Desktop\output.csv；该值非空，绕过了前端默认路径推断，
        ''' 后端又无条件 CreateDirectory，于是磁盘上真被建出 C:\Users\Yourusername\Desktop\。
        ''' 规则：
        '''   ① X:\Users\&lt;非当前用户&gt;\... → 把该段替换为当前用户（保留 Desktop/Documents 等子路径）；
        '''   ② 命中占位符标记（Yourusername / 用户名 / &lt;username&gt; / ... 等）→ 改锚"真实桌面 + 原文件名"；
        '''   ③ 合法路径（含当前用户目录）原样返回。
        '''   ① 优先于 ②：C:\Users\YourUsername\Documents\report.pdf 应修正为
        '''   C:\Users\Administrator\Documents\report.pdf，而不是被一刀切扔到桌面。
        ''' </summary>
        ''' <param name="filePath">待净化的输出文件路径（可含环境变量）</param>
        ''' <returns>净化后的绝对路径</returns>
        Public Shared Function SanitizeUserOutputPath(filePath As String) As String
            If String.IsNullOrWhiteSpace(filePath) Then Return filePath
            Dim p As String = NormalizeUserPathInput(filePath)
            Dim fixedPath As String = TryFixUserSegment(p)
            If fixedPath IsNot Nothing Then Return fixedPath
            If PlaceholderPathRegex.IsMatch(p) Then Return BuildDesktopFileFallback(p)
            Return p
        End Function

        ''' <summary>
        ''' 净化"模型/用户给出的**目录**输出路径"（拆分导出、图表导出、分块保存等的 outputDir）。
        ''' 与 SanitizeUserOutputPath 同源，但目录场景不做"桌面 + 文件名"拼装。
        ''' </summary>
        ''' <param name="dirPath">待净化的输出目录</param>
        ''' <returns>净化后的目录；命中占位符时落到真实桌面</returns>
        Public Shared Function SanitizeUserDirectory(dirPath As String) As String
            If String.IsNullOrWhiteSpace(dirPath) Then Return dirPath
            Dim p As String = NormalizeUserPathInput(dirPath)
            Dim fixedPath As String = TryFixUserSegment(p)
            If fixedPath IsNot Nothing Then Return fixedPath
            If PlaceholderPathRegex.IsMatch(p) Then Return GetRealDesktop()
            Return p
        End Function

        ''' <summary>去空白/去包裹引号，并展开 %USERPROFILE% 之类环境变量与开头的 ~。</summary>
        Private Shared Function NormalizeUserPathInput(raw As String) As String
            Dim p As String = raw.Trim()
            p = p.Trim(""""c).Trim("'"c).Trim()
            Try
                p = Environment.ExpandEnvironmentVariables(p)
            Catch
            End Try
            ' 【防假成功】展开开头的 ~（Unix 风格 home 路径）。不展开时 .NET 会把 ~ 当成
            ' 普通相对目录名，Directory.CreateDirectory("~/Desktop") 会在进程 CWD 下
            ' 真造出字面名为 "~" 的文件夹（实测落盘 Documents\~\Desktop\），写入"成功"
            ' 但用户在真实桌面找不到文件。
            If p = "~" OrElse p.StartsWith("~/") OrElse p.StartsWith("~\") Then
                Dim rest As String = If(p.Length > 1, p.Substring(2), "")
                Dim profile As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
                If String.IsNullOrWhiteSpace(profile) Then profile = GetRealDesktop()
                If rest.Length > 0 Then
                    p = profile.TrimEnd("\"c, "/"c) & "\" & rest.TrimStart("\"c, "/"c)
                Else
                    p = profile
                End If
            End If
            Return p
        End Function

        ''' <summary>真实桌面目录（拿不到时回落到临时目录）。</summary>
        Private Shared Function GetRealDesktop() As String
            Dim desktop As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            If String.IsNullOrWhiteSpace(desktop) Then desktop = System.IO.Path.GetTempPath()
            Return desktop
        End Function

        ''' <summary>
        ''' 用户目录归属校验：X:\Users\&lt;seg&gt;\... / /Users/&lt;seg&gt;/... / /home/&lt;seg&gt;/...
        ''' 若 &lt;seg&gt; 不是当前用户，则替换为当前用户并保留其后的子路径（Desktop/Downloads/…）。
        ''' 返回 Nothing 表示"无需或无法修正，按原样处理"。
        ''' </summary>
        Private Shared Function TryFixUserSegment(p As String) As String
            Dim m As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(
                p,
                "^(?<root>[A-Za-z]:[\\/]Users[\\/]|/Users/|/home/)(?<user>[^\\/]+)(?<rest>[\\/].*)?$",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            If Not m.Success Then Return Nothing

            Dim realUser As String = Environment.GetEnvironmentVariable("USERNAME")
            If String.IsNullOrEmpty(realUser) Then realUser = Environment.UserName
            If String.IsNullOrEmpty(realUser) Then Return Nothing
            If String.Equals(m.Groups("user").Value, realUser, StringComparison.OrdinalIgnoreCase) Then Return Nothing

            ' 只认"当前用户 profile 目录真实存在"这一个条件，绝不为纠正再造陌生目录；
            ' 不校验子路径是否存在——输出目录本就常常还不存在。
            Dim profileRoot As String = m.Groups("root").Value & realUser
            If Not System.IO.Directory.Exists(profileRoot) Then Return Nothing

            Dim rest As String = If(m.Groups("rest").Success, m.Groups("rest").Value, "")
            Return profileRoot & rest
        End Function

        ''' <summary>真实桌面目录 + 原文件名（文件名缺失时回落 output.bin）。</summary>
        Private Shared Function BuildDesktopFileFallback(filePath As String) As String
            Dim fileName As String = ""
            Try
                fileName = System.IO.Path.GetFileName(filePath)
            Catch
            End Try
            If String.IsNullOrWhiteSpace(fileName) Then fileName = "output.bin"
            fileName = GetSafeFileName(fileName)
            Return System.IO.Path.Combine(GetRealDesktop(), fileName)
        End Function

    End Class

End Namespace