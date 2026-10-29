' ============================================================================
' HostPathHelper - 宿主（PPT/Word）侧的路径净化助手
'
' 为什么要单独一份（2026-10-05）：
'   PptBridge / WordBridge 原本直接调用 Excel 主项目的 Utils.PathHelper，但两个宿主
'   工程并没有编译那个文件（也不该为它拖进 ConfigConstants/ConfigurationManager
'   等一串依赖），于是 `Imports ExcelAIVSTOAddIn.Utils` 直接编译不过——
'   表现为"源码改了、DLL 却停在旧时间戳"（22:43 构建失败，DLL 仍是 13:42）。
'   本文件只依赖 BCL，两个宿主工程 link 它即可编译通过。
'
' 语义与 Utils.PathHelper 保持一致（前端 tool-registry.util.js 的
' isPlaceholderPath 与此对齐，形成双层防护）：
'   ① X:\Users\<非当前用户>\... → 该段替换为当前用户（保留 Desktop/Documents 等子路径）
'   ② 命中占位符标记（Yourusername / <username> / %username% / ...）→ 改锚"真实桌面 + 原文件名"
'   ③ 合法路径原样返回
'
' 注意：Module 的成员隐式 Shared，方法/字段**不要**再写 Shared 关键字（VB 会报 BC30433）。
' ============================================================================
Option Strict Off
Option Explicit On

Imports System
Imports System.IO
Imports System.Text.RegularExpressions

Public Module HostPathHelper

    ''' <summary>去掉路径末尾的连续分隔符。</summary>
    Private Function TrimTrailingSep(p As String) As String
        If String.IsNullOrEmpty(p) Then Return p
        Dim s As String = p
        Do While s.Length > 0 AndAlso (s.EndsWith("\") OrElse s.EndsWith("/"))
            s = s.Substring(0, s.Length - 1)
        Loop
        Return s
    End Function

    ''' <summary>当前登录用户的 profile 目录（如 C:\Users\Administrator）。</summary>
    Public Function GetRealUserProfile() As String
        Dim p As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
        If String.IsNullOrWhiteSpace(p) Then p = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
        If String.IsNullOrWhiteSpace(p) Then
            p = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "..")
        End If
        Return TrimTrailingSep(p)
    End Function

    ''' <summary>当前登录用户的真实桌面（宿主导出文件的默认落点）。</summary>
    Public Function GetRealDesktop() As String
        Dim d As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        If String.IsNullOrWhiteSpace(d) Then d = Path.Combine(GetRealUserProfile(), "Desktop")
        Return TrimTrailingSep(d)
    End Function

    ''' <summary>当前登录用户的"下载"目录；取不到时退回真实桌面。</summary>
    Public Function GetRealDownloads() As String
        Try
            Dim dl As String = Path.Combine(GetRealUserProfile(), "Downloads")
            If Directory.Exists(dl) Then Return dl
        Catch
        End Try
        Return GetRealDesktop()
    End Function

    ''' <summary>
    ''' 宿主 Config 目录（与 Excel PathHelper.ConfigDirectory 优先级一致：
    ''' 运行时 bin\...\Config → 开发源码 Config → 最终回退运行时路径并创建）。
    ''' 供 custom_ppt_themes.json 等与 Excel 主工程对齐的文件读写。
    ''' </summary>
    Public Function GetHostConfigDirectory() As String
        Dim assemblyDir As String = ""
        Try
            assemblyDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
        Catch
        End Try
        If String.IsNullOrWhiteSpace(assemblyDir) Then
            assemblyDir = AppDomain.CurrentDomain.BaseDirectory
        End If
        Dim runtimeConfig As String = Path.Combine(assemblyDir, "Config")
        If Directory.Exists(runtimeConfig) Then Return runtimeConfig
        Try
            Dim devConfig As String = Path.GetFullPath(Path.Combine(assemblyDir, "..", "..", "Config"))
            If Directory.Exists(devConfig) Then Return devConfig
        Catch
        End Try
        Try
            If Not Directory.Exists(runtimeConfig) Then Directory.CreateDirectory(runtimeConfig)
        Catch
        End Try
        Return runtimeConfig
    End Function

    ''' <summary>移除文件名里的非法字符。</summary>
    Public Function GetSafeFileName(fileName As String) As String
        If String.IsNullOrEmpty(fileName) Then Return String.Empty
        Dim s As String = fileName
        For Each c As Char In Path.GetInvalidFileNameChars()
            s = s.Replace(c, "_"c)
        Next
        Return s
    End Function

    ''' <summary>
    ''' 净化"模型/用户给出的文件输出路径"，避免把文件写进编造出来的目录。
    ''' 动机（实测）：模型常把用户名当占位符，传 C:\Users\Yourusername\Desktop\out.pdf，
    ''' 而后端又会无条件 CreateDirectory，磁盘上就真被建出该目录。
    ''' </summary>
    Public Function SanitizeUserOutputPath(filePath As String) As String
        If String.IsNullOrWhiteSpace(filePath) Then Return filePath

        Dim p As String = NormalizePathInput(filePath)
        Dim fixedPath As String = TryFixUserSegment(p)
        If fixedPath IsNot Nothing Then Return fixedPath
        If PlaceholderPathRegex.IsMatch(p) Then Return DesktopFileFallback(p)
        Return p
    End Function

    ''' <summary>目录版净化：命中占位符时直接落到真实桌面（不做文件名拼装）。</summary>
    Public Function SanitizeUserDirectory(dirPath As String) As String
        If String.IsNullOrWhiteSpace(dirPath) Then Return dirPath
        Dim p As String = NormalizePathInput(dirPath)
        Dim fixedPath As String = TryFixUserSegment(p)
        If fixedPath IsNot Nothing Then Return fixedPath
        If PlaceholderPathRegex.IsMatch(p) Then Return GetRealDesktop()
        Return p
    End Function

    ''' <summary>
    ''' 把 X:\Users\<段>\... 里的用户名段纠正为**当前登录用户**。
    ''' C:\Users\YourUsername\Documents\report.pdf → C:\Users\<当前用户>\Documents\report.pdf
    ''' 已是当前用户的合法路径原样返回（返回 Nothing 表示"无需修正"）。
    ''' </summary>
    Private Function TryFixUserSegment(p As String) As String
        If String.IsNullOrWhiteSpace(p) Then Return Nothing
        Dim m As Match = Regex.Match(p, "^([A-Za-z]:\\Users\\)([^\\]+)(\\.*)$", RegexOptions.IgnoreCase)
        If Not m.Success Then Return Nothing

        Dim seg As String = m.Groups(2).Value
        If String.Equals(seg, Environment.UserName, StringComparison.OrdinalIgnoreCase) Then Return Nothing
        If String.Equals(seg, "Public", StringComparison.OrdinalIgnoreCase) Then Return Nothing

        ' 域账号/重定向场景：用户目录名可能不等于 UserName，用 profile 实际路径的末段再比一次
        Dim profile As String = GetRealUserProfile()
        Dim profileLeaf As String = profile
        Dim idx As Integer = profile.LastIndexOf("\"c)
        If idx > 0 Then profileLeaf = profile.Substring(idx + 1)
        If String.Equals(seg, profileLeaf, StringComparison.OrdinalIgnoreCase) Then Return Nothing

        Return TrimTrailingSep(profile) & m.Groups(3).Value
    End Function

    ''' <summary>占位符路径改锚到"真实桌面 + 原文件名"。</summary>
    Private Function DesktopFileFallback(p As String) As String
        ' 不用 Path.GetFileName：路径里带 <username> 这类尖括号时它会抛 ArgumentException，
        ' 结果文件名整个丢失（实测 C:\path\to\<username>\x.pdf → export）。
        ' 这里纯字符串切末段，对畸形路径同样稳。
        Dim name As String = ""
        Try
            Dim i As Integer = p.LastIndexOfAny(New Char() {"\"c, "/"c})
            If i >= 0 AndAlso i < p.Length - 1 Then name = p.Substring(i + 1)
        Catch
        End Try
        If String.IsNullOrWhiteSpace(name) Then name = "export"
        name = GetSafeFileName(name)
        Return Path.Combine(GetRealDesktop(), name)
    End Function

    ''' <summary>去空白/去包裹引号，并展开 %USERPROFILE% 之类环境变量与开头的 ~。</summary>
    Private Function NormalizePathInput(raw As String) As String
        If raw Is Nothing Then Return String.Empty
        ' Trim('"') 会让 VB 在重载解析上出错（'"' 是 1 字符字符串而非 Char 字面量），用 ChrW(34)
        Dim s As String = raw.Trim().Trim(ChrW(34))
        If s.Length = 0 Then Return s
        Try
            s = Environment.ExpandEnvironmentVariables(s)
        Catch
        End Try
        ' 【防假成功】展开开头的 ~（Unix 风格 home 路径）。
        ' 不展开时 .NET 会把 ~ 当成普通相对目录名：Directory.CreateDirectory("~/Desktop")
        ' 会在进程 CWD 下真造出字面名为 "~" 的文件夹（实测落盘 Documents\~\Desktop\），
        ' 写入"成功"但用户在真实桌面找不到文件。
        If s = "~" OrElse s.StartsWith("~/") OrElse s.StartsWith("~\") Then
            Dim rest As String = If(s.Length > 1, s.Substring(2), "")
            Dim profile As String = GetRealUserProfile()
            s = If(rest.Length > 0, TrimTrailingSep(profile) & "\" & rest.TrimStart("\"c, "/"c), profile)
        End If
        Return s.Trim()
    End Function

    ''' <summary>
    ''' 输出路径里"编造的用户名/占位符"判定（与 Utils.PathHelper.PlaceholderPathRegex 同源）。
    ''' </summary>
    Private ReadOnly PlaceholderPathRegex As New Regex(
        "(?i)(" &
        "your[\s_\-]?user\w*|yourname|your[\s_\-]?name|some[\s_\-]?user|test[\s_\-]?user|" &
        "user[\s_\-]?name|<[^<>/]{1,24}>|\{[^{}/]{1,24}\}|%username%|" &
        "用户名|用户目录|示例用户|placeholder|dummy|" &
        "your[\s_\-]?path|path[\\/]+to[\\/]|example\.(com|org)|" &
        "xxx+|\.\.\.|…" &
        ")",
        RegexOptions.Compiled)

End Module
