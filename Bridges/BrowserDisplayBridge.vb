' 本文件大量使用 Shell.Application late binding（COM 互操作），
' 显式 Option Strict Off 保证与宿主工程配置无关。
Option Strict Off
Option Explicit On
Option Infer On

Imports System.IO
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.Security.Cryptography
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Bridges

    ''' <summary>
    ''' BrowserDisplayBridge - HTML 浏览器展示桥梁类
    ''' 将 HTML 内容写入临时文件并用系统默认浏览器打开
    ''' </summary>
    <ComVisible(True)>
    <ClassInterface(ClassInterfaceType.AutoDispatch)>
    <ProgId("ExcelAIVSTOAddIn.BrowserDisplayBridge")>
    Public Class BrowserDisplayBridge
        Implements IBridge

        ' 单例实例
        Private Shared ReadOnly _lazyInstance As New Lazy(Of BrowserDisplayBridge)(Function() New BrowserDisplayBridge())

        ' 临时文件目录
        Private ReadOnly _tempDir As String = Path.Combine(Path.GetTempPath(), "ExcelAI_Display")

        ' 幂等缓存：防止同一 HTML 在短时间内被重复写入+打开
        ' key = SHA1(title + htmlContent), value = (filePath, lastAccessUtc)
        ' 30 秒内重复调用同一内容 → 直接复用已有文件（不创建新副本）
        Private ReadOnly _recentCache As New Dictionary(Of String, (filePath As String, ts As DateTime))
        Private ReadOnly _cacheLock As New Object()
        Private Const IdleTtlSeconds As Integer = 30

        ' ── L2 title 级去重（5 分钟窗口，跨轮次） ──
        ' 场景：用户第一次触发"生成销售看板"→ tab1 开；
        '       3 分钟后再触发同一 title（可能数据更新了，SHA1 不同）→ 旧 SHA1 缓存不命中
        '       → 旧 L1 也过了（>10s）→ 又 Process.Start 开 tab2；
        '       再触发几次 → tab3、tab4...
        '
        ' 解法：5 分钟内同一 title 只 Process.Start 一次。
        ' 第二次调用时：覆盖写文件 + 返回"数据已更新，浏览器刷新即可" + 不 Process.Start。
        ' 用户在已开的那个 tab 里 Ctrl+R 就能看到新数据。
        Private ReadOnly _titleLastOpen As New Dictionary(Of String, (filePath As String, ts As DateTime))
        Private ReadOnly _titleLock As New Object()
        Private Const TitleWindowMinutes As Integer = 5
        Private _lastCall As (title As String, contentLen As Integer, ts As DateTime, filePath As String) =
            ("", 0, DateTime.MinValue, "")
        Private ReadOnly _lastCallLock As New Object()
        Private Const DebounceWindowMs As Integer = 10000  ' 10 秒窗口（覆盖 DashboardGenerator 双调的 2-4s 间隔）

#Region "单例模式"

        Public Shared ReadOnly Property Instance As BrowserDisplayBridge
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        Private Sub New()
            Try
                ' 确保临时目录存在
                If Not Directory.Exists(_tempDir) Then
                    Directory.CreateDirectory(_tempDir)
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("BrowserDisplayBridge 初始化异常: " & ex.Message)
            End Try
        End Sub

#End Region

#Region "HTML 浏览器展示 API"

        ''' <summary>
        ''' 将 HTML 内容写入临时文件并用系统默认浏览器打开
        ''' </summary>
        ''' <param name="htmlContent">完整 HTML 内容（含 html/head/body 标签）</param>
        ''' <param name="title">页面标题，也用于生成文件名</param>
        ''' <returns>操作结果 JSON</returns>
        <ComVisible(True)>
        Public Function OpenHtmlInBrowser(htmlContent As String, Optional title As String = "ExcelAI报表") As String
            Try
                If String.IsNullOrWhiteSpace(htmlContent) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "HTML 内容不能为空"})
                End If

                ' ── 调用溯源日志 ──
                ' 记录调用方（MCP tool / Skill Bridge / 前端 WebView2 COM），
                ' 便于排查"为什么同一轮渲染两次 HTML"。
                Try
                    Dim trace = New StackTrace(False)
                    Dim frames = trace.GetFrames()
                    Dim callerNames As New List(Of String)()
                    For Each f In frames
                        Dim m = f.GetMethod()
                        If m Is Nothing Then Continue For
                        Dim fullName = m.DeclaringType?.Name & "." & m.Name
                        ' 跳过本类内部帧，保留外部调用者
                        If Not fullName.StartsWith("BrowserDisplayBridge.") Then
                            callerNames.Add(fullName)
                            If callerNames.Count >= 5 Then Exit For
                        End If
                    Next
                    Dim callerTag = String.Join(" ← ", callerNames)
                    CallSiteLogger.Log("OpenHtmlInBrowser", callerTag, title, htmlContent.Length)
                    System.Diagnostics.Debug.WriteLine($"[BrowserDisplayBridge] 调用: title=""{title}"", len={htmlContent.Length}, 调用链: {callerTag}")
                Catch
                End Try

                ' ── L1 瞬时防抖层（10 秒窗口，按 title + len） ──
                ' 已证实：DashboardGenerator 每次渲染被调两次，间隔 2-4 秒。
                ' 即使 SHA1 不同（内容有时间戳差异），title + len 通常相同。
                SyncLock _lastCallLock
                    Dim now = DateTime.Now
                    Dim elapsed = (now - _lastCall.ts).TotalMilliseconds
                    Dim titleMatch = (_lastCall.title = title)
                    Dim lenMatch = (_lastCall.contentLen = htmlContent.Length)
                    Dim fileOk = (Not String.IsNullOrEmpty(_lastCall.filePath)) AndAlso File.Exists(_lastCall.filePath)
                    Dim inWindow = (elapsed < DebounceWindowMs)
                    Dim willHit = inWindow AndAlso titleMatch AndAlso lenMatch AndAlso fileOk

                    CallSiteLogger.Log("L1_Check",
                        $"elapsed={elapsed:F0}ms<{DebounceWindowMs}={inWindow}; titleMatch={titleMatch}; lenMatch={lenMatch}; fileOk={fileOk} → hit={willHit}",
                        title, htmlContent.Length)

                    If willHit Then
                        System.Diagnostics.Debug.WriteLine($"[BrowserDisplayBridge] L1防抖命中，跳过写文件+Process.Start: {_lastCall.filePath}")
                        CallSiteLogger.Log("L1_Hit", _lastCall.filePath, title, htmlContent.Length)
                        Dim reopenJson = ReopenHtmlFile(_lastCall.filePath)
                        Return reopenJson
                    End If
                End SyncLock

                ' ── 幂等缓存：30 秒内重复调用同一内容 → 复用已有文件 ──
                Dim cacheKey = ComputeCacheKey(title, htmlContent)
                SyncLock _cacheLock
                    ' 清理过期条目（顺便做懒清理）
                    Dim now = DateTime.UtcNow
                    Dim expired = _recentCache.Where(Function(kv) (now - kv.Value.ts).TotalSeconds > IdleTtlSeconds).Select(Function(kv) kv.Key).ToList()
                    For Each k In expired
                        _recentCache.Remove(k)
                    Next
                    ' 命中缓存且文件还在 → 直接 Reopen，不创建新文件
                    Dim hit As (filePath As String, ts As DateTime) = ("", DateTime.MinValue)
                    If _recentCache.TryGetValue(cacheKey, hit) AndAlso File.Exists(hit.filePath) Then
                        _recentCache(cacheKey) = (hit.filePath, now)
                        ' SHA1 幂等命中也更新 _lastCall，让 L1 防抖层下一次能挡住
                        SyncLock _lastCallLock
                            _lastCall = (title, htmlContent.Length, DateTime.Now, hit.filePath)
                        End SyncLock
                        Dim reopenJson = ReopenHtmlFile(hit.filePath)
                        System.Diagnostics.Debug.WriteLine($"[BrowserDisplayBridge] 幂等命中，复用: {hit.filePath}")
                        ' 把 reopen 结果里 success=True 透传回去，保持返回结构一致
                        Dim reopenResp = JsonConvert.DeserializeObject(Of JObject)(reopenJson)
                        If reopenResp IsNot Nothing AndAlso reopenResp("success")?.Value(Of Boolean)() Then
                            Return JsonConvert.SerializeObject(New With {
                                .success = True,
                                .message = "HTML 报表已在浏览器中打开（复用最近文件）",
                                .filePath = hit.filePath,
                                .title = title,
                                .idempotent = True
                            })
                        End If
                    End If
                End SyncLock

                ' 清理标题作为文件名的一部分
                Dim safeTitle = If(String.IsNullOrWhiteSpace(title), "ExcelAI报表", title)
                safeTitle = String.Join("_", safeTitle.Split(Path.GetInvalidFileNameChars()))
                If safeTitle.Length > 50 Then safeTitle = safeTitle.Substring(0, 50)

                ' ╔═══════════════════════════════════════════════════════════════╗
                ' ║ 文件名策略：按 title 固定命名（不加时间戳），直接覆盖。          ║
                ' ║                                                               ║
                ' ║ 为什么：                                                       ║
                ' ║ 1) Agent 常因多路径/重试/轮次间隔生成"同名不同内容"的 HTML，    ║
                ' ║    时间戳策略会让 Temp 目录堆满近似重复文件；                    ║
                ' ║ 2) 浏览器打开 file:// 不会锁定 HTML（只读访问），覆盖安全；      ║
                ' ║ 3) 真正需要历史版本的用户 → 另存为，不是 AI 生成器的职责。       ║
                ' ╚═══════════════════════════════════════════════════════════════╝
                Dim fileName = $"ExcelAI_{safeTitle}.html"
                Dim filePath = Path.Combine(_tempDir, fileName)

                ' 如果目标文件已被浏览器锁定（极少见）→ 退化为带序号的唯一文件名
                If File.Exists(filePath) Then
                    Try
                        Dim testStream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
                        testStream.Close()
                    Catch
                        ' 被独占 → 加随机后缀避免冲突
                        fileName = $"ExcelAI_{safeTitle}_{Guid.NewGuid().ToString("N").Substring(0, 6)}.html"
                        filePath = Path.Combine(_tempDir, fileName)
                    End Try
                End If

                ' 如果用户只提供了 body 内容，自动包装成完整 HTML
                Dim fullHtml = htmlContent
                If Not htmlContent.Trim().ToLower().StartsWith("<html") AndAlso
                   Not htmlContent.Trim().ToLower().StartsWith("<!doctype") Then
                    fullHtml = $"<!DOCTYPE html>" & vbCrLf &
                               $"<html lang=""zh-CN"">" & vbCrLf &
                               $"<head>" & vbCrLf &
                               $"<meta charset=""UTF-8"">" & vbCrLf &
                               $"<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">" & vbCrLf &
                               $"<title>{safeTitle}</title>" & vbCrLf &
                               $"</head>" & vbCrLf &
                               $"<body style=""margin:0;padding:20px;font-family:'Segoe UI',Roboto,Helvetica,Arial,sans-serif;background:#f5f5f5;"">" & vbCrLf &
                               $"{htmlContent}" & vbCrLf &
                               $"</body></html>"
                End If

                ' 离线优先：把 __ECHARTS_LOCAL__ 占位替换为本地内联的 echarts，
                ' 避免系统浏览器以 file:// 打开临时 HTML 时 CDN 加载失败导致图表全白。
                If fullHtml.Contains("__ECHARTS_LOCAL__") Then
                    Dim echartsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "UI", "lib", "echarts.min.js")
                    Dim echartsInline = ""
                    Try
                        If File.Exists(echartsPath) Then
                            echartsInline = File.ReadAllText(echartsPath, System.Text.Encoding.UTF8)
                        End If
                    Catch
                        echartsInline = ""
                    End Try
                    If Not String.IsNullOrEmpty(echartsInline) Then
                        fullHtml = fullHtml.Replace("<script src=""__ECHARTS_LOCAL__""></script>", "<script>" & echartsInline & "</script>")
                    Else
                        ' 本地文件缺失 → 回退 CDN，保证联网时仍可用
                        fullHtml = fullHtml.Replace("<script src=""__ECHARTS_LOCAL__""></script>", "<script src=""https://cdn.jsdelivr.net/npm/echarts@5.5.0/dist/echarts.min.js""></script>")
                    End If
                End If

                ' 写入文件（UTF-8 无 BOM）
                File.WriteAllText(filePath, fullHtml, New System.Text.UTF8Encoding(False))

                ' ── L2 title 级去重（5 分钟窗口，跨轮次） ──
                Dim shouldProcessStart = True
                Dim alreadyOpenInBrowser = False
                Dim hitTitleWindow = False
                Try
                    ' L2a: title 级时间窗口检查（5 分钟内同一 title 只 Process.Start 一次）
                    SyncLock _titleLock
                        Dim titleKey = title
                        Dim elapsedMin As Double
                        If _titleLastOpen.TryGetValue(titleKey, Nothing) Then
                            Dim existing = _titleLastOpen(titleKey)
                            elapsedMin = (DateTime.Now - existing.ts).TotalMinutes
                            If elapsedMin < TitleWindowMinutes Then
                                hitTitleWindow = True
                                ' 覆盖已有 filePath（可能用户换了数据但 title 相同）
                                _titleLastOpen(titleKey) = (filePath, DateTime.Now)
                            Else
                                _titleLastOpen(titleKey) = (filePath, DateTime.Now)
                            End If
                        Else
                            _titleLastOpen(titleKey) = (filePath, DateTime.Now)
                        End If
                    End SyncLock

                    ' L2b: ShellWindows COM 检查（浏览器是否已打开这个 filePath）
                    If Not hitTitleWindow Then
                        Try
                            alreadyOpenInBrowser = IsFileOpenInBrowser(filePath)
                        Catch
                        End Try
                    End If

                    shouldProcessStart = Not hitTitleWindow AndAlso Not alreadyOpenInBrowser
                Catch
                    shouldProcessStart = True  ' 任何异常兜底：照常打开
                End Try

                If shouldProcessStart Then
                    Dim psi As New ProcessStartInfo(filePath) With {
                        .UseShellExecute = True
                    }
                    Process.Start(psi)
                    System.Diagnostics.Debug.WriteLine("[BrowserDisplayBridge] Process.Start 新打开: " & filePath)
                    CallSiteLogger.Log("ProcessStart", "新打开", title, htmlContent.Length)
                Else
                    Dim skipReason = If(hitTitleWindow, "L2_title窗口命中(5min内同title只开一次)", "L3_ShellWindows已打开")
                    System.Diagnostics.Debug.WriteLine("[BrowserDisplayBridge] 跳过 Process.Start: " & skipReason)
                    CallSiteLogger.Log("ProcessStart_Skipped", skipReason, title, htmlContent.Length)
                End If

                ' ── 写入幂等缓存（仅在成功时）──
                SyncLock _cacheLock
                    _recentCache(cacheKey) = (filePath, DateTime.UtcNow)
                End SyncLock

                ' ── 更新 L1 防抖状态 ──
                SyncLock _lastCallLock
                    _lastCall = (title, htmlContent.Length, DateTime.Now, filePath)
                End SyncLock

                Dim msg = If(shouldProcessStart,
                    "HTML 报表已在浏览器中打开",
                    If(hitTitleWindow,
                        "数据已更新（覆盖写入 " & Path.GetFileName(filePath) & "），浏览器中已有该页面，请按 Ctrl+R 刷新查看最新内容",
                        "HTML 报表已在浏览器中打开（跳过 Process.Start）"))
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .message = msg,
                    .filePath = filePath,
                    .title = safeTitle,
                    .openedNewTab = shouldProcessStart,
                    .skipReason = If(shouldProcessStart, "", If(hitTitleWindow, "L2_title_window", "L3_already_open"))
                })

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[BrowserDisplayBridge] 打开浏览器失败: {ex.Message}")
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 重新打开已存在的 HTML 文件（用于页面关闭后恢复）
        ''' </summary>
        ''' <param name="filePath">HTML 文件的完整路径</param>
        ''' <returns>操作结果 JSON</returns>
        <ComVisible(True)>
        Public Function ReopenHtmlFile(filePath As String) As String
            Try
                If String.IsNullOrWhiteSpace(filePath) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "文件路径不能为空"})
                End If

                If Not File.Exists(filePath) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "文件不存在: " & filePath})
                End If

                ' 用默认浏览器重新打开
                Dim psi As New ProcessStartInfo(filePath) With {
                    .UseShellExecute = True
                }
                Process.Start(psi)

                System.Diagnostics.Debug.WriteLine($"[BrowserDisplayBridge] HTML 已重新打开: {filePath}")

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .message = "HTML 报表已在浏览器中重新打开",
                    .filePath = filePath
                })

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[BrowserDisplayBridge] 重新打开失败: {ex.Message}")
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 使用内置 ECharts 模板生成交互式图表页面并在浏览器中打开
        ''' </summary>
        ''' <param name="chartType">图表类型: bar/line/pie/scatter/radar/heatmap</param>
        ''' <param name="dataJson">图表数据 JSON，包含 labels 和 datasets</param>
        ''' <param name="title">图表标题</param>
        ''' <param name="theme">主题: light 或 dark</param>
        ''' <returns>操作结果 JSON</returns>
        <ComVisible(True)>
        Public Function OpenChartInBrowser(chartType As String, dataJson As String, Optional title As String = "数据图表", Optional theme As String = "light") As String
            Try
                Dim data = JsonConvert.DeserializeObject(Of JObject)(dataJson)
                Dim labels = If(data("labels")?.ToObject(Of List(Of String))(), New List(Of String)())
                Dim datasets = If(data("datasets")?.ToObject(Of List(Of JObject))(), New List(Of JObject)())

                Dim bgColor = If(theme.ToLower() = "dark", "#1a1a2e", "#ffffff")
                Dim textColor = If(theme.ToLower() = "dark", "#e0e0e0", "#333333")

                ' 构建 ECharts 配置
                Dim seriesList As New List(Of String)()
                For Each ds In datasets
                    Dim seriesName = If(ds("name")?.ToString(), "数据系列")
                    Dim values = ds("values")?.ToObject(Of List(Of Object))()
                    Dim seriesJson = $"{{name:'{seriesName}',type:'{chartType}',data:{JsonConvert.SerializeObject(values)}}}"
                    seriesList.Add(seriesJson)
                Next

                Dim html = $"<!DOCTYPE html>
<html lang=""zh-CN"">
<head>
<meta charset=""UTF-8"">
<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">
<title>{title}</title>
<script src=""https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js""></script>
<style>
body{{margin:0;padding:20px;font-family:'Segoe UI',Roboto,Helvetica,Arial,sans-serif;background:{bgColor};color:{textColor};}}
.container{{max-width:1200px;margin:0 auto;background:{If(theme="dark","#16213e","#fff")};border-radius:8px;box-shadow:0 2px 12px rgba(0,0,0,0.1);padding:24px;}}
h1{{margin:0 0 20px 0;font-size:22px;font-weight:600;}}
#chart{{width:100%;height:500px;}}
.footer{{margin-top:16px;font-size:12px;color:#999;text-align:right;}}
</style>
</head>
<body>
<div class=""container"">
<h1>{title}</h1>
<div id=""chart""></div>
<div class=""footer"">Generated by Excel AI Plugin</div>
</div>
<script>
var chart = echarts.init(document.getElementById('chart'), '{If(theme="dark","dark","")}');
var option = {{
    tooltip: {{ trigger: 'axis' }},
    legend: {{ data: {JsonConvert.SerializeObject(datasets.Select(Function(d) d("name")?.ToString()).ToList())} }},
    grid: {{ left: '3%', right: '4%', bottom: '3%', containLabel: true }},
    xAxis: {{ type: 'category', data: {JsonConvert.SerializeObject(labels)} }},
    yAxis: {{ type: 'value' }},
    series: [{String.Join(",", seriesList)}]
}};
chart.setOption(option);
window.addEventListener('resize', function(){{ chart.resize(); }});
</script>
</body>
</html>"

                Return OpenHtmlInBrowser(html, title)

            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = $"生成图表失败: {ex.Message}"})
            End Try
        End Function

#End Region

#Region "IBridge 生命周期"

        Public ReadOnly Property BridgeName As String Implements IBridge.BridgeName
            Get
                Return "BrowserDisplayBridge"
            End Get
        End Property

        Private _disposed As Boolean = False
        Public Sub Dispose() Implements IBridge.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not _disposed Then
                _disposed = True
            End If
        End Sub

        ''' <summary>幂等缓存 key：SHA1(title + \u0000 + htmlContent)。
        ''' SHA1 对这种非安全场景够用且快；title 参与 key 防止不同标题的同内容被错误去重。</summary>
        Private Shared Function ComputeCacheKey(title As String, htmlContent As String) As String
            Try
                Dim raw = Encoding.UTF8.GetBytes(If(title, "") & ChrW(0) & htmlContent)
                Using sha = SHA1.Create()
                    Dim hash = sha.ComputeHash(raw)
                    Dim sb As New StringBuilder(hash.Length * 2)
                    For Each b In hash
                        sb.Append(b.ToString("x2"))
                    Next
                    Return sb.ToString()
                End Using
            Catch
                Return "len_" & If(title, "").Length & "_" & htmlContent.Length.ToString()
            End Try
        End Function

        ''' <summary>
        ''' 检测指定本地 HTML 文件是否已在浏览器标签页中打开。
        ''' 通过 Shell.Application COM 枚举所有 ShellWindow（IE/Edge/Chrome 均注册），
        ''' 对每个 ShellWindow.LocationURL 检查是否以 file:// 开头且包含该文件名。
        '''
        ''' 注意：late binding（Option Strict Off 兼容），不需要 SHDocVw PIA。
        ''' </summary>
        Private Shared Function IsFileOpenInBrowser(filePath As String) As Boolean
            Try
                ' Shell.Application 枚举所有已打开的浏览器/Explorer 窗口
                Dim shell = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"))
                If shell Is Nothing Then Return False
                Dim windows = shell.Namespace(0).Parent.Windows()
                For Each win In windows
                    Try
                        Dim url As String = ""
                        Try
                            url = CStr(win.LocationURL)
                        Catch
                            ' 某些窗口（如文件资源管理器）没有 LocationURL → 跳过
                            Continue For
                        End Try
                        If String.IsNullOrEmpty(url) Then Continue For
                        ' file:// URL: 小写化路径名匹配
                        Dim lowerFile = filePath.Replace("\", "/").ToLower()
                        Dim lowerUrl = url.ToLower()
                        If lowerUrl.StartsWith("file://") AndAlso lowerUrl.Contains(lowerFile) Then
                            Marshal.ReleaseComObject(win)
                            Marshal.ReleaseComObject(windows)
                            Marshal.ReleaseComObject(shell)
                            Return True
                        End If
                    Catch
                    End Try
                Next
                Marshal.ReleaseComObject(windows)
                Marshal.ReleaseComObject(shell)
            Catch
                ' COM 异常兜底
            End Try
            Return False
        End Function

#End Region

    End Class

    ''' <summary>
    ''' 轻量调用溯源日志：用于排查"同一轮渲染两次 HTML"类问题。
    ''' 写到 %TEMP%\ExcelAI_CallTrace.log，一行一次调用：
    '''   [时间戳] 方法名 | 调用链 | title | contentLength | hashCode
    ''' </summary>
    Friend Module CallSiteLogger
        Private ReadOnly _logPath As String =
            Path.Combine(Path.GetTempPath(), "ExcelAI_CallTrace.log")
        Private ReadOnly _lock As New Object()

        Public Sub Log(method As String, callChain As String, title As String, contentLen As Integer)
            Try
                SyncLock _lock
                    Dim line = $"[{DateTime.Now:HH:mm:ss.fff}] {method} | {callChain} | title=""{If(title, "")}"" | len={contentLen}"
                    File.AppendAllText(_logPath, line & Environment.NewLine)
                End SyncLock
            Catch
            End Try
        End Sub
    End Module

End Namespace