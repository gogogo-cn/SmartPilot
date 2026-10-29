Option Strict On
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports ExcelAIVSTOAddIn.Services
Imports ExcelAIVSTOAddIn.Services.Mcp

Namespace Services.Mcp

    Partial Public Class ComputerAgentTools

        Private Shared ReadOnly _cache As New Dictionary(Of String, MethodInfo)(StringComparer.OrdinalIgnoreCase)
        Private Shared _initialized As Boolean = False
        ' 默认开启：LLM 不得未经确认执行任意命令 / 杀进程 / 打开任意路径。
        ' computer_set_safe_mode(enabled=false) 可关闭；关闭后 execute_command 仍只允许极短白名单。
        Private Shared _safe As Boolean = True

        ' ── 工具定义 ──

        Public Shared Function GetToolDefinitions() As List(Of McpTool)
            Dim t As New List(Of McpTool)()
            t.Add(New McpTool("computer_screen_capture", "截图返回 base64 或保存路径", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"x", New McpToolProperty("integer", "左上X，可选")}, {"y", New McpToolProperty("integer", "左上Y")},
                {"width", New McpToolProperty("integer", "宽")}, {"height", New McpToolProperty("integer", "高")},
                {"returnFormat", New McpToolProperty("string", "base64 或 path")}}}))
            t.Add(New McpTool("computer_screen_get_info", "获取屏幕分辨率信息", New McpToolInputSchema()))
            t.Add(New McpTool("computer_screen_ocr", "截图+OCR返回识别文字", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty)()}))
            t.Add(New McpTool("computer_describe", "AI视觉描述屏幕内容", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"prompt", New McpToolProperty("string", "描述指令")}}}))
            t.Add(New McpTool("computer_mouse_move", "移动鼠标到指定坐标", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"x", New McpToolProperty("integer", "X")}, {"y", New McpToolProperty("integer", "Y")}}, .Required = New List(Of String) From {"x", "y"}}))
            t.Add(New McpTool("computer_mouse_click", "点击鼠标按钮", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"x", New McpToolProperty("integer", "X")}, {"y", New McpToolProperty("integer", "Y")},
                {"button", New McpToolProperty("string", "left/right/middle")}, {"clicks", New McpToolProperty("integer", "次数")}}}))
            t.Add(New McpTool("computer_mouse_drag", "拖动鼠标从起点到终点", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"fromX", New McpToolProperty("integer", "起点X")}, {"fromY", New McpToolProperty("integer", "起点Y")},
                {"toX", New McpToolProperty("integer", "终点X")}, {"toY", New McpToolProperty("integer", "终点Y")}},
                .Required = New List(Of String) From {"fromX", "fromY", "toX", "toY"}}))
            t.Add(New McpTool("computer_mouse_scroll", "鼠标滚轮滚动", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"amount", New McpToolProperty("integer", "量")}}, .Required = New List(Of String) From {"amount"}}))
            t.Add(New McpTool("computer_key_press", "按下并释放一个键", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"key", New McpToolProperty("string", "键名如 Enter/Tab/F1")}}, .Required = New List(Of String) From {"key"}}))
            t.Add(New McpTool("computer_key_hotkey", "组合键", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"keys", New McpToolProperty("string", "如 Ctrl+C")}}, .Required = New List(Of String) From {"keys"}}))
            t.Add(New McpTool("computer_type_text", "输入文字", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"text", New McpToolProperty("string", "文字")}}, .Required = New List(Of String) From {"text"}}))
            t.Add(New McpTool("computer_locate_text", "OCR定位屏幕上指定文字", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"text", New McpToolProperty("string", "搜索目标")}}, .Required = New List(Of String) From {"text"}}))
            t.Add(New McpTool("computer_locate_color", "定位屏幕上指定颜色", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"color", New McpToolProperty("string", "#RRGGBB")}, {"tolerance", New McpToolProperty("integer", "容差0-100")}},
                .Required = New List(Of String) From {"color"}}))
            t.Add(New McpTool("computer_click_text", "OCR定位并点击屏幕上指定文字（候选词逗号分隔按优先级）", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"texts", New McpToolProperty("string", "候选词，逗号分隔，如 播放,continue 播放,开始播放")},
                {"region_x", New McpToolProperty("integer", "区域X可选")}, {"region_y", New McpToolProperty("integer", "区域Y可选")},
                {"region_w", New McpToolProperty("integer", "区域宽可选")}, {"region_h", New McpToolProperty("integer", "区域高可选")}},
                .Required = New List(Of String) From {"texts"}}))
            t.Add(New McpTool("computer_clipboard_read", "读取剪贴板文字", New McpToolInputSchema()))
            t.Add(New McpTool("computer_clipboard_write", "写入剪贴板文字", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"text", New McpToolProperty("string", "文字")}}, .Required = New List(Of String) From {"text"}}))
            t.Add(New McpTool("computer_open_application", "打开应用程序或文件", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"path", New McpToolProperty("string", "路径")}, {"arguments", New McpToolProperty("string", "参数")}},
                .Required = New List(Of String) From {"path"}}))
            t.Add(New McpTool("computer_execute_command", "执行只读系统查询命令（安全模式默认拒绝；关闭后仅允许 hostname/whoami/ver 等白名单，禁止任意 cmd）", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"command", New McpToolProperty("string", "命令")}}, .Required = New List(Of String) From {"command"}}))
            t.Add(New McpTool("computer_wait", "等待毫秒数", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"ms", New McpToolProperty("integer", "毫秒")}}, .Required = New List(Of String) From {"ms"}}))
            t.Add(New McpTool("computer_set_safe_mode", "切换安全模式", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"enabled", New McpToolProperty("boolean", "true/false")}}, .Required = New List(Of String) From {"enabled"}}))
            ' ── 定位与感知增强 ──
            t.Add(New McpTool("computer_mouse_position_get", "获取当前鼠标光标坐标", New McpToolInputSchema()))
            t.Add(New McpTool("computer_window_list", "列出当前可见顶层窗口标题", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"keyword", New McpToolProperty("string", "按标题关键词过滤，可选")}}}))
            t.Add(New McpTool("computer_process_list", "列出运行中的进程", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"name", New McpToolProperty("string", "按进程名过滤，可选")}}}))
            t.Add(New McpTool("computer_process_kill", "结束指定进程", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"name", New McpToolProperty("string", "进程名如 notepad/notepad.exe")}}, .Required = New List(Of String) From {"name"}}))
            t.Add(New McpTool("computer_close_application", "按应用名（口语/别名）关闭程序，确保真正退出", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"app", New McpToolProperty("string", "应用名/别名，如 音乐、浏览器、QQ音乐")}}, .Required = New List(Of String) From {"app"}}))
            t.Add(New McpTool("computer_play_media", "聚焦应用并触发播放，OCR验证是否正在播放", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"app", New McpToolProperty("string", "应用名/别名，如 QQ音乐、音乐")}}, .Required = New List(Of String) From {"app"}}))
            ' ── 等待与校验 ──
            t.Add(New McpTool("computer_wait_text", "等待屏幕出现指定文字（轮询OCR直到超时或命中）", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"text", New McpToolProperty("string", "要等待出现的文字")},
                {"timeout_ms", New McpToolProperty("integer", "超时毫秒，默认10000")},
                {"region_x", New McpToolProperty("integer", "区域X可选")}, {"region_y", New McpToolProperty("integer", "区域Y可选")},
                {"region_w", New McpToolProperty("integer", "区域宽可选")}, {"region_h", New McpToolProperty("integer", "区域高可选")}},
                .Required = New List(Of String) From {"text"}}))
            t.Add(New McpTool("computer_assert_text", "校验屏幕是否出现/消失指定文字（返回found）", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"text", New McpToolProperty("string", "待校验文字")},
                {"expect", New McpToolProperty("string", "present 或 absent，默认 present")},
                {"region_x", New McpToolProperty("integer", "区域X可选")}, {"region_y", New McpToolProperty("integer", "区域Y可选")},
                {"region_w", New McpToolProperty("integer", "区域宽可选")}, {"region_h", New McpToolProperty("integer", "区域高可选")}},
                .Required = New List(Of String) From {"text"}}))
            t.Add(New McpTool("computer_focus_window", "按应用名/别名将已打开窗口切换置前（不打开新窗口）", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"app", New McpToolProperty("string", "应用名/别名，如 QQ音乐、浏览器、计算器")}}, .Required = New List(Of String) From {"app"}}))
            t.Add(New McpTool("computer_capture_window", "按应用名/别名截图其窗口区域（非整屏），返回 base64 或路径", New McpToolInputSchema() With {.Properties = New Dictionary(Of String, McpToolProperty) From {
                {"app", New McpToolProperty("string", "应用名/别名，如 QQ音乐、浏览器")},
                {"returnFormat", New McpToolProperty("string", "base64 或 path，默认 base64")}}, .Required = New List(Of String) From {"app"}}))
            Return t
        End Function

        ''' <summary>
        ''' 工具名 → 中文关键词映射（给 registry.Match() 用，决定 AI 能不能自动命中）
        ''' </summary>
        Private Shared ReadOnly _kwMap As New Dictionary(Of String, String())(StringComparer.OrdinalIgnoreCase) From {
            {"computer_screen_capture",       {"截屏","截图","屏幕截图","拍照","截屏一下","截个屏"}},
            {"computer_screen_get_info",      {"屏幕信息","分辨率","分辨率是多少"}},
            {"computer_screen_ocr",           {"截图识别","OCR识别","识别屏幕文字","截屏识别","屏幕文字识别"}},
            {"computer_describe",             {"描述屏幕","屏幕内容","看看屏幕","屏幕是什么","AI看屏"}},
            {"computer_mouse_move",           {"移动鼠标","鼠标移动","鼠标移到"}},
            {"computer_mouse_click",          {"点击鼠标","鼠标点击","单击","双击","点一下"}},
            {"computer_mouse_drag",           {"拖动鼠标","鼠标拖动","拖拽"}},
            {"computer_mouse_scroll",         {"滚动鼠标","滚轮","鼠标滚轮","上下滚动"}},
            {"computer_key_press",            {"按键","按下键","按键盘"}},
            {"computer_key_hotkey",           {"组合键","快捷键","Ctrl","Alt","Shift","Win键"}},
            {"computer_type_text",            {"输入文字","打字","输入","粘贴"}},
            {"computer_locate_text",          {"定位文字","找到文字","屏幕上的文字在哪"}},
            {"computer_locate_color",         {"定位颜色","找到颜色"}},
            {"computer_click_text",           {"点击播放","点播放","播放按钮","播放一下","继续播放","按播放"}},
            {"computer_clipboard_read",       {"读剪贴板","剪贴板内容","看看剪贴板"}},
            {"computer_clipboard_write",      {"写剪贴板","复制到剪贴板","剪贴板写入"}},
            {"computer_open_application",     {"打开应用","启动程序","打开浏览器","打开Chrome","打开Excel","打开文件","打开文件夹"}},
            {"computer_execute_command",      {"执行命令","运行命令","cmd","命令行"}},
            {"computer_wait",                 {"等待","等一下","暂停","sleep"}},
            {"computer_set_safe_mode",        {"安全模式","safe mode"}},
            {"computer_mouse_position_get",   {"鼠标坐标","光标位置","鼠标在哪","当前鼠标","光标在哪"}},
            {"computer_window_list",          {"窗口列表","列出窗口","有哪些窗口","窗口标题"}},
            {"computer_process_list",         {"进程列表","列出进程","有哪些进程","正在运行什么"}},
            {"computer_process_kill",         {"结束进程","关闭进程","杀掉进程","终止进程","kill进程"}},
            {"computer_close_application",     {"退出应用","关闭应用","退出程序","关闭程序","退出QQ音乐","关闭QQ音乐","退出音乐","关闭音乐","退出浏览器","关闭浏览器"}},
            {"computer_play_media",           {"播放音乐","放音乐","播放一下","放首歌","放个歌","开始播放","继续播"}},
            {"computer_wait_text",            {"等文字出现","等待文字","等到出现","等屏幕出现","等待结果"}},
            {"computer_assert_text",          {"校验文字","检查是否出现","确认出现","验证显示","有没有出现","文字是否在"}},
            {"computer_focus_window",         {"切到窗口","切换到窗口","聚焦窗口","前置窗口","激活窗口","切到应用","切换到应用","切到X","切到某某"}},
            {"computer_capture_window",       {"截该窗口","截图窗口","截图该窗口","截应用窗口","截图应用","截某窗口图片"}}
        }

        Public Shared Sub RegisterAllTools()
            For Each tool In GetToolDefinitions()
                ' 查中文关键词（命中则用，未命中则只用工具名兜底）
                Dim kwList As New List(Of String) From {tool.Name}
                If _kwMap.ContainsKey(tool.Name) Then
                    kwList.AddRange(_kwMap(tool.Name))
                End If
                McpCapabilityRegistry.Instance.RegisterMcpTool(tool,
                    displayName:=tool.Name,
                    category:=CapabilityCategory.ComputerAgent,
                    keywords:=kwList,
                    isWriteOperation:=tool.Name.StartsWith("computer_mouse") OrElse tool.Name.StartsWith("computer_key") OrElse tool.Name = "computer_type_text" OrElse
                        tool.Name = "computer_open_application" OrElse tool.Name = "computer_process_kill" OrElse tool.Name = "computer_close_application" OrElse tool.Name = "computer_execute_command",
                    isDangerous:=(tool.Name = "computer_execute_command" OrElse tool.Name = "computer_process_kill" OrElse
                        tool.Name = "computer_close_application" OrElse tool.Name = "computer_open_application"),
                    source:="computer-local")
            Next
        End Sub

        Public Shared Function ExecuteTool(toolName As String, args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim mn As String = "Execute" & ToolSuffix(toolName)
            EnsureCache()
            Dim mi As MethodInfo = Nothing
            If _cache.TryGetValue(mn, mi) Then
                Try
                    Return DirectCast(mi.Invoke(Nothing, New Object() {If(args, New Dictionary(Of String, Object)())}), ToolsCallResult)
                Catch ex As TargetInvocationException
                    Return Err(If(ex.InnerException IsNot Nothing, ex.InnerException.Message, ex.Message))
                End Try
            End If
            Return Err("未找到执行方法: " & mn)
        End Function

        Private Shared Function SafeBlocked(op As String) As ToolsCallResult
            Return Err("安全模式已开启，已拦截" & op & "。")
        End Function

        Private Shared ReadOnly _protectedProcesses As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            "excel", "winword", "powerpnt", "outlook", "onenote", "msaccess",
            "explorer", "dwm", "csrss", "winlogon", "lsass", "services", "smss",
            "system", "idle", "svchost", "csrss", "wininit", "fontdrvhost",
            "cursor", "devenv", "wps", "et", "wpp", "wpspdf"
        }

        Private Shared ReadOnly _blockedLaunchNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            "cmd", "cmd.exe", "powershell", "powershell.exe", "pwsh", "pwsh.exe",
            "wscript", "wscript.exe", "cscript", "cscript.exe", "mshta", "mshta.exe",
            "rundll32", "rundll32.exe", "regsvr32", "regsvr32.exe",
            "bitsadmin", "bitsadmin.exe", "certutil", "certutil.exe",
            "msiexec", "msiexec.exe", "ftp", "ftp.exe", "curl", "curl.exe",
            "wget", "wget.exe", "bash", "bash.exe", "wsl", "wsl.exe",
            "命令行", "命令提示符"
        }

        Private Shared ReadOnly _safeCmdAllow As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            "hostname", "whoami", "ver", "date /t", "time /t"
        }

        Private Shared Function IsProtectedProcess(name As String) As Boolean
            If String.IsNullOrWhiteSpace(name) Then Return True
            Dim n = name.Trim().ToLowerInvariant()
            If n.EndsWith(".exe", StringComparison.Ordinal) Then n = n.Substring(0, n.Length - 4)
            Return _protectedProcesses.Contains(n)
        End Function

        Private Shared Function IsBlockedLaunchTarget(target As String) As Boolean
            If String.IsNullOrWhiteSpace(target) Then Return True
            Dim t = target.Trim()
            If t.StartsWith("http://", StringComparison.OrdinalIgnoreCase) OrElse
               t.StartsWith("https://", StringComparison.OrdinalIgnoreCase) Then Return False
            Dim fn = Path.GetFileName(t)
            If String.IsNullOrWhiteSpace(fn) Then fn = t
            Dim baseName = Path.GetFileNameWithoutExtension(fn)
            If _blockedLaunchNames.Contains(fn) OrElse _blockedLaunchNames.Contains(baseName) Then Return True
            Dim ext = Path.GetExtension(fn).ToLowerInvariant()
            If ext = ".bat" OrElse ext = ".cmd" OrElse ext = ".ps1" OrElse ext = ".vbs" OrElse
               ext = ".js" OrElse ext = ".jse" OrElse ext = ".wsf" OrElse ext = ".hta" OrElse
               ext = ".scr" OrElse ext = ".com" OrElse ext = ".msi" Then Return True
            Return False
        End Function

        Private Shared Function CommandLooksSafe(cmd As String) As Boolean
            If String.IsNullOrWhiteSpace(cmd) Then Return False
            Dim c = cmd.Trim()
            If c.IndexOfAny(New Char() {"|"c, "&"c, ";"c, "`"c, "$"c, "("c, ")"c, "<"c, ">"c, ControlChars.Lf, ControlChars.Cr}) >= 0 Then Return False
            If c.Contains("&&") OrElse c.Contains("||") Then Return False
            Return _safeCmdAllow.Contains(c)
        End Function

        Private Shared Function ToolSuffix(name As String) As String
            Dim s = name.Replace("computer_", "")
            Dim parts = s.Split(New Char() {"_"c}, StringSplitOptions.RemoveEmptyEntries)
            Dim sb As New System.Text.StringBuilder()
            For Each p In parts
                If p.Length > 0 Then
                    sb.Append(Char.ToUpperInvariant(p(0)))
                    If p.Length > 1 Then sb.Append(p.Substring(1))
                End If
            Next
            Return sb.ToString()
        End Function

        Private Shared Sub EnsureCache()
            If _initialized Then Return
            SyncLock _cache
                If _initialized Then Return
                Dim flags = BindingFlags.NonPublic Or BindingFlags.Static Or BindingFlags.DeclaredOnly
                Dim dictType = GetType(Dictionary(Of String, Object))
                For Each mi In GetType(ComputerAgentTools).GetMethods(flags)
                    If mi.ReturnType Is GetType(ToolsCallResult) AndAlso mi.GetParameters().Length = 1 AndAlso
                       mi.GetParameters()(0).ParameterType Is dictType Then
                        _cache(mi.Name) = mi
                    End If
                Next
                _initialized = True
            End SyncLock
        End Sub

        ' ── 共享辅助 ──

        Private Shared Function G(Of T)(a As Dictionary(Of String, Object), n As String, d As T) As T
            If a Is Nothing Then Return d
            Dim v As Object = Nothing
            If a.TryGetValue(n, v) AndAlso v IsNot Nothing Then
                Try
                    Return CType(v, T)
                Catch
                    Try
                        Return DirectCast(Convert.ChangeType(v, GetType(T)), T)
                    Catch
                    End Try
                End Try
            End If
            Return d
        End Function

        Private Shared Function OK(s As String) As ToolsCallResult
            Return New ToolsCallResult() With {.Content = New List(Of McpContentItem) From {New McpTextContent With {.Text = s}}}
        End Function

        Private Shared Function Err(s As String) As ToolsCallResult
            Return New ToolsCallResult() With {.IsError = True, .Content = New List(Of McpContentItem) From {New McpTextContent With {.Text = s}}}
        End Function

        Private Shared Function FullRect() As Rectangle
            Return Rectangle.FromLTRB(
                SystemInformation.VirtualScreen.Left, SystemInformation.VirtualScreen.Top,
                SystemInformation.VirtualScreen.Right, SystemInformation.VirtualScreen.Bottom)
        End Function

        Private Shared Sub CaptureBmp(bmp As Bitmap, rect As Rectangle)
            ' 改用 Graphics.CopyFromScreen 而非手动 GetDC/BitBlt：
            ' 在 VSTO 插件进程里 GetDC(0)+BitBlt 常拿不到可用屏幕 DC，会静默失败输出整张全白图（前端看似"秒出空白"）。
            ' CopyFromScreen 对交互式桌面会话更可靠，且失败会抛出 Win32Exception，由外层 catch 返回明确错误。
            Using g = Graphics.FromImage(bmp)
                g.CopyFromScreen(rect.Left, rect.Top, 0, 0, bmp.Size, CopyPixelOperation.SourceCopy)
            End Using
        End Sub

        Private Shared Function BmpToPngBase64(bmp As Bitmap) As String
            Using ms As New MemoryStream()
                bmp.Save(ms, ImageFormat.Png)
                Return Convert.ToBase64String(ms.ToArray())
            End Using
        End Function

        Private Shared Function SaveBmpToTempPng(bmp As Bitmap) As String
            Dim p = Path.Combine(Path.GetTempPath(), "ca_" & Guid.NewGuid().ToString("N") & ".png")
            bmp.Save(p, ImageFormat.Png)
            Return p
        End Function

        ' ── DllImport ──

        <DllImport("user32.dll")>
        Private Shared Function SetCursorPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
        End Function

        <DllImport("user32.dll")>
        Private Shared Sub mouse_event(ByVal f As UInteger, ByVal dx As UInteger, ByVal dy As UInteger, ByVal d As UInteger, ByVal ex As IntPtr)
        End Sub

        <DllImport("user32.dll")>
        Private Shared Sub keybd_event(ByVal b As Byte, ByVal s As Byte, ByVal f As UInteger, ByVal ex As IntPtr)
        End Sub

        <DllImport("gdi32.dll")>
        Private Shared Function BitBlt(ByVal dd As IntPtr, ByVal dx As Integer, ByVal dy As Integer, ByVal dw As Integer, ByVal dh As Integer, ByVal sd As IntPtr, ByVal sx As Integer, ByVal sy As Integer, ByVal r As UInteger) As Boolean
        End Function

        <DllImport("user32.dll")>
        Private Shared Function GetDC(ByVal h As IntPtr) As IntPtr
        End Function

        <DllImport("user32.dll")>
        Private Shared Function ReleaseDC(ByVal h As IntPtr, ByVal d As IntPtr) As Integer
        End Function

        <DllImport("user32.dll")>
        Private Shared Function GetCursorPos(ByRef lpPt As POINT) As Boolean
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
        Private Shared Function EnumWindows(ByVal lpEnumFunc As EnumWindowsProc, ByVal lParam As IntPtr) As Boolean
        End Function

        Private Delegate Function EnumWindowsProc(ByVal hWnd As IntPtr, ByVal lParam As IntPtr) As Boolean

        <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
        Private Shared Function GetWindowText(ByVal hWnd As IntPtr, ByVal lpString As System.Text.StringBuilder, ByVal nMaxCount As Integer) As Integer
        End Function

        <DllImport("user32.dll")>
        Private Shared Function IsWindowVisible(ByVal hWnd As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll")>
        Private Shared Function GetWindowThreadProcessId(ByVal hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
        End Function

        <DllImport("user32.dll")>
        Private Shared Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll")>
        Private Shared Function ShowWindow(ByVal hWnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
        End Function

        <DllImport("user32.dll")>
        Private Shared Function IsIconic(ByVal hWnd As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll")>
        Private Shared Function GetWindowRect(ByVal hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
        End Function

        Private Structure RECT
            Public Left As Integer
            Public Top As Integer
            Public Right As Integer
            Public Bottom As Integer
        End Structure

        ' 按应用取窗口矩形。QQ音乐等浮窗/迷你窗是有真实边界的顶层窗口，用 GetWindowRect 框取。
        Private Shared Function WindowRectOf(hwnd As IntPtr) As Rectangle
            Dim rc As RECT
            If GetWindowRect(hwnd, rc) Then
                Return Rectangle.FromLTRB(rc.Left, rc.Top, rc.Right, rc.Bottom)
            End If
            Return Rectangle.Empty
        End Function

        ' 按应用名/别名解析候选进程名（与关闭/播放同口径）：别名 exe 基名 + 原始词基名，UWP calc 特例一并补齐。
        Private Shared Function ProcessCandidates(raw As String, low As String) As List(Of String)
            Dim cand As New List(Of String)
            Dim aliases As String() = Nothing
            If _appAliases.TryGetValue(low, aliases) Then
                For Each al In aliases
                    Dim base = Path.GetFileNameWithoutExtension(al)
                    If base.Length > 0 Then
                        Dim b = base.ToLowerInvariant()
                        If cand.Contains(b) = False Then cand.Add(b)
                    End If
                Next
            End If
            Dim baseSelf = Path.GetFileNameWithoutExtension(low)
            If baseSelf.Length > 0 AndAlso baseSelf <> low Then
                Dim bb = baseSelf.ToLowerInvariant()
                If cand.Contains(bb) = False Then cand.Add(bb)
            End If
            If cand.Count = 0 Then cand.Add(low.Replace(".exe", ""))
            If cand.Contains("calc") Then cand.Add("calculatorapp")
            Return cand
        End Function

        ' 按应用解析主窗口（轮询至多约4s，覆盖刚打开/启动未就绪）；未运行返回 IntPtr.Zero。
        Private Shared Function FindMainWindowByApp(cand As List(Of String)) As IntPtr
            For i = 1 To 8
                For Each p In Process.GetProcesses()
                    Dim pn = p.ProcessName.ToLowerInvariant()
                    If cand.Contains(pn) AndAlso p.MainWindowHandle <> IntPtr.Zero Then
                        Return p.MainWindowHandle
                    End If
                Next
                Thread.Sleep(500)
            Next
            Return IntPtr.Zero
        End Function

        ' 切换窗口：按应用名/别名将已打开窗口置前（不打开新窗口），供"切换到X窗口"使用。
        Private Shared Function ExecuteFocusWindow(a As Dictionary(Of String, Object)) As ToolsCallResult
            Try
                Dim raw = G(a, "app", "")
                If String.IsNullOrWhiteSpace(raw) Then Return Err("需要指定 app（应用名/别名）")
                Dim low = raw.Trim().ToLowerInvariant()
                Dim cand = ProcessCandidates(raw, low)
                Dim hwnd = FindMainWindowByApp(cand)
                If hwnd = IntPtr.Zero Then Return OK("not_found：未检测到正在运行的 " & raw.Trim())
                If IsIconic(hwnd) Then ShowWindow(hwnd, 9) ' SW_RESTORE
                SetForegroundWindow(hwnd)
                Return OK("已切换前置 " & raw.Trim())
            Catch ex As Exception
                Return Err(ex.Message)
            End Try
        End Function

        ' 截图窗口：按应用名找主窗口，只裁该窗口区域返回 base64/path（非整屏）。
        Private Shared Function ExecuteCaptureWindow(a As Dictionary(Of String, Object)) As ToolsCallResult
            Try
                Dim raw = G(a, "app", "")
                If String.IsNullOrWhiteSpace(raw) Then Return Err("需要指定 app（应用名/别名）")
                Dim low = raw.Trim().ToLowerInvariant()
                Dim cand = ProcessCandidates(raw, low)
                Dim hwnd = FindMainWindowByApp(cand)
                If hwnd = IntPtr.Zero Then Return OK("not_found：未检测到正在运行的 " & raw.Trim())
                If IsIconic(hwnd) Then ShowWindow(hwnd, 9)
                Thread.Sleep(150)
                Dim rc = WindowRectOf(hwnd)
                If rc.IsEmpty OrElse rc.Width <= 0 OrElse rc.Height <= 0 Then Return Err("无法获取窗口区域: " & raw.Trim())
                Dim fmt = G(a, "returnFormat", "base64")
                Using bmp As New Bitmap(rc.Width, rc.Height, PixelFormat.Format24bppRgb)
                    CaptureBmp(bmp, rc)
                    If fmt = "path" Then
                        Return OK("path=" & SaveBmpToTempPng(bmp))
                    End If
                    Return OK("base64=" & BmpToPngBase64(bmp))
                End Using
            Catch ex As Exception
                Return Err(ex.Message)
            End Try
        End Function

        Private Structure POINT
            Public X As Integer
            Public Y As Integer
        End Structure

        Private Const SW_RESTORE As Integer = 9

        Private Const MLD As UInteger = 2
        Private Const MLU As UInteger = 4
        Private Const MRD As UInteger = 8
        Private Const MRU As UInteger = 16
        Private Const MMD As UInteger = 32
        Private Const MMU As UInteger = 64
        Private Const MW As UInteger = 2048
        Private Const KU As UInteger = 2
        Private Const SRCCOPY As UInteger = 133693760

        Private Shared ReadOnly _vkMap As Dictionary(Of String, Byte) = BuildVkMap()

        Private Shared Function BuildVkMap() As Dictionary(Of String, Byte)
            Dim d As New Dictionary(Of String, Byte)(StringComparer.OrdinalIgnoreCase)
            d("enter") = 13 : d("return") = 13 : d("tab") = 9 : d("escape") = 27 : d("esc") = 27
            d("space") = 32 : d("backspace") = 8 : d("delete") = 46 : d("del") = 46
            d("home") = 36 : d("end") = 35 : d("pageup") = 33 : d("pagedown") = 34
            d("arrowup") = 38 : d("up") = 38 : d("arrowdown") = 40 : d("down") = 40
            d("arrowleft") = 37 : d("left") = 37 : d("arrowright") = 39 : d("right") = 39
            d("f1") = 112 : d("f2") = 113 : d("f3") = 114 : d("f4") = 115
            d("f5") = 116 : d("f6") = 117 : d("f7") = 118 : d("f8") = 119
            d("f9") = 120 : d("f10") = 121 : d("f11") = 122 : d("f12") = 123
            d("ctrl") = 17 : d("control") = 17 : d("alt") = 18 : d("shift") = 16
            d("win") = 91 : d("windows") = 91
            ' 媒体键（系统级媒体控制；播放/暂停/下一首/上一首/停止，对 QQ音乐/网易云等生效）
            d("media_play_pause") = &HB3 : d("media_play") = &HB3 : d("media_pause") = &HB3 : d("playpause") = &HB3
            d("media_next") = &HB0 : d("media_prev") = &HB1 : d("media_stop") = &HB2
            Return d
        End Function

        ' ── Execute 实现 ──

        Private Shared Function ExecuteScreenCapture(a As Dictionary(Of String, Object)) As ToolsCallResult
            Try
                Dim x = G(a, "x", -1), y = G(a, "y", -1), w = G(a, "width", 0), h = G(a, "height", 0)
                Dim fmt = G(a, "returnFormat", "base64")
                Dim rect = If(x >= 0 AndAlso w > 0, New Rectangle(x, y, w, h), FullRect())
                System.Diagnostics.Debug.WriteLine("[ComputerAgent] ExecuteScreenCapture rect=" & rect.ToString() & " fmt=" & fmt)
                Using bmp As New Bitmap(rect.Width, rect.Height, PixelFormat.Format24bppRgb)
                    CaptureBmp(bmp, rect)
                    System.Diagnostics.Debug.WriteLine("[ComputerAgent] CaptureBmp OK, bmp.Size=" & bmp.Size.ToString())
                    If fmt = "path" Then
                        Dim p = SaveBmpToTempPng(bmp)
                        System.Diagnostics.Debug.WriteLine("[ComputerAgent] saved path=" & p)
                        Return OK("path=" & p)
                    Else
                        Dim b64 = BmpToPngBase64(bmp)
                        System.Diagnostics.Debug.WriteLine("[ComputerAgent] base64 len=" & b64.Length)
                        Return OK("base64=" & b64)
                    End If
                End Using
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[ComputerAgent] ExecuteScreenCapture FAILED: " & ex.GetType().FullName & ": " & ex.Message & vbCrLf & ex.StackTrace)
                Return Err("截图失败 [" & ex.GetType().Name & "]: " & ex.Message)
            End Try
        End Function

        Private Shared Function ExecuteScreenGetInfo(a As Dictionary(Of String, Object)) As ToolsCallResult
            Dim vs = SystemInformation.VirtualScreen
            Return OK("virtual_w=" & vs.Width & " virtual_h=" & vs.Height & " primary_w=" & System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width &
                      " primary_h=" & System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height & " monitors=" & System.Windows.Forms.Screen.AllScreens.Count().ToString())
        End Function

        Private Shared Function ExecuteScreenOcr(a As Dictionary(Of String, Object)) As ToolsCallResult
            Try
                Dim rect = BuildRegion(a, FullRect())
                Using bmp As New Bitmap(rect.Width, rect.Height, PixelFormat.Format24bppRgb)
                    CaptureBmp(bmp, rect)
                    Dim tempPath = SaveBmpToTempPng(bmp)
                    Try
                        Dim ocrMgr = OcrServiceManager.Instance
                        If Not ocrMgr.IsEngineAvailable(OcrEngineType.PaddleOcr) Then Return Err("OCR 引擎未就绪")
                        Dim text = ocrMgr.RecognizeText(tempPath).GetAwaiter().GetResult()
                        Return OK(text)
                    Finally
                        If File.Exists(tempPath) Then File.Delete(tempPath)
                    End Try
                End Using
            Catch ex As Exception
                Return Err(ex.Message)
            End Try
        End Function

        Private Shared Function ExecuteDescribe(a As Dictionary(Of String, Object)) As ToolsCallResult
            Try
                Dim rect = BuildRegion(a, FullRect())
                Using bmp As New Bitmap(rect.Width, rect.Height, PixelFormat.Format24bppRgb)
                    CaptureBmp(bmp, rect)
                    Dim base64 = BmpToPngBase64(bmp)
                    Dim prompt = G(a, "prompt", "描述屏幕内容")
                    Dim llm = MultimodalLLMEngine.SharedInstance
                    Dim resp = llm.ProcessImageAndTextAsync(base64, prompt).GetAwaiter().GetResult()
                    If resp Is Nothing OrElse resp.Content Is Nothing Then Return Err("LLM 返回空")
                    Return OK(resp.Content)
                End Using
            Catch ex As Exception
                Return Err(ex.Message)
            End Try
        End Function

        Private Shared Function ExecuteMouseMove(a As Dictionary(Of String, Object)) As ToolsCallResult
            SetCursorPos(G(a, "x", 0), G(a, "y", 0))
            Return OK("OK")
        End Function

        Private Shared Function ExecuteMouseClick(a As Dictionary(Of String, Object)) As ToolsCallResult
            Dim x = G(a, "x", -1), y = G(a, "y", -1)
            If x >= 0 Then SetCursorPos(x, y)
            Thread.Sleep(20)
            Dim d As UInteger = MLD, u As UInteger = MLU
            Dim btn = G(a, "button", "left").ToLowerInvariant()
            If btn = "right" Then d = MRD : u = MRU
            If btn = "middle" Then d = MMD : u = MMU
            Dim times = G(a, "clicks", 1)
            For i = 1 To times
                mouse_event(d, 0, 0, 0, IntPtr.Zero)
                Thread.Sleep(30)
                mouse_event(u, 0, 0, 0, IntPtr.Zero)
            Next
            Return OK("OK")
        End Function

        Private Shared Function ExecuteMouseDrag(a As Dictionary(Of String, Object)) As ToolsCallResult
            Dim fx = G(a, "fromX", 0), fy = G(a, "fromY", 0)
            Dim tx = G(a, "toX", 0), ty = G(a, "toY", 0)
            SetCursorPos(fx, fy)
            Thread.Sleep(20)
            mouse_event(MLD, 0, 0, 0, IntPtr.Zero)
            Thread.Sleep(50)
            For i = 1 To 10
                SetCursorPos(CInt(fx + (tx - fx) * i / 10), CInt(fy + (ty - fy) * i / 10))
                Thread.Sleep(30)
            Next
            mouse_event(MLU, 0, 0, 0, IntPtr.Zero)
            Return OK("OK")
        End Function

        Private Shared Function ExecuteMouseScroll(a As Dictionary(Of String, Object)) As ToolsCallResult
            ' amount>0 向上滚，amount<0 向下滚；滚轮 delta 为有符号值，直接按位转成 UInteger 传给 mouse_event
            Dim amount = G(a, "amount", 0)
            mouse_event(MW, 0, 0, CUInt(amount), IntPtr.Zero)
            Return OK("OK")
        End Function

        Private Shared Function ExecuteKeyPress(a As Dictionary(Of String, Object)) As ToolsCallResult
            Dim key = G(a, "key", "").Trim()
            Dim vk As Byte = 0
            If _vkMap.TryGetValue(key, vk) Then
                keybd_event(vk, 0, 0, IntPtr.Zero)
                Thread.Sleep(20)
                keybd_event(vk, 0, KU, IntPtr.Zero)
            ElseIf key.Length = 1 Then
                Dim ch = CByte(AscW(key(0)))
                keybd_event(ch, 0, 0, IntPtr.Zero)
                Thread.Sleep(20)
                keybd_event(ch, 0, KU, IntPtr.Zero)
            Else
                Return Err("未知按键: " & key)
            End If
            Return OK("OK")
        End Function

        Private Shared Function ExecuteKeyHotkey(a As Dictionary(Of String, Object)) As ToolsCallResult
            Dim parts = G(a, "keys", "").Split(New Char() {"+"c}, StringSplitOptions.RemoveEmptyEntries)
            Dim vks As New List(Of Byte)()
            For Each k In parts
                Dim kTrim = k.Trim()
                Dim vk As Byte = 0
                If _vkMap.TryGetValue(kTrim, vk) Then
                    vks.Add(vk)
                ElseIf kTrim.Length = 1 Then
                    vks.Add(CByte(AscW(kTrim(0))))
                Else
                    Return Err("未知按键: " & k)
                End If
            Next
            For Each vkItem In vks
                keybd_event(vkItem, 0, 0, IntPtr.Zero)
                Thread.Sleep(10)
            Next
            Thread.Sleep(50)
            For i = vks.Count - 1 To 0 Step -1
                keybd_event(vks(i), 0, KU, IntPtr.Zero)
                Thread.Sleep(10)
            Next
            Return OK("OK")
        End Function

        Private Shared Function ExecuteTypeText(a As Dictionary(Of String, Object)) As ToolsCallResult
            My.Computer.Clipboard.SetText(G(a, "text", ""))
            Thread.Sleep(50)
            keybd_event(CByte(17), 0, 0, IntPtr.Zero)
            Thread.Sleep(20)
            keybd_event(CByte(86), 0, 0, IntPtr.Zero)
            Thread.Sleep(20)
            keybd_event(CByte(86), 0, KU, IntPtr.Zero)
            keybd_event(CByte(17), 0, KU, IntPtr.Zero)
            Return OK("OK")
        End Function

        Private Shared Function ExecuteLocateText(a As Dictionary(Of String, Object)) As ToolsCallResult
            Try
                Dim target = G(a, "text", "")
                If String.IsNullOrWhiteSpace(target) Then Return OK("found=false")  ' 空目标防假阳性
                Dim rect = BuildRegion(a, FullRect())
                Using bmp As New Bitmap(rect.Width, rect.Height, PixelFormat.Format24bppRgb)
                    CaptureBmp(bmp, rect)
                    Dim tempPath = SaveBmpToTempPng(bmp)
                    Try
                        Dim ocrMgr = OcrServiceManager.Instance
                        If Not ocrMgr.IsEngineAvailable(OcrEngineType.PaddleOcr) Then Return Err("OCR 引擎未就绪")
                        Dim doc = ocrMgr.RecognizeTextWithPositionAsync(tempPath).GetAwaiter().GetResult()
                        If doc Is Nothing OrElse doc.Elements Is Nothing Then Return OK("found=false")
                        ' 优先返回精确一致的行（整行等于目标），其次接受目标出现在行内的子串匹配，
                        ' 避免"确定"这类短词命中止步按钮等无关行。
                        Dim substringHit As OcrElement = Nothing
                        For Each r As OcrElement In doc.Elements
                            If r.Text Is Nothing Then Continue For
                            If String.Equals(r.Text.Trim(), target, StringComparison.OrdinalIgnoreCase) Then
                                Return OK("found x=" & (rect.Left + CenterX(r)) & " y=" & (rect.Top + CenterY(r)))
                            End If
                            If substringHit Is Nothing AndAlso r.Text.IndexOf(target, StringComparison.OrdinalIgnoreCase) >= 0 Then
                                substringHit = r
                            End If
                        Next
                        If substringHit IsNot Nothing Then
                            Return OK("found x=" & (rect.Left + CenterX(substringHit)) & " y=" & (rect.Top + CenterY(substringHit)))
                        End If
                        Return OK("found=false")
                    Finally
                        If File.Exists(tempPath) Then File.Delete(tempPath)
                    End Try
                End Using
            Catch ex As Exception
                Return Err(ex.Message)
            End Try
        End Function

        Private Shared Function ExecuteClickText(a As Dictionary(Of String, Object)) As ToolsCallResult
            Try
                ' 候选词用英文逗号分隔，按优先级排列
                Dim raw = G(a, "texts", "")
                Dim candidates = raw.Split(New Char() {","c}, StringSplitOptions.RemoveEmptyEntries) _
                    .Select(Function(s) s.Trim()) _
                    .Where(Function(s) s.Length > 0).ToList()
                If candidates.Count = 0 Then Return OK("not_found=no_candidates")
                Dim rect = BuildRegion(a, FullRect())
                Using bmp As New Bitmap(rect.Width, rect.Height, PixelFormat.Format24bppRgb)
                    CaptureBmp(bmp, rect)
                    Dim tempPath = SaveBmpToTempPng(bmp)
                    Try
                        Dim ocrMgr = OcrServiceManager.Instance
                        If Not ocrMgr.IsEngineAvailable(OcrEngineType.PaddleOcr) Then Return Err("OCR 引擎未就绪")
                        Dim doc = ocrMgr.RecognizeTextWithPositionAsync(tempPath).GetAwaiter().GetResult()
                        If doc Is Nothing OrElse doc.Elements Is Nothing Then Return OK("not_found=no_ocr_result")
                        ' 对候选词按顺序匹配：精确一致 > 行内子串。优先返回第一个命中的候选词所在行。
                        Dim primaryRect As OcrBoundingBox = Nothing
                        Dim foundCandidate As String = Nothing
                        For Each cand In candidates
                            Dim substringHit As OcrElement = Nothing
                            Dim exactHit As OcrElement = Nothing
                            For Each r As OcrElement In doc.Elements
                                If r.Text Is Nothing Then Continue For
                                If String.Equals(r.Text.Trim(), cand, StringComparison.OrdinalIgnoreCase) Then
                                    If exactHit Is Nothing Then exactHit = r
                                ElseIf r.Text.IndexOf(cand, StringComparison.OrdinalIgnoreCase) >= 0 Then
                                    If substringHit Is Nothing Then substringHit = r
                                End If
                            Next
                            If exactHit IsNot Nothing Then
                                foundCandidate = cand : primaryRect = exactHit.BoundingBox
                                Exit For
                            End If
                            If substringHit IsNot Nothing Then
                                foundCandidate = cand : primaryRect = substringHit.BoundingBox
                                Exit For
                            End If
                        Next
                        If foundCandidate Is Nothing Then
                            Return OK("not_found=no_text_match")
                        End If
                        Dim cx = rect.Left + CInt((primaryRect.Left + primaryRect.Right) \ 2)
                        Dim cy = rect.Top + CInt((primaryRect.Top + primaryRect.Bottom) \ 2)
                        SetCursorPos(cx, cy) : Thread.Sleep(30)
                        mouse_event(MLD, 0, 0, 0, IntPtr.Zero) : Thread.Sleep(30)
                        mouse_event(MLU, 0, 0, 0, IntPtr.Zero)
                        Return OK("clicked target=""" & foundCandidate & """ x=" & cx & " y=" & cy)
                    Finally
                        If File.Exists(tempPath) Then File.Delete(tempPath)
                    End Try
                End Using
            Catch ex As Exception
                Return Err(ex.Message)
            End Try
        End Function

        Private Shared Function CenterX(r As OcrElement) As Integer
            If r.BoundingBox Is Nothing Then Return 0
            Return (r.BoundingBox.Left + r.BoundingBox.Right) \ 2
        End Function

        Private Shared Function CenterY(r As OcrElement) As Integer
            If r.BoundingBox Is Nothing Then Return 0
            Return (r.BoundingBox.Top + r.BoundingBox.Bottom) \ 2
        End Function

        Private Shared Function ExecuteLocateColor(a As Dictionary(Of String, Object)) As ToolsCallResult
            Try
                Dim hex = G(a, "color", "#FFFFFF")
                If Not hex.StartsWith("#"c) Then hex = "#" & hex
                Dim tol = G(a, "tolerance", 10)
                Dim target = ColorTranslator.FromHtml(hex)
                Dim rect = BuildRegion(a, FullRect())
                Using bmp As New Bitmap(rect.Width, rect.Height, PixelFormat.Format24bppRgb)
                    CaptureBmp(bmp, rect)
                    Dim tol2 = tol * tol * 3
                    For px = 0 To bmp.Width - 1 Step 2
                        For py = 0 To bmp.Height - 1 Step 2
                            Dim c = bmp.GetPixel(px, py)
                            Dim dr = c.R - target.R, dg = c.G - target.G, db = c.B - target.B
                            If dr * dr + dg * dg + db * db <= tol2 Then
                                Return OK("found x=" & (rect.Left + px) & " y=" & (rect.Top + py))
                            End If
                        Next
                    Next
                    Return OK("found=false")
                End Using
            Catch ex As Exception
                Return Err(ex.Message)
            End Try
        End Function

        Private Shared Function ExecuteClipboardRead(a As Dictionary(Of String, Object)) As ToolsCallResult
            Try
                Return OK(My.Computer.Clipboard.GetText())
            Catch ex As Exception
                Return Err(ex.Message)
            End Try
        End Function

        Private Shared Function ExecuteClipboardWrite(a As Dictionary(Of String, Object)) As ToolsCallResult
            Try
                My.Computer.Clipboard.SetText(G(a, "text", ""))
                Return OK("OK")
            Catch ex As Exception
                Return Err(ex.Message)
            End Try
        End Function

        ' ── 打开任意应用 ──
        ' 解析链（数据驱动 + 回退，避免"找不到 chrome.exe"这类裸 exe 直启失败）：
        '   别名候选（浏览器 chrome→edge 回退）→ 标准安装目录 → 开始菜单快捷方式 → PATH where.exe
        '   → 裸 exe 交给 ShellExecute 尝试；浏览器仍失败则唤起系统默认浏览器。
        Private Shared Function ExecuteOpenApplication(a As Dictionary(Of String, Object)) As ToolsCallResult
            Dim appName = G(a, "app", "")
            Dim explicitPath = G(a, "path", "")
            If _safe AndAlso Not String.IsNullOrWhiteSpace(explicitPath) Then
                Return Err("安全模式已开启，已拦截任意路径打开。请使用应用名（如 Excel、Chrome），不要传绝对路径。")
            End If
            If _safe AndAlso _blockedLaunchNames.Contains(If(appName, "").Trim()) Then
                Return SafeBlocked("打开命令行/脚本宿主")
            End If

            Dim target = explicitPath
            If String.IsNullOrWhiteSpace(target) Then
                target = ResolveAppTarget(appName)
            End If
            If IsBlockedLaunchTarget(target) Then
                Return SafeBlocked("打开命令行/脚本宿主或脚本文件")
            End If
            Dim isBrowser = IsBrowserIntent(appName)

            ' 浏览器且未解析到可执行文件 → 唤起系统默认浏览器
            If String.IsNullOrWhiteSpace(target) AndAlso isBrowser Then
                target = DefaultBrowserUrl
            End If
            If String.IsNullOrWhiteSpace(target) Then
                Return Err("未指定要打开的应用（缺 path 或 app），且无法按名称解析到已安装程序")
            End If

            Dim psi As New ProcessStartInfo()
            psi.FileName = target
            psi.Arguments = G(a, "arguments", "")
            psi.UseShellExecute = True
            Try
                Dim pp = Process.Start(psi)
                Return OK("pid=" & If(pp IsNot Nothing, pp.Id, 0).ToString() & " target=" & target)
            Catch ex As Exception
                ' 浏览器解析/启动失败 → 仍然唤起系统默认浏览器，满足"打开浏览器"
                If isBrowser AndAlso Not target.StartsWith("http") Then
                    Try
                        Dim pp = Process.Start(DefaultBrowserUrl)
                        Return OK("使用系统默认浏览器打开（pid=" & If(pp IsNot Nothing, pp.Id, 0).ToString() & "）")
                    Catch
                    End Try
                End If
                Return Err("打开应用失败 [" & target & "]: " & ex.Message & "（可用 path 指定完整绝对路径重试）")
            End Try
        End Function

        Private Const DefaultBrowserUrl As String = "https://www.google.com"

        Private Shared Function IsBrowserIntent(appName As String) As Boolean
            If String.IsNullOrWhiteSpace(appName) Then Return False
            Dim low = appName.Trim().ToLowerInvariant()
            Return low = "浏览器" OrElse low = "browser" OrElse _
                   low.Contains("chrome") OrElse low.Contains("edge") OrElse _
                   low.Contains("火狐") OrElse low.Contains("firefox") OrElse low.Contains("谷歌")
        End Function

        Private Shared Function ResolveAppTarget(appName As String) As String
            If String.IsNullOrWhiteSpace(appName) Then Return ""
            Dim name = appName.Trim()
            If name.Length = 0 Then Return ""

            Dim low = name.ToLowerInvariant()
            Dim candidates As New List(Of String)
            Dim aliases As String() = Nothing
            If _appAliases.TryGetValue(low, aliases) Then
                candidates.AddRange(aliases)
            End If
            If candidates.Count = 0 Then candidates.Add(name)

            For Each cand In candidates
                Dim t = LocateLaunchTarget(cand, low)
                If t <> "" Then Return t
            Next
            Return ""
        End Function

        ' 把可执行名解析成可启动的绝对目标；找不到返回 ""
        Private Shared Function LocateLaunchTarget(exeName As String, lowName As String) As String
            ' ① 已含路径 → 校验存在
            If exeName.Contains("\") OrElse exeName.Contains("/") Then
                If File.Exists(exeName) Then Return exeName
            End If
            ' ② 标准安装目录（参考数据；键为不含扩展名的进程名）
            Dim baseName = Path.GetFileNameWithoutExtension(exeName).ToLowerInvariant()
            Dim known As List(Of String) = Nothing
            If _exeDirs.TryGetValue(baseName, known) Then
                For Each d In known
                    Dim probe = Path.Combine(d, exeName)
                    If File.Exists(probe) Then Return probe
                Next
            End If
            ' ③ 开始菜单快捷方式（按 exe 基名 / 用户原词 匹配；.lnk / .url 均可由 Process.Start 执行）
            Dim lnk = FindStartMenuShortcut(exeName, lowName)
            If lnk <> "" Then Return lnk
            ' ④ PATH where.exe
            If Not exeName.Contains(" ") AndAlso Not exeName.Contains("\") Then
                Try
                    Dim pi As New ProcessStartInfo("where.exe", exeName)
                    pi.UseShellExecute = False
                    pi.RedirectStandardOutput = True
                    pi.CreateNoWindow = True
                    Using p = Process.Start(pi)
                        Dim line = p.StandardOutput.ReadLine()
                        p.WaitForExit(3000)
                        If Not String.IsNullOrWhiteSpace(line) Then
                            Dim trimmed = line.Trim()
                            If File.Exists(trimmed) Then Return trimmed
                        End If
                    End Using
                Catch
                End Try
            End If
            ' ⑤ 兜底：裸 exe 名交给 ShellExecute（系统 App Paths / System32 可能能解析，如 calc/notepad）
            If Not exeName.Contains("\") Then Return exeName
            Return ""
        End Function

        Private Shared Function FindStartMenuShortcut(exeName As String, lowName As String) As String
            Dim menus() As String = {
                Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu),
                Environment.GetFolderPath(Environment.SpecialFolder.StartMenu)
            }
            Dim targetName = Path.GetFileNameWithoutExtension(exeName).ToLowerInvariant()
            For Each menu In menus
                If String.IsNullOrEmpty(menu) OrElse Not Directory.Exists(menu) Then Continue For
                Dim progs = Path.Combine(menu, "Programs")
                If Not Directory.Exists(progs) Then Continue For
                Try
                    For Each f In Directory.EnumerateFiles(progs, "*.*", SearchOption.AllDirectories)
                        Dim ext = Path.GetExtension(f).ToLowerInvariant()
                        If ext <> ".lnk" AndAlso ext <> ".url" Then Continue For
                        Dim fname = Path.GetFileNameWithoutExtension(f).ToLowerInvariant()
                        If fname = targetName OrElse fname = lowName Then Return f
                        If lowName.Length >= 2 AndAlso fname.Contains(lowName) Then Return f
                    Next
                Catch
                End Try
            Next
            Return ""
        End Function

        Private Shared ReadOnly _exeDirs As Dictionary(Of String, List(Of String)) = BuildExeDirs()

        Private Shared Function BuildExeDirs() As Dictionary(Of String, List(Of String))
            Dim local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
            Dim pf = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
            Dim pf86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
            Dim d As New Dictionary(Of String, List(Of String))(StringComparer.OrdinalIgnoreCase)
            d("chrome") = New List(Of String) From {
                Path.Combine(local, "Google\Chrome\Application"),
                Path.Combine(pf, "Google\Chrome\Application"),
                Path.Combine(pf86, "Google\Chrome\Application")
            }
            d("msedge") = New List(Of String) From {
                Path.Combine(pf86, "Microsoft\Edge\Application"),
                Path.Combine(pf, "Microsoft\Edge\Application")
            }
            d("firefox") = New List(Of String) From {
                Path.Combine(pf, "Mozilla Firefox"),
                Path.Combine(pf86, "Mozilla Firefox")
            }
            d("qqmusic") = New List(Of String) From {
                Path.Combine(local, "Tencent\QQMusic"),
                Path.Combine(pf, "Tencent\QQMusic"),
                Path.Combine(pf86, "Tencent\QQMusic")
            }
            Return d
        End Function

        Private Shared ReadOnly _appAliases As Dictionary(Of String, String()) = BuildAppAliases()

        Private Shared Function BuildAppAliases() As Dictionary(Of String, String())
            ' ⚠ 硬编码约定（勿破坏）：此表只用来"消解歧义默认值"，不是"打开应用"的必经机制。
            '   新增软件一律靠 ResolveAppTarget 的开始菜单快捷方式扫描 / where(PATH) 动态解析，
            '   【不要】为每个新软件往这里加条目——否则别名表会无限膨胀且无法覆盖未知应用。
            '   仅当某个口语词存在歧义、系统无法唯一推断时才填默认映射（音乐→QQ音乐、表格→Excel 等）。
            ' 值为有序候选；浏览器 chrome→edge 回退，全部缺失时 ExecuteOpenApplication 走系统默认浏览器兜底
            Dim d As New Dictionary(Of String, String())(StringComparer.OrdinalIgnoreCase)
            d("浏览器") = {"chrome.exe", "msedge.exe"}
            d("browser") = {"chrome.exe", "msedge.exe"}
            d("chrome") = {"chrome.exe", "msedge.exe"}
            d("chrom") = {"chrome.exe", "msedge.exe"}
            d("谷歌浏览器") = {"chrome.exe", "msedge.exe"}
            d("googlechrome") = {"chrome.exe", "msedge.exe"}
            d("edge") = {"msedge.exe"}
            d("msedge") = {"msedge.exe"}
            d("edge浏览器") = {"msedge.exe"}
            d("火狐") = {"firefox.exe"}
            d("firefox") = {"firefox.exe"}
            d("记事本") = {"notepad.exe"}
            d("notepad") = {"notepad.exe"}
            d("计算器") = {"calc.exe"}
            d("calculator") = {"calc.exe"}
            d("画图") = {"mspaint.exe"}
            d("mspaint") = {"mspaint.exe"}
            d("cmd") = {"cmd.exe"}
            d("命令行") = {"cmd.exe"}
            d("命令提示符") = {"cmd.exe"}
            d("控制面板") = {"control.exe"}
            d("资源管理器") = {"explorer.exe"}
            d("explorer") = {"explorer.exe"}
            d("微信") = {"WeChat.exe"}
            d("wechat") = {"WeChat.exe"}
            d("weixin") = {"WeChat.exe"}
            d("qq音乐") = {"QQMusic.exe"}
            d("qqmusic") = {"QQMusic.exe"}
            d("音乐") = {"QQMusic.exe"}
            d("excel") = {"excel.exe"}
            d("excel表格") = {"excel.exe"}
            d("表格") = {"excel.exe"}
            d("word") = {"winword.exe"}
            d("word文档") = {"winword.exe"}
            d("ppt") = {"powerpnt.exe"}
            d("powerpoint") = {"powerpnt.exe"}
            d("幻灯片") = {"powerpnt.exe"}
            Return d
        End Function

        Private Shared Function ExecuteExecuteCommand(a As Dictionary(Of String, Object)) As ToolsCallResult
            If _safe Then Return SafeBlocked("命令执行")
            Dim cmd = G(a, "command", "")
            If Not CommandLooksSafe(cmd) Then
                Return Err("已拒绝任意命令。关闭安全模式后也仅允许 hostname / whoami / ver / date /t / time /t，禁止管道、重定向与复合命令。")
            End If
            Try
                Dim psi As New ProcessStartInfo("cmd.exe", "/c " & cmd)
                psi.UseShellExecute = False
                psi.RedirectStandardOutput = True
                psi.RedirectStandardError = True
                psi.CreateNoWindow = True
                Dim p = Process.Start(psi)
                If p Is Nothing Then Return Err("未能启动命令")
                If Not p.WaitForExit(8000) Then
                    Try : p.Kill() : Catch : End Try
                    Return Err("命令超时已终止")
                End If
                Dim so = p.StandardOutput.ReadToEnd()
                Dim se = p.StandardError.ReadToEnd()
                If so.Length > 4000 Then so = so.Substring(0, 4000) & "…"
                If se.Length > 1000 Then se = se.Substring(0, 1000) & "…"
                Return OK("code=" & p.ExitCode.ToString() & " out=" & so & " err=" & se)
            Catch ex As Exception
                Return Err(ex.Message)
            End Try
        End Function

        Private Shared Function ExecuteWait(a As Dictionary(Of String, Object)) As ToolsCallResult
            Dim ms = G(a, "ms", 1000)
            If ms < 0 Then ms = 0
            If ms > 5000 Then ms = 5000
            Thread.Sleep(ms)
            Return OK("waited=" & ms.ToString())
        End Function

        Private Shared Function ExecuteSetSafeMode(a As Dictionary(Of String, Object)) As ToolsCallResult
            If Not G(a, "enabled", True) Then
                Return Err("不能通过工具关闭安全模式。")
            End If
            _safe = True
            Return OK("safeMode=True")
        End Function

        ' ══════════════════ 定位与感知增强 ══════════════════

        Private Shared Function ExecuteMousePositionGet(a As Dictionary(Of String, Object)) As ToolsCallResult
            Try
                Dim pt As POINT
                If GetCursorPos(pt) Then
                    Return OK("x=" & pt.X & " y=" & pt.Y)
                End If
                Return Err("获取鼠标位置失败")
            Catch ex As Exception
                Return Err(ex.Message)
            End Try
        End Function

        Private Shared Function ExecuteWindowList(a As Dictionary(Of String, Object)) As ToolsCallResult
            Try
                Dim keyword = G(a, "keyword", "")
                Dim names As New List(Of String)()
                EnumWindows(Function(h As IntPtr, l As IntPtr) As Boolean
                                If Not IsWindowVisible(h) Then Return True
                                Dim sb As New System.Text.StringBuilder(512)
                                GetWindowText(h, sb, sb.Capacity)
                                Dim title = sb.ToString()
                                Dim pid = 0
                                GetWindowThreadProcessId(h, pid)
                                If title.Length > 0 Then
                                    If keyword.Length = 0 OrElse title.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0 Then
                                        names.Add(title & " [pid=" & pid & "]")
                                    End If
                                End If
                                Return True
                            End Function, IntPtr.Zero)
                Return OK(If(names.Count = 0, "no windows", String.Join(vbCrLf, names)))
            Catch ex As Exception
                Return Err(ex.Message)
            End Try
        End Function

        Private Shared Function ExecuteProcessList(a As Dictionary(Of String, Object)) As ToolsCallResult
            Try
                Dim nameF = G(a, "name", "").ToLowerInvariant()
                Dim rows As New List(Of String)()
                For Each p In Process.GetProcesses()
                    Dim pn = If(p.ProcessName, "")
                    If nameF.Length = 0 OrElse pn.ToLowerInvariant().Contains(nameF) Then
                        rows.Add(pn & " [pid=" & p.Id.ToString() & "]")
                    End If
                Next
                Return OK(If(rows.Count = 0, "no process", String.Join(vbCrLf, rows)))
            Catch ex As Exception
                Return Err(ex.Message)
            End Try
        End Function

        Private Shared Function ExecuteProcessKill(a As Dictionary(Of String, Object)) As ToolsCallResult
            If _safe Then Return SafeBlocked("进程结束")
            Try
                Dim nameF = G(a, "name", "").Trim().ToLowerInvariant()
                If nameF.EndsWith(".exe") Then nameF = nameF.Substring(0, nameF.Length - 4)
                If String.IsNullOrWhiteSpace(nameF) Then Return Err("缺少 name 参数")
                If IsProtectedProcess(nameF) Then
                    Return Err("已拒绝结束受保护进程: " & nameF)
                End If
                Dim killed As Integer = 0
                For Each p In Process.GetProcesses()
                    If p.ProcessName.ToLowerInvariant() = nameF Then
                        p.Kill()
                        killed += 1
                    End If
                Next
                Return OK(If(killed > 0, "killed " & killed.ToString() & " processes", "not_found"))
            Catch ex As Exception
                Return Err(ex.Message)
            End Try
        End Function

        ' 按应用名（口语/别名）关闭：复用 _appAliases 解析成可执行基名，匹配运行进程后强制结束，
        ' 确保 QQ音乐 等托盘应用"真正退出"（CloseMainWindow 往往只缩回托盘）。
        Private Shared Function ExecuteCloseApplication(a As Dictionary(Of String, Object)) As ToolsCallResult
            If _safe Then Return SafeBlocked("关闭应用")
            Try
                Dim raw = G(a, "app", "")
                If String.IsNullOrWhiteSpace(raw) Then Return Err("需要指定应用名（app）")
                Dim low = raw.Trim().ToLowerInvariant()
                Dim cand As New List(Of String)
                Dim aliases As String() = Nothing
                If _appAliases.TryGetValue(low, aliases) Then
                    For Each al In aliases
                        cand.Add(Path.GetFileNameWithoutExtension(al).ToLowerInvariant())
                    Next
                End If
                Dim baseSelf = Path.GetFileNameWithoutExtension(low)
                If baseSelf.Length > 0 AndAlso baseSelf <> low AndAlso cand.Contains(baseSelf) = False Then
                    cand.Add(baseSelf)
                End If
                If cand.Count = 0 Then cand.Add(low.Replace(".exe", ""))
                ' UWP 内置应用"启动器名 ≠ 真实进程名"特例补丁：calc.exe 只是 Windows 计算器的 UWP 启动器，
                ' 拉起后真实 UI 进程为 CalculatorApp.exe（由 ApplicationFrameHost 承载），calc 进程会退出。
                ' 属已知内置应用特例，勿据此为任意软件硬编码——通用软件仍走开始菜单/PATH 动态解析。
                If cand.Contains("calc") Then cand.Add("calculatorapp")
                Dim killedNames As New List(Of String)
                Dim foundNames As New List(Of String)
                For Each p In Process.GetProcesses()
                    Dim pn = p.ProcessName.ToLowerInvariant()
                    If cand.Contains(pn) Then
                        If IsProtectedProcess(pn) Then Continue For
                        foundNames.Add(pn)
                        Try
                            p.Kill()
                            killedNames.Add(pn)
                        Catch
                            ' 权限等导致 Kill 失败：不吞错，用 foundNames 与返回文案如实说明
                        End Try
                    End If
                Next
                If killedNames.Count > 0 Then
                    Return OK("已退出 " & raw.Trim() & "（关闭 " & killedNames.Count.ToString() & " 个进程）")
                End If
                If foundNames.Count > 0 Then
                    Return OK("检测到正在运行的 " & raw.Trim() & "，但无法结束（可能权限受限）：" & String.Join("、", foundNames.Distinct()))
                End If
                Return OK("not_found：未检测到正在运行的 " & raw.Trim())
            Catch ex As Exception
                Return Err(ex.Message)
            End Try
        End Function

        ' 播放：聚焦应用窗口 → OCR 判断是否已在播 → 空格切换 → OCR 复验。
        ' 刻意不用"坐标点按钮"：OCR 坐标点 QQ音乐等应用的播放控件并不可靠（可能点错/点到单曲列表），
        ' 空格键是所有桌面播放器通用的播放切换；用"界面变播放态"（暂停/正在播放）做实据验证，未验证即如实返回。
        Private Shared Function ExecutePlayMedia(a As Dictionary(Of String, Object)) As ToolsCallResult
            Try
                Dim raw = G(a, "app", "")
                If String.IsNullOrWhiteSpace(raw) Then Return Err("需要指定 app（应用名/别名）")
                Dim low = raw.Trim().ToLowerInvariant()
                Dim cand As New List(Of String)
                Dim aliases As String() = Nothing
                If _appAliases.TryGetValue(low, aliases) Then
                    For Each al In aliases
                        cand.Add(Path.GetFileNameWithoutExtension(al).ToLowerInvariant())
                    Next
                End If
                Dim baseSelf = Path.GetFileNameWithoutExtension(low)
                If baseSelf.Length > 0 AndAlso baseSelf <> low AndAlso cand.Contains(baseSelf) = False Then
                    cand.Add(baseSelf.ToLowerInvariant())
                End If
                If cand.Count = 0 Then cand.Add(low.Replace(".exe", ""))
                Dim hwnd As IntPtr = IntPtr.Zero
                For i = 1 To 8
                    Dim pn As String = ""
                    For Each p In Process.GetProcesses()
                        pn = p.ProcessName.ToLowerInvariant()
                        If cand.Contains(pn) AndAlso p.MainWindowHandle <> IntPtr.Zero Then
                            hwnd = p.MainWindowHandle
                            Exit For
                        End If
                    Next
                    If hwnd <> IntPtr.Zero Then Exit For
                    Thread.Sleep(500)
                Next
                If hwnd = IntPtr.Zero Then Return Err("未找到正在运行的应用主窗口: " & raw.Trim())
                If IsIconic(hwnd) Then ShowWindow(hwnd, 9) ' SW_RESTORE
                SetForegroundWindow(hwnd)
                Thread.Sleep(1200)
                Dim startText = ScreenOcrText()
                If startText.Length > 0 AndAlso LooksPlaying(startText) Then
                    Return OK("已正在播放，无需触发")
                End If
                SendVk(32) ' VK_SPACE
                Thread.Sleep(900)
                Dim verifyText = ScreenOcrText()
                If LooksPlaying(verifyText) Then
                    Return OK("已触发播放并验证正在播放（空格键）")
                End If
                ' 空格已发送成功（触发即视为已执行）；若界面播放键是纯图形按钮，OCR 读不出"暂停/正在播放"
                ' 文字 → 无法凭界面文字确认 → 如实转中性成功，绝不误报失败（假阴性）。
                Return OK("已触发播放（界面为图形按钮，无法凭OCR文字确认播放态，请以听感为准）")
            Catch ex As Exception
                Return Err(ex.Message)
            End Try
        End Function

        Private Shared Function ScreenOcrText() As String
            Try
                Dim rect = FullRect()
                Using bmp As New Bitmap(rect.Width, rect.Height, PixelFormat.Format24bppRgb)
                    CaptureBmp(bmp, rect)
                    Dim tempPath = SaveBmpToTempPng(bmp)
                    Try
                        Dim ocrMgr = OcrServiceManager.Instance
                        If Not ocrMgr.IsEngineAvailable(OcrEngineType.PaddleOcr) Then Return ""
                        Return ocrMgr.RecognizeText(tempPath).GetAwaiter().GetResult()
                    Finally
                        If File.Exists(tempPath) Then File.Delete(tempPath)
                    End Try
                End Using
            Catch ex As Exception
                ' 原先裸 Catch 静默吞掉 OCR 异常，调用方拿到空串会误判"无文字"。
                System.Diagnostics.Debug.WriteLine($"[ComputerAgent] 截图 OCR 失败: {ex.GetType().FullName}: {ex.Message}")
                Return ""
            End Try
        End Function

        Private Shared Function LooksPlaying(s As String) As Boolean
            Return s.Contains("暂停") OrElse s.Contains("正在播放") OrElse s.Contains("播放中") OrElse s.Contains("Pause") OrElse s.Contains("PAUSE")
        End Function

        Private Shared Sub SendVk(vk As Byte)
            keybd_event(vk, 0, 0, IntPtr.Zero)
            Thread.Sleep(20)
            keybd_event(vk, 0, KU, IntPtr.Zero)
        End Sub

        ' ── 区域截图辅助：优先全屏，或给定区域 ──
        Private Shared Function BuildRegion(a As Dictionary(Of String, Object), defaultRect As Rectangle) As Rectangle
            Dim x = G(a, "region_x", -999999)
            If x = -999999 Then x = G(a, "x", -999999)
            If x = -999999 Then Return defaultRect
            Dim y = G(a, "region_y", -999999)
            If y = -999999 Then y = G(a, "y", -999999)
            Dim w = G(a, "region_w", 0)
            If w = 0 Then w = G(a, "width", 0)
            Dim h = G(a, "region_h", 0)
            If h = 0 Then h = G(a, "height", 0)
            If x < 0 OrElse w <= 0 Then Return defaultRect
            Return New Rectangle(x, y, w, h)
        End Function

        Private Shared Function OcrFindText(rect As Rectangle, target As String) As String
            ' 返回 "found x=.. y=.." 或 "found=false"
            ' 空目标必须显式失败，避免 IndexOf("") 恒命中首个元素造成假阳性
            If String.IsNullOrWhiteSpace(target) Then Return "found=false"
            Using bmp As New Bitmap(rect.Width, rect.Height, PixelFormat.Format24bppRgb)
                CaptureBmp(bmp, rect)
                Dim tempPath = SaveBmpToTempPng(bmp)
                Try
                    Dim ocrMgr = OcrServiceManager.Instance
                    If Not ocrMgr.IsEngineAvailable(OcrEngineType.PaddleOcr) Then Return "found=false"
                    Dim doc = ocrMgr.RecognizeTextWithPositionAsync(tempPath).GetAwaiter().GetResult()
                    If doc Is Nothing OrElse doc.Elements Is Nothing Then Return "found=false"
                    For Each r As OcrElement In doc.Elements
                        If r.Text IsNot Nothing AndAlso r.Text.IndexOf(target, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            Dim cx = 0, cy = 0
                            If r.BoundingBox IsNot Nothing Then cx = (r.BoundingBox.Left + r.BoundingBox.Right) \ 2 : cy = (r.BoundingBox.Top + r.BoundingBox.Bottom) \ 2
                            Return "found x=" & (rect.Left + cx) & " y=" & (rect.Top + cy)
                        End If
                    Next
                    Return "found=false"
                Finally
                    If File.Exists(tempPath) Then File.Delete(tempPath)
                End Try
            End Using
        End Function

        Private Shared Function ExecuteWaitText(a As Dictionary(Of String, Object)) As ToolsCallResult
            Try
                Dim target = G(a, "text", "")
                Dim timeout = G(a, "timeout_ms", 10000)
                If timeout < 0 Then timeout = 0
                If timeout > 20000 Then timeout = 20000
                Dim rect = BuildRegion(a, FullRect())
                Dim sw = System.Diagnostics.Stopwatch.StartNew()
                Do While sw.ElapsedMilliseconds < timeout
                    Dim res = OcrFindText(rect, target)
                    If res.StartsWith("found ") Then Return OK(res)
                    Thread.Sleep(300)
                Loop
                Return OK("timeout")
            Catch ex As Exception
                Return Err(ex.Message)
            End Try
        End Function

        Private Shared Function ExecuteAssertText(a As Dictionary(Of String, Object)) As ToolsCallResult
            Try
                Dim target = G(a, "text", "")
                Dim expect = G(a, "expect", "present").ToLowerInvariant()
                Dim rect = BuildRegion(a, FullRect())
                Dim res = OcrFindText(rect, target)
                Dim foundNow = res.StartsWith("found ")
                If expect = "absent" Then
                    Return OK(If(foundNow, "assert_failed text_present", "assert_ok (text absent) " & res))
                Else
                    If foundNow Then Return OK("assert_ok " & res)
                    Return OK("assert_failed text_not_found")
                End If
            Catch ex As Exception
                Return Err(ex.Message)
            End Try
        End Function

    End Class

End Namespace

