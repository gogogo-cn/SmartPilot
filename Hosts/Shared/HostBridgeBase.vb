' ============================================================================
' HostBridgeBase - 暴露给 WebView2 前端的桥对象基类（注册名固定为 excelBridge，
' 与 Excel 版一致，前端 200+ 处调用无需改动）
'
' 子类（WordBridge / PptBridge）只需重写宿主相关能力；
' Excel 专属能力在此给出安全兜底，避免前端调用到不存在的方法而抛异常。
' ============================================================================
Option Strict Off
Option Explicit On

Imports System
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Reflection
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

<ComVisible(True)>
<ClassInterface(ClassInterfaceType.AutoDual)>
Public Class HostBridgeBase

    Protected _kind As OfficeHostKind
    Protected _app As Object
    Protected _version As String
    Protected _mcp As HostMcpTools

    Public Sub New()
    End Sub

    Public Sub New(kind As OfficeHostKind, app As Object)
        _kind = kind
        _app = app
        _version = "1.0.0"
        _mcp = New HostMcpTools(kind, Me)
    End Sub

    ' ------------------------------------------------------------------
    ' 宿主信息
    ' ------------------------------------------------------------------
    <ComVisible(True)>
    Public Overridable Function GetHostInfo() As String
        Dim name = If(_kind = OfficeHostKind.Word, "word",
                   If(_kind = OfficeHostKind.PowerPoint, "powerpoint", "excel"))
        Return "{" &
            """host"":""" & name & """," &
            """" & "hostKind" & """:" & CInt(_kind) & "," &
            """version"":""" & Safe(_version) & """" &
            "}"
    End Function

    ' ------------------------------------------------------------------
    ' 配置（前端直连大模型，配置只需本地存取）
    ' ------------------------------------------------------------------
    <ComVisible(True)>
    Public Function GetAIConfig() As String
        Return HostConfig.GetAIConfigJson()
    End Function

    <ComVisible(True)>
    Public Function SaveAIConfig(configJson As String) As String
        Return HostConfig.SaveAIConfigJson(configJson)
    End Function

    ''' <summary>同步 WebView 界面语言到本机，并刷新功能区标签。</summary>
    <ComVisible(True)>
    Public Function SaveUiLocale(locale As String) As String
        Try
            HostRibbonLocale.SaveLocale(locale)
            HostRibbonLocale.Invalidate()
            Try
                Dim panes = HostPaneManager.Current
                If panes IsNot Nothing Then
                    panes.ExecuteScriptAll(HostRibbonLocale.ApplyLocaleScript())
                    panes.SetPaneTitle("ppttheme", HostRibbonLocale.PptThemePaneTitle())
                End If
            Catch
            End Try
            Return "{""success"":true,""locale"":""" & HostRibbonLocale.GetLocale() & """}"
        Catch ex As Exception
            Return "{""success"":false,""error"":""" & Safe(ex.Message) & """}"
        End Try
    End Function

    <ComVisible(True)>
    Public Function GetUiLocale() As String
        Try
            Return "{""success"":true,""locale"":""" & HostRibbonLocale.GetLocale() & """}"
        Catch ex As Exception
            Return "{""success"":false,""error"":""" & Safe(ex.Message) & """}"
        End Try
    End Function

    <ComVisible(True)>
    Public Function GetAvailableLLMProviders() As String
        Return "{""success"":true,""providers"":[""openai"",""deepseek"",""qwen"",""moonshot"",""zhipu"",""siliconflow"",""ollama"",""custom""]}"
    End Function

    <ComVisible(True)>
    Public Function GetAvailableOcrEngines() As String
        Return "{""success"":true,""engines"":[{""id"":""none"",""name"":""未启用""}]}"
    End Function

    <ComVisible(True)>
    Public Function GetOcrStatus() As String
        Try
            Return Ocr.GetOcrStatus()
        Catch ex As Exception
            Return "{""success"":true,""available"":false,""engine"":""none""}"
        End Try
    End Function

    <ComVisible(True)>
    Public Overridable Function GetOcrEngineStatus(engine As String) As String
        Try
            Return Ocr.GetOcrEngineStatus(engine)
        Catch ex As Exception
            Return "{""success"":false,""error"":""" & Safe(ex.Message) & """}"
        End Try
    End Function

    <ComVisible(True)>
    Public Function ForceCheckRapidOcrStatus() As String
        Return GetOcrStatus()
    End Function

    ' ------------------------------------------------------------------
    ' OCR 识别：统一委托公共核心库（SmartPilot.Core）的 OcrBridge 单例。
    ' 这样 PPT/Word 与 Excel 走的是同一套状态机与数据结构，行为完全一致
    ' （识别、格式化、结构化标注、原始坐标、解析脚本、公式 OCR 全部同源）。
    ' 例外：SaveBase64Image 由本基类实现（OcrBridge 未含该方法）。
    ' ------------------------------------------------------------------
    Private Shared ReadOnly Property Ocr As Global.ExcelAIVSTOAddIn.Bridges.OcrBridge
        Get
            Return Global.ExcelAIVSTOAddIn.Bridges.OcrBridge.Instance
        End Get
    End Property

    <ComVisible(True)>
    Public Overridable Function StartOcrRecognitionWithFormat(imagePath As String, recognitionMode As String, outputFormat As String) As String
        Try
            Return Ocr.StartOcrRecognitionWithFormat(imagePath, recognitionMode, outputFormat)
        Catch ex As Exception
            Return "{""started"":false,""error"":""" & Safe(ex.Message) & """}"
        End Try
    End Function

    <ComVisible(True)>
    Public Overridable Function GetOcrResultWithFormat(format As String) As String
        Try
            Return Ocr.GetOcrResultWithFormat(format)
        Catch ex As Exception
            Return "{""ready"":true,""success"":false,""error"":""" & Safe(ex.Message) & """}"
        End Try
    End Function

    <ComVisible(True)>
    Public Overridable Function GetOcrStructuredResult() As String
        Try
            Return Ocr.GetOcrStructuredResult()
        Catch ex As Exception
            Return "{""ready"":true,""success"":false,""error"":""" & Safe(ex.Message) & """}"
        End Try
    End Function

    <ComVisible(True)>
    Public Overridable Function GetOcrRawData() As String
        Try
            Return Ocr.GetOcrRawData()
        Catch ex As Exception
            Return "{""ready"":true,""success"":false,""error"":""" & Safe(ex.Message) & """}"
        End Try
    End Function

    <ComVisible(True)>
    Public Overridable Function CheckFormulaOcrStatus() As String
        Try
            Return Ocr.CheckFormulaOcrStatus()
        Catch ex As Exception
            Return "{""success"":false,""enabled"":false}"
        End Try
    End Function

    <ComVisible(True)>
    Public Overridable Function StartFormulaOcrInstall() As String
        Try
            Return Ocr.StartFormulaOcrInstall()
        Catch ex As Exception
            Return "{""success"":false,""error"":""" & Safe(ex.Message) & """}"
        End Try
    End Function

    <ComVisible(True)>
    Public Overridable Function GetFormulaInstallStatus() As String
        Try
            Return Ocr.GetFormulaInstallStatus()
        Catch ex As Exception
            Return "{""success"":true,""installed"":false,""running"":false}"
        End Try
    End Function

    <ComVisible(True)>
    Public Overridable Function StartScreenshotOcr() As String
        Return "{""success"":false,""error"":""当前宿主未启用截图 OCR""}"
    End Function

    <ComVisible(True)>
    Public Overridable Sub CloseOcrPanel()
    End Sub

    <ComVisible(True)>
    Public Overridable Sub OpenOcrPanel()
        Try
            Dim panes = HostPaneManager.Current
            If panes IsNot Nothing Then
                panes.Show("ocr", "ocr-panel.html", HostRibbonLocale.Text("Capture to sheet", "识图入表"))
            End If
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("[HostBridge] OpenOcrPanel: " & ex.Message)
        End Try
    End Sub

    <ComVisible(True)>
    Public Overridable Sub OpenSettingsPanel(tab As String)
    End Sub

    <ComVisible(True)>
    Public Overridable Sub OpenMobilePanel()
        Try
            Dim panes = HostPaneManager.Current
            If panes IsNot Nothing Then
                panes.Show("mobile", "mobile-control-panel.html", HostRibbonLocale.Text("Mobile Control", "手机遥控"))
            End If
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("[HostBridge] OpenMobilePanel: " & ex.Message)
        End Try
    End Sub

    <ComVisible(True)>
    Public Overridable Sub EnsureLocalBridge()
    End Sub

    <ComVisible(True)>
    Public Overridable Function StartMobileControl(Optional port As Integer = -1) As String
        Return "{""success"":false,""started"":false,""error"":""当前宿主没有手机遥控服务""}"
    End Function

    <ComVisible(True)>
    Public Overridable Sub OpenOcrSettingsPanel()
        Try
            Dim panes = HostPaneManager.Current
            If panes IsNot Nothing Then
                panes.Show("ocrsettings", "ocr-settings-panel.html", HostRibbonLocale.Text("OCR Model Settings", "OCR 模型设置"))
            End If
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("[HostBridge] OpenOcrSettingsPanel: " & ex.Message)
        End Try
    End Sub

    <ComVisible(True)>
    Public Overridable Function ClearLocalUserData() As String
        Return HostConfig.ClearStoredApiKey()
    End Function

    <ComVisible(True)>
    Public Overridable Function LicenseRegister(bodyJson As String) As String
        Return HostLicense.Register(bodyJson)
    End Function

    <ComVisible(True)>
    Public Overridable Function LicenseLogin(bodyJson As String) As String
        Return HostLicense.Login(bodyJson)
    End Function

    <ComVisible(True)>
    Public Overridable Function LicenseLogout() As String
        Return HostLicense.Logout()
    End Function

    <ComVisible(True)>
    Public Overridable Function LicenseGetEntitlement() As String
        Return HostLicense.GetEntitlement()
    End Function

    <ComVisible(True)>
    Public Overridable Function LicenseIsSubscriptionActive() As String
        Return HostLicense.IsSubscriptionActive()
    End Function

    <ComVisible(True)>
    Public Overridable Function LicenseStartCheckout(bodyJson As String) As String
        Return HostLicense.StartCheckout(bodyJson)
    End Function

    <ComVisible(True)>
    Public Overridable Function LicenseOpenBillingPortal() As String
        Return HostLicense.OpenBillingPortal()
    End Function

    <ComVisible(True)>
    Public Overridable Function LicenseSetApiBaseUrl(url As String) As String
        Return HostLicense.SetApiBaseUrl(url)
    End Function

    <ComVisible(True)>
    Public Overridable Function LicenseGetApiBaseUrl() As String
        Return HostLicense.GetApiBaseUrl()
    End Function

        <ComVisible(True)>
        Public Overridable Function LicensePingHealth() As String
            Return HostLicense.PingHealth()
        End Function

        ''' <summary>离线激活码激活（路线①：无授权服务器也能卖）。body: { key }。</summary>
        <ComVisible(True)>
        Public Overridable Function LicenseActivateKey(bodyJson As String) As String
            Return HostLicense.ActivateKey(bodyJson)
        End Function

        ''' <summary>本地离线激活状态（不打网络）。</summary>
        <ComVisible(True)>
        Public Overridable Function LicenseGetLocalActivation() As String
            Return HostLicense.GetLocalActivation()
        End Function

        <ComVisible(True)>
        Public Overridable Function LicenseClearLocalActivation() As String
            Return HostLicense.ClearLocalActivation()
        End Function

        ''' <summary>
        ''' 「注销并申请退款」：清除本机激活 + 该 key 在本机永久失效 + 返回消 key 码。
        ''' 用户把码发给客服，运营台验签通过才退款。
        ''' </summary>
        <ComVisible(True)>
        Public Overridable Function LicenseDeactivateForRefund() As String
            Return HostLicense.DeactivateForRefund()
        End Function

        ''' <summary>查询某张激活码是否已在本机注销过。body: { key }。</summary>
        <ComVisible(True)>
        Public Overridable Function LicenseIsKeyDeactivated(bodyJson As String) As String
            Return HostLicense.IsKeyDeactivated(bodyJson)
        End Function

    ''' <summary>打开 PPT 主题模板编辑器（注入 excelBridge 的任务窗格/停靠窗）</summary>
    <ComVisible(True)>
    Public Overridable Sub OpenPptThemeEditor()
        Try
            Dim panes = HostPaneManager.Current
            If panes IsNot Nothing Then
                panes.Show("ppttheme", "ppt-theme-editor.html", HostRibbonLocale.PptThemePaneTitle())
            End If
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("[HostBridge] OpenPptThemeEditor: " & ex.Message)
        End Try
    End Sub

    <ComVisible(True)>
    Public Overridable Sub ClosePptThemeEditor()
        Try
            HostPaneManager.Current?.Hide("ppttheme")
        Catch
        End Try
    End Sub

    ' ------------------------------------------------------------------
    ' PPT 自定义主题 CRUD（写宿主 Config\custom_ppt_themes.json）
    ' Excel 主桥走 PptGenerationService；PPT/Word 宿主无该服务，在此用 JSON 文件对齐。
    ' ------------------------------------------------------------------
    Private Shared Function CustomPptThemesPath() As String
        Return Path.Combine(HostPathHelper.GetHostConfigDirectory(), "custom_ppt_themes.json")
    End Function

    <ComVisible(True)>
    Public Overridable Function SaveCustomPptTheme(themeJson As String) As String
        Try
            If String.IsNullOrWhiteSpace(themeJson) Then
                Return "{""success"":false,""error"":""主题内容为空""}"
            End If
            Dim t As JObject = JObject.Parse(themeJson)
            Dim themeId As String = If(t.Value(Of String)("id"), "").Trim()
            If String.IsNullOrWhiteSpace(themeId) Then
                themeId = "custom_" & Guid.NewGuid().ToString("N").Substring(0, 8)
            End If
            t("id") = themeId
            If String.IsNullOrWhiteSpace(t.Value(Of String)("name")) Then t("name") = themeId
            If String.IsNullOrWhiteSpace(t.Value(Of String)("font")) Then t("font") = "微软雅黑"
            If Not (TypeOf t("colors") Is JObject) Then t("colors") = New JObject()

            Dim customPath As String = CustomPptThemesPath()
            Dim dir As String = Path.GetDirectoryName(customPath)
            If Not String.IsNullOrEmpty(dir) AndAlso Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)

            Dim root As JObject = New JObject()
            If File.Exists(customPath) Then
                Try
                    root = JObject.Parse(File.ReadAllText(customPath))
                Catch
                    root = New JObject()
                End Try
            End If
            Dim arr As JArray = TryCast(root("themes"), JArray)
            If arr Is Nothing Then
                arr = New JArray()
                root("themes") = arr
            End If
            Dim existing As JToken = arr.FirstOrDefault(
                Function(x) String.Equals(x.Value(Of String)("id"), themeId, StringComparison.OrdinalIgnoreCase))
            If existing IsNot Nothing Then
                existing.Replace(t)
            Else
                arr.Add(t)
            End If
            root("version") = If(root("version") Is Nothing, 1, root("version"))
            root("lastUpdated") = DateTime.Now.ToString("yyyy-MM-dd")
            File.WriteAllText(customPath, root.ToString(Formatting.Indented))
            Return JsonConvert.SerializeObject(New With {Key .success = True, Key .id = themeId, Key .message = "主题已保存并生效"})
        Catch ex As Exception
            Return JsonConvert.SerializeObject(New With {Key .success = False, Key .error = ex.Message})
        End Try
    End Function

    <ComVisible(True)>
    Public Overridable Function DeleteCustomPptTheme(themeId As String) As String
        Try
            If String.IsNullOrWhiteSpace(themeId) Then
                Return "{""success"":false,""error"":""主题 id 为空""}"
            End If
            Dim customPath As String = CustomPptThemesPath()
            If Not File.Exists(customPath) Then
                Return "{""success"":true,""message"":""无自定义主题""}"
            End If
            Dim root As JObject = JObject.Parse(File.ReadAllText(customPath))
            Dim arr As JArray = TryCast(root("themes"), JArray)
            If arr Is Nothing OrElse arr.Count = 0 Then
                Return "{""success"":true,""message"":""无自定义主题""}"
            End If
            Dim kept As New JArray()
            For Each x As JToken In arr
                If Not String.Equals(x.Value(Of String)("id"), themeId, StringComparison.OrdinalIgnoreCase) Then
                    kept.Add(x)
                End If
            Next
            root("themes") = kept
            File.WriteAllText(customPath, root.ToString(Formatting.Indented))
            Return JsonConvert.SerializeObject(New With {Key .success = True, Key .message = "已删除主题：" & themeId})
        Catch ex As Exception
            Return JsonConvert.SerializeObject(New With {Key .success = False, Key .error = ex.Message})
        End Try
    End Function

    <ComVisible(True)>
    Public Overridable Function ListCustomPptThemes() As String
        Try
            Dim customPath As String = CustomPptThemesPath()
            Dim arr As New JArray()
            If File.Exists(customPath) Then
                Try
                    Dim root As JObject = JObject.Parse(File.ReadAllText(customPath))
                    Dim t As JArray = TryCast(root("themes"), JArray)
                    If t IsNot Nothing Then arr = t
                Catch
                End Try
            End If
            Return JsonConvert.SerializeObject(New With {Key .success = True, Key .themes = arr})
        Catch ex As Exception
            Return JsonConvert.SerializeObject(New With {Key .success = False, Key .error = ex.Message, Key .themes = New JArray()})
        End Try
    End Function

    ' ------------------------------------------------------------------
    ' Excel 专属 / 尚未实现能力的兜底（返回结构化空结果，不抛异常）
    ' ------------------------------------------------------------------
    <ComVisible(True)>
    Public Overridable Function InsertTableData(rowsJson As String) As String
        Return NotSupported("InsertTableData")
    End Function

    <ComVisible(True)>
    Public Overridable Function InsertParsedData(rowsJson As String) As String
        Return NotSupported("InsertParsedData")
    End Function

    ' 【修复】前端解析器期望的是对象而非裸数组：
    '   GetParseScripts            → {success, scripts:[...]}
    '   GetParseScriptCategories   → {success, categories:[...]}
    '   GetParseScript             → {success, script:{...}}
    ' 之前返回裸 "[]"/"{}" 会让 result.success === undefined，
    ' 前端据此走到失败分支弹出「加载脚本失败: 未知错误」。
    ' PPT/Word 未内置 ParseScriptManager（会级联整个技能引擎），故返回空脚本集。
    <ComVisible(True)>
    Public Function GetParseScripts() As String
        Try
            Return Ocr.GetParseScripts()
        Catch ex As Exception
            Return "{""success"":true,""scripts"":[]}"
        End Try
    End Function

    <ComVisible(True)>
    Public Function GetParseScriptCategories() As String
        Try
            Return Ocr.GetParseScriptCategories()
        Catch ex As Exception
            Return "{""success"":true,""categories"":[]}"
        End Try
    End Function

    <ComVisible(True)>
    Public Function GetParseScript(scriptId As String) As String
        Try
            Return Ocr.GetParseScript(scriptId)
        Catch ex As Exception
            Return "{""success"":false,""error"":""" & Safe(ex.Message) & """}"
        End Try
    End Function

    <ComVisible(True)>
    Public Function SaveParseScript(scriptJson As String) As String
        Try
            Return Ocr.SaveParseScript(scriptJson)
        Catch ex As Exception
            Return NotSupported("SaveParseScript")
        End Try
    End Function

    <ComVisible(True)>
    Public Function DeleteParseScript(scriptId As String) As String
        Try
            Return Ocr.DeleteParseScript(scriptId)
        Catch ex As Exception
            Return "{""success"":true}"
        End Try
    End Function

    <ComVisible(True)>
    Public Function DuplicateParseScript(scriptId As String, newName As String) As String
        Try
            Return Ocr.DuplicateParseScript(scriptId, newName)
        Catch ex As Exception
            Return NotSupported("DuplicateParseScript")
        End Try
    End Function

    <ComVisible(True)>
    Public Function ResetBuiltinScript(scriptId As String) As String
        Try
            Return Ocr.ResetBuiltinScript(scriptId)
        Catch ex As Exception
            Return "{""success"":true}"
        End Try
    End Function

    <ComVisible(True)>
    Public Function ImportParseScriptFromMarkdown(markdownContent As String) As String
        Try
            Return Ocr.ImportParseScriptFromMarkdown(markdownContent)
        Catch ex As Exception
            Return NotSupported("ImportParseScriptFromMarkdown")
        End Try
    End Function

    <ComVisible(True)>
    Public Function ClearParseScriptCache() As String
        Try
            Return Ocr.ClearParseScriptCache()
        Catch ex As Exception
            Return "{""success"":false,""error"":""" & Safe(ex.Message) & """}"
        End Try
    End Function

    ' 【修复】前端期望 {success, items:[...]}（之前返回裸 "[]" 等价于 success=undefined）
    <ComVisible(True)>
    Public Function GetOcrElementsJson() As String
        Try
            Return Ocr.GetOcrElementsJson()
        Catch ex As Exception
            Return "{""success"":false,""items"":[]}"
        End Try
    End Function

    <ComVisible(True)>
    Public Function ExecuteSkill(skillId As String, argsJson As String) As String
        Return NotSupported("ExecuteSkill")
    End Function

    <ComVisible(True)>
    Public Function GetOllamaModels() As String
        Try
            ' 后端 Ollama 模型检测（绕过前端 CORS）。OcrBridge 只暴露异步版，这里同步等待。
            Return Ocr.GetOllamaModelsAsync().GetAwaiter().GetResult()
        Catch ex As Exception
            Return "{""success"":false,""models"":[],""error"":""" & Safe(ex.Message) & """}"
        End Try
    End Function

    <ComVisible(True)>
    Public Overridable Function GetAvailableMcpTools() As String
        Try
            If _mcp Is Nothing Then _mcp = New HostMcpTools(_kind, Me)
            Dim json = _mcp.GetToolsJson()
            Try
                Dim o = JObject.Parse(json)
                LogBridge("GetAvailableMcpTools 返回 " & If(o("count"), 0) & " 个工具")
            Catch
            End Try
            Return json
        Catch ex As Exception
            LogBridge("GetAvailableMcpTools 异常: " & ex.Message)
            Return "{""success"":false,""tools"":[],""count"":0,""error"":""" & Safe(ex.Message) & """}"
        End Try
    End Function

    <ComVisible(True)>
    Public Overridable Function CallLocalTool(toolName As String, argumentsJson As String) As String
        Try
            If _mcp Is Nothing Then _mcp = New HostMcpTools(_kind, Me)
            Return _mcp.CallTool(toolName, argumentsJson)
        Catch ex As Exception
            Return "{""success"":false,""error"":""" & Safe(ex.Message) & """}"
        End Try
    End Function

    ''' <summary>MCP 配置（前端启动时读取是否自动重连外部 MCP；本地宿主不启用外部客户端）</summary>
    <ComVisible(True)>
    Public Function GetMcpConfig() As String
        Return "{""success"":true,""config"":{""autoStartClients"":false}}"
    End Function

    ''' <summary>设置面板关闭（设置页 HTML 会调用 excelBridge.CloseSettingsPanel）</summary>
    <ComVisible(True)>
    Public Function CloseSettingsPanel() As String
        Try
            Dim panes = HostPaneManager.Current
            If panes IsNot Nothing Then
                panes.Hide("aisettings")
                panes.Hide("ocrsettings")
            End If
            Return "{""success"":true}"
        Catch ex As Exception
            Return "{""success"":false,""error"":""" & Safe(ex.Message) & """}"
        End Try
    End Function

    ' ------------------------------------------------------------------
    ' 技能：各宿主专属子集（来自 Config\<kind>_skills.json，由 gen_host_skills.py 生成）
    ' ------------------------------------------------------------------
    Protected Sub LogBridge(msg As String)
        Try
            Dim dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelAIVSTOAddIn")
            If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
            File.AppendAllText(Path.Combine(dir, "host_bridge.log"),
                DateTime.Now.ToString("HH:mm:ss") & " [" & _kind.ToString() & "] " & msg & vbCrLf)
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' 崩溃面包屑：进程被 wwlib 强退时，最后一步会留在 crash_last_step.txt（同步 Flush）。
    ''' </summary>
    Protected Sub CrashStep(stepName As String)
        Try
            Dim dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelAIVSTOAddIn")
            If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
            Dim line = DateTime.Now.ToString("HH:mm:ss.fff") & " [" & _kind.ToString() & "] " & stepName
            File.WriteAllText(Path.Combine(dir, "crash_last_step.txt"), line & vbCrLf)
            File.AppendAllText(Path.Combine(dir, "host_bridge.log"), line & vbCrLf)
        Catch
        End Try
    End Sub

    Private Function LoadHostSkills() As JArray
        Dim arr As New JArray()
        Try
            Dim fileName = If(_kind = OfficeHostKind.Word, "word_skills.json", "ppt_skills.json")
            Dim asmDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
            Dim baseDir = AppDomain.CurrentDomain.BaseDirectory
            Dim candidates As New List(Of String)()
            For Each d As String In New String() {asmDir, baseDir, Path.GetDirectoryName(asmDir)}
                If Not String.IsNullOrEmpty(d) Then
                    candidates.Add(Path.Combine(d, "Config", fileName))
                    candidates.Add(Path.Combine(d, fileName))
                End If
            Next
            Dim found As String = Nothing
            For Each c As String In candidates
                If System.IO.File.Exists(c) Then found = c : Exit For
            Next
            LogBridge("LoadHostSkills loc=" & asmDir & " file=" & fileName & " found=" & (found IsNot Nothing) & If(found IsNot Nothing, " path=" & found, ""))
            If found IsNot Nothing Then
                Dim tok = JToken.Parse(System.IO.File.ReadAllText(found))
                If tok.Type = JTokenType.Array Then arr = DirectCast(tok, JArray)
            End If
            ' 应用用户停用名单（按宿主持久化在 %LOCALAPPDATA%\ExcelAIVSTOAddIn\<kind>_disabled_skills.json）
            Dim disabled = LoadDisabledSkills()
            If disabled.Count > 0 Then
                For Each s As JToken In arr
                    Dim nm = If(s("name"), s("Name"))
                    If nm IsNot Nothing AndAlso disabled.Contains(nm.ToString()) Then
                        s("isEnabled") = False
                    End If
                Next
            End If
        Catch ex As Exception
            LogBridge("LoadHostSkills 失败: " & ex.Message)
            System.Diagnostics.Debug.WriteLine("[HostBridgeBase] 技能清单加载失败: " & ex.Message)
        End Try
        Return arr
    End Function

    ' 用户停用名单：按宿主持久化到 %LOCALAPPDATA%\ExcelAIVSTOAddIn\<kind>_disabled_skills.json
    ' （静态 <kind>_skills.json 由 gen_host_skills.py 生成、只读，停用状态单独存放，避免污染生成物）
    Private Function GetDisabledSkillsPath() As String
        Dim fileName = If(_kind = OfficeHostKind.Word, "word_disabled_skills.json", "ppt_disabled_skills.json")
        Dim dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelAIVSTOAddIn")
        If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
        Return Path.Combine(dir, fileName)
    End Function

    Private Function LoadDisabledSkills() As HashSet(Of String)
        Dim set0 As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Try
            Dim p = GetDisabledSkillsPath()
            If System.IO.File.Exists(p) Then
                Dim tok = JToken.Parse(System.IO.File.ReadAllText(p))
                If tok IsNot Nothing AndAlso tok.Type = JTokenType.Array Then
                    For Each t As JToken In tok
                        Dim n = t.ToString().Trim()
                        If n <> "" Then set0.Add(n)
                    Next
                End If
            End If
        Catch ex As Exception
            LogBridge("LoadDisabledSkills 失败: " & ex.Message)
        End Try
        Return set0
    End Function

    Private Sub SaveDisabledSkills(set0 As HashSet(Of String))
        Try
            Dim p = GetDisabledSkillsPath()
            Dim dir = Path.GetDirectoryName(p)
            If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
            Dim arr As New JArray()
            For Each n As String In set0
                arr.Add(n)
            Next
            System.IO.File.WriteAllText(p, arr.ToString(Formatting.None))
        Catch ex As Exception
            LogBridge("SaveDisabledSkills 失败: " & ex.Message)
        End Try
    End Sub

    <ComVisible(True)>
    Public Overridable Function GetAllSkills() As String
        Dim arr = LoadHostSkills()
        LogBridge("GetAllSkills 返回 " & arr.Count & " 个技能")
        Dim root As New JObject()
        root("success") = True
        root("skills") = arr
        Return root.ToString(Formatting.None)
    End Function

    <ComVisible(True)>
    Public Overridable Function GetQuickShortcutSkills() As String
        Dim qsArr As New JArray()
        For Each s As JToken In LoadHostSkills()
            Dim qs = s("isQuickShortcut")
            If qs IsNot Nothing AndAlso qs.ToString().ToLowerInvariant() = "true" Then qsArr.Add(s)
        Next
        Dim root As New JObject()
        root("success") = True
        root("skills") = qsArr
        Return root.ToString(Formatting.None)
    End Function

    <ComVisible(True)>
    Public Function GetSkillList() As String
        Return LoadHostSkills().ToString(Formatting.None)
    End Function

    <ComVisible(True)>
    Public Function GetSheets() As String
        Return "[]"
    End Function

    <ComVisible(True)>
    Public Function GetSelectedSheetNames() As String
        Return "[]"
    End Function

    <ComVisible(True)>
    Public Function GetActiveSheetName() As String
        Return ""
    End Function

    <ComVisible(True)>
    Public Function GetRangeData(rangeAddress As String) As String
        Return "{""success"":false,""error"":""当前宿主不支持单元格区域""}"
    End Function

    <ComVisible(True)>
    Public Function LogToHost(message As String) As String
        Try
            System.Diagnostics.Debug.WriteLine("[JS] " & message)
        Catch
        End Try
        Return "{""success"":true}"
    End Function

    ' ------------------------------------------------------------------
    ' 工具
    ' ------------------------------------------------------------------
    Protected Function NotSupported(methodName As String) As String
        Return "{""success"":false,""error"":""" & methodName & " 在当前宿主尚未实现""}"
    End Function

    Protected Function Ok(payload As String) As String
        Return "{""success"":true,""data"":" & payload & "}"
    End Function

    Protected Function Err(message As String) As String
        Return "{""success"":false,""error"":""" & Safe(message) & """}"
    End Function

    Protected Function Safe(s As String) As String
        If s Is Nothing Then Return ""
        Return s.Replace("\", "\\").Replace("""", "\""").Replace(vbCrLf, "\n").Replace(vbLf, "\n").Replace(vbTab, " ")
    End Function

    ' ------------------------------------------------------------------
    ' 颜色 / 对齐：原为 PptBridge 与 WordBridge 各自私有的两份重复实现，
    ' 现上提到基类共享。
    ' ------------------------------------------------------------------

    ''' <summary>#RRGGBB 或十进制字符串 → OLE 颜色（BGR）。三宿主语义一致，故共享。</summary>
    Protected Function OleColorFromArgb(s As String) As Integer
        If s Is Nothing Then Return 0
        s = s.Trim()
        If s.StartsWith("#") AndAlso s.Length = 7 Then
            Dim r = Convert.ToInt32(s.Substring(1, 2), 16)
            Dim g = Convert.ToInt32(s.Substring(3, 2), 16)
            Dim b = Convert.ToInt32(s.Substring(5, 2), 16)
            Return (b << 16) Or (g << 8) Or r
        End If
        Dim n As Integer
        If Integer.TryParse(s, n) Then Return n
        Return 0
    End Function

    ''' <summary>把 #RRGGBB 向白色混合 blend∈[0,1] 比例，返回更浅的同色调色（用于隔行底纹 zebra）。
    ''' 无法解析时回退浅灰 #F2F2F2。三宿主共享。</summary>
    Protected Function MakeZebraTint(hex As String, Optional blend As Double = 0.80) As String
        If hex Is Nothing Then Return "#F2F2F2"
        hex = hex.Trim()
        If Not (hex.StartsWith("#") AndAlso hex.Length = 7) Then Return "#F2F2F2"
        Try
            Dim r = Convert.ToInt32(hex.Substring(1, 2), 16)
            Dim g = Convert.ToInt32(hex.Substring(3, 2), 16)
            Dim b = Convert.ToInt32(hex.Substring(5, 2), 16)
            r = CInt(r + (255 - r) * blend)
            g = CInt(g + (255 - g) * blend)
            b = CInt(b + (255 - b) * blend)
            Return "#" & r.ToString("X2") & g.ToString("X2") & b.ToString("X2")
        Catch
            Return "#F2F2F2"
        End Try
    End Function

    ''' <summary>
    ''' 段落对齐：left/right/center → 各 Office 应用的对齐枚举值。
    ''' 【注意】不能统一成一个返回值 —— PowerPoint 与 Word 的对齐枚举本就不同：
    '''   PowerPoint  PpParagraphAlignment: 左=1 中=2 右=3
    '''   Word        WdParagraphAlignment: 左=0 中=1 右=2
    ''' 因此这里只给中性默认（居中=1，与 Word 一致），由各宿主子类按自身枚举重写。
    ''' </summary>
    Protected Overridable Function MapAlign(s As String) As Integer
        Select Case If(s, "").ToLower()
            Case "left" : Return 0
            Case "right" : Return 2
            Case Else : Return 1
        End Select
    End Function

    ''' <summary>前端传入的 JSON 数组字符串 → 二维字符串数组（极简解析，够用于插入表格/文本）</summary>
    Protected Function ParseRows(rowsJson As String) As String(,)
        ' 支持 [["a","b"],["c","d"]] 或 [{"k":"v"}]
        Dim list As New System.Collections.Generic.List(Of String())
        ' 先尝试对象数组
        Dim s = If(rowsJson, "").Trim()
        If s.StartsWith("[") Then
            Dim depth = 0
            Dim startIdx = -1
            For i As Integer = 0 To s.Length - 1
                Dim c = s(i)
                If c = "["c Then
                    depth += 1
                    If depth = 2 Then startIdx = i + 1
                ElseIf c = "]"c Then
                    depth -= 1
                    If depth = 1 AndAlso startIdx >= 0 Then
                        list.Add(ParseJsonArrayToStrings(s.Substring(startIdx, i - startIdx)))
                        startIdx = -1
                    End If
                End If
            Next
        End If
        If list.Count = 0 Then Return Nothing
        Dim cols = 0
        For Each r As String() In list
            If r.Length > cols Then cols = r.Length
        Next
        Dim result(cols - 1, list.Count - 1) As String
        For r As Integer = 0 To list.Count - 1
            For c As Integer = 0 To list(r).Length - 1
                result(c, r) = list(r)(c)
            Next
        Next
        Return result
    End Function

    Private Function ParseJsonArrayToStrings(inner As String) As String()
        Dim vals As New System.Collections.Generic.List(Of String)()
        Dim cur As New System.Text.StringBuilder()
        Dim inStr = False
        Dim esc = False
        For i As Integer = 0 To inner.Length - 1
            Dim c = inner(i)
            If inStr Then
                If esc Then
                    cur.Append(c)
                    esc = False
                ElseIf c = "\"c Then
                    esc = True
                ElseIf c = """"c Then
                    inStr = False
                Else
                    cur.Append(c)
                End If
            Else
                If c = """"c Then
                    inStr = True
                ElseIf c = ","c Then
                    vals.Add(cur.ToString())
                    cur.Clear()
                ElseIf c = "{"c OrElse c = "}"c OrElse c = "["c OrElse c = "]"c Then
                    ' 嵌套结构忽略（仅取标量值）
                ElseIf c = ":"c Then
                    ' 对象值：忽略键，保留值
                    cur.Clear()
                ElseIf Not Char.IsWhiteSpace(c) Then
                    cur.Append(c)
                End If
            End If
        Next
        vals.Add(cur.ToString())
        Return vals.ToArray()
    End Function


    ' ---- auto-filled Excel-only COM method safe defaults (unsupported on non-Excel hosts; avoids DISP_E_UNKNOWNNAME) ----
<ComVisible(True)>
Public Overridable Function AddMcpClientAsync(clientConfigJson As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function AutoReconnectMcpClientsAsync() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function BrowseLocalFile() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function CallMcpToolAsync(clientName As String, toolName As String, argumentsJson As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function CaptureAndRegisterUndo(label As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function CheckPythonEnvironment() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function CreateKnowledgeBaseAsync(name As String, description As String, Optional optionsJson As String = "") As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function DebugExecuteVBA(vbaCode As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function DeleteDocumentAsync(docId As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function DeleteKnowledgeBaseAsync(kbId As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function DiagnosePaddleOcr() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function DisconnectMcpClientAsync(name As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function ExecutePython(pythonCode As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function ExecuteSkillAsync(skillName As String, inputsJson As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function ExecuteVBA(vbaCode As String, Optional skillName As String = "") As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function FillDataToExcel(dataJson As String, dataType As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetActiveWorkbookName() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetAvailableSkills() As String
    Dim enabled As New JArray()
    For Each s As JToken In LoadHostSkills()
        Dim en = s("isEnabled")
        Dim isEn = True
        If en IsNot Nothing Then isEn = (en.ToString().ToLowerInvariant() = "true")
        If isEn Then enabled.Add(s)
    Next
    Dim root As New JObject()
    root("success") = True
    root("skills") = enabled
    Return root.ToString(Formatting.None)
End Function

<ComVisible(True)>
Public Overridable Function GetCellValue(address As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetCloudOcrConfig() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetCurrentOcrEngine() As String
    Try
        Return Ocr.GetCurrentOcrEngine()
    Catch
        Return "rapidocr"
    End Try
End Function

<ComVisible(True)>
Public Overridable Function GetDesktopPath() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetDocumentsAsync(kbId As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetHostWorkbook() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetIntentConfig() As String
    ' Word/PPT/Excel 共用 Config\intent_config.json；缺文件时前端走内嵌兜底
    Try
        Dim configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config", "intent_config.json")
        If File.Exists(configPath) Then
            Dim raw = File.ReadAllText(configPath)
            If Not String.IsNullOrWhiteSpace(raw) AndAlso raw.TrimStart().StartsWith("{") Then
                Return "{""success"":true,""config"":" & raw & "}"
            End If
        End If
    Catch ex As Exception
        System.Diagnostics.Debug.WriteLine("[HostBridge] GetIntentConfig: " & ex.Message)
    End Try
    Return "{""success"":false,""error"":""intent_config not found""}"
End Function

<ComVisible(True)>
Public Overridable Function GetKnowledgeBasesAsync() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetMcpStatus() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetMobileControlStatus() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetMultiSelection() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetOcrResult() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetOllamaModelsAsync() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetPaddleOcrAdvancedConfig() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetPaddleOcrInstallProgress() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetPaddleOcrStatus() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetPptTemplates() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

''' <summary>
''' 卡片类型配置（Excel 侧从 Skill 读识别类型）。Word/PPT 无此概念：返回空列表，
''' 避免前端 bridge.GetRecognitionTypes() 打到不存在的成员 → 0x80020006 未处理拒绝。
''' </summary>
<ComVisible(True)>
Public Overridable Function GetRecognitionTypes() As String
    Return "{""success"":true,""recognitionTypes"":[]}"
End Function

<ComVisible(True)>
Public Overridable Function GetQdrantInstallStatus() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetRagConfigAsync() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetRedoStack() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetSelectedRangeData() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetSelection() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

    ''' <summary>
    ''' 采集「当前选中的对象」结构化信息（幻灯片上的表格/形状、文档中的表格等）。
    ''' 用途：让 AI 知道用户选中的是具体对象，而不是"无选区"——这是 PPT/Word 此前
    ''' 只采集选中文本、却不报选中形状/表格，导致意图分析判"未指定操作对象"的根因。
    ''' Excel 走 Range 选区（GetSelection/GetMultiSelection），本方法仅 PPT/Word 重写。
    ''' </summary>
<ComVisible(True)>
Public Overridable Function GetSelectionInfo() As String
    Return "{""success"":false,""error"":""GetSelectionInfo 在当前宿主尚未实现""}"
End Function

<ComVisible(True)>
Public Overridable Function GetSidebarWidth() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetUIPath() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetUndoStack() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetUsedRangeData() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetWindowsOcrStatus() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetWorkbookPath(fileName As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

''' <summary>
''' 获取当前 Office 应用打开的文档/工作簿/演示文稿列表（Excel 专属；PPT/Word 返回空列表让前端优雅降级）。
''' 前端 excel-connection-widget.js 在 VSTO WebView2 加载时无条件调此方法，
''' 若宿主没有兜底 → WebView2 hostObject 内部抛 `TypeError: e is not a function`，
''' 且被 BridgeAccessor 代理异步 reject 后最终变成"未处理 Promise 拒绝"刷屏。
''' 这里返回与 ExcelDataBridge 同构的 {success, count, workbooks} 结构，前端按空工作簿列表处理。
''' </summary>
<ComVisible(True)>
Public Overridable Function GetOpenWorkbooks() As String
    Return "{""success"":true,""count"":0,""workbooks"":[]}"
End Function

''' <summary>
''' 激活指定名称的工作簿（Excel 专属；Word/PPT 返回失败让前端优雅降级）。
''' 前端 excel-connection-widget.js 在所有 VSTO 宿主加载时都会尝试调此方法，
''' 若宿主没有 stub → WebView2 hostObject 抛 0x80020006 未知名称 → 未处理 Promise 拒绝刷屏。
''' </summary>
<ComVisible(True)>
Public Overridable Function ActivateWorkbook(name As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

''' <summary>
''' 激活指定名称的工作表（Excel 专属；Word/PPT 返回失败让前端优雅降级）。
''' 同上，前端 switchSheet() 会调 ActivateSheet，宿主必须有 stub 才能避免 WebView2 的 0x80020006。
''' </summary>
<ComVisible(True)>
Public Overridable Function ActivateSheet(sheetName As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function ImportDocumentAsync(ossUrl As String, kbId As String, Optional title As String = "", Optional configJson As String = "") As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function ImportLocalDocumentAsync(localPath As String, kbId As String, Optional title As String = "") As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function InsertDataSmart(payloadJson As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function InstallOcrEngine() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function InstallPaddleOcr() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function InstallQdrantAsync() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function InstallQdrantFromLocalFileAsync(zipFilePath As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function Invoke(method As String, arg As Object) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Sub LogMessage(level As String, message As String)
    ' 【可观测性】前端 ChatLog 只写 WebView2 console，生产环境不可见；agent-loop / chat-orchestrator /
    ' mcp-dispatcher 的关键判定都经 LogMessage 落盘。Excel 版（ExcelBridge）写的是 ChatLogger → chat_log_*.txt，
    ' 而共享宿主此前是空实现，导致 PPT/Word 侧的关键决策全部静默丢弃（chat_log 里没有任何 PPT/Word 记录），
    ' 排查"工具为什么没被调用"只能靠翻代码猜。这里统一落到 host_bridge.log（与 LogBridge 同格式）。
    Try
        LogBridge("FRONTEND." & If(String.IsNullOrEmpty(level), "INFO", level) & " " & If(String.IsNullOrEmpty(message), "", message))
    Catch
    End Try
End Sub

<ComVisible(True)>
Public Overridable Function MatchCapability(userMessage As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function ParseDocumentAsync(fileName As String, base64Data As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function PickExcelFile(Optional filter As String = "", Optional title As String = "", Optional initialDir As String = "") As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function ProcessExcelFileAsync(base64Data As String, Optional sheetName As String = Nothing) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function QueryDocumentAsync(question As String, chunksJson As String, Optional topK As Integer = 5) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function RecognizeTextWithOCR(imagePath As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function RedoLastAction() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function ReloadAllSkills() As String
    ' 技能来自静态文件，每次调用均重新读取，重载即成功
    Return "{""success"":true}"
End Function

<ComVisible(True)>
Public Overridable Function GetSkillByName(name As String) As String
    If name Is Nothing Then name = ""
    Dim found As JToken = Nothing
    For Each s As JToken In LoadHostSkills()
        If s("name") IsNot Nothing AndAlso s("name").ToString() = name Then
            found = s
            Exit For
        End If
    Next
    If found Is Nothing Then
        Dim nf As New JObject()
        nf("success") = False
        nf("error") = "skill not found: " & name
        Return nf.ToString(Formatting.None)
    End If
    Dim root As New JObject()
    root("success") = True
    root("skill") = found
    Return root.ToString(Formatting.None)
End Function

<ComVisible(True)>
Public Overridable Function OpenSkillSourceFile(name As String) As String
    ' 宿主加载项不部署技能 MD 源文件，打开源文件不适用
    Return "{""success"":false,""error"":""该宿主不支持打开技能源文件""}"
End Function

<ComVisible(True)>
Public Overridable Function ToggleSkillEnabled(name As String, enabled As Boolean) As String
    If String.IsNullOrWhiteSpace(name) Then
        Return "{""success"":false,""error"":""skillName 为空""}"
    End If
    Dim disabled = LoadDisabledSkills()
    If enabled Then
        disabled.Remove(name)
    Else
        disabled.Add(name)
    End If
    SaveDisabledSkills(disabled)
    LogBridge("ToggleSkillEnabled " & name & " -> " & (If(enabled, "启用", "停用")))
    Return "{""success"":true}"
End Function

<ComVisible(True)>
Public Overridable Function ToggleQuickShortcut(name As String, isQuick As Boolean) As String
    Return "{""success"":true}"
End Function

<ComVisible(True)>
Public Overridable Function GetCustomSkillDebugInfo() As String
    ' 宿主无自定义技能持久化，返回空信息
    Return "{""success"":true,""customSkills"":[],""debug"":{}}"
End Function

<ComVisible(True)>
Public Overridable Function RemoveMcpClientAsync(name As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function ResetRangeContext() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Sub RunPaddleOCRInstaller()
End Sub

<ComVisible(True)>
Public Overridable Function SaveBase64Image(base64Image As String, fileName As String) As String
    Try
        ' 去除 Data URI 前缀（如 data:image/png;base64,）
        Dim base64Data As String = base64Image
        If base64Data.Contains(",") Then
            base64Data = base64Data.Substring(base64Data.IndexOf(","c) + 1)
        End If

        Dim bytes = Convert.FromBase64String(base64Data)
        Dim tempDir = IO.Path.Combine(IO.Path.GetTempPath(), "ExcelAIOCR")
        If Not IO.Directory.Exists(tempDir) Then IO.Directory.CreateDirectory(tempDir)

        ' 时间戳防重名
        Dim safeName = IO.Path.GetFileNameWithoutExtension(fileName) & "_" &
                       DateTime.Now.ToString("yyyyMMddHHmmss") & IO.Path.GetExtension(fileName)
        Dim filePath = IO.Path.Combine(tempDir, safeName)
        IO.File.WriteAllBytes(filePath, bytes)

        Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {.success = True, .path = filePath})
    Catch ex As Exception
        Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
    End Try
End Function

<ComVisible(True)>
Public Overridable Sub SaveCloudOcrConfig(configJson As String)
End Sub

<ComVisible(True)>
Public Overridable Function SaveOcrConfig(apiKey As String, secretKey As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function SaveRagConfigAsync(configJson As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function SetCellValue(address As String, value As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Sub SetOcrEngine(engine As String)
    Try
        Ocr.SetOcrEngine(engine)
    Catch
    End Try
End Sub

<ComVisible(True)>
Public Overridable Sub SetPaddleOcrAdvancedConfig(configJson As String)
End Sub

<ComVisible(True)>
Public Overridable Function SetSidebarWidth(width As Object) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Sub SetUndoSuppressed(flag As Boolean)
End Sub

<ComVisible(True)>
Public Overridable Function ShowOpenFileDialog() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function StartMcpServerAsync(Optional port As Integer = -1) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function StartOcrRecognition(imagePath As String, Optional recognitionMode As String = "auto") As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function StartQdrantAsync() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Sub StartSelectionMonitor()
End Sub

<ComVisible(True)>
Public Overridable Function StopMcpServerAsync() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Sub StopSelectionMonitor()
End Sub

<ComVisible(True)>
Public Overridable Function TestCloudOcrConfig(provider As String, configJson As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function TestEmbeddingConnectionAsync(provider As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function TestMcpClientConnectionAsync(configOrName As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function TestQdrantConnectionAsync() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function TestRagQueryAsync(question As String, kbId As String, Optional topK As Integer = 5) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function UndoLastAction() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

' ── 高频 Excel 专属兜底（前端可能无条件调用；缺失时 WebView2 内部抛 TypeError: e is not a function，
' 被 BridgeAccessor 包装后仍能变成未处理 Promise 拒绝刷屏） ──
<ComVisible(True)>
Public Overridable Function GetMcpServerStatus() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetMcpServerConfig() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function SaveMcpServerConfigAsync(configJson As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetMcpClientTools(name As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetAllCapabilities() As String
    Return "{""success"":true,""capabilities"":[]}"
End Function

<ComVisible(True)>
Public Overridable Function CallMcpTool(clientName As String, toolName As String, argumentsJson As String) As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Sub ShutdownMcpClients()
End Sub

<ComVisible(True)>
Public Overridable Function GetTempPath() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

<ComVisible(True)>
Public Overridable Function GetDownloadsPath() As String
    Return "{""success"":false,""error"":""operation not supported on this host""}"
End Function

''' <summary>
''' 在资源管理器中打开路径（文件定位选中 / 打开目录）。导出成功卡「打开文件夹」共用。
''' </summary>
<ComVisible(True)>
Public Overridable Function OpenPathInExplorer(path As String) As String
    Try
        If String.IsNullOrWhiteSpace(path) Then
            Return "{""success"":false,""error"":""路径为空""}"
        End If
        Dim p = path.Trim().Replace("/"c, "\"c)
        If System.IO.File.Exists(p) Then
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo() With {
                .FileName = "explorer.exe",
                .Arguments = "/select,""" & p & """",
                .UseShellExecute = True
            })
            Return "{""success"":true,""mode"":""select""}"
        End If
        If System.IO.Directory.Exists(p) Then
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(p) With {
                .UseShellExecute = True
            })
            Return "{""success"":true,""mode"":""dir""}"
        End If
        Dim parent = System.IO.Path.GetDirectoryName(p)
        If Not String.IsNullOrEmpty(parent) AndAlso System.IO.Directory.Exists(parent) Then
            System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(parent) With {
                .UseShellExecute = True
            })
            Return "{""success"":true,""mode"":""parent""}"
        End If
        Return "{""success"":false,""error"":""路径不存在""}"
    Catch ex As Exception
        Return "{""success"":false,""error"":""" & ex.Message.Replace("""", "'") & """}"
    End Try
End Function

<ComVisible(True)>
Public Overridable Function CloseMobilePanel() As String
    Return "{""success"":true}"
End Function

End Class

