Option Strict On
Option Explicit On
Option Infer On

Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading.Tasks
Imports System.Diagnostics
Imports System.Reflection
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json.Serialization
Imports ExcelAIVSTOAddIn.Skills
Imports ExcelAIVSTOAddIn.Utils
Imports ExcelAIVSTOAddIn.Services
Imports ExcelAIVSTOAddIn.Services.Mcp
Imports ExcelAIVSTOAddIn.Bridges
Imports Microsoft.Office.Interop.Excel

Namespace ExcelAIVSTOAddIn

    ''' <summary>
    ''' ExcelBridge - 前端与后端交互的桥梁门面）(Facade Pattern)
    ''' 
    ''' 重构说明:
    ''' 原始 ExcelBridge 承担了过多职责（2843行），现已拆分为以下职责单一的子类：
    '''   - SkillBridge: Skill 系统 API
    '''   - ExcelDataBridge: Excel 数据操作 API
    '''   - ConfigBridge: 配置管理 API (LLM/OCR)
    '''   - OcrBridge: OCR 引擎状态和测试 API
    '''   - UndoBridge: 撤销功能 API
    '''   - FormatBridge: 格式化与渲染 API
    '''   - McpBridge: MCP 协议完整 API
    '''   - RagBridge: RAG 知识）API
    '''   - DocumentBridge: 文档解析 API
    '''   - IntentBridge: 意图识别 API
    '''   - ThinkingBridge: 思考与动作引擎 API
    ''' 
    ''' 本类作为 Facade 门面，保持原）COM 接口不变）    ''' 内部将调用委托给各子类实例，确保前端 JavaScript 调用完全兼容）    ''' 
    ''' 保留在本类的核心功能:
    '''   - VBA 代码执行（与 Excel Application 紧密耦合）    '''   - 辅助方法（ShowMessage、LogMessage 等）
    ''' </summary>
    <ComVisible(True)>
    <ClassInterface(ClassInterfaceType.AutoDispatch)>
    <ProgId("ExcelAIVSTOAddIn.ExcelBridge")>
    Public Class ExcelBridge

        ' 单例实例 - 使用 Lazy(Of T) 确保线程安全
        Private Shared ReadOnly _lazyInstance As New Lazy(Of ExcelBridge)(Function() New ExcelBridge(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        ' ===== 子桥接器引用（通过 BridgeRegistry 统一管理）====
        Private ReadOnly _registry As BridgeRegistry = BridgeRegistry.Instance

        Private _skillBridge As SkillBridge
        Private _excelDataBridge As ExcelDataBridge
        Private _configBridge As ConfigBridge
        Private _licenseBridge As LicenseBridge
        Private _ocrBridge As OcrBridge
        Private _undoBridge As UndoBridge
        Private _formatBridge As FormatBridge
        Private _mcpBridge As McpBridge
        Private _ragBridge As RagBridge
        Private _documentBridge As DocumentBridge
        Private _intentBridge As IntentBridge
        Private _thinkingBridge As ThinkingBridge
        Private _browserDisplayBridge As BrowserDisplayBridge
        Private _mobileControlBridge As MobileControlBridge
        Private _vbaExecutionBridge As VbaExecutionBridge

        ' 配置管理
        Private _config As ConfigurationManager

        ' 频率限制
        Private _dataWriteLimiter As Utils.RateLimiter = Utils.RateLimiter.ForDataWrite
        Private _generalLimiter As Utils.RateLimiter = Utils.RateLimiter.ForGeneral

#Region "单例模式"

        ''' <summary>
        ''' 获取 ExcelBridge 单例实例（线程安全 Lazy 初始化）
        ''' </summary>
        Public Shared ReadOnly Property Instance As ExcelBridge
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        ''' <summary>
        ''' 私有构造函数（单例模式）
        ''' </summary>
        Private Sub New()
            Try
                System.Diagnostics.Debug.WriteLine("========== ExcelBridge (Facade) 构造函数开始==========")
                _config = ConfigurationManager.Instance
                System.Diagnostics.Debug.WriteLine("ConfigurationManager 初始化完成")
                System.Diagnostics.Debug.WriteLine("子桥接器将采用延迟初始化")
                System.Diagnostics.Debug.WriteLine("========== ExcelBridge (Facade) 构造函数完成==========")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("ExcelBridge 构造函数异常: " & ex.Message)
                Throw
            End Try
        End Sub

#End Region

#Region "子桥接器访问器（私有）"

        Private ReadOnly Property SkillBridge As SkillBridge
            Get
                If _skillBridge Is Nothing Then _skillBridge = SkillBridge.Instance
                Return _skillBridge
            End Get
        End Property

        Private ReadOnly Property ExcelDataBridge As ExcelDataBridge
            Get
                If _excelDataBridge Is Nothing Then _excelDataBridge = ExcelDataBridge.Instance
                Return _excelDataBridge
            End Get
        End Property

        Private ReadOnly Property ConfigBridge As ConfigBridge
            Get
                If _configBridge Is Nothing Then _configBridge = ConfigBridge.Instance
                Return _configBridge
            End Get
        End Property

        Private ReadOnly Property LicenseBridge As LicenseBridge
            Get
                If _licenseBridge Is Nothing Then _licenseBridge = LicenseBridge.Instance
                Return _licenseBridge
            End Get
        End Property

        Private ReadOnly Property OcrBridge As OcrBridge
            Get
                If _ocrBridge Is Nothing Then _ocrBridge = OcrBridge.Instance
                Return _ocrBridge
            End Get
        End Property

        Private ReadOnly Property UndoBridge As UndoBridge
            Get
                If _undoBridge Is Nothing Then _undoBridge = UndoBridge.Instance
                Return _undoBridge
            End Get
        End Property

        Private ReadOnly Property FormatBridge As FormatBridge
            Get
                If _formatBridge Is Nothing Then _formatBridge = FormatBridge.Instance
                Return _formatBridge
            End Get
        End Property

        Private ReadOnly Property McpBridge As McpBridge
            Get
                If _mcpBridge Is Nothing Then _mcpBridge = McpBridge.Instance
                Return _mcpBridge
            End Get
        End Property

        Private ReadOnly Property RagBridge As RagBridge
            Get
                If _ragBridge Is Nothing Then _ragBridge = RagBridge.Instance
                Return _ragBridge
            End Get
        End Property

        Private ReadOnly Property DocumentBridge As DocumentBridge
            Get
                If _documentBridge Is Nothing Then _documentBridge = DocumentBridge.Instance
                Return _documentBridge
            End Get
        End Property

        Private ReadOnly Property IntentBridge As IntentBridge
            Get
                If _intentBridge Is Nothing Then _intentBridge = IntentBridge.Instance
                Return _intentBridge
            End Get
        End Property

        Private ReadOnly Property ThinkingBridge As ThinkingBridge
            Get
                If _thinkingBridge Is Nothing Then _thinkingBridge = ThinkingBridge.Instance
                Return _thinkingBridge
            End Get
        End Property

        Private ReadOnly Property BrowserDisplayBridge As BrowserDisplayBridge
            Get
                If _browserDisplayBridge Is Nothing Then _browserDisplayBridge = BrowserDisplayBridge.Instance
                Return _browserDisplayBridge
            End Get
        End Property

        Private ReadOnly Property MobileControlBridge As MobileControlBridge
            Get
                If _mobileControlBridge Is Nothing Then _mobileControlBridge = MobileControlBridge.Instance
                Return _mobileControlBridge
            End Get
        End Property

        Private ReadOnly Property VbaExecutionBridge As VbaExecutionBridge
            Get
                If _vbaExecutionBridge Is Nothing Then
                    _vbaExecutionBridge = VbaExecutionBridge.Instance
                    _registry.Register(_vbaExecutionBridge)
                End If
                Return _vbaExecutionBridge
            End Get
        End Property

#End Region

#Region "Skill API (委托）SkillBridge)"

        <ComVisible(True)>
        Public Function GetAvailableSkills() As String
            Return SkillBridge.GetAvailableSkills()
        End Function

        <ComVisible(True)>
        Public Function GetSkillDetails(skillName As String) As String
            Return SkillBridge.GetSkillDetails(skillName)
        End Function

        <ComVisible(True)>
        Public Function GetSkillByName(skillName As String) As String
            Return SkillBridge.GetSkillByName(skillName)
        End Function

        <ComVisible(True)>
        Public Function GetQuickShortcutSkills() As String
            Return SkillBridge.GetQuickShortcutSkills()
        End Function

        <ComVisible(True)>
        Public Function ToggleQuickShortcut(skillName As String, isQuick As Boolean) As String
            Return SkillBridge.ToggleQuickShortcut(skillName, isQuick)
        End Function

        <ComVisible(True)>
        Public Function ToggleSkillEnabled(skillName As String, isEnabled As Boolean) As String
            Return SkillBridge.ToggleSkillEnabled(skillName, isEnabled)
        End Function

        <ComVisible(True)>
        Public Function ExecuteSkill(skillName As String, inputsJson As String) As String
            Return SkillBridge.ExecuteSkill(skillName, inputsJson)
        End Function

        <ComVisible(True)>
        Public Async Function ExecuteSkillAsync(skillName As String, inputsJson As String) As Task(Of String)
            Return Await SkillBridge.ExecuteSkillAsync(skillName, inputsJson)
        End Function

        <ComVisible(True)>
        Public Function GetAllSkills() As String
            Return SkillBridge.GetAllSkills()
        End Function

        <ComVisible(True)>
        Public Function GetRecognitionTypes() As String
            Return SkillBridge.GetRecognitionTypes()
        End Function

        <ComVisible(True)>
        Public Function ReloadAllSkills() As String
            Return SkillBridge.ReloadAllSkills()
        End Function

        <ComVisible(True)>
        Public Function DeleteSkill(skillName As String) As String
            Return SkillBridge.DeleteSkill(skillName)
        End Function

        <ComVisible(True)>
        Public Function SaveSkill(skillJson As String) As String
            Return SkillBridge.SaveSkill(skillJson)
        End Function

        <ComVisible(True)>
        Public Function OpenSkillSourceFile(skillName As String) As String
            Return SkillBridge.OpenSkillSourceFile(skillName)
        End Function

        ''' <summary>
        ''' 在资源管理器中打开路径：文件则定位选中，目录则打开；供导出成功卡「打开文件夹」使用。
        ''' </summary>
        <ComVisible(True)>
        Public Function OpenPathInExplorer(path As String) As String
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

#End Region

#Region "配置 API (委托）ConfigBridge)"

        <ComVisible(True)>
        Public Function GetSettings() As String
            Return ConfigBridge.GetSettings()
        End Function

        <ComVisible(True)>
        Public Function SaveSettings(settingsJson As String) As String
            Return ConfigBridge.SaveSettings(settingsJson)
        End Function

        <ComVisible(True)>
        Public Function GetLLMConfig() As String
            Return ConfigBridge.GetLLMConfig()
        End Function

        <ComVisible(True)>
        Public Function SaveLLMConfig(configJson As String) As String
            Return ConfigBridge.SaveLLMConfig(configJson)
        End Function

        <ComVisible(True)>
        Public Function GetOcrConfig() As String
            Return ConfigBridge.GetOcrConfig()
        End Function

        <ComVisible(True)>
        Public Function SaveOcrConfig(apiKey As String, secretKey As String) As String
            Return ConfigBridge.SaveOcrConfig(apiKey, secretKey)
        End Function

        <ComVisible(True)>
        Public Function SaveAIConfig(configJson As String) As String
            Return ConfigBridge.SaveAIConfig(configJson)
        End Function

        <ComVisible(True)>
        Public Function GetAIConfig() As String
            Return ConfigBridge.GetAIConfig()
        End Function

#End Region

#Region "订阅授权 API (委托 LicenseBridge，BYOK 席位)"

        <ComVisible(True)>
        Public Function LicenseRegister(bodyJson As String) As String
            Return LicenseBridge.Register(bodyJson)
        End Function

        <ComVisible(True)>
        Public Function LicenseLogin(bodyJson As String) As String
            Return LicenseBridge.Login(bodyJson)
        End Function

        <ComVisible(True)>
        Public Function LicenseLogout() As String
            Return LicenseBridge.Logout()
        End Function

        <ComVisible(True)>
        Public Function LicenseGetEntitlement() As String
            Return LicenseBridge.GetEntitlement()
        End Function

        <ComVisible(True)>
        Public Function LicenseIsSubscriptionActive() As String
            Return LicenseBridge.IsSubscriptionActive()
        End Function

        <ComVisible(True)>
        Public Function LicenseStartCheckout(bodyJson As String) As String
            Return LicenseBridge.StartCheckout(bodyJson)
        End Function

        <ComVisible(True)>
        Public Function LicenseOpenBillingPortal() As String
            Return LicenseBridge.OpenBillingPortal()
        End Function

        <ComVisible(True)>
        Public Function LicenseSetApiBaseUrl(url As String) As String
            Return LicenseBridge.SetLicenseApiBaseUrl(url)
        End Function

        <ComVisible(True)>
        Public Function LicenseGetApiBaseUrl() As String
            Return LicenseBridge.GetLicenseApiBaseUrl()
        End Function

        <ComVisible(True)>
        Public Function LicensePingHealth() As String
            Return LicenseBridge.PingHealth()
        End Function

        <ComVisible(True)>
        Public Function LicenseActivateKey(bodyJson As String) As String
            Return LicenseBridge.ActivateKey(bodyJson)
        End Function

        <ComVisible(True)>
        Public Function LicenseGetLocalActivation() As String
            Return LicenseBridge.GetLocalActivation()
        End Function

        <ComVisible(True)>
        Public Function LicenseClearLocalActivation() As String
            Return LicenseBridge.ClearLocalActivation()
        End Function

        ''' <summary>注销并申请退款（消 key 闭环）。激活记录与 Word/PPT 共享，注销同样全端生效。</summary>
        <ComVisible(True)>
        Public Function LicenseDeactivateForRefund() As String
            Return LicenseBridge.DeactivateForRefund()
        End Function

        ''' <summary>查询激活码是否已在本机注销过（查共享注销名单）。</summary>
        <ComVisible(True)>
        Public Function LicenseIsKeyDeactivated(bodyJson As String) As String
            Return LicenseBridge.IsKeyDeactivated(bodyJson)
        End Function

#End Region

#Region "PPT 自定义主题"

        <ComVisible(True)>
        Public Function SaveCustomPptTheme(themeJson As String) As String
            Return PptGenerationService.SaveCustomPptTheme(themeJson)
        End Function

        <ComVisible(True)>
        Public Function DeleteCustomPptTheme(themeId As String) As String
            Return PptGenerationService.DeleteCustomPptTheme(themeId)
        End Function

        <ComVisible(True)>
        Public Function ListCustomPptThemes() As String
            Return PptGenerationService.ListCustomPptThemes()
        End Function

#End Region

#Region "OCR API (委托）OcrBridge)"

        <ComVisible(True)>
        Public Function GetAvailableOcrEngines() As String
            Return OcrBridge.GetAvailableOcrEngines()
        End Function

        <ComVisible(True)>
        Public Function GetAvailableLLMProviders() As String
            Return OcrBridge.GetAvailableLLMProviders()
        End Function

        <ComVisible(True)>
        Public Function GetOcrStatus() As String
            Return OcrBridge.GetOcrStatus()
        End Function

        <ComVisible(True)>
        Public Function GetOcrEngineStatus(engine As String) As String
            Return OcrBridge.GetOcrEngineStatus(engine)
        End Function

        <ComVisible(True)>
        Public Function ForceCheckRapidOcrStatus() As String
            Return OcrBridge.ForceCheckRapidOcrStatus()
        End Function

        <ComVisible(True)>
        Public Function TestCloudOcrConfig(provider As String, configJson As String) As String
            Return OcrBridge.TestCloudOcrConfig(provider, configJson)
        End Function

        <ComVisible(True)>
        Public Async Function GetOllamaModelsAsync() As Task(Of String)
            Return Await OcrBridge.GetOllamaModelsAsync()
        End Function

        <ComVisible(True)>
        Public Sub SetOcrEngine(engine As String)
            OcrBridge.SetOcrEngine(engine)
        End Sub

        <ComVisible(True)>
        Public Sub SaveCloudOcrConfig(configJson As String)
            OcrBridge.SaveCloudOcrConfig(configJson)
        End Sub

        <ComVisible(True)>
        Public Function GetCurrentOcrEngine() As String
            Return OcrBridge.GetCurrentOcrEngine()
        End Function

        ''' <summary>
        ''' 启动截图识别功能
        ''' </summary>
        <ComVisible(True)>
        Public Sub StartScreenshotOcr()
            Try
                Dim addIn As ThisAddIn = Globals.ThisAddIn
                If addIn IsNot Nothing Then
                    addIn.StartScreenshotOcr()
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("ExcelBridge.StartScreenshotOcr 异常: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' 检查公式识别增强包（pix2tex）是否已安装
        ''' </summary>
        <ComVisible(True)>
        Public Function CheckFormulaOcrStatus() As String
            Return OcrBridge.CheckFormulaOcrStatus()
        End Function

        ''' <summary>
        ''' 启动公式识别增强包后台安装（非阻塞，立即返回）        ''' </summary>
        <ComVisible(True)>
        Public Function StartFormulaOcrInstall() As String
            Return OcrBridge.StartFormulaOcrInstall()
        End Function

        ''' <summary>
        ''' 获取公式增强包安装状态（轮询用）
        ''' </summary>
        <ComVisible(True)>
        Public Function GetFormulaInstallStatus() As String
            Return OcrBridge.GetFormulaInstallStatus()
        End Function

        ''' <summary>
        ''' 安装公式识别增强包（pix2tex / LaTeX-OCR）——同步版）        ''' </summary>
        <ComVisible(True)>
        Public Function InstallFormulaOcr() As String
            Return OcrBridge.InstallFormulaOcr()
        End Function

        <ComVisible(True)>
        Public Function GetCloudOcrConfig() As String
            Return OcrBridge.GetCloudOcrConfig()
        End Function

        <ComVisible(True)>
        Public Function GetPaddleOcrAdvancedConfig() As String
            Return OcrBridge.GetPaddleOcrAdvancedConfig()
        End Function

        <ComVisible(True)>
        Public Sub SetPaddleOcrAdvancedConfig(configJson As String)
            OcrBridge.SetPaddleOcrAdvancedConfig(configJson)
        End Sub

        ''' <summary>
        ''' ）Base64 图片保存到临时目录，返回 JSON { success, path }
        ''' </summary>
        <ComVisible(True)>
        Public Function SaveBase64Image(base64Image As String, fileName As String) As String
            Try
                ' 去除 Data URI 前缀（如 data:image/png;base64,）
                Dim base64Data As String = base64Image
                If base64Data.Contains(",") Then
                    base64Data = base64Data.Substring(base64Data.IndexOf(",") + 1)
                End If

                Dim bytes = Convert.FromBase64String(base64Data)
                Dim tempDir = Path.Combine(Path.GetTempPath(), "ExcelAIOCR")
                If Not Directory.Exists(tempDir) Then Directory.CreateDirectory(tempDir)

                ' 用时间戳防止文件名冲突
                Dim safeName = Path.GetFileNameWithoutExtension(fileName) & "_" & DateTime.Now.ToString("yyyyMMddHHmmss") & Path.GetExtension(fileName)
                Dim filePath = Path.Combine(tempDir, safeName)
                File.WriteAllBytes(filePath, bytes)

                Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                    .success = True,
                    .path = filePath
                })
            Catch ex As Exception
                Return Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ' OCR 异步任务状态已迁移）OcrBridge，消）Facade 层状态耦合

        ''' <summary>
        ''' 启动 OCR 识别（异步，立即返回，不阻塞 COM 线程）        ''' 委托）OcrBridge 处理
        ''' </summary>
        <ComVisible(True)>
        Public Function StartOcrRecognition(imagePath As String, Optional recognitionMode As String = "auto") As String
            Return OcrBridge.StartOcrRecognition(imagePath, recognitionMode)
        End Function

        ''' <summary>
        ''' 启动 OCR 识别并指定输出格式（异步，立即返回，不阻）COM 线程）        ''' 委托）OcrBridge 处理
        ''' </summary>
        <ComVisible(True)>
        Public Function StartOcrRecognitionWithFormat(imagePath As String, recognitionMode As String, outputFormat As String) As String
            Return OcrBridge.StartOcrRecognitionWithFormat(imagePath, recognitionMode, outputFormat)
        End Function

        ''' <summary>
        ''' 启动 PDF OCR 识别（异步，立即返回，不阻塞 COM 线程）
        ''' 引擎侧完成 PDF → 高分辨率图片 → 逐页 OCR → 按页分组
        ''' </summary>
        <ComVisible(True)>
        Public Function StartPdfOcrRecognition(pdfPath As String,
                                                  Optional recognitionMode As String = "auto",
                                                  Optional dpi As Integer = 300) As String
            Return OcrBridge.StartPdfOcrRecognition(pdfPath, recognitionMode, dpi)
        End Function

        ''' <summary>
        ''' 轮询 PDF OCR 结果（非阻塞，返回多页 JSON）
        ''' </summary>
        <ComVisible(True)>
        Public Function GetPdfOcrResult() As String
            Return OcrBridge.GetPdfOcrResult()
        End Function

        ''' <summary>
        ''' 轮询 OCR 识别结果（非阻塞）        ''' 委托）OcrBridge 处理
        ''' </summary>
        <ComVisible(True)>
        Public Function GetOcrResult() As String
            Return OcrBridge.GetOcrResult()
        End Function

        ''' <summary>
        ''' 轮询 OCR 识别结果（非阻塞，指定输出格式）
        ''' 委托）OcrBridge 处理
        ''' </summary>
        <ComVisible(True)>
        Public Function GetOcrResultWithFormat(format As String) As String
            Return OcrBridge.GetOcrResultWithFormat(format)
        End Function

        ''' <summary>
        ''' 获取 OCR 结构化结果（）box 坐标，用于图片标注）
        ''' 委托）OcrBridge 处理
        ''' </summary>
        <ComVisible(True)>
        Public Function GetOcrStructuredResult() As String
            Return OcrBridge.GetOcrStructuredResult()
        End Function

        ''' <summary>
        ''' 获取 PaddleOCR 原始返回数据（含坐标、文字、置信度）        ''' 委托）OcrBridge 处理
        ''' </summary>
        <ComVisible(True)>
        Public Function GetOcrRawData() As String
            Return OcrBridge.GetOcrRawData()
        End Function

        ''' <summary>
        ''' 按指定格式返）OCR 识别结果（plaintext/markdown/html/excel）        ''' 委托）OcrBridge 处理
        ''' </summary>
        <ComVisible(True)>
        Public Function RecognizeTextWithFormat(imagePath As String, format As String, Optional fileName As String = "") As String
            Return OcrBridge.RecognizeTextWithFormat(imagePath, format, fileName)
        End Function

#End Region

#Region "Excel 数据操作 API (委托）ExcelDataBridge)"

        <ComVisible(True)>
        Public Function FillDataToExcel(dataJson As String, dataType As String) As String
            Return ExcelDataBridge.FillDataToExcel(dataJson, dataType)
        End Function

        <ComVisible(True)>
        Public Function InsertParsedData(dataJson As String) As String
            Return ExcelDataBridge.InsertParsedData(dataJson)
        End Function

        <ComVisible(True)>
        Public Function GetSelection() As String
            Return ExcelDataBridge.GetSelection()
        End Function

        Public Function GetMultiSelection() As String
            Return ExcelDataBridge.GetMultiSelection()
        End Function

        <ComVisible(True)>
        Public Function GetCellValue(address As String) As String
            Return ExcelDataBridge.GetCellValue(address)
        End Function

        <ComVisible(True)>
        Public Function SetCellValue(address As String, value As String) As String
            Return ExcelDataBridge.SetCellValue(address, value)
        End Function

        <ComVisible(True)>
        Public Function DeleteRows(sheetName As String, rowIndex As Integer, count As Integer) As String
            Return ExcelDataBridge.DeleteRows(sheetName, rowIndex, count)
        End Function

        <ComVisible(True)>
        Public Function InsertRows(sheetName As String, rowIndex As Integer, count As Integer) As String
            Return ExcelDataBridge.InsertRows(sheetName, rowIndex, count)
        End Function

        <ComVisible(True)>
        Public Function GetActiveSheetName() As String
            Return ExcelDataBridge.GetActiveSheetName()
        End Function

        <ComVisible(True)>
        Public Function GetSheets() As String
            Return ExcelDataBridge.GetSheets()
        End Function

        <ComVisible(True)>
        Public Function GetSelectedSheetNames() As String
            Return ExcelDataBridge.GetSelectedSheetNames()
        End Function

        <ComVisible(True)>
        Public Function GetActiveCell() As String
            Return ExcelDataBridge.GetActiveCell()
        End Function

        <ComVisible(True)>
        Public Function GetActiveWorkbookName() As String
            Return ExcelDataBridge.GetActiveWorkbookName()
        End Function

        <ComVisible(True)>
        Public Function GetWorkbookPath(fileName As String) As String
            Return ExcelDataBridge.GetWorkbookPath(fileName)
        End Function

        <ComVisible(True)>
        Public Function GetOpenWorkbooks() As String
            Return ExcelDataBridge.GetOpenWorkbooks()
        End Function

        <ComVisible(True)>
        Public Function ActivateWorkbook(name As String) As String
            Return ExcelDataBridge.ActivateWorkbook(name)
        End Function

        <ComVisible(True)>
        Public Function ActivateSheet(sheetName As String) As String
            Return ExcelDataBridge.ActivateSheet(sheetName)
        End Function

        <ComVisible(True)>
        Public Function GetDesktopPath() As String
            Return Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        End Function

        <ComVisible(True)>
        Public Function GetDownloadsPath() As String
            Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads")
        End Function

        <ComVisible(True)>
        Public Function GetTempPath() As String
            Return Path.GetTempPath()
        End Function

        <ComVisible(True)>
        Public Function GetSelectedRangeData() As String
            Return ExcelDataBridge.GetSelectedRangeData()
        End Function

        <ComVisible(True)>
        Public Function GetRangeData(rangeAddress As String) As String
            Return ExcelDataBridge.GetRangeData(rangeAddress)
        End Function

        <ComVisible(True)>
        Public Sub StartSelectionMonitor()
            Try
                Dim addIn As ThisAddIn = Globals.ThisAddIn
                If addIn IsNot Nothing Then
                    addIn.StartSelectionMonitor()
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("ExcelBridge.StartSelectionMonitor 异常: " & ex.Message)
            End Try
        End Sub

        <ComVisible(True)>
        Public Sub StopSelectionMonitor()
            Try
                Dim addIn As ThisAddIn = Globals.ThisAddIn
                If addIn IsNot Nothing Then
                    addIn.StopSelectionMonitor()
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("ExcelBridge.StopSelectionMonitor 异常: " & ex.Message)
            End Try
        End Sub

        <ComVisible(True)>
        Public Function IsSelectionMonitorRunning() As Boolean
            Try
                Dim addIn As ThisAddIn = Globals.ThisAddIn
                If addIn IsNot Nothing Then
                    ' 通过反射访问私有字段，避免修）ThisAddIn 接口
                    Dim field = GetType(ThisAddIn).GetField("_selectionMonitorActive", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Instance)
                    If field IsNot Nothing Then
                        Return CBool(field.GetValue(addIn))
                    End If
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("ExcelBridge.IsSelectionMonitorRunning 异常: " & ex.Message)
            End Try
            Return False
        End Function

#End Region

#Region "撤销功能 (委托）UndoBridge)"

        <ComVisible(True)>
        Public Function GetUndoStack() As String
            Return UndoBridge.GetUndoStack()
        End Function

        <ComVisible(True)>
        Public Function UndoLastAction() As String
            Return UndoBridge.UndoLastAction()
        End Function

        <ComVisible(True)>
        Public Sub SetUndoSuppressed(flag As Boolean)
            Bridges.UndoBridge.UndoSuppressed = flag
        End Sub

        <ComVisible(True)>
        Public Function CaptureAndRegisterUndo(label As String) As String
            Try
                Dim app = Globals.ThisAddIn.Application
                Dim wb = app.ActiveWorkbook
                If wb Is Nothing Then Return "无活动工作簿"
                Dim ws = TryCast(wb.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                If ws Is Nothing Then Return "无活动工作表"
                ' 整表备份：把当前工作表整体拍快照压入撤销栈，撤销时整张还原
                Bridges.UndoBridge.Instance.AddSheetBackup(ws.Name)
                Return "ok"
            Catch ex As Exception
                Return "登记失败: " & ex.Message
            End Try
        End Function

        <ComVisible(True)>
        Public Sub ClearUndoStack()
            UndoBridge.ClearUndoStack()
        End Sub

        <ComVisible(True)>
        Public Function GetRedoStack() As String
            Return UndoBridge.GetRedoStack()
        End Function

        <ComVisible(True)>
        Public Function RedoLastAction() As String
            Return UndoBridge.RedoLastAction()
        End Function

        <ComVisible(True)>
        Public Sub ClearRedoStack()
            UndoBridge.ClearRedoStack()
        End Sub

#End Region

#Region "格式化与渲染 API (委托）FormatBridge)"

        <ComVisible(True)>
        Public Function RenderFormattedText(markdownText As String, Optional startCell As String = "") As String
            Return FormatBridge.RenderFormattedText(markdownText, startCell)
        End Function

        <ComVisible(True)>
        Public Function InsertTableData(rowsJson As String, Optional writeNumbers As Boolean = False) As String
            Return FormatBridge.InsertTableData(rowsJson, writeNumbers)
        End Function

        <ComVisible(True)>
        Public Function InsertDataSmart(payloadJson As String) As String
            Return FormatBridge.InsertDataSmart(payloadJson)
        End Function

#End Region

#Region "HTML 浏览器展）API (委托）BrowserDisplayBridge)"

        <ComVisible(True)>
        Public Function OpenHtmlInBrowser(htmlContent As String, Optional title As String = "ExcelAI报表") As String
            Return BrowserDisplayBridge.OpenHtmlInBrowser(htmlContent, title)
        End Function

        <ComVisible(True)>
        Public Function OpenChartInBrowser(chartType As String, dataJson As String, Optional title As String = "数据图表", Optional theme As String = "light") As String
            Return BrowserDisplayBridge.OpenChartInBrowser(chartType, dataJson, title, theme)
        End Function

        <ComVisible(True)>
        Public Function ReopenHtmlFile(filePath As String) As String
            Return BrowserDisplayBridge.ReopenHtmlFile(filePath)
        End Function

#End Region

#Region "新架构：思考与动作引擎 (委托）ThinkingBridge)"

        <ComVisible(True)>
        Public Async Function AnalyzeAndThink(
            message As String,
            chatHistoryJson As String,
            excelInfoJson As String,
            selectedCellsJson As String) As Task(Of String)

            Return Await ThinkingBridge.AnalyzeAndThink(message, chatHistoryJson, excelInfoJson, selectedCellsJson)
        End Function

        <ComVisible(True)>
        Public Async Function ExecuteAction(actionJson As String) As Task(Of String)
            Return Await ThinkingBridge.ExecuteAction(actionJson)
        End Function

        <ComVisible(True)>
        Public Function GetAvailableActions() As String
            Return ThinkingBridge.GetAvailableActions()
        End Function

#End Region

#Region "分层意图识别 API (委托）IntentBridge)"

        <ComVisible(True)>
        Public Function RecognizeIntent(input As String, Optional hasAttachment As Boolean = False) As String
            Return IntentBridge.RecognizeIntent(input, hasAttachment)
        End Function

        <ComVisible(True)>
        Public Function GetQuickResponse(input As String) As String
            Return IntentBridge.GetQuickResponse(input)
        End Function

        <ComVisible(True)>
        Public Function ReloadRules() As String
            Return IntentBridge.ReloadRules()
        End Function

#End Region

#Region "MCP API (委托）McpBridge)"

        <ComVisible(True)>
        Public Function GetMcpStatus() As String
            Return McpBridge.GetMcpStatus()
        End Function

        <ComVisible(True)>
        Public Function GetMcpServerStatus() As String
            Return McpBridge.GetMcpServerStatus()
        End Function

        <ComVisible(True)>
        Public Async Function StartMcpServerAsync(Optional port As Integer = -1) As Task(Of String)
            Return Await McpBridge.StartMcpServerAsync(port)
        End Function

        <ComVisible(True)>
        Public Async Function StopMcpServerAsync() As Task(Of String)
            Return Await McpBridge.StopMcpServerAsync()
        End Function

        <ComVisible(True)>
        Public Function GetMcpServerConfig() As String
            Return McpBridge.GetMcpServerConfig()
        End Function

        <ComVisible(True)>
        Public Async Function SaveMcpServerConfigAsync(configJson As String) As Task(Of String)
            Return Await McpBridge.SaveMcpServerConfigAsync(configJson)
        End Function

        <ComVisible(True)>
        Public Async Function AddMcpClientAsync(clientConfigJson As String) As Task(Of String)
            Return Await McpBridge.AddMcpClientAsync(clientConfigJson)
        End Function

        <ComVisible(True)>
        Public Async Function RemoveMcpClientAsync(name As String) As Task(Of String)
            Return Await McpBridge.RemoveMcpClientAsync(name)
        End Function

        <ComVisible(True)>
        Public Async Function DisconnectMcpClientAsync(name As String) As Task(Of String)
            Return Await McpBridge.DisconnectMcpClientAsync(name)
        End Function

        <ComVisible(True)>
        Public Function GetMcpClientTools(name As String) As String
            Return McpBridge.GetMcpClientTools(name)
        End Function

        <ComVisible(True)>
        Public Async Function TestMcpClientConnectionAsync(configOrName As String) As Task(Of String)
            Return Await McpBridge.TestMcpClientConnectionAsync(configOrName)
        End Function

        <ComVisible(True)>
        Public Async Function CallMcpToolAsync(clientName As String, toolName As String, argumentsJson As String) As Task(Of String)
            Return Await McpBridge.CallMcpToolAsync(clientName, toolName, argumentsJson)
        End Function

        <ComVisible(True)>
        Public Function CallMcpTool(clientName As String, toolName As String, argumentsJson As String) As String
            Return McpBridge.CallMcpTool(clientName, toolName, argumentsJson)
        End Function

        <ComVisible(True)>
        Public Function CallLocalTool(toolName As String, argumentsJson As String) As String
            Return McpBridge.CallLocalTool(toolName, argumentsJson)
        End Function

        <ComVisible(True)>
        Public Function ResetRangeContext() As String
            Return McpBridge.ResetRangeContext()
        End Function

        <ComVisible(True)>
        Public Function GetAvailableMcpTools() As String
            Return McpBridge.GetAvailableMcpTools()
        End Function

        ''' <summary>
        ''' 获取意图识别关键词配置（唯一事实源 Config/intent_config.json）
        ''' 返回 { success, config } JSON；加载失败返回 { success:false }
        ''' </summary>
        <ComVisible(True)>
        Public Function GetIntentConfig() As String
            Try
                Dim json = Services.Mcp.ExcelMcpTools.GetIntentConfigJson()
                If String.IsNullOrEmpty(json) Then
                    Return "{""success"":false,""error"":""intent_config.json 未加载""}"
                End If
                Return "{""success"":true,""config"":" & json & "}"
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[ExcelBridge] GetIntentConfig 异常: {ex.Message}")
                Return "{""success"":false,""error"":""" & ex.Message.Replace("""", "'") & """}"
            End Try
        End Function

        ''' <summary>
        ''' 根据用户消息匹配最佳能力（无参版本，WebView2 COM 不支持 Optional 参数）
        ''' </summary>
        <ComVisible(True)>
        Public Function MatchCapability(userMessage As String) As String
            Return McpBridge.MatchCapability(userMessage, 5)
        End Function

        ''' <summary>
        ''' 获取所有能力 JSON
        ''' </summary>
        <ComVisible(True)>
        Public Function GetAllCapabilities() As String
            Return McpBridge.GetAllCapabilities()
        End Function

        <ComVisible(True)>
        Public Function GetMcpConfig() As String
            Return McpBridge.GetMcpConfig()
        End Function

        <ComVisible(True)>
        Public Async Function AutoReconnectMcpClientsAsync() As Task(Of String)
            Return Await McpBridge.AutoReconnectMcpClientsAsync()
        End Function

        ''' <summary>
        ''' 关闭时释放所）MCP 客户端连）        ''' </summary>
        Public Sub ShutdownMcpClients()
            McpBridge.ShutdownMcpClients()
        End Sub

#End Region

#Region "统一生命周期管理"

        ''' <summary>
        ''' 统一释放所有子 Bridge 资源
        ''' 通过 BridgeRegistry 统一管理，新）Bridge 无需修改此方）        ''' </summary>
        Public Sub ShutdownAll()
            Try
                System.Diagnostics.Debug.WriteLine("[ExcelBridge] ShutdownAll 委托）BridgeRegistry...")
                _registry.ShutdownAll()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[ExcelBridge] ShutdownAll 异常: {ex.Message}")
            End Try
        End Sub

#End Region

#Region "文档解析 API (委托到 DocumentBridge)"

        <ComVisible(True)>
        Public Async Function ParseDocumentAsync(fileName As String, base64Data As String) As Task(Of String)
            Return Await DocumentBridge.ParseDocumentAsync(fileName, base64Data)
        End Function

        ''' <summary>
        ''' 处理 Excel 文件（前端 batch-processor / app-file-batch 调用入口）
        ''' 委托到 DocumentBridge.ParseDocumentAsync，自动路由到 ExcelFileParser
        ''' </summary>
        ''' <param name="base64Data">Base64 编码的文件内容</param>
        ''' <param name="sheetName">指定工作表名称（可选，当前未使用，保留兼容）</param>
        ''' <returns>解析结果JSON</returns>
        <ComVisible(True)>
        Public Async Function ProcessExcelFileAsync(base64Data As String, Optional sheetName As String = Nothing) As Task(Of String)
            Try
                System.Diagnostics.Debug.WriteLine($"[ExcelBridge.ProcessExcelFileAsync] 开始处理 Excel 文件, sheetName={If(sheetName, "(全部)")}")

                ' 使用固定文件名让 DocumentRouter 路由到 ExcelFileParser
                Dim fileName = If(sheetName, "uploaded.xlsx")
                If Not fileName.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) AndAlso Not fileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) Then
                    fileName = "uploaded.xlsx"
                End If

                Dim result = Await DocumentBridge.ParseDocumentAsync(fileName, base64Data)
                System.Diagnostics.Debug.WriteLine($"[ExcelBridge.ProcessExcelFileAsync] 处理完成, result={result.Substring(0, Math.Min(200, result.Length))}")
                Return result
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[ExcelBridge.ProcessExcelFileAsync] 异常: {ex.Message}")
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        <ComVisible(True)>
        Public Async Function ChunkDocumentAsync(
            fullText As String,
            Optional structureJson As String = "",
            Optional optionsJson As String = "") As Task(Of String)

            Return Await DocumentBridge.ChunkDocumentAsync(fullText, structureJson, optionsJson)
        End Function

        <ComVisible(True)>
        Public Async Function QueryDocumentAsync(
            question As String,
            chunksJson As String,
            Optional topK As Integer = 5) As Task(Of String)

            Return Await DocumentBridge.QueryDocumentAsync(question, chunksJson, topK)
        End Function

        <ComVisible(True)>
        Public Function GetSupportedDocFormats() As String
            Return DocumentBridge.GetSupportedDocFormats()
        End Function

#End Region

#Region "RAG 知识库管）API (委托）RagBridge)"

        ''' <summary>
        ''' 打开本地文件选择对话）        ''' </summary>
        <ComVisible(True)>
        Public Function BrowseLocalFile() As String
            Return RagBridge.BrowseLocalFile()
        End Function

        <ComVisible(True)>
        Public Function ImportDocumentAsync(ossUrl As String, kbId As String, Optional title As String = "", Optional configJson As String = "") As String
            Return RagBridge.ImportDocumentAsync(ossUrl, kbId, title, configJson)
        End Function

        ''' <summary>
        ''' 导入本地文件到知识库（接收本地文件路径）
        ''' </summary>
        <ComVisible(True)>
        Public Function ImportLocalDocumentAsync(localPath As String, kbId As String, Optional title As String = "") As String
            Return RagBridge.ImportLocalDocumentAsync(localPath, kbId, title)
        End Function

        <ComVisible(True)>
        Public Function SearchKnowledgeAsync(query As String, kbId As String, Optional topK As Integer = 5, Optional optionsJson As String = "") As String
            Return RagBridge.SearchKnowledgeAsync(query, kbId, topK, optionsJson)
        End Function

        <ComVisible(True)>
        Public Function GetKnowledgeBasesAsync() As String
            Return RagBridge.GetKnowledgeBasesAsync()
        End Function

        <ComVisible(True)>
        Public Function CreateKnowledgeBaseAsync(name As String, description As String, Optional optionsJson As String = "") As String
            Return RagBridge.CreateKnowledgeBaseAsync(name, description, optionsJson)
        End Function

        <ComVisible(True)>
        Public Function DeleteKnowledgeBaseAsync(kbId As String) As String
            Return RagBridge.DeleteKnowledgeBaseAsync(kbId)
        End Function

        <ComVisible(True)>
        Public Function GetDocumentsAsync(kbId As String) As String
            Return RagBridge.GetDocumentsAsync(kbId)
        End Function

        <ComVisible(True)>
        Public Function DeleteDocumentAsync(docId As String) As String
            Return RagBridge.DeleteDocumentAsync(docId)
        End Function

        ''' <summary>
        ''' 重试失败的知识库文档（重置状态并重新后台处理）
        ''' </summary>
        <ComVisible(True)>
        Public Function RetryDocumentAsync(docId As String) As String
            Return RagBridge.RetryDocumentAsync(docId)
        End Function

        ''' <summary>
        ''' 同步"待同步"状态的知识库（Qdrant 就绪后重试建集合）
        ''' </summary>
        <ComVisible(True)>
        Public Function SyncKnowledgeBaseStatusAsync() As String
            Return RagBridge.SyncKnowledgeBaseStatusAsync()
        End Function

        ''' <summary>
        ''' RAG 对话增强：检索知识库返回可注入 LLM 的上下文本（未命中 shouldFallback=True）
        ''' </summary>
        <ComVisible(True)>
        Public Function EnhanceWithRagAsync(question As String, Optional kbId As String = "") As String
            Return RagBridge.EnhanceWithRagAsync(question, kbId)
        End Function

        <ComVisible(True)>
        Public Function TestRagQueryAsync(question As String, kbId As String, Optional topK As Integer = 5) As String
            Return RagBridge.TestRagQueryAsync(question, kbId, topK)
        End Function

        <ComVisible(True)>
        Public Function GetRagConfigAsync() As String
            Return RagBridge.GetRagConfigAsync()
        End Function

        <ComVisible(True)>
        Public Function SaveRagConfigAsync(configJson As String) As String
            Return RagBridge.SaveRagConfigAsync(configJson)
        End Function

        <ComVisible(True)>
        Public Function TestQdrantConnectionAsync() As String
            Return RagBridge.TestQdrantConnectionAsync()
        End Function

        <ComVisible(True)>
        Public Function TestEmbeddingConnectionAsync(provider As String) As String
            Return RagBridge.TestEmbeddingConnectionAsync(provider)
        End Function

        <ComVisible(True)>
        Public Function GetQdrantInstallStatus() As String
            Return RagBridge.GetQdrantInstallStatus()
        End Function

        <ComVisible(True)>
        Public Function ShowOpenFileDialog() As String
            Return RagBridge.ShowOpenFileDialog()
        End Function

        <ComVisible(True)>
        Public Async Function InstallQdrantAsync() As Task(Of String)
            Return Await RagBridge.InstallQdrantAsync()
        End Function

        <ComVisible(True)>
        Public Function InstallQdrantFromLocalFileAsync(zipFilePath As String) As String
            Return RagBridge.InstallQdrantFromLocalFileAsync(zipFilePath)
        End Function

        <ComVisible(True)>
        Public Function StartQdrantAsync() As String
            Return RagBridge.StartQdrantAsync()
        End Function

        <ComVisible(True)>
        Public Function StopQdrantAsync() As String
            Return RagBridge.StopQdrantAsync()
        End Function

        ''' <summary>检测 RAG 本地环境（Ollama 安装/运行 + embedding 模型就绪）。前端: await excelBridge.CheckRagEnvironmentAsync()</summary>
        <ComVisible(True)>
        Public Async Function CheckRagEnvironmentAsync() As Task(Of String)
            Return Await RagBridge.CheckRagEnvironmentAsync()
        End Function

        ''' <summary>一键拉起 embedding 模型（ollama pull）。前端: await excelBridge.StartModelPull()</summary>
        <ComVisible(True)>
        Public Function StartModelPull() As String
            Return RagBridge.StartModelPull()
        End Function

        ''' <summary>查询模型拉起进度。前端: await excelBridge.GetModelPullProgress()</summary>
        <ComVisible(True)>
        Public Function GetModelPullProgress() As String
            Return RagBridge.GetModelPullProgress()
        End Function

        ''' <summary>取消正在进行的模型拉起。前端: await excelBridge.StopModelPull()</summary>
        <ComVisible(True)>
        Public Function StopModelPull() As String
            Return RagBridge.StopModelPull()
        End Function

        ''' <summary>用系统默认浏览器打开外部链接（如 Ollama 下载页）。前端: await excelBridge.OpenExternalUrl(url)</summary>
        <ComVisible(True)>
        Public Function OpenExternalUrl(url As String) As String
            Return RagBridge.OpenExternalUrl(url)
        End Function

        ''' <summary>
        ''' 弹出 Windows 原生文件选择框，供前端选目标 Excel 文件。
        ''' 可选参数：
        '''  - filter: 过滤字符串，格式如 "Excel 文件|*.xlsx;*.xls;*.xlsm|所有文件|*.*"；为空默认 Excel 过滤
        '''  - title: 对话框标题；为空默认"选择目标 Excel 文件"
        '''  - initialDir: 初始目录；为空默认桌面
        ''' 返回 JSON: {success, filePath, message, error}
        ''' </summary>
        <ComVisible(True)>
        Public Function PickExcelFile(Optional filter As String = "", Optional title As String = "", Optional initialDir As String = "") As String
            Try
                Using dlg As New OpenFileDialog()
                    If String.IsNullOrEmpty(title) Then title = "选择目标 Excel 文件"
                    If String.IsNullOrEmpty(filter) Then filter = "Excel 文件 (*.xlsx;*.xls;*.xlsm;*.csv)|*.xlsx;*.xls;*.xlsm;*.csv|Excel 工作簿 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*"
                    If String.IsNullOrEmpty(initialDir) Then
                        initialDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                    End If
                    dlg.Title = title
                    dlg.Filter = filter
                    dlg.CheckFileExists = False
                    dlg.CheckPathExists = True
                    Try
                        If Directory.Exists(initialDir) Then dlg.InitialDirectory = initialDir
                    Catch : End Try

                    If dlg.ShowDialog() = DialogResult.OK Then
                        Return JsonConvert.SerializeObject(New With {
                            .success = True,
                            .filePath = dlg.FileName
                        })
                    Else
                        Return JsonConvert.SerializeObject(New With {
                            .success = False,
                            .message = "用户取消选择"
                        })
                    End If
                End Using
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

#End Region

#Region "移动端遥）API (委托）MobileControlBridge)"

        ''' <summary>
        ''' 启动移动端遥）HTTP 服务
        ''' </summary>
        <ComVisible(True)>
        Public Function StartMobileControl(Optional port As Integer = -1) As String
            Return MobileControlBridge.StartServer(port)
        End Function

        ''' <summary>
        ''' 停止移动端遥）HTTP 服务
        ''' </summary>
        <ComVisible(True)>
        Public Function StopMobileControl() As String
            Return MobileControlBridge.StopServer()
        End Function

        ''' <summary>
        ''' 获取移动端遥控服务状）        ''' </summary>
        <ComVisible(True)>
        Public Function GetMobileControlStatus() As String
            Return MobileControlBridge.GetStatus()
        End Function

        ''' <summary>
        ''' 打开手机遥控面板（代理方法，转发）ThisAddIn）        ''' </summary>
        <ComVisible(True)>
        Public Sub OpenMobilePanel()
            Try
                Globals.ThisAddIn.ShowMobilePanel()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("打开手机遥控面板失败: " & ex.Message)
            End Try
        End Sub

        ''' <summary>用户打开独立网页版时再拉起本地桥（9527），不在 Excel 启动时启动。</summary>
        <ComVisible(True)>
        Public Sub EnsureLocalBridge()
            Try
                Services.BridgeServerLauncher.EnsureStarted()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("启动本地桥失败: " & ex.Message)
            End Try
        End Sub

        <ComVisible(True)>
        Public Sub OpenOcrSettingsPanel()
            Try
                Globals.ThisAddIn.ShowOcrSettingsPanel()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("打开 OCR 模型设置失败: " & ex.Message)
            End Try
        End Sub

        <ComVisible(True)>
        Public Function ClearLocalUserData() As String
            Return ConfigBridge.ClearLocalUserData()
        End Function

        ''' <summary>
        ''' 关闭手机遥控面板（代理方法，转发）ThisAddIn）        ''' </summary>
        <ComVisible(True)>
        Public Sub CloseMobilePanel()
            Try
                Globals.ThisAddIn.CloseMobilePanel()
            Catch ex As Exception  ' 
            End Try
        End Sub

#End Region

#Region "LLM 流式响应 API"

        ''' <summary>
        ''' 流式调用 LLM ））token 推送到前端
        ''' 前端通过 window.__onLLMStream 回调接收 chunk
        ''' </summary>
        <ComVisible(True)>
        Public Async Function StreamLLMChatAsync(prompt As String, Optional streamId As String = "") As Task(Of String)
            Try
                If String.IsNullOrWhiteSpace(prompt) Then
                    Return ApiResponse.Fail("prompt 不能为空")
                End If

                Dim engine = Services.MultimodalLLMEngine.SharedInstance
                Dim contents = New List(Of Services.MultimodalContent) From {
                    Services.MultimodalContent.CreateText(prompt)
                }

                Dim result = Await engine.ProcessContentStreamingAsync(contents, "", If(String.IsNullOrEmpty(streamId), Nothing, streamId))

                If result.Success Then
                    Return ApiResponse.Ok(New With {
                        .content = result.Content,
                        .usage = result.Usage,
                        .processingTime = result.ProcessingTime.TotalSeconds
                    }, "流式响应完成")
                Else
                    Return ApiResponse.Fail(result.ErrorMessage)
                End If
            Catch ex As Exception
                Return ApiResponse.Fail(ex.Message)
            End Try
        End Function

#End Region

        ''' <summary>
        ''' 后端原生生成 PPT（不生成/不编译 VBA）
        ''' 读选区数据、用 PowerPoint COM 直出固定主题演示文稿。前端入口。
        ''' </summary>
        <ComVisible(True)>
        Public Function GeneratePPTFromData(argsJson As String) As String
            Return PptGenerationService.Generate(argsJson)
        End Function

        ''' <summary>
        ''' 获取 PPT 模板配置（主题 + 版式白名单），供前端动态加载主题与版式。
        ''' 数据源：Config/ppt_templates.json（内置）+ Config/custom_ppt_themes.json（用户导入）合并。
        ''' 成功时返回合并后的 JSON；失败时返回 {"success":false,"error":"..."}。
        ''' </summary>
        <ComVisible(True)>
        Public Function GetPptTemplates() As String
            Try
                Dim merged As String = PptGenerationService.GetMergedTemplatesJson()
                If String.IsNullOrWhiteSpace(merged) OrElse merged = "{}" Then
                    Return JsonConvert.SerializeObject(New With {
                        Key .success = False,
                        Key .error = "ppt_templates.json 不存在：" & Path.Combine(PathHelper.ConfigDirectory, "ppt_templates.json")
                    })
                End If
                Return merged
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    Key .success = False,
                    Key .error = "读取 PPT 模板配置失败：" & ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 从已有 PPT 提取主题（配色 + 字体）并注册为自定义主题。
        ''' 成功返回 {"success":true,"themeId","themeName","font","colors"}；失败返回 {"success":false,"error"}。
        ''' </summary>
        <ComVisible(True)>
        Public Function ImportPptTheme(pptxPath As String) As String
            Return PptGenerationService.ImportThemeFromPptx(pptxPath)
        End Function

        ''' <summary>
        ''' 打开原生文件选择对话框，选择公司模板 .pptx。
        ''' 成功返回 {"success":true,"filePath":"..."}；用户取消返回 {"success":false,"message":"用户取消选择"}。
        ''' </summary>
        <ComVisible(True)>
        Public Function BrowsePptFile() As String
            Return PptGenerationService.BrowsePptFile()
        End Function

        ''' <summary>
        ''' 数据感知设计推荐：分析选区数据，返回推荐版式/图表/洞察（供前端配置卡预填）。
        ''' 数据源：前端传入 headers/rows（与 GeneratePPTFromData 同契约）。
        ''' 成功时返回 {"success":true,"slides":[...],"insights":[...],"profile":{...}}。
        ''' </summary>
        <ComVisible(True)>
        Public Function RecommendPptDesign(argsJson As String) As String
            Try
                Dim args As JObject = Nothing
                If Not String.IsNullOrWhiteSpace(argsJson) Then
                    Try
                        args = JObject.Parse(argsJson)
                    Catch ex As Exception  ' 
                        args = Nothing
                    End Try
                End If

                Dim headers As List(Of String) = Nothing
                Dim rows As List(Of Dictionary(Of String, Object)) = Nothing
                If args IsNot Nothing Then
                    Try
                        headers = args("headers")?.ToObject(Of List(Of String))()
                    Catch ex As Exception  ' 
                        headers = Nothing
                    End Try
                    Try
                        rows = args("rows")?.ToObject(Of List(Of Dictionary(Of String, Object)))()
                    Catch ex As Exception  ' 
                        rows = Nothing
                    End Try
                End If
                If headers Is Nothing OrElse headers.Count = 0 Then
                    Return JsonConvert.SerializeObject(New With {
                        Key .success = False,
                        Key .error = "缺少选区表头（headers），请先在 Excel 中选中数据区域"
                    })
                End If

                Dim reportTitle As String = If(args Is Nothing, "", args.Value(Of String)("reportTitle"))
                If String.IsNullOrWhiteSpace(reportTitle) Then reportTitle = "数据分析报告"
                Dim sheetName As String = If(args Is Nothing, "", args.Value(Of String)("sheetName"))
                Dim addr As String = If(args Is Nothing, "", args.Value(Of String)("address"))
                Dim src As String = If(String.IsNullOrWhiteSpace(sheetName), "", sheetName)
                If Not String.IsNullOrWhiteSpace(addr) Then src = If(src = "", addr, src & " · " & addr)
                If String.IsNullOrWhiteSpace(src) Then src = "Excel 选区"

                ' LLM 列理解提示（可选）：前端让大模型思考"该拿什么列"后下发，
                ' 覆盖规则引擎的启发式选列（主角指标/分类/时间/ID 列）
                Dim columnHints As JObject = Nothing
                If args IsNot Nothing AndAlso args("columnHints") IsNot Nothing Then
                    Try
                        columnHints = TryCast(args("columnHints"), JObject)
                    Catch ex As Exception  ' 
                        columnHints = Nothing
                    End Try
                End If

                ' 用户需求原文（可选）：后端据此做结构偏置（趋势/对比/构成/明细/图表类型）
                Dim requirement As String = If(args Is Nothing, "", args.Value(Of String)("requirement"))

                Dim profile As PptDataProfile = PptDesignEngine.AnalyzeData(headers, rows)
                PptDesignEngine.ApplyColumnHints(profile, columnHints)
                Dim slides As JArray = PptDesignEngine.BuildDesign(profile, reportTitle, src, requirement)
                ' 结构化洞察（kind=kpi/trend/composition/note），供前端"洞察预览"分区渲染
                Dim insights As New JArray()
                For Each obj As JObject In PptDesignEngine.GenerateInsightObjects(profile)
                    insights.Add(obj)
                Next

                Dim profileSummary As New JObject()
                profileSummary("rowCount") = profile.RowCount
                Dim numericNames As New JArray()
                For Each cp As PptColumnProfile In profile.NumericColumns
                    numericNames.Add(cp.Name)
                Next
                profileSummary("numericColumns") = numericNames
                profileSummary("timeSeriesColumn") = If(profile.TimeSeriesColumn Is Nothing, "", profile.TimeSeriesColumn.Name)
                profileSummary("mainCategoryColumn") = If(profile.MainCategoryColumn Is Nothing, "", profile.MainCategoryColumn.Name)

                Dim result As New JObject()
                result("success") = True
                result("slides") = slides
                result("insights") = insights
                result("profile") = profileSummary
                result("orientation") = PptDesignEngine.RecommendOrientation(profile)
                Return result.ToString()
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    Key .success = False,
                    Key .error = "设计推荐失败：" & ex.Message
                })
            End Try
        End Function

#Region "VBA 执行 API (委托）VbaExecutionBridge)"

        ''' <summary>
        ''' 执行 VBA 代码（前端调用入口）
        ''' 委托）VbaExecutionBridge，保）COM 接口兼容
        ''' </summary>
        <ComVisible(True)>
        Public Function ExecuteVBA(vbaCode As String, Optional skillName As String = "") As String
            Return VbaExecutionBridge.ExecuteVBA(vbaCode, skillName)
        End Function

        ''' <summary>
        ''' 调试执行 VBA 代码：一次性执行并返回每行执行后的变量状态日）        ''' 委托）VbaExecutionBridge，保）COM 接口兼容
        ''' </summary>
        <ComVisible(True)>
        Public Function DebugExecuteVBA(vbaCode As String) As String
            Return VbaExecutionBridge.DebugExecuteVBA(vbaCode)
        End Function

#End Region

#Region "辅助方法 (保留）ExcelBridge ）"

        <ComVisible(True)>
        Public Sub ShowMessage(message As String, title As String)
            Try
                MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("显示消息框失） " & ex.Message)
            End Try
        End Sub

        <ComVisible(True)>
        Public Sub CloseOcrPanel()
            Try
                Globals.ThisAddIn.CloseOcrPanel()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("关闭 OCR 面板失败: " & ex.Message)
            End Try
        End Sub

        <ComVisible(True)>
        Public Sub OpenOcrPanel()
            Try
                Globals.ThisAddIn.ShowOcrPanel()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("打开识图入表面板失败: " & ex.Message)
            End Try
        End Sub

        <ComVisible(True)>
        Public Sub OpenSettingsPanel(Optional tabName As String = "")
            Try
                Globals.ThisAddIn.ShowSettingsPanel(tabName)
                System.Diagnostics.Debug.WriteLine("打开设置面板成功" & If(String.IsNullOrEmpty(tabName), "", "，切换到标签: " & tabName))
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("打开设置面板失败: " & ex.Message)
            End Try
        End Sub

        <ComVisible(True)>
        Public Sub CloseSettingsPanel()
            Try
                Globals.ThisAddIn.CloseSettingsPanel()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("关闭设置面板失败: " & ex.Message)
            End Try
        End Sub

        <ComVisible(True)>
        Public Sub OpenAiSettingsPanel()
            Try
                Globals.ThisAddIn.ShowAiSettingsPanel()
                System.Diagnostics.Debug.WriteLine("打开 AI 模型设置面板成功")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("打开 AI 模型设置面板失败: " & ex.Message)
            End Try
        End Sub

        ''' <summary>打开 PPT 主题模板编辑器任务窗格（注入 excelBridge，保存/列表可直连）</summary>
        <ComVisible(True)>
        Public Sub OpenPptThemeEditor()
            Try
                Globals.ThisAddIn.ShowPptThemeEditorPanel()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("打开 PPT 主题编辑器失败: " & ex.Message)
            End Try
        End Sub

        <ComVisible(True)>
        Public Sub ClosePptThemeEditor()
            Try
                Globals.ThisAddIn.ClosePptThemeEditorPanel()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("关闭 PPT 主题编辑器失败: " & ex.Message)
            End Try
        End Sub

        <ComVisible(True)>
        Public Sub LogMessage(level As String, message As String)
            Try
                System.Diagnostics.Debug.WriteLine($"[Frontend.{level}] {message}")
                Utils.ChatLogger.Instance?.LogInfo("FRONTEND", $"[{level}] {message}")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"LogMessage 自身错误: {ex.Message}")
            End Try
        End Sub

        <ComVisible(True)>
        Public Function GetLogs(count As Integer) As String
            Try
                Dim chatLogger = Utils.ChatLogger.Instance
                If chatLogger Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False, .error = "日志记录器未初始化"})
                End If

                Dim logText As String = chatLogger.GetRecentLogs(Math.Max(1, Math.Min(count, 500)))
                If String.IsNullOrWhiteSpace(logText) OrElse logText = "日志文件不存在" Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = True, .count = 0, .logs = New List(Of String)()})
                End If

                Dim lines = logText.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries).ToList()
                Return JsonConvert.SerializeObject(New With {
                    .success = True, .count = lines.Count, .logs = lines
                })
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[GetLogs] 错误: {ex.Message}")
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        <ComVisible(True)>
        Public Function ClearLogs() As String
            Try
                Dim chatLogger = Utils.ChatLogger.Instance
                If chatLogger IsNot Nothing Then
                    chatLogger.ClearAllLogs()
                End If
                Return JsonConvert.SerializeObject(New With {.success = True, .message = "日志已清除"})
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[ClearLogs] 错误: {ex.Message}")
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

#Region "OCR解析脚本 API"

        Public Function GetParseScripts() As String
            Return OcrBridge.GetParseScripts()
        End Function

        Public Function ClearParseScriptCache() As String
            Return OcrBridge.ClearParseScriptCache()
        End Function

        Public Function GetParseScript(scriptId As String) As String
            Return OcrBridge.GetParseScript(scriptId)
        End Function

        Public Function SaveParseScript(scriptJson As String) As String
            Return OcrBridge.SaveParseScript(scriptJson)
        End Function

        Public Function DeleteParseScript(scriptId As String) As String
            Return OcrBridge.DeleteParseScript(scriptId)
        End Function

        Public Function ResetBuiltinScript(scriptId As String) As String
            Return OcrBridge.ResetBuiltinScript(scriptId)
        End Function

        Public Function DuplicateParseScript(sourceId As String, newName As String) As String
            Return OcrBridge.DuplicateParseScript(sourceId, newName)
        End Function

        Public Function ExecuteParseScript(scriptId As String, ocrText As String) As String
            Return OcrBridge.ExecuteParseScript(scriptId, ocrText)
        End Function

        Public Function GetParseScriptCategories() As String
            Return OcrBridge.GetParseScriptCategories()
        End Function

        Public Function SearchParseScripts(keyword As String) As String
            Return OcrBridge.SearchParseScripts(keyword)
        End Function

        Public Function ImportParseScriptFromMarkdown(markdownContent As String) As String
            Return OcrBridge.ImportParseScriptFromMarkdown(markdownContent)
        End Function

        <ComVisible(True)>
        Public Function GetDomainPacks() As String
            Return OcrBridge.GetDomainPacks()
        End Function

        <ComVisible(True)>
        Public Function InstallDomainPackFromMarkdown(markdown As String) As String
            Return OcrBridge.InstallDomainPackFromMarkdown(markdown)
        End Function

        <ComVisible(True)>
        Public Function SetDomainPackEnabled(id As String, enabled As Boolean) As String
            Return OcrBridge.SetDomainPackEnabled(id, enabled)
        End Function

        <ComVisible(True)>
        Public Function UninstallDomainPack(id As String) As String
            Return OcrBridge.UninstallDomainPack(id)
        End Function

        <ComVisible(True)>
        Public Function SetDomainPackOption(id As String, key As String, value As String) As String
            Return OcrBridge.SetDomainPackOption(id, key, value)
        End Function

        <ComVisible(True)>
        Public Function SaveMechanicalCorrection(json As String) As String
            Return OcrBridge.SaveMechanicalCorrection(json)
        End Function

#End Region

#End Region

        ''' <summary>
        ''' 设置自动化执行免确认开关（由前端「自动化免确认」开关调用）。
        ''' 开启后，VBA / Python 高风险代码执行跳过用户确认弹窗，供自动化任务自动执行。
        ''' </summary>
        <ComVisible(True)>
        Public Sub SetAutoExecutionBypass(enabled As Boolean)
            Global.ExcelAIVSTOAddIn.Security.AutoExecutionSettings.BypassConfirmation = enabled
            System.Diagnostics.Debug.WriteLine($"[ExcelBridge] 自动化执行免确认已切换: {enabled}")
        End Sub

        ''' <summary>
        ''' 获取 UI 文件夹路径（供前端构造独立页面 file:// URL）
        ''' </summary>
        Public Function GetUIPath() As String
            Try
                Dim assemblyDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                Dim uiFolder = Path.Combine(assemblyDir, "UI")
                Return If(Directory.Exists(uiFolder), uiFolder, "")
            Catch ex As Exception
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' 获取当前宿主（Excel 正打开）的工作簿路径信息。
        ''' 返回 JSON: {success, hasHost, path, directory, name}
        '''   - hasHost:  是否存在已保存的宿主工作簿（未保存的新建簿 hasHost=false）
        '''   - path:     工作簿完整路径（不存在则空）
        '''   - directory:工作簿所在目录（不存在则空）
        '''   - name:     工作簿名称（不存在时使用 ActiveWorkbook.Name）
        ''' </summary>
        <ComVisible(True)>
        Public Function GetHostWorkbook() As String
            Try
                Dim app = Globals.ThisAddIn?.Application
                Dim wb As Workbook = If(app IsNot Nothing, app.ActiveWorkbook, Nothing)
                If wb Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = True,
                        .hasHost = False,
                        .path = "",
                        .directory = "",
                        .name = ""
                    })
                End If

                Dim wbName As String = ""
                Try : wbName = wb.Name : Catch : End Try

                Dim fullPath As String = ""
                Try : fullPath = wb.FullName : Catch : End Try

                ' 新建未保存的簿 FullName 等于 Name（不含路径分隔符），视为无宿主
                Dim isSaved As Boolean = False
                Try : isSaved = Not String.IsNullOrEmpty(fullPath) AndAlso fullPath.Contains(Path.DirectorySeparatorChar.ToString()) : Catch : End Try

                Dim dir As String = ""
                If isSaved Then
                    Try : dir = Path.GetDirectoryName(fullPath) : Catch : End Try
                End If

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .hasHost = isSaved,
                    .path = If(isSaved, fullPath, ""),
                    .directory = If(dir, ""),
                    .name = wbName
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message,
                    .hasHost = False,
                    .path = "",
                    .directory = "",
                    .name = ""
                })
            End Try
        End Function

    End Class

End Namespace