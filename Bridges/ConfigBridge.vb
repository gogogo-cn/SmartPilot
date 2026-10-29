Option Explicit On
Option Infer On

Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Diagnostics
Imports System.Security.AccessControl
Imports System.Security.Principal
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ExcelAIVSTOAddIn.Utils

Namespace Bridges

    ''' <summary>
    ''' 配置管理桥接类 — 从 ExcelBridge 中提取的配置 API
    ''' </summary>
    <ComVisible(True)>
    <ClassInterface(ClassInterfaceType.AutoDispatch)>
    Public Class ConfigBridge
        Implements IBridge

        ' 单例实例 - 使用 Lazy(Of T) 确保线程安全
        Private Shared ReadOnly _lazyInstance As New Lazy(Of ConfigBridge)(Function() New ConfigBridge(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

#Region "单例模式"

        ''' <summary>
        ''' 获取 ConfigBridge 单例实例（线程安全 Lazy 初始化）
        ''' </summary>
        Public Shared ReadOnly Property Instance As ConfigBridge
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        ''' <summary>
        ''' 私有构造函数（单例模式）
        ''' </summary>
        Private Sub New()
        End Sub

#End Region

        ''' <summary>
        ''' 获取当前设置
        ''' </summary>
        <ComVisible(True)>
        Public Function GetSettings() As String
            Try
                Dim config = ConfigurationManager.Instance
                If config Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.error = "ConfigurationManager 未初始化"})
                End If

                Return JsonConvert.SerializeObject(config.GetAllSettings())
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 保存设置
        ''' </summary>
        <ComVisible(True)>
        Public Function SaveSettings(settingsJson As String) As String
            Try
                Dim config = ConfigurationManager.Instance
                If config Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "ConfigurationManager 未初始化"})
                End If

                Dim settings = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(settingsJson)
                config.SaveSettings(settings)

                Return JsonConvert.SerializeObject(New With {.success = True})
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 获取 LLM 配置（返回脱敏后的 API Key 给前端）
        ''' </summary>
        <ComVisible(True)>
        Public Function GetLLMConfig() As String
            Try
                Dim config = ConfigurationManager.Instance
                If config Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.error = "ConfigurationManager 未初始化"})
                End If

                Dim apiKey As String = ReadLLMConfig(config, Utils.ConfigConstants.LLM.API_KEY, "LLM_API_KEY", "")

                Dim llmConfig = New With {
                    .provider = ReadLLMConfig(config, Utils.ConfigConstants.LLM.PROVIDER, "LLM_PROVIDER", "ollama"),
                    .model = ReadLLMConfig(config, Utils.ConfigConstants.LLM.MODEL, "LLM_MODEL", ""),
                    .apiKey = MaskSensitiveValue(apiKey),
                    .apiUrl = ReadLLMConfig(config, Utils.ConfigConstants.LLM.API_URL, "LLM_API_URL", ""),
                    .hasApiKey = Not String.IsNullOrEmpty(apiKey)
                }
                Return JsonConvert.SerializeObject(llmConfig)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 保存 LLM 配置
        ''' </summary>
        <ComVisible(True)>
        Public Function SaveLLMConfig(configJson As String) As String
            Try
                Dim config = ConfigurationManager.Instance
                If config Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "ConfigurationManager 未初始化"})
                End If

                Dim settings = JsonConvert.DeserializeObject(Of Dictionary(Of String, String))(configJson)
                
                ' 兼容 apiType 和 provider 两种字段名
                Dim provider As String = Nothing
                If settings.ContainsKey("provider") Then
                    provider = settings("provider")
                ElseIf settings.ContainsKey("apiType") Then
                    provider = settings("apiType")
                End If
                If provider IsNot Nothing Then
                    config.SetSetting(Utils.ConfigConstants.LLM.PROVIDER, provider)
                End If

                If settings.ContainsKey("model") Then
                    config.SetSetting(Utils.ConfigConstants.LLM.MODEL, settings("model"))
                End If
                If settings.ContainsKey("apiKey") Then
                    config.SetSetting(Utils.ConfigConstants.LLM.API_KEY, settings("apiKey"))
                End If
                If settings.ContainsKey("apiUrl") Then
                    config.SetSetting(Utils.ConfigConstants.LLM.API_URL, settings("apiUrl"))
                End If

                Return JsonConvert.SerializeObject(New With {.success = True})
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 保存 AI 配置
        ''' </summary>
        <ComVisible(True)>
        Public Function SaveAIConfig(configJson As String) As String
            Try
                Dim saved = SaveLLMConfig(configJson)
                ' 【独立执行】配置保存成功后，导出明文 LLM 配置供独立服务（Excel 关闭后）读取
                Try
                    ExportLLMConfigForStandalone()
                Catch exp As Exception
                    System.Diagnostics.Debug.WriteLine("[ConfigBridge] ExportLLMConfigForStandalone 失败: " & exp.Message)
                End Try
                Return saved
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 获取 AI 配置（返回完整未脱敏配置，供聊天面板读取）
        ''' </summary>
        <ComVisible(True)>
        Public Function GetAIConfig() As String
            Try
                Dim config = ConfigurationManager.Instance
                If config Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.error = "ConfigurationManager 未初始化"})
                End If

                ' 🆕 首次访问时如果内部为空，尝试从独立服务写入的 llm_config.json 反向同步
                Try
                    ImportLLMConfigIfEmpty()
                Catch : End Try

                Dim provider = ReadLLMConfig(config, Utils.ConfigConstants.LLM.PROVIDER, "LLM_PROVIDER", "ollama")
                Dim apiKey = ReadLLMConfig(config, Utils.ConfigConstants.LLM.API_KEY, "LLM_API_KEY", "")

                If provider.Equals("ollama", StringComparison.OrdinalIgnoreCase) AndAlso String.IsNullOrEmpty(apiKey) Then
                    apiKey = "local"
                End If

                Dim aiConfig = New With {
                    .provider = provider,
                    .model = ReadLLMConfig(config, Utils.ConfigConstants.LLM.MODEL, "LLM_MODEL", ""),
                    .apiKey = apiKey,
                    .apiUrl = ReadLLMConfig(config, Utils.ConfigConstants.LLM.API_URL, "LLM_API_URL", "")
                }
                Return JsonConvert.SerializeObject(aiConfig)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 反向同步：仅当 ConfigurationManager 里 LLM 核心字段（provider/model）都为空时，
        ''' 从磁盘权威位置（ConfigDirectory\llm_config.json）回灌。
        ''' 目的：用户在独立任务列表页保存的配置（通过 standalone_agent /save-config 写入磁盘），
        ''' 下次启动 Excel 时会自动导入到 VSTO 内部存储，实现「两端双向共用」。
        ''' 注意：若 .NET 内部已有任意 provider 或 model 值，视为用户已在 Excel 内显式配置 → 不覆盖。
        ''' </summary>
        <ComVisible(True)>
        Public Function ImportLLMConfigIfEmpty() As String
            Try
                Dim config = ConfigurationManager.Instance
                If config Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .skipped = True, .reason = "ConfigurationManager 未初始化"})
                End If
                ' 仅当两边核心字段都为空才回灌
                Dim curProvider = ReadLLMConfig(config, Utils.ConfigConstants.LLM.PROVIDER, "LLM_PROVIDER", "")
                Dim curModel = ReadLLMConfig(config, Utils.ConfigConstants.LLM.MODEL, "LLM_MODEL", "")
                If Not String.IsNullOrWhiteSpace(curProvider) OrElse Not String.IsNullOrWhiteSpace(curModel) Then
                    Return JsonConvert.SerializeObject(New With {.success = True, .skipped = True, .reason = "内部已有配置，不覆盖（provider=" & curProvider & ", model=" & curModel & "）"})
                End If
                Dim outDir = PathHelper.ConfigDirectory
                If String.IsNullOrWhiteSpace(outDir) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .skipped = True, .reason = "ConfigDirectory 为空"})
                End If
                Dim cfgPath = Path.Combine(outDir, "llm_config.json")
                If Not File.Exists(cfgPath) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .skipped = True, .reason = "磁盘没有 llm_config.json"})
                End If
                Dim json As String = File.ReadAllText(cfgPath, System.Text.Encoding.UTF8)
                If String.IsNullOrWhiteSpace(json) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .skipped = True, .reason = "llm_config.json 为空"})
                End If
                Dim j = JObject.Parse(json)
                Dim provider As String = If(j("provider")?.ToString(), If(j("apiType")?.ToString(), ""))
                Dim model As String = If(j("model")?.ToString(), "")
                Dim apiKey As String = If(j("apiKey")?.ToString(), "")
                Dim apiUrl As String = If(j("apiUrl")?.ToString(), If(j("baseUrl")?.ToString(), ""))
                Dim importedAny As Boolean = False
                If Not String.IsNullOrWhiteSpace(provider) Then
                    config.SetSetting(Utils.ConfigConstants.LLM.PROVIDER, provider)
                    importedAny = True
                End If
                If Not String.IsNullOrWhiteSpace(model) Then
                    config.SetSetting(Utils.ConfigConstants.LLM.MODEL, model)
                    importedAny = True
                End If
                If Not String.IsNullOrWhiteSpace(apiKey) AndAlso Not String.Equals(apiKey, "local", StringComparison.OrdinalIgnoreCase) Then
                    config.SetSetting(Utils.ConfigConstants.LLM.API_KEY, apiKey)
                    importedAny = True
                End If
                If Not String.IsNullOrWhiteSpace(apiUrl) Then
                    config.SetSetting(Utils.ConfigConstants.LLM.API_URL, apiUrl)
                    importedAny = True
                End If
                If importedAny Then config.FlushSave()
                Return JsonConvert.SerializeObject(New With {
                    .success = importedAny,
                    .imported = importedAny,
                    .path = cfgPath,
                    .provider = provider,
                    .model = model,
                    .message = If(importedAny, "已从 llm_config.json 反向同步进 VSTO 内部存储", "llm_config.json 里没读到可导入的 provider/model")
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 把 llm_config.json 的 NTFS 权限收紧为当前用户，避免同机其他账户读到 API 密钥。
        ''' 失败时不阻断导出：独立服务仍需能读到文件。
        ''' </summary>
        Private Shared Sub RestrictConfigFileToCurrentUser(filePath As String)
            Try
                If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then Return
                Dim fi As New FileInfo(filePath)
                Dim sec = fi.GetAccessControl()
                sec.SetAccessRuleProtection(True, False)
                Dim rule As New FileSystemAccessRule(
                    WindowsIdentity.GetCurrent().User,
                    FileSystemRights.FullControl,
                    InheritanceFlags.None,
                    PropagationFlags.None,
                    AccessControlType.Allow)
                sec.AddAccessRule(rule)
                fi.SetAccessControl(sec)
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' 导出明文 LLM 配置到独立服务可读的 llm_config.json（多路径双写 = 同一 SSOT 多副本）
        ''' 供 Excel 关闭后运行的独立执行服务读取（复用 VSTO 配置好的模型 KEY）。
        ''' 注意：此文件含 API Key，仅写入本机，并把 NTFS 权限收紧为当前用户。禁止打进安装包或提交到仓库。
        ''' 写入策略（保证无论通过 MobileControlBridge --config、find_config() 还是 /save-config 都能读到同一份内容）：
        '''   ① PathHelper.ConfigDirectory/llm_config.json（= bin\Debug|Release\Config，MobileControlBridge 启动时 --config 必传）
        '''   ② %APPDATA%\ExcelAIVSTOAddIn\Config\llm_config.json（独立服务的 ssot_write_path 首选，也是 find_config 权威首候选）
        '''   ③ 同步到另一种编译配置目录（Debug/Release 另一目录存在时也写一份，切构建不丢配置）
        ''' </summary>
        <ComVisible(True)>
        Public Function ExportLLMConfigForStandalone() As String
            Try
                Dim config = ConfigurationManager.Instance
                If config Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "ConfigurationManager 未初始化"})
                End If

                ' 🆕 先做反向同步：如果内部配置是空的但磁盘上有（独立页里填过并保存），先回灌到 .NET，
                '    避免把空值写出去覆盖了用户已有的配置。
                Try
                    ImportLLMConfigIfEmpty()
                Catch : End Try

                Dim provider = ReadLLMConfig(config, Utils.ConfigConstants.LLM.PROVIDER, "LLM_PROVIDER", "")
                Dim model = ReadLLMConfig(config, Utils.ConfigConstants.LLM.MODEL, "LLM_MODEL", "")
                Dim apiKey = ReadLLMConfig(config, Utils.ConfigConstants.LLM.API_KEY, "LLM_API_KEY", "")
                Dim apiUrl = ReadLLMConfig(config, Utils.ConfigConstants.LLM.API_URL, "LLM_API_URL", "")

                If provider.Equals("ollama", StringComparison.OrdinalIgnoreCase) AndAlso String.IsNullOrEmpty(apiKey) Then
                    apiKey = "local"
                End If

                Dim payload = New With {
                    .provider = provider,
                    .apiType = provider,
                    .model = model,
                    .apiUrl = apiUrl,
                    .baseUrl = apiUrl,
                    .apiKey = apiKey
                }

                Dim json = JsonConvert.SerializeObject(payload, Formatting.Indented)

                Dim outDir = PathHelper.ConfigDirectory
                If String.IsNullOrWhiteSpace(outDir) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "ConfigDirectory 为空（PathHelper 未初始化）"})
                End If
                If Not Directory.Exists(outDir) Then Directory.CreateDirectory(outDir)

                ' ── 收集所有目标路径，并多写一份（保持一致，避免"不同入口读到不同配置"）──
                Dim written As New List(Of String)()
                Dim primaryPath = Path.Combine(outDir, "llm_config.json")
                File.WriteAllText(primaryPath, json, New System.Text.UTF8Encoding(False))
                RestrictConfigFileToCurrentUser(primaryPath)
                written.Add(primaryPath)

                ' ② %APPDATA%\ExcelAIVSTOAddIn\Config\llm_config.json（独立服务 /save-config 首选写入位置，也是 find_config 首候选）
                Try
                    Dim appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
                    If Not String.IsNullOrWhiteSpace(appData) Then
                        Dim appDataDir = Path.Combine(appData, "ExcelAIVSTOAddIn", "Config")
                        If Not Directory.Exists(appDataDir) Then Directory.CreateDirectory(appDataDir)
                        Dim appDataPath = Path.Combine(appDataDir, "llm_config.json")
                        If Not String.Equals(appDataPath, primaryPath, StringComparison.OrdinalIgnoreCase) Then
                            File.WriteAllText(appDataPath, json, New System.Text.UTF8Encoding(False))
                            RestrictConfigFileToCurrentUser(appDataPath)
                            written.Add(appDataPath)
                        End If
                    End If
                Catch : End Try

                ' ③ 另一构建目录（例如当前是 Debug 写了，Release 目录也同步一份；反之亦然）
                Try
                    Dim asmDir = Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly().Location)
                    If Not String.IsNullOrWhiteSpace(asmDir) Then
                        Dim parentDir = Directory.GetParent(asmDir)?.FullName
                        If Not String.IsNullOrWhiteSpace(parentDir) Then
                            For Each cfg In {"Debug", "Release"}
                                Dim siblingDir = Path.Combine(parentDir, cfg, "Config")
                                If Directory.Exists(siblingDir) Then
                                    Dim siblingPath = Path.Combine(siblingDir, "llm_config.json")
                                    Dim alreadyWritten As Boolean = False
                                    For Each w In written
                                        If String.Equals(w, siblingPath, StringComparison.OrdinalIgnoreCase) Then
                                            alreadyWritten = True : Exit For
                                        End If
                                    Next
                                    If Not alreadyWritten Then
                                        File.WriteAllText(siblingPath, json, New System.Text.UTF8Encoding(False))
                                        RestrictConfigFileToCurrentUser(siblingPath)
                                        written.Add(siblingPath)
                                    End If
                                End If
                            Next
                        End If
                    End If
                Catch : End Try

                ' 清理历史遗留的冗余副本（如果有），避免"多个 llm_config.json"造成困惑与过期副本误导
                Dim obsolete As New List(Of String)
                Try
                    Dim srcRoot = Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly().Location)
                    obsolete.Add(Path.Combine(srcRoot, "StandaloneAgent", "llm_config.json"))
                    obsolete.Add(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "StandaloneAgent", "llm_config.json"))
                    Try
                        Dim devRoot = Directory.GetParent(srcRoot).Parent.FullName
                        obsolete.Add(Path.Combine(devRoot, "StandaloneAgent", "llm_config.json"))
                    Catch : End Try
                Catch : End Try
                For Each p In obsolete
                    Try
                        Dim isWrittenPath As Boolean = False
                        For Each w In written
                            If String.Equals(w, p, StringComparison.OrdinalIgnoreCase) Then
                                isWrittenPath = True : Exit For
                            End If
                        Next
                        If File.Exists(p) AndAlso Not isWrittenPath Then
                            File.Delete(p)
                        End If
                    Catch : End Try
                Next

                Dim msg = "[ConfigBridge] LLM 配置（多路径双写 SSOT）已导出到 " & written.Count & " 处：" & String.Join(" | ", written)
                System.Diagnostics.Debug.WriteLine(msg)
                Return JsonConvert.SerializeObject(New With {.success = True, .path = primaryPath, .allPaths = written})
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function
        <ComVisible(True)>
        Public Function GetOcrConfig() As String
            Try
                Dim config = ConfigurationManager.Instance
                If config Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.error = "ConfigurationManager 未初始化"})
                End If

                Dim apiKey As String = config.GetSetting("OCR_API_KEY", "")
                Dim secretKey As String = config.GetSetting("OCR_SECRET_KEY", "")

                Dim ocrConfig = New With {
                    .engine = config.GetSetting("OCR_ENGINE", "paddleocr"),
                    .apiKey = MaskSensitiveValue(apiKey),
                    .secretKey = MaskSensitiveValue(secretKey),
                    .hasApiKey = Not String.IsNullOrEmpty(apiKey),
                    .hasSecretKey = Not String.IsNullOrEmpty(secretKey)
                }
                Return JsonConvert.SerializeObject(ocrConfig)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 保存 OCR 配置
        ''' </summary>
        <ComVisible(True)>
        Public Function SaveOcrConfig(apiKey As String, secretKey As String) As String
            Try
                Dim config = ConfigurationManager.Instance
                If config Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "ConfigurationManager 未初始化"})
                End If

                config.SetSetting("OCR_API_KEY", apiKey)
                config.SetSetting("OCR_SECRET_KEY", secretKey)

                Return JsonConvert.SerializeObject(New With {.success = True})
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 获取云端 OCR 配置
        ''' </summary>
        <ComVisible(True)>
        Public Function GetCloudOcrConfig() As String
            Try
                Return GetOcrConfig()
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 保存云端 OCR 配置
        ''' </summary>
        <ComVisible(True)>
        Public Sub SaveCloudOcrConfig(configJson As String)
            Try
                Dim config = ConfigurationManager.Instance
                If config Is Nothing Then Return
                If String.IsNullOrWhiteSpace(configJson) Then Return

                Debug.WriteLine("[SaveCloudOcrConfig] 保存云端 OCR 配置")
                Dim settings = JsonConvert.DeserializeObject(Of Dictionary(Of String, String))(configJson)

                If settings.ContainsKey("provider") Then
                    config.SetSetting("OCR_CLOUD_PROVIDER", settings("provider"))
                End If
                If settings.ContainsKey("apiKey") Then
                    config.SetSetting("OCR_CLOUD_API_KEY", settings("apiKey"))
                End If
                If settings.ContainsKey("secretKey") Then
                    config.SetSetting("OCR_CLOUD_SECRET_KEY", settings("secretKey"))
                End If
                ' 腾讯云额外字段 (P1修复: 补全 secretId 支持)
                If settings.ContainsKey("secretId") Then
                    config.SetSetting("OCR_TENCENT_SECRET_ID", settings("secretId"))
                End If
                config.FlushSave()

                ' 同步到 CloudOcrService 配置文件（用于实际 OCR 调用）
                Try
                    Dim cloudConfig = Services.CloudOcrConfig.Load()
                    If settings.ContainsKey("provider") Then cloudConfig.Provider = settings("provider")
                    If settings.ContainsKey("apiKey") Then cloudConfig.BaiduApiKey = settings("apiKey")
                    If settings.ContainsKey("secretKey") Then cloudConfig.BaiduSecretKey = settings("secretKey")
                    If settings.ContainsKey("secretId") Then cloudConfig.TencentSecretId = settings("secretId")
                    cloudConfig.Save()
                Catch ocrEx As Exception
                    Debug.WriteLine($"[SaveCloudOcrConfig] CloudOcrConfig 同步失败: {ocrEx.Message}")
                End Try

                ' 重载 CloudOcrService 使配置立即生效
                Try
                    Services.OcrServiceManager.Instance.ReloadCloudOcrConfig()
                Catch reloadEx As Exception
                    Debug.WriteLine($"[SaveCloudOcrConfig] OcrServiceManager 重载失败: {reloadEx.Message}")
                End Try

                Debug.WriteLine("[SaveCloudOcrConfig] 配置已保存")
            Catch ex As Exception
                Debug.WriteLine($"[SaveCloudOcrConfig] 保存失败: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' 获取 OCR 状态（兼容方法）
        ''' </summary>
        <ComVisible(True)>
        Public Function GetOcrStatus() As String
            Try
                Dim ocrMgr = Services.OcrServiceManager.Instance
                Dim currentEngine As Services.OcrEngineType = Services.OcrEngineType.RapidOCR
                If ocrMgr IsNot Nothing Then currentEngine = ocrMgr.CurrentEngine
                Dim available = If(ocrMgr IsNot Nothing, ocrMgr.IsEngineAvailable(currentEngine), False)
                Return JsonConvert.SerializeObject(New With {
                    .status = If(available, "ready", "not_installed"),
                    .engine = Services.OcrServiceManager.GetEngineDisplayName(currentEngine),
                    .engineKey = currentEngine.ToString(),
                    .isAvailable = available,
                    .currentVersion = Nothing,
                    .needsUpgrade = False,
                    .recommendedVersion = "latest"
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .status = "error",
                    .engine = "",
                    .isAvailable = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 获取 OCR 引擎状态
        ''' </summary>
        <ComVisible(True)>
        Public Function GetOcrEngineStatus(engine As String) As String
            Try
                Dim config = ConfigurationManager.Instance
                Select Case engine.ToLower()
                    Case "paddleocr"
                        ' 检查 PaddleOCR 是否安装
                        Dim paddleStatus = CheckPaddleOcrInstallation()
                        Return JsonConvert.SerializeObject(paddleStatus)
                    Case "cloud"
                        ' 检查云端 OCR 是否已配置 API Key
                        Dim cloudApiKey As String = If(config IsNot Nothing, config.GetSetting("OCR_CLOUD_API_KEY", ""), "")
                        Dim hasCloudConfig As Boolean = Not String.IsNullOrEmpty(cloudApiKey)
                        Return JsonConvert.SerializeObject(New With {
                            .available = hasCloudConfig,
                            .engine = engine,
                            .message = If(hasCloudConfig, "云端 OCR 已配置", "云端 OCR 未配置，请在设置中填写 API Key")
                        })
                    Case Else
                        Return JsonConvert.SerializeObject(New With {
                            .available = False,
                            .engine = engine,
                            .message = "未知引擎类型"
                        })
                End Select
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 测试云端 OCR 配置（验证 API Key 有效性，通过检查引擎可用性）
        ''' </summary>
        <ComVisible(True)>
        Public Function TestCloudOcrConfig(provider As String, configJson As String) As String
            Try
                If String.IsNullOrWhiteSpace(provider) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "未指定 OCR 提供商"})
                End If
                If String.IsNullOrWhiteSpace(configJson) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "未提供配置信息"})
                End If

                Debug.WriteLine($"[TestCloudOcrConfig] 测试提供商: {provider}")
                Dim config = JsonConvert.DeserializeObject(Of Dictionary(Of String, String))(configJson)

                ' 使用 OcrServiceManager 检查云端 OCR 引擎是否可配置/可用
                Dim ocrManager As Services.OcrServiceManager = Services.OcrServiceManager.Instance

                ' 验证必填字段是否存在（API Key 是云端 OCR 的核心）
                If Not config.ContainsKey("apiKey") OrElse String.IsNullOrWhiteSpace(config("apiKey")) Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = "API Key 未配置或为空"
                    })
                End If

                ' 保存测试配置并验证引擎状态
                Try
                    ocrManager.CurrentEngine = Services.OcrEngineType.CloudOCR
                    Dim isAvailable = ocrManager.IsEngineAvailable(Services.OcrEngineType.CloudOCR)
                    Return JsonConvert.SerializeObject(New With {
                        .success = isAvailable,
                        .message = If(isAvailable, $"与 {provider} 配置已生效", $"{provider} 连接不可用，请检查 API Key 和网络")
                    })
                Catch engineEx As Exception
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = $"配置验证失败: {engineEx.Message}"
                    })
                End Try
            Catch ex As Exception
                Debug.WriteLine($"[TestCloudOcrConfig] 错误: {ex.Message}")
                Return JsonConvert.SerializeObject(New With {.success = False, .error = $"测试失败: {ex.Message}"})
            End Try
        End Function

        ''' <summary>
        ''' 获取当前 OCR 引擎
        ''' </summary>
        <ComVisible(True)>
        Public Function GetCurrentOcrEngine() As String
            Try
                ' 以实际运行引擎（OcrServiceManager.CurrentEngine）为权威，避免"配置默认值"与"实际引擎"不符
                Dim ocrMgr = Services.OcrServiceManager.Instance
                If ocrMgr IsNot Nothing Then
                    Return EngineKeyFromType(ocrMgr.CurrentEngine)
                End If
                Dim config = ConfigurationManager.Instance
                If config IsNot Nothing Then
                    Dim cfgEngine = config.GetSetting("OCR_ENGINE", "")
                    If Not String.IsNullOrWhiteSpace(cfgEngine) Then Return cfgEngine
                End If
                Return "rapidocr"
            Catch ex As Exception
                Return "rapidocr"
            End Try
        End Function

        ''' <summary>
        ''' 设置 OCR 引擎
        ''' </summary>
        <ComVisible(True)>
        Public Sub SetOcrEngine(engine As String)
            Try
                If String.IsNullOrWhiteSpace(engine) Then Return

                Dim config = ConfigurationManager.Instance
                If config IsNot Nothing Then
                    config.SetSetting("OCR_ENGINE", engine)
                    config.FlushSave()
                End If

                ' 同步到 OcrServiceManager，使引擎切换立即生效
                Try
                    Dim ocrMgr = Services.OcrServiceManager.Instance
                    If ocrMgr IsNot Nothing Then
                        Select Case engine.ToLowerInvariant()
                            Case "paddleocr", "paddle"
                                ocrMgr.CurrentEngine = Services.OcrEngineType.PaddleOCR
                            Case "rapidocr", "rapid"
                                ocrMgr.CurrentEngine = Services.OcrEngineType.RapidOCR
                            Case "cloudocr", "cloud", "azure", "google", "baidu", "tencent"
                                ' 所有云服务商统一使用 CloudOCR 引擎，具体提供商由 CloudOcrConfig.Provider 控制
                                ocrMgr.CurrentEngine = Services.OcrEngineType.CloudOCR
                            Case Else
                                Debug.WriteLine($"[SetOcrEngine] 未知引擎类型: {engine}")
                        End Select
                        Debug.WriteLine($"[SetOcrEngine] OCR引擎已切换为: {engine}")
                    End If
                Catch syncEx As Exception
                    Debug.WriteLine($"[SetOcrEngine] OcrServiceManager 同步失败: {syncEx.Message}")
                End Try
            Catch ex As Exception
                Debug.WriteLine($"[SetOcrEngine] 设置失败: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' 获取可用的 OCR 引擎列表
        ''' </summary>
        <ComVisible(True)>
        Public Function GetAvailableOcrEngines() As String
            Try
                ' 与 OcrEngineType 枚举 / SetOcrEngine 支持清单保持一致
                Dim engines = New List(Of String) From {"PaddleOCR", "RapidOCR", "CloudOCR"}
                Return JsonConvert.SerializeObject(engines)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 获取可用的 LLM 提供商列表
        ''' </summary>
        <ComVisible(True)>
        Public Function GetAvailableLLMProviders() As String
            Try
                Return JsonConvert.SerializeObject(ConfigConstants.LLM.AVAILABLE_PROVIDERS)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.error = ex.Message})
            End Try
        End Function

#Region "辅助方法"

        ''' <summary>
        ''' 敏感值脱敏 — 只显示前3后4字符，中间用***替代
        ''' 用于返回给前端的 API Key/Secret Key 等敏感信息
        ''' </summary>
        ''' <param name="value">原始值</param>
        ''' <returns>脱敏后的值，如 "sk-***f456"；空值返回空字符串</returns>
        Private Shared Function MaskSensitiveValue(value As String) As String
            If String.IsNullOrEmpty(value) Then
                Return String.Empty
            End If

            If value.Length <= 8 Then
                ' 太短的值：只显示首字符
                Return value.Substring(0, 1) & "***"
            End If

            ' 显示前3字符 + *** + 后4字符
            Return value.Substring(0, 3) & "***" & value.Substring(value.Length - 4)
        End Function

        ''' <summary>
        ''' 检查 PaddleOCR 安装状态（真实探测，不再返回"开发中"占位）
        ''' </summary>
        Private Function CheckPaddleOcrInstallation() As Object
            Try
                Dim ocrMgr = Services.OcrServiceManager.Instance
                Dim available As Boolean = If(ocrMgr IsNot Nothing, ocrMgr.IsEngineAvailable(Services.OcrEngineType.PaddleOCR), False)
                Return New With {
                    .available = available,
                    .engine = "paddleocr",
                    .message = If(available, "PaddleOCR 已安装并可用", "PaddleOCR 未就绪（需 Python + paddleocr 环境）")
                }
            Catch ex As Exception
                Return New With {
                    .available = False,
                    .engine = "paddleocr",
                    .error = ex.Message
                }
            End Try
        End Function

        ''' <summary>
        ''' 将 OcrEngineType 转换为前端约定的引擎 key（rapidocr/paddleocr/cloud）
        ''' </summary>
        Private Shared Function EngineKeyFromType(engine As Services.OcrEngineType) As String
            Select Case engine
                Case Services.OcrEngineType.RapidOCR
                    Return "rapidocr"
                Case Services.OcrEngineType.PaddleOCR
                    Return "paddleocr"
                Case Services.OcrEngineType.CloudOCR
                    Return "cloud"
                Case Else
                    Return "rapidocr"
            End Select
        End Function

#End Region

        ''' <summary>
        ''' 读取 LLM 配置，兼容旧版下划线键名（LLM_PROVIDER 等）和新版驼峰键名（LLMProvider 等）
        ''' </summary>
        Public Function ReadLLMConfig(config As ConfigurationManager, key As String, legacyKey As String, defaultValue As String) As String
            Dim value = config.GetSetting(key, "")
            If String.IsNullOrEmpty(value) Then
                value = config.GetSetting(legacyKey, defaultValue)
            End If
            Return value
        End Function

        ''' <summary>
        ''' 删除本机模型密钥、云 OCR 密钥，以及独立服务可读的 llm_config.json。
        ''' 不删除工作簿内容，也不清除账户登录令牌。
        ''' </summary>
        <ComVisible(True)>
        Public Function ClearLocalUserData() As String
            Try
                Dim removed As New List(Of String)()
                Dim config = ConfigurationManager.Instance
                If config IsNot Nothing Then
                    Dim snapshot = config.GetAllSettings()
                    For Each key In snapshot.Keys
                        Dim lower = If(key, "").ToLowerInvariant()
                        If lower.Contains("apikey") OrElse lower.Contains("secret") OrElse
                           lower.Contains("password") OrElse lower.Contains("privatekey") Then
                            config.RemoveSetting(key)
                            removed.Add("setting:" & key)
                        End If
                    Next
                    config.FlushSave()
                End If

                For Each path In LlmConfigFileCandidates()
                    Try
                        If File.Exists(path) Then
                            File.Delete(path)
                            removed.Add(path)
                        End If
                    Catch ex As Exception
                        Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message, .removed = removed})
                    End Try
                Next

                Dim hostClear = HostConfig.ClearStoredApiKey()
                Dim hostJo = Newtonsoft.Json.Linq.JObject.Parse(hostClear)
                If hostJo.Value(Of Boolean?)("success") <> True Then
                    Return hostClear
                End If

                Return JsonConvert.SerializeObject(New With {.success = True, .removed = removed})
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        Private Shared Function LlmConfigFileCandidates() As IEnumerable(Of String)
            Dim paths As New List(Of String)()
            Try
                Dim primary = Path.Combine(PathHelper.ConfigDirectory, "llm_config.json")
                paths.Add(primary)
            Catch
            End Try
            Try
                Dim appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
                If Not String.IsNullOrWhiteSpace(appData) Then
                    paths.Add(Path.Combine(appData, "ExcelAIVSTOAddIn", "Config", "llm_config.json"))
                End If
            Catch
            End Try
            Try
                Dim asmDir = Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly().Location)
                Dim parentDir = If(String.IsNullOrWhiteSpace(asmDir), Nothing, Directory.GetParent(asmDir)?.FullName)
                If Not String.IsNullOrWhiteSpace(parentDir) Then
                    paths.Add(Path.Combine(parentDir, "Debug", "Config", "llm_config.json"))
                    paths.Add(Path.Combine(parentDir, "Release", "Config", "llm_config.json"))
                End If
            Catch
            End Try
            Return paths
        End Function

#Region "IBridge 生命周期"

        Public ReadOnly Property BridgeName As String Implements IBridge.BridgeName
            Get
                Return "ConfigBridge"
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

#End Region

    End Class

End Namespace