Imports System
Imports System.Collections.Concurrent
Imports System.Globalization
Imports System.IO
Imports Newtonsoft.Json.Linq

Namespace Utils

    ''' <summary>
    ''' 国际化服务 - 提供多语言支持
    ''' 支持中文（默认）、英文、日文等语言切换
    ''' </summary>
    Public Class LocalizationService

        ' 单例实例
        Private Shared ReadOnly _lazyInstance As New Lazy(Of LocalizationService)(
            Function() New LocalizationService(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        ' 当前语言
        Private _currentLanguage As String = "zh-CN"

        ' 语言资源缓存
        Private ReadOnly _resources As New ConcurrentDictionary(Of String, JObject)

        ' 支持的语言列表
        Private ReadOnly _supportedLanguages As New List(Of String) From {
            "zh-CN", "en-US", "ja-JP"
        }

        ' 语言资源目录
        Private ReadOnly _resourcesFolder As String

        ' 日志服务
        Private ReadOnly _logger As ILogService

        ''' <summary>
        ''' 获取单例实例
        ''' </summary>
        Public Shared ReadOnly Property Instance As LocalizationService
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        ''' <summary>
        ''' 当前语言代码
        ''' </summary>
        Public Property CurrentLanguage As String
            Get
                Return _currentLanguage
            End Get
            Set(value As String)
                If _supportedLanguages.Contains(value) AndAlso value <> _currentLanguage Then
                    _currentLanguage = value
                    LoadLanguageResources(value)
                    _logger?.LogInfo("LOCALE", String.Format("语言已切换为: {0}", value))
                End If
            End Set
        End Property

        ''' <summary>
        ''' 支持的语言列表
        ''' </summary>
        Public ReadOnly Property SupportedLanguages As List(Of String)
            Get
                Return _supportedLanguages
            End Get
        End Property

        ''' <summary>
        ''' 私有构造函数
        ''' </summary>
        Private Sub New()
            _logger = LogServiceManager.Instance

            ' 设置资源目录
            _resourcesFolder = Path.Combine(
                GetAssemblyDirectory(),
                "Config",
                "Locales")

            ' 确保目录存在
            If Not Directory.Exists(_resourcesFolder) Then
                Directory.CreateDirectory(_resourcesFolder)
                CreateDefaultResources()
            End If

            ' 加载默认语言
            LoadLanguageResources(_currentLanguage)

            ' 尝试从系统设置获取语言
            TryDetectSystemLanguage()
        End Sub

        ''' <summary>
        ''' 获取程序集目录
        ''' </summary>
        Private Function GetAssemblyDirectory() As String
            Try
                Dim codeBase = System.Reflection.Assembly.GetExecutingAssembly().CodeBase
                Dim uri = New Uri(codeBase)
                Return Path.GetDirectoryName(uri.LocalPath)
            Catch ex As Exception  ' 
                Return Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            End Try
        End Function

        ''' <summary>
        ''' 尝试检测系统语言
        ''' </summary>
        Private Sub TryDetectSystemLanguage()
            Try
                Dim systemLang = CultureInfo.CurrentUICulture.Name
                If _supportedLanguages.Contains(systemLang) Then
                    _currentLanguage = systemLang
                    LoadLanguageResources(systemLang)
                End If
            Catch ex As Exception  ' 
            End Try
        End Sub

        ''' <summary>
        ''' 加载语言资源
        ''' </summary>
        Private Sub LoadLanguageResources(language As String)
            Try
                Dim resourcePath = Path.Combine(_resourcesFolder, String.Format("{0}.json", language))

                If File.Exists(resourcePath) Then
                    Dim json = File.ReadAllText(resourcePath, System.Text.Encoding.UTF8)
                    Dim jObj = JObject.Parse(json)
                    _resources.AddOrUpdate(language, jObj, Function(k, v) jObj)
                    _logger?.LogInfo("LOCALE", String.Format("语言资源已加载: {0}", language))
                Else
                    _logger?.LogWarning("LOCALE", String.Format("语言资源文件不存在: {0}", resourcePath))
                End If
            Catch ex As Exception
                _logger?.LogError("LOCALE", String.Format("加载语言资源失败: {0}", language), ex)
            End Try
        End Sub

        ''' <summary>
        ''' 获取本地化字符串
        ''' </summary>
        ''' <param name="key">资源键（格式：category.key）</param>
        ''' <param name="defaultValue">默认值（当资源不存在时返回）</param>
        ''' <returns>本地化字符串</returns>
        Public Function GetString(key As String, Optional defaultValue As String = "") As String
            If String.IsNullOrEmpty(key) Then Return defaultValue

            Dim resource = _resources.GetOrAdd(_currentLanguage, Function(k) New JObject())

            ' 支持嵌套键（如 "ui.buttons.submit"）
            Dim parts = key.Split("."c)
            Dim current = resource

            For i = 0 To parts.Length - 2
                Dim child = current(parts(i))
                If child Is Nothing Then Return defaultValue
                current = DirectCast(child, JObject)
            Next

            Dim value = current(parts.Last())
            If value Is Nothing Then Return defaultValue

            Return value.ToString()
        End Function

        ''' <summary>
        ''' 获取本地化字符串（带参数替换）
        ''' </summary>
        ''' <param name="key">资源键</param>
        ''' <param name="args">替换参数</param>
        ''' <param name="defaultValue">默认值</param>
        ''' <returns>格式化后的本地化字符串</returns>
        Public Function GetString(key As String, args As Object(), Optional defaultValue As String = "") As String
            Dim template = GetString(key, defaultValue)
            If String.IsNullOrEmpty(template) Then Return defaultValue

            Try
                Return String.Format(template, args)
            Catch ex As Exception  ' 
                Return template
            End Try
        End Function

        ''' <summary>
        ''' 获取所有本地化字符串（用于前端）
        ''' </summary>
        Public Function GetAllStrings() As String
            Dim resource = _resources.GetOrAdd(_currentLanguage, Function(k) New JObject())
            Return resource.ToString()
        End Function

        ''' <summary>
        ''' 创建默认资源文件
        ''' </summary>
        Private Sub CreateDefaultResources()
            ' 中文资源
            Dim zhCN = New JObject()
            zhCN("ui") = New JObject From {
                {"title", "SmartPilot 智航"},
                {"buttons", New JObject From {
                    {"submit", "发送"},
                    {"cancel", "取消"},
                    {"clear", "清空"},
                    {"settings", "设置"}
                }},
                {"labels", New JObject From {
                    {"inputPlaceholder", "输入您的问题或指令..."},
                    {"loading", "正在加载..."},
                    {"error", "发生错误"}
                }}
            }
            zhCN("ocr") = New JObject From {
                {"title", "识图入表"},
                {"buttons", New JObject From {
                    {"upload", "上传图片"},
                    {"capture", "截图识别"},
                    {"paste", "粘贴图片"}
                }},
                {"labels", New JObject From {
                    {"result", "识别结果"},
                    {"copy", "复制结果"},
                    {"insert", "插入到单元格"}
                }}
            }
            zhCN("settings") = New JObject From {
                {"title", "模型配置"},
                {"tabs", New JObject From {
                    {"llm", "LLM 配置"},
                    {"ocr", "OCR 配置"},
                    {"mcp", "MCP 配置"},
                    {"rag", "RAG 配置"}
                }}
            }
            zhCN("errors") = New JObject From {
                {"networkError", "网络连接失败"},
                {"timeoutError", "请求超时"},
                {"unknownError", "未知错误"}
            }

            File.WriteAllText(
                Path.Combine(_resourcesFolder, "zh-CN.json"),
                zhCN.ToString(),
                System.Text.Encoding.UTF8)

            ' 英文资源
            Dim enUS = New JObject()
            enUS("ui") = New JObject From {
                {"title", "SmartPilot"},
                {"buttons", New JObject From {
                    {"submit", "Send"},
                    {"cancel", "Cancel"},
                    {"clear", "Clear"},
                    {"settings", "Settings"}
                }},
                {"labels", New JObject From {
                    {"inputPlaceholder", "Enter your question or instruction..."},
                    {"loading", "Loading..."},
                    {"error", "Error occurred"}
                }}
            }
            enUS("ocr") = New JObject From {
                {"title", "Image Recognition"},
                {"buttons", New JObject From {
                    {"upload", "Upload Image"},
                    {"capture", "Screenshot"},
                    {"paste", "Paste Image"}
                }},
                {"labels", New JObject From {
                    {"result", "Recognition Result"},
                    {"copy", "Copy Result"},
                    {"insert", "Insert to Cell"}
                }}
            }
            enUS("settings") = New JObject From {
                {"title", "Model Configuration"},
                {"tabs", New JObject From {
                    {"llm", "LLM Config"},
                    {"ocr", "OCR Config"},
                    {"mcp", "MCP Config"},
                    {"rag", "RAG Config"}
                }}
            }
            enUS("errors") = New JObject From {
                {"networkError", "Network connection failed"},
                {"timeoutError", "Request timeout"},
                {"unknownError", "Unknown error"}
            }

            File.WriteAllText(
                Path.Combine(_resourcesFolder, "en-US.json"),
                enUS.ToString(),
                System.Text.Encoding.UTF8)

            _logger?.LogInfo("LOCALE", "默认语言资源文件已创建")
        End Sub

    End Class

End Namespace