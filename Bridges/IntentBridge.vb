Option Explicit On
Option Infer On

Imports System.Runtime.InteropServices
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Serialization
Imports ExcelAIVSTOAddIn.Utils

Namespace Bridges

    ''' <summary>
    ''' IntentBridge - 分层意图识别 API 桥梁类
    ''' 提供 COM 可见的意图识别方法供前端 JavaScript 调用
    ''' </summary>
    <ComVisible(True)>
    <ClassInterface(ClassInterfaceType.AutoDispatch)>
    <ProgId("ExcelAIVSTOAddIn.IntentBridge")>
    Public Class IntentBridge
        Implements IBridge

        ' 单例实例 - 使用 Lazy(Of T) 确保线程安全
        Private Shared ReadOnly _lazyInstance As New Lazy(Of IntentBridge)(Function() New IntentBridge(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        ' 意图识别引擎和规则管理器
        Private _intentEngine As IntentRecognitionEngine
        Private _rulesManager As RulesManager
        Private _initialized As Boolean = False
        Private _initializationLock As New Object()

#Region "单例模式"

        ''' <summary>
        ''' 获取 IntentBridge 单例实例（线程安全 Lazy 初始化）
        ''' </summary>
        Public Shared ReadOnly Property Instance As IntentBridge
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

#Region "初始化"

        ''' <summary>
        ''' 确保意图识别组件已初始化（延迟加载）
        ''' </summary>
        Private Sub EnsureInitialized()
            If _initialized Then Return

            SyncLock _initializationLock
                If _initialized Then Return

                Try
                    System.Diagnostics.Debug.WriteLine("[IntentBridge] 开始延迟初始化意图识别组件...")

                    ' 初始化规则管理器（单例模式）
                    _rulesManager = RulesManager.Instance
                    If _rulesManager Is Nothing Then
                        System.Diagnostics.Debug.WriteLine("[IntentBridge] RulesManager.Instance 返回 Nothing")
                        Return
                    End If
                    System.Diagnostics.Debug.WriteLine("[IntentBridge] RulesManager 实例获取成功")

                    ' 初始化意图识别引擎
                    _intentEngine = New IntentRecognitionEngine()
                    If _intentEngine Is Nothing Then
                        System.Diagnostics.Debug.WriteLine("[IntentBridge] IntentRecognitionEngine 初始化失败")
                        Return
                    End If
                    System.Diagnostics.Debug.WriteLine("[IntentBridge] IntentRecognitionEngine 初始化成功")

                    _initialized = True
                    System.Diagnostics.Debug.WriteLine("[IntentBridge] 意图识别组件延迟初始化成功")
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("[IntentBridge] 意图识别组件初始化失败: " & ex.Message)

                    ' 确保在失败时设置为 Nothing，以便调用方知道初始化失败
                    _intentEngine = Nothing
                    _rulesManager = Nothing
                    Throw
                End Try
            End SyncLock
        End Sub

#End Region

#Region "分层意图识别 API"

        ''' <summary>
        ''' 分层意图识别（借鉴 Trae 设计理念）
        ''' Level 1: 本地规则匹配（小于 1ms）
        ''' Level 2: Skill 关键词匹配（小于 10ms）
        ''' Level 3: 需要 LLM 分析（大于 500ms）
        ''' </summary>
        ''' <param name="input">用户输入文本</param>
        ''' <param name="hasAttachment">是否有附件</param>
        ''' <returns>JSON 格式的识别结果</returns>
        <ComVisible(True)>
        Public Function RecognizeIntent(input As String, Optional hasAttachment As Boolean = False) As String
            Try
                EnsureInitialized()
                If _intentEngine Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = "意图识别引擎未初始化",
                        .shouldCallLLM = True
                    })
                End If

                Dim result = _intentEngine.RecognizeIntent(input, hasAttachment)

                ' 使用 camelCase 序列化
                Dim settings = New JsonSerializerSettings With {
                    .ContractResolver = New CamelCasePropertyNamesContractResolver(),
                    .Formatting = Formatting.Indented
                }

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .level = result.Level,
                    .levelName = result.Level.ToString(),
                    .intent = result.Intent,
                    .confidence = result.Confidence,
                    .skillName = result.SkillName,
                    .shouldCallLLM = result.ShouldCallLLM,
                    .source = result.Source,
                    .response = result.Response,
                    .matchedKeyword = result.MatchedKeyword,
                    .needsConfirmation = result.NeedsConfirmation,
                    .processingTimeMs = result.ProcessingTime.TotalMilliseconds
                }, settings)

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[IntentBridge] RecognizeIntent 错误: {ex.Message}")
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message,
                    .shouldCallLLM = True
                })
            End Try
        End Function

        ''' <summary>
        ''' 获取快速响应（用于 Level 1 本地规则匹配）
        ''' </summary>
        ''' <param name="input">用户输入文本</param>
        ''' <returns>JSON 格式的快速响应结果</returns>
        <ComVisible(True)>
        Public Function GetQuickResponse(input As String) As String
            Try
                EnsureInitialized()
                If _rulesManager Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "RulesManager 未初始化"})
                End If

                Dim response = _rulesManager.GetQuickResponse(input)

                If response IsNot Nothing Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = True,
                        .hasResponse = True,
                        .response = response,
                        .source = "local_rules"
                    })
                Else
                    Return JsonConvert.SerializeObject(New With {
                        .success = True,
                        .hasResponse = False
                    })
                End If

            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 重新加载规则文件
        ''' </summary>
        ''' <returns>JSON 格式的操作结果</returns>
        <ComVisible(True)>
        Public Function ReloadRules() As String
            Try
                EnsureInitialized()
                If _rulesManager Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "RulesManager 未初始化"})
                End If

                _rulesManager.ReloadRules()
                _intentEngine.ClearCache()

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .message = "规则已重新加载"
                })

            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

#End Region

#Region "IBridge 生命周期"

        Public ReadOnly Property BridgeName As String Implements IBridge.BridgeName
            Get
                Return "IntentBridge"
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