Option Strict On
Option Explicit On
Option Infer On

Imports System.Runtime.InteropServices
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ExcelAIVSTOAddIn.Services
Imports ExcelAIVSTOAddIn.Services.Mcp

Namespace Bridges

    ''' <summary>
    ''' ThinkingBridge - 新架构思考与动作引擎 API 桥梁类
    ''' 提供 COM 可见的思考流程和动作执行方法供前端 JavaScript 调用
    ''' </summary>
    <ComVisible(True)>
    <ClassInterface(ClassInterfaceType.AutoDispatch)>
    <ProgId("ExcelAIVSTOAddIn.ThinkingBridge")>
    Public Class ThinkingBridge
        Implements IBridge, IThinkingBridge

        Private Shared ReadOnly _lazyInstance As New Lazy(Of ThinkingBridge)(Function() New ThinkingBridge(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        Private _thinkingEngine As ThinkingEngine
        Private _actionEngine As ActionEngine
        Private _config As Utils.ConfigurationManager

#Region "单例模式"

        Public Shared ReadOnly Property Instance As ThinkingBridge
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        Private Sub New()
        End Sub

#End Region

#Region "初始化"

        Public Sub Initialize(bridge As ExcelAIVSTOAddIn.ExcelBridge)
            Try
                System.Diagnostics.Debug.WriteLine("[ThinkingBridge] 开始初始化...")

                _config = Utils.ConfigurationManager.Instance

                If _thinkingEngine Is Nothing Then
                    _thinkingEngine = ThinkingEngine.Instance
                    _thinkingEngine.Initialize(bridge)
                    System.Diagnostics.Debug.WriteLine("[ThinkingBridge] ThinkingEngine 初始化成功并已注入 ExcelBridge 引用")
                End If

                If _actionEngine Is Nothing Then
                    _actionEngine = ActionEngine.Instance
                    _actionEngine.Initialize(bridge)
                    System.Diagnostics.Debug.WriteLine("[ThinkingBridge] ActionEngine 初始化成功并已注入 ExcelBridge 引用")
                End If

                Dim llmEngine = MultimodalLLMEngine.SharedInstance
                llmEngine.LoadFromConfiguration(_config)
                System.Diagnostics.Debug.WriteLine("[ThinkingBridge] LLM Engine 配置已初始化")

                System.Diagnostics.Debug.WriteLine("[ThinkingBridge] 初始化完成")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[ThinkingBridge] 初始化失败: " & ex.Message)
                Throw
            End Try
        End Sub

        #End Region

    #Region "新架构：思考与动作引擎 API"

        <ComVisible(True)>
        Public Async Function AnalyzeAndThink(
            message As String,
            chatHistoryJson As String,
            excelInfoJson As String,
            selectedCellsJson As String,
            Optional thinkingMode As String = "AUTO"
        ) As Task(Of String) Implements IThinkingBridge.AnalyzeAndThink

            Try
                If _thinkingEngine Is Nothing Then
                    _thinkingEngine = ThinkingEngine.Instance
                End If
                If _actionEngine Is Nothing Then
                    _actionEngine = ActionEngine.Instance
                End If

                Dim historyMessages As List(Of Dictionary(Of String, String)) = Nothing
                Dim contextMessages As List(Of String) = New List(Of String)()
                Try
                    historyMessages = JsonConvert.DeserializeObject(Of List(Of Dictionary(Of String, String)))(chatHistoryJson)
                    If historyMessages IsNot Nothing Then
                        For Each m In historyMessages
                            Dim role As String = If(m IsNot Nothing AndAlso m.ContainsKey("role"), If(m("role"), ""), "")
                            Dim content As String = If(m IsNot Nothing AndAlso m.ContainsKey("content"), If(m("content"), ""), "")
                            contextMessages.Add(role & ": " & content)
                        Next
                    End If
                Catch ex As Exception  ' 
                    contextMessages = New List(Of String)()
                End Try

                System.Diagnostics.Debug.WriteLine("[ThinkingBridge] 准备调用v3.0 AnalyzeAndThink, 用户指定模式: " & thinkingMode)

                Dim thinkingResult As ThinkingResponse = Await _thinkingEngine.AnalyzeAndThink(
                    message,
                    hasImage:=message.Contains("data:image"),
                    contextHistory:=String.Join("; ", contextMessages),
                    externalMode:=thinkingMode
                )

                System.Diagnostics.Debug.WriteLine("[ThinkingBridge] v3.0执行完成, 模式: " & thinkingResult.Mode.ToString() & ", 步骤数: " & thinkingResult.ThinkingSteps.Count.ToString())

                Dim resultObj As New Dictionary(Of String, Object) From {
                    {"success", True},
                    {"thinkingMode", thinkingResult.Mode.ToString()},
                    {"thinkingProcess", thinkingResult.ThinkingSteps},
                    {"answer", thinkingResult.FinalAnswer},
                    {"actions", thinkingResult.Actions},
                    {"confidence", thinkingResult.Confidence},
                    {"conversationId", If(thinkingResult.ConversationId, "")}
                }
                Return JsonConvert.SerializeObject(resultObj)

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[ThinkingBridge] AnalyzeAndThink 错误: " & ex.Message)
                Dim errObj As New Dictionary(Of String, Object) From {
                    {"success", False},
                    {"error", ex.Message}
                }
                Return JsonConvert.SerializeObject(errObj)
            End Try
        End Function

        <ComVisible(True)>
        Public Async Function ExecuteAction(actionJson As String) As Task(Of String) Implements IThinkingBridge.ExecuteAction
            Try
                If _actionEngine Is Nothing Then
                    _actionEngine = ActionEngine.Instance
                End If

                Dim action As ActionItem = JsonConvert.DeserializeObject(Of ActionItem)(actionJson)
                Dim result As OperationResult = Await _actionEngine.ExecuteSingleAction(action)

                Dim resultObj As New Dictionary(Of String, Object) From {
                    {"success", result.Success},
                    {"message", If(result.Message, "")},
                    {"data", If(result.Data, Nothing)}
                }
                Return JsonConvert.SerializeObject(resultObj)

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[ThinkingBridge] ExecuteAction 错误: " & ex.Message)
                Dim errObj As New Dictionary(Of String, Object) From {
                    {"success", False},
                    {"error", ex.Message}
                }
                Return JsonConvert.SerializeObject(errObj)
            End Try
        End Function

        <ComVisible(True)>
        Public Function GetAvailableActions() As String Implements IThinkingBridge.GetAvailableActions
            Try
                If _actionEngine Is Nothing Then
                    _actionEngine = ActionEngine.Instance
                End If

                Dim actions = _actionEngine.GetAvailableActions()
                Dim resultObj As New Dictionary(Of String, Object) From {
                    {"success", True},
                    {"actions", actions}
                }
                Return JsonConvert.SerializeObject(resultObj)

            Catch ex As Exception
                Dim errObj As New Dictionary(Of String, Object) From {
                    {"success", False},
                    {"error", ex.Message}
                }
                Return JsonConvert.SerializeObject(errObj)
            End Try
        End Function

#End Region

#Region "IBridge 生命周期"

        Public ReadOnly Property BridgeName As String Implements IBridge.BridgeName
            Get
                Return "ThinkingBridge"
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