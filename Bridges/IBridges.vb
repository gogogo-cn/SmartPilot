Option Explicit On
Option Infer On

Imports System.Threading.Tasks

Namespace Bridges

    ''' <summary>
    ''' IBridge 基础接口 - 所有 Bridge 子类的统一生命周期契约
    ''' 继承 IDisposable，确保资源可以被统一释放
    ''' </summary>
    Public Interface IBridge
        Inherits IDisposable
        ''' <summary>
        ''' 桥接器名称，用于日志和调试
        ''' </summary>
        ReadOnly Property BridgeName As String
    End Interface

    ''' <summary>
    ''' 撤销动作记录类 - 统一定义在 Bridges 命名空间中供所有 Bridge 子类使用
    ''' </summary>
    Public Class UndoAction
        Public Property ActionType As String
        Public Property Address As String
        Public Property OldValue As Object
        Public Property NewValue As Object
        Public Property SheetName As String
        Public Property Timestamp As DateTime
        Public Property Payload As String

        Public Sub New(actionType As String, address As String, oldValue As Object, newValue As Object, sheetName As String, Optional payload As String = "")
            Me.ActionType = actionType
            Me.Address = address
            Me.OldValue = oldValue
            Me.NewValue = newValue
            Me.SheetName = sheetName
            Me.Payload = payload
            Me.Timestamp = DateTime.Now
        End Sub
    End Class

    ''' <summary>
    ''' ISkillBridge 接口 - Skill 系统 API 契约
    ''' 定义 Skill 管理的核心操作，支持依赖注入和单元测试
    ''' </summary>
    Public Interface ISkillBridge
        Function GetAvailableSkills() As String
        Function GetSkillDetails(skillName As String) As String
        Function GetSkillByName(skillName As String) As String
        Function GetQuickShortcutSkills() As String
        Function ToggleQuickShortcut(skillName As String, isQuick As Boolean) As String
        Function ToggleSkillEnabled(skillName As String, isEnabled As Boolean) As String
        Function ExecuteSkill(skillName As String, inputsJson As String) As String
        Function ExecuteSkillAsync(skillName As String, inputsJson As String) As Task(Of String)
        Function GetAllSkills() As String
        Function DeleteSkill(skillName As String) As String
        Function SaveSkill(skillJson As String) As String
    End Interface

    ''' <summary>
    ''' IMcpBridge 接口 - MCP (Model Context Protocol) API 契约
    ''' 定义 MCP 协议的核心操作，支持依赖注入和单元测试
    ''' </summary>
    Public Interface IMcpBridge
        Function GetMcpStatus() As String
        Function GetMcpServerStatus() As String
        Function StartMcpServerAsync(Optional port As Integer = -1) As Task(Of String)
        Function StopMcpServerAsync() As Task(Of String)
        Function GetMcpServerConfig() As String
        Function SaveMcpServerConfigAsync(configJson As String) As Task(Of String)
        Function AddMcpClientAsync(clientConfigJson As String) As Task(Of String)
        Function RemoveMcpClientAsync(name As String) As Task(Of String)
        Function GetMcpClientTools(name As String) As String
        Function TestMcpClientConnectionAsync(name As String) As Task(Of String)
        Function CallMcpToolAsync(clientName As String, toolName As String, argumentsJson As String) As Task(Of String)
        Function CallMcpTool(clientName As String, toolName As String, argumentsJson As String) As String
        Function GetAvailableMcpTools() As String
        Function GetMcpConfig() As String
        Function AutoReconnectMcpClientsAsync() As Task(Of String)
        Function CallLocalTool(toolName As String, argumentsJson As String) As String
        Sub ShutdownMcpClients()
    End Interface

    ''' <summary>
    ''' IExcelDataReader 接口 - 最小作用域数据读取契约
    ''' 只暴露服务层实际依赖的 Excel 数据读取能力，避免曝光 ExcelDataBridge 全貌
    ''' </summary>
    Public Interface IExcelDataReader
        Function GetRangeData(rangeAddress As String) As String
        Function GetSelectedRangeData() As String
        Function GetUsedRangeData() As String
    End Interface

    ''' <summary>
    ''' IUndoBridge 接口 - 撤销/重做管理 API 契约
    ''' 定义撤销栈的核心操作，供 Excel 数据桥等消费者经依赖注入按接口访问
    ''' </summary>
    Public Interface IUndoBridge
        Function GetUndoStack() As String
        Function UndoLastAction() As String
        Sub ClearUndoStack()
        Function GetRedoStack() As String
        Function RedoLastAction() As String
        Sub ClearRedoStack()
        Sub AddUndoAction(action As UndoAction)
        ''' <summary>
        ''' 整表备份：在"增删改"操作执行【之前】调用，把受影响工作表整体拍快照压入撤销栈
        ''' </summary>
        Sub AddSheetBackup(sheetName As String)
    End Interface

    ''' <summary>
    ''' IThinkingBridge 接口 - 思考与动作引擎 API 契约
    ''' 定义 AI 思考流程和动作执行的核心操作
    ''' </summary>
    Public Interface IThinkingBridge
        Function AnalyzeAndThink(message As String, chatHistoryJson As String, excelInfoJson As String, selectedCellsJson As String, Optional thinkingMode As String = "AUTO") As Task(Of String)
        Function ExecuteAction(actionJson As String) As Task(Of String)
        Function GetAvailableActions() As String
    End Interface

End Namespace