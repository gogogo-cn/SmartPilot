Option Strict On
Option Explicit On
Option Infer On

Namespace Security

    ''' <summary>
    ''' 自动化执行设置（全局静态标志）。
    ''' 由前端「自动化免确认」开关通过 ExcelBridge.SetAutoExecutionBypass 控制。
    ''' 开启后，VBA / Python 高风险代码执行跳过用户确认弹窗（用户已显式授权）。
    ''' </summary>
    Public NotInheritable Class AutoExecutionSettings

        ''' <summary>
        ''' 是否跳过高风险代码执行的用户确认（默认关闭，保持安全确认）。
        ''' 属于进程内运行时状态：Excel 重启后恢复为默认 False，
        ''' 前端初始化时会按 localStorage 记忆的开关状态重新设置。
        ''' </summary>
        Public Shared Property BypassConfirmation As Boolean = False

    End Class

End Namespace
