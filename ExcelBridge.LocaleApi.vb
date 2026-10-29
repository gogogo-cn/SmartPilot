Option Strict On
Option Explicit On
Option Infer On

Imports System
Imports System.Runtime.InteropServices
Imports Newtonsoft.Json

Namespace ExcelAIVSTOAddIn

    ' ExcelBridge 的「界面语言 / 功能区本地化」分部实现（与 Word/PPT 行为对齐）。
    '
    ' 为什么需要这个文件（2026-10-11 定位）：
    '   Word/PPT 的桥对象继承 Hosts\Shared\HostBridgeBase.vb，那里自带
    '   SaveUiLocale / GetUiLocale；而 Excel 的 ExcelBridge 是一个不继承该基类的门面类，
    '   于是前端 App.i18n.setLocale() → bridge.SaveUiLocale(locale) 在 Excel 里永远失败。
    '   更隐蔽的是：WebView2 的 hostObject 对「任意方法名」的 typeof 都是 function
    '   （见 UI\js\bridge-accessor.js 的 _wrapHostBridge 注释），BridgeAccessor 又把失败
    '   吞成结构化 JSON → 表现为「WebView 里中英文切了，功能区按钮纹丝不动」。
    '
    ' 语言文件与功能区共用：%LocalAppData%\SmartPilot\ui-locale.txt（内容 zh / en），
    ' 由 Hosts\Shared\HostRibbon.vb 的 HostRibbonLocale 读写。
    ' （类级 XML 文档只在主文件 ExcelBridge.vb 保留一份，避免 BC42314 分部重复）
    Partial Public Class ExcelBridge

        ''' <summary>
        ''' 同步 WebView 界面语言到本机，并立即刷新功能区标签（触发 getLabel 回调重取）。
        ''' 返回 JSON 字符串：{"success":true,"locale":"zh"} 或 {"success":false,"error":"..."}
        ''' </summary>
        <ComVisible(True)>
        Public Function SaveUiLocale(locale As String) As String
            Try
                HostRibbonLocale.SaveLocale(locale)
                HostRibbonLocale.Invalidate()
                BroadcastUiLocale()
                Return "{""success"":true,""locale"":""" & HostRibbonLocale.GetLocale() & """}"
            Catch ex As Exception
                Return "{""success"":false,""error"":" & JsonConvert.SerializeObject(ex.Message) & "}"
            End Try
        End Function

        ''' <summary>
        ''' 主聊天和任务窗格是各自的 WebView2，localStorage 不共享。
        ''' 语言落盘后把脚本推进每一个已打开的页面，PPT 主题模板才会跟着变。
        ''' </summary>
        Private Shared Sub BroadcastUiLocale()
            Dim script As String = HostRibbonLocale.ApplyLocaleScript()
            Try
                Globals.ThisAddIn.ExecuteScriptOnWebView(script)
            Catch
            End Try
            Try
                Dim panes As TaskPaneManager = TaskPaneManager.Current
                If panes IsNot Nothing Then
                    panes.ExecuteScriptAll(script)
                    panes.SetPaneTitle("PptTheme", HostRibbonLocale.PptThemePaneTitle())
                End If
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' 读取当前宿主界面语言（zh / en）。
        ''' 供 WebView 首屏对齐使用：宿主功能区是英文时，WebView 不要停在中文。
        ''' </summary>
        <ComVisible(True)>
        Public Function GetUiLocale() As String
            Try
                Return "{""success"":true,""locale"":""" & HostRibbonLocale.GetLocale() & """}"
            Catch ex As Exception
                Return "{""success"":false,""error"":" & JsonConvert.SerializeObject(ex.Message) & "}"
            End Try
        End Function

    End Class

End Namespace
