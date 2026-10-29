Option Strict Off
Option Explicit On

Imports System
Imports System.Drawing
Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core

<ComVisible(True)>
Public Class Ribbon
    Inherits HostRibbonBase
    Implements IRibbonExtensibility

    ' ------------------------------------------------------------------
    ' 面板显示：本宿主唯一的差异点（Word 的 ThisAddIn 用 HostPaneManager）。
    ' 其余 5 个功能区回调与 4 个图标回调均由共享基类 HostRibbonBase 统一提供。
    ' ------------------------------------------------------------------
    Protected Overrides Sub TogglePanel(key As String, html As String, title As String)
        Try
            ' 【必须显式声明类型】本工程未开启 Option Infer，若写 "Dim addIn = Globals.ThisAddIn"
            ' 会把 addIn 推断为 Object → 走后期绑定 → 只能访问 Public 成员，
            ' 而 ThisAddIn 的面板方法是 Friend，会报「未找到类型 ThisAddIn 的公共成员 ToggleChatPanel」。
            Dim addIn As ThisAddIn = Globals.ThisAddIn
            If addIn Is Nothing Then Return
            Select Case key
                Case "main"
                    addIn.ToggleChatPanel()
                Case "aisettings"
                    addIn.ShowAiSettingsPanel()
                Case "ocr"
                    addIn.ToggleOcrPanel()
                Case "ocrsettings"
                    addIn.ShowOcrSettingsPanel()
                Case "mobile"
                    addIn.ToggleMobilePanel()
            End Select
        Catch ex As Exception
            System.Windows.Forms.MessageBox.Show("切换「" & title & "」面板失败: " & ex.Message, "SmartPilot",
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements IRibbonExtensibility.GetCustomUI
        Dim asm As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
        Dim candidates As String() = {"WordAIVSTOAddIn.Ribbon.xml", "Ribbon.xml"}
        For Each name As String In candidates
            Try
                Dim stream As System.IO.Stream = asm.GetManifestResourceStream(name)
                If stream IsNot Nothing Then
                    Using reader As New System.IO.StreamReader(stream)
                        Return reader.ReadToEnd()
                    End Using
                End If
            Catch
            End Try
        Next
        Return String.Empty
    End Function

    ' 图标回调（GetChatImage / GetSettingImage / GetOcrImage / GetMobileControlImage）
    ' 已由共享基类 HostRibbonBase 统一提供，此处不再重复维护。
End Class

