Imports System
Imports System.Drawing
Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports ExcelAIVSTOAddIn.Forms

<ComVisible(True)>
Public Class Ribbon
    Inherits HostRibbonBase
    Implements IRibbonExtensibility
    ' ------------------------------------------------------------------
    ' 面板显示：本宿主唯一的差异点（Excel 走 ThisAddIn / TaskPaneManager）。
    ' 其余 5 个功能区回调与 4 个图标回调均由共享基类 HostRibbonBase 统一提供，
    ' 不再在此重复维护（此前 Excel 还另有一套基于内嵌 PNG 的图标加载机制）。
    ' ------------------------------------------------------------------
    Protected Overrides ReadOnly Property IsExcelHost As Boolean
        Get
            Return True
        End Get
    End Property

    Protected Overrides Sub TogglePanel(key As String, html As String, title As String)
        Try
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
            MessageBox.Show("切换「" & title & "」面板失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements IRibbonExtensibility.GetCustomUI
        Try
            Dim resourceName As String = "ExcelAIVSTOAddIn.Ribbon.xml"
            Dim asm As Assembly = Assembly.GetExecutingAssembly()

            Dim stream As Stream = asm.GetManifestResourceStream(resourceName)
            If stream IsNot Nothing Then
                Using reader As New StreamReader(stream)
                    Return reader.ReadToEnd()
                End Using
            End If

            System.Diagnostics.Debug.WriteLine("[Ribbon] 错误: 找不到嵌入资源! resourceName=" & resourceName)
            Return String.Empty

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("[Ribbon] GetCustomUI 异常: " & ex.ToString())
            Return String.Empty
        End Try
    End Function

    ' OnLoad 与 OnToggleSidebar 已由共享基类 HostRibbonBase 提供。

    ' ------------------------------------------------------------------
    ' 图标回调（GetChatImage / GetSettingImage / GetOcrImage / GetMobileControlImage）
    ' 已统一由共享基类 HostRibbonBase 提供（GDI+ 程序化绘制）。
    ' 原先 Excel 独有的「从内嵌 PNG 资源加载」机制（LoadImageFromResource /
    ' ImageToPictureDisp）随之移除 —— 它曾是「图标空白」问题的根源（资源名不匹配），
    ' 且是三端中唯一的一套，不利于统一维护。
    ' ------------------------------------------------------------------

End Class
