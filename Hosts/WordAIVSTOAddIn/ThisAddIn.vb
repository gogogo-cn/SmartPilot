Option Strict Off
Option Explicit On

Imports System
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports Microsoft.Office.Tools

Public Class ThisAddIn

    Private _bridge As WordBridge
    Private _panes As HostPaneManager
    Private _ribbon As Ribbon

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Try
            Dim uiDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "UI")
            _bridge = New WordBridge(Me.Application)

            _panes = New HostPaneManager(
                Function(uc As UserControl, title As String) As IPaneHost
                    Dim pane = Globals.ThisAddIn.CustomTaskPanes.Add(uc, title)
                    pane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
                    pane.Width = 480
                    Return New CustomTaskPaneHost(pane)
                End Function,
                uiDir, _bridge, "word")

            System.Diagnostics.Debug.WriteLine("[WordAddIn] SmartPilot 已启动，UI 目录: " & uiDir)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("[WordAddIn] 启动失败: " & ex.ToString())
            MessageBox.Show("SmartPilot 启动失败: " & ex.Message, "SmartPilot", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        Try
            If _panes IsNot Nothing Then _panes.Dispose()
        Catch
        End Try
    End Sub

    ' ---------------- Ribbon 回调入口 ----------------
    Friend Sub ToggleChatPanel()
        _panes.Toggle("main", "index.html", "SmartPilot 智航")
    End Sub

    Friend Sub ToggleOcrPanel()
        _panes.Toggle("ocr", "ocr-panel.html", "识图入表")
    End Sub

    Friend Sub ShowAiSettingsPanel()
        _panes.Toggle("aisettings", "ai-settings-panel.html?tab=ai", "AI 模型设置")
    End Sub

    Friend Sub ShowOcrSettingsPanel()
        _panes.Toggle("ocrsettings", "ocr-settings-panel.html", "OCR 模型设置")
    End Sub

    Friend Sub ToggleMobilePanel()
        _panes.Toggle("mobile", "mobile-control-panel.html", "手机遥控")
    End Sub

    Protected Overrides Function RequestService(serviceGuid As Guid) As Object
        If serviceGuid = GetType(IRibbonExtensibility).GUID Then
            If _ribbon Is Nothing Then _ribbon = New Ribbon()
            Return _ribbon
        End If
        Return MyBase.RequestService(serviceGuid)
    End Function
End Class
