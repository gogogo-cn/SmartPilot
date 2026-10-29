Option Strict On
Option Infer On

Imports System
Imports ExcelAIVSTOAddIn.Utils

Partial Public NotInheritable Class ThisAddIn

#Region "配置热重载"

    Private _configWatcher As Utils.ConfigFileWatcher

    ''' <summary>
    ''' 初始化配置热重载
    ''' </summary>
    Private Sub InitializeConfigHotReload()
        Try
            _configWatcher = Utils.ConfigFileWatcher.Instance
            AddHandler _configWatcher.ConfigFileChanged, AddressOf OnConfigFileChanged
            _chatLogger?.LogInfo("ADDIN", "Config hot reload enabled")
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("[ThisAddIn] Config hot reload init failed: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 配置文件变更处理
    ''' </summary>
    Private Sub OnConfigFileChanged(sender As Object, e As EventArgs)
        Try
            _chatLogger?.LogInfo("ADDIN", "Config file change detected, reloading...")
            ConfigurationManager.Instance.Reload()
            SendConfigToUI()
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("[ThisAddIn] Config file change handler failed: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 发送配置到UI
    ''' </summary>
    Private Sub SendConfigToUI()
        Try
            Dim configJson As String = ConfigurationManager.Instance.GetConfigJson()
            ' 新WPF主面板不使用WebMessage通道;配置同步使用ExecutionViewModel公共属性绑定
            System.Diagnostics.Debug.WriteLine("[ThisAddIn] Config updated, WPF panel config sync pending: " & configJson.Length.ToString() & " chars")
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("[ThisAddIn] Send config to UI failed: " & ex.Message)
        End Try
    End Sub

#End Region

End Class