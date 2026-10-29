' ============================================================================
' Connect - PowerPoint COM 加载项（IDTExtensibility2 + IRibbonExtensibility）
'
' 为什么不用 VSTO：Visual Studio 的 VSTO 只提供 Excel / Word / Outlook 三套
' 宿主程序集（Microsoft.Office.Tools.{Excel,Word,Outlook}），PowerPoint 没有
' 对应的 ApplicationFactory / AddInBase，因此 PPT 侧采用经典 COM 加载项：
'   · 实现 Extensibility.IDTExtensibility2 接收 Application
'   · 实现 IRibbonExtensibility 提供功能区 XML 与回调
'   · 由 register.ps1 用 RegAsm 注册并写入 HKCU\...\PowerPoint\Addins
' ============================================================================
Option Strict Off
Option Explicit On

Imports System
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports AddInDesignerObjects
Imports Microsoft.Office.Core

' ⚠️ 必须是 AutoDual：Office 通过 IDispatch 按方法名调用 Ribbon 回调（onAction 等）。
'    若设成 None，CCW 只暴露接口成员，回调方法对 Office 不可见 → 点击按钮无任何反应。
<ComVisible(True)>
<ProgId("PptAIVSTOAddIn.Connect")>
<Guid("3F9A7C64-5B21-4E88-9A10-7D2C6E4B8F15")>
<ClassInterface(ClassInterfaceType.AutoDual)>
Public Class Connect
    Inherits HostRibbonBase
    Implements IDTExtensibility2
    Implements IRibbonExtensibility

    Private _app As Object
    Private _bridge As PptBridge
    Private _panes As HostPaneManager

    ' ---------------- 诊断日志 ----------------
    Public Sub New()
        Log("Connect ctor (COM instance created)")
    End Sub

    Private Shared Sub Log(msg As String)
        Try
            Dim dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SmartPilotOffice")
            If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
            File.AppendAllText(Path.Combine(dir, "ppt_addin.log"),
                Date.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") & " [" & Process.GetCurrentProcess().Id & "] " & msg & vbCrLf)
        Catch
        End Try
    End Sub

    ' ---------------- IDTExtensibility2 ----------------
    Public Sub OnConnection(ByVal application As Object, ByVal connectMode As ext_ConnectMode,
                            ByVal addInInst As Object, ByRef custom As System.Array) Implements IDTExtensibility2.OnConnection
        Log("OnConnection mode=" & CInt(connectMode))
        Try
            _app = application
            ' 若 PowerPoint 已在运行后加载（ext_cm_AfterStartup），此处即可初始化
            If connectMode <> ext_ConnectMode.ext_cm_Startup Then
                Initialize()
            End If
        Catch ex As Exception
            Log("OnConnection FAILED: " & ex.ToString())
        End Try
    End Sub

    Public Sub OnStartupComplete(ByRef custom As System.Array) Implements IDTExtensibility2.OnStartupComplete
        Log("OnStartupComplete")
        Try
            Initialize()
        Catch ex As Exception
            Log("OnStartupComplete FAILED: " & ex.ToString())
        End Try
    End Sub

    Public Sub OnDisconnection(ByVal RemoveMode As ext_DisconnectMode, ByRef custom As System.Array) Implements IDTExtensibility2.OnDisconnection
        Try
            If _panes IsNot Nothing Then _panes.Dispose()
        Catch
        End Try
        _panes = Nothing
        _app = Nothing
    End Sub

    Public Sub OnAddInsUpdate(ByRef custom As System.Array) Implements IDTExtensibility2.OnAddInsUpdate
    End Sub

    Public Sub OnBeginShutdown(ByRef custom As System.Array) Implements IDTExtensibility2.OnBeginShutdown
        Try
            If _panes IsNot Nothing Then _panes.Dispose()
        Catch
        End Try
    End Sub

    Private Sub Initialize()
        Try
            If _panes IsNot Nothing Then
                Log("Initialize skipped (already initialized)")
                Exit Sub
            End If
            HostPaneManager.Logger = AddressOf Log
            DockedFormHost.Logger = AddressOf Log
            Log("Initialize begin")
            Dim baseDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
            Dim uiDir = Path.Combine(baseDir, "UI")
            Log("UI dir exists=" & Directory.Exists(uiDir) & " -> " & uiDir)
            _bridge = New PptBridge(_app)
            Log("PptBridge created")

            _panes = New HostPaneManager(
                Function(uc As UserControl, title As String) As IPaneHost
                    Dim topOffset As Integer = 0
                    Dim refHwnd = GetPowerPointWindowHandle(_app, topOffset)
                    Dim ownerHwnd As IntPtr = IntPtr.Zero
                    Try
                        ownerHwnd = Process.GetCurrentProcess().MainWindowHandle
                    Catch
                    End Try
                    If ownerHwnd = IntPtr.Zero Then ownerHwnd = refHwnd
                    Log("Pane factory: ref=" & refHwnd.ToString() & " owner=" & ownerHwnd.ToString() & " topOffset=" & topOffset)
                    Return New DockedFormHost(ownerHwnd, refHwnd, uc, title, 480, topOffset, followRef:=False)
                End Function,
                uiDir, _bridge, "powerpoint")

            Log("SmartPilot initialized OK")
        Catch ex As Exception
            Log("Initialize FAILED: " & ex.ToString())
        End Try
    End Sub

    ''' <summary>
    ''' 取 PowerPoint 窗口句柄：优先 Application.HWND → ActiveWindow.HWND → 本进程主窗口；
    ''' 再尝试下钻到内容区子窗口（MDIClient），使面板贴在幻灯片区右侧而不是盖住功能区。
    ''' </summary>
    ''' <param name="topOffset">输出：父窗口为主窗口时用于避开功能区的顶部留白</param>
    Friend Shared Function GetPowerPointWindowHandle(app As Object, ByRef topOffset As Integer) As IntPtr
        topOffset = 0
        Dim h As IntPtr = IntPtr.Zero

        If app IsNot Nothing Then
            Try
                Dim v = app.HWND
                If v IsNot Nothing Then h = New IntPtr(Convert.ToInt64(v))
                Log("HWND prop -> " & h.ToString())
            Catch ex As Exception
                Log("Application.HWND failed: " & ex.Message)
            End Try
            If h = IntPtr.Zero Then
                Try
                    Dim v = app.ActiveWindow.HWND
                    If v IsNot Nothing Then h = New IntPtr(Convert.ToInt64(v))
                    Log("ActiveWindow.HWND -> " & h.ToString())
                Catch ex As Exception
                    Log("ActiveWindow.HWND failed: " & ex.Message)
                End Try
            End If
        End If

        If h = IntPtr.Zero Then
            Try
                h = Process.GetCurrentProcess().MainWindowHandle
                Log("Process.MainWindowHandle -> " & h.ToString())
            Catch ex As Exception
                Log("Process.MainWindowHandle failed: " & ex.Message)
            End Try
        End If

        If h = IntPtr.Zero Then Return IntPtr.Zero

        ' 下钻到内容区（MDIClient）：面板贴幻灯片区域右侧，不遮挡功能区
        Try
            Dim mdi = Win32.FindWindowEx(h, IntPtr.Zero, "MDIClient", Nothing)
            If mdi <> IntPtr.Zero Then
                Log("Found MDIClient -> " & mdi.ToString())
                ' MDIClient 的矩形不一定等于可见内容区（可能含状态栏等），
                ' 取其最大的可见子窗口（幻灯片编辑区）作为更精确的定位参考
                Dim content = FindLargestVisibleChild(mdi)
                If content <> IntPtr.Zero Then
                    Log("Content child -> " & content.ToString())
                    topOffset = 0
                    Return content
                End If
                topOffset = 0
                Return mdi
            End If
        Catch ex As Exception
            Log("FindWindowEx(MDIClient) failed: " & ex.Message)
        End Try

        ' 回退为主窗口：顶部留出功能区高度
        topOffset = 130
        Return h
    End Function

    ''' <summary>枚举直接子窗口，返回可见且面积最大的那个</summary>
    Private Shared Function FindLargestVisibleChild(hwndParent As IntPtr) As IntPtr
        Dim best As IntPtr = IntPtr.Zero
        Dim bestArea As Long = 0
        Try
            Win32.EnumChildWindows(hwndParent,
                Function(child As IntPtr, lParam As IntPtr) As Boolean
                    Try
                        If Win32.IsWindowVisible(child) Then
                            Dim r As Win32.RECT = Nothing
                            If Win32.GetWindowRect(child, r) Then
                                Dim area As Long = CLng(r.Right - r.Left) * CLng(r.Bottom - r.Top)
                                If area > bestArea Then
                                    bestArea = area
                                    best = child
                                End If
                            End If
                        End If
                    Catch
                    End Try
                    Return True
                End Function, IntPtr.Zero)
        Catch ex As Exception
            Log("EnumChildWindows failed: " & ex.Message)
        End Try
        Return best
    End Function

    ' ---------------- IRibbonExtensibility ----------------
    Public Function GetCustomUI(ByVal ribbonID As String) As String Implements IRibbonExtensibility.GetCustomUI
        Try
            Dim asm = Assembly.GetExecutingAssembly()
            Dim candidates As String() = {"PptAIVSTOAddIn.Ribbon.xml", "Ribbon.xml"}
            For Each name As String In candidates
                Dim stream As Stream = asm.GetManifestResourceStream(name)
                If stream IsNot Nothing Then
                    Using reader As New StreamReader(stream)
                        Dim xml = reader.ReadToEnd()
                        Log("GetCustomUI ok: " & name & " len=" & xml.Length)
                        Return xml
                    End Using
                End If
            Next
        Catch ex As Exception
            Log("GetCustomUI FAILED: " & ex.ToString())
        End Try
        Return String.Empty
    End Function

    ' ------------------------------------------------------------------
    ' 面板显示：本宿主唯一的差异点。
    ' 其余 5 个功能区回调与 4 个图标回调均由共享基类 HostRibbonBase 统一提供，
    ' 不再在此重复维护（PPT 走 IDTExtensibility2 + HostPaneManager）。
    ' ------------------------------------------------------------------
    Protected Overrides Sub TogglePanel(key As String, html As String, title As String)
        Log("TogglePanel: " & key & " -> " & html)
        Try
            Initialize()
            _panes.Toggle(key, html, title)
            Log("TogglePanel issued: " & key)
        Catch ex As Exception
            Log("TogglePanel FAILED (" & key & "): " & ex.ToString())
            MessageBox.Show("切换「" & title & "」面板失败: " & ex.Message, "SmartPilot", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class
