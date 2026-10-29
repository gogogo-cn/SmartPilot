' ============================================================================
' HostRibbon - 三个宿主（Excel / PowerPoint / Word）共用的功能区基础设施
'
' 目的：消除「一处逻辑、三份维护」。此前：
'   - 5 个面板回调（OnToggleSidebar / AiModel / OcrPanel / OcrSettings / MobilePanel）
'     在 Excel Ribbon.vb、PPT Connect.vb、Word Ribbon.vb 各写一遍
'   - 4 个图标回调（GetChatImage / GetSettingImage / GetOcrImage / GetMobileControlImage）
'     同样三份
'   - 图标工厂 HostRibbonIcons + HostPictureConverter 在 PPT/Word 各重复一份
'
' 现在：回调与图标统一在本文件的 HostRibbonBase 中，各宿主只重写 TogglePanel
' （因为 Excel 走 TaskPaneManager、PPT/Word 走 HostPaneManager，显示面板的方式不同）。
' ============================================================================
Option Strict Off
Option Explicit On

Imports System
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Globalization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports Microsoft.Office.Core

''' <summary>
''' 宿主功能区回调基类。
''' 各宿主的 Ribbon 类继承本类，只需实现 TogglePanel（面板显示方式）与
''' GetCustomUI（嵌入资源清单名随 RootNamespace 不同而不同）。
''' </summary>
<ComVisible(True)>
Public MustInherit Class HostRibbonBase

    Protected ribbonUI As IRibbonUI

    Public Overridable Sub OnLoad(ByVal ribbon As IRibbonUI)
        Try
            ribbonUI = ribbon
            HostRibbonLocale.Active = Me
            ' 首次加载时若尚无语言文件，按系统 UI 语言写入，保证 getLabel 稳定可读
            HostRibbonLocale.EnsureLocaleFile()
            InvalidateLabels()
        Catch
        End Try
    End Sub

    ''' <summary>界面语言变更后刷新功能区标签（由 SaveUiLocale 触发）。</summary>
    Public Sub InvalidateLabels()
        Try
            If ribbonUI IsNot Nothing Then ribbonUI.Invalidate()
        Catch
        End Try
    End Sub

    ' ------------------------------------------------------------------
    ' 面板显示：宿主差异点，由子类实现
    ' ------------------------------------------------------------------
    Protected MustOverride Sub TogglePanel(key As String, html As String, title As String)

    ' ------------------------------------------------------------------
    ' 宿主标识：用于「仅某宿主可见」的功能区按钮（如 Excel 专属面板）。
    ' 共享 Ribbon.xml 中的 getVisible="GetExcelOnlyVisible" 据此决定是否显示，
    ' 从而在不分裂 XML 文件的前提下保留各宿主独有功能。
    ' ------------------------------------------------------------------
    Protected Overridable ReadOnly Property IsExcelHost As Boolean
        Get
            Return False
        End Get
    End Property

    Public Overridable Function GetExcelOnlyVisible(ByVal control As IRibbonControl) As Boolean
        Return IsExcelHost
    End Function

    ' ------------------------------------------------------------------
    ' 功能区标签 / 提示（随 ui-locale 切换中英文）
    ' ------------------------------------------------------------------
    Public Overridable Function GetControlLabel(ByVal control As IRibbonControl) As String
        Dim en As Boolean = HostRibbonLocale.IsEnglish()
        Select Case control.Id
            Case "tabSmartPilot"
                Return If(en, "SmartPilot", "SmartPilot 智航")
            Case "groupMain"
                Return If(en, "Main", "主功能")
            Case "groupSettings"
                Return If(en, "Settings", "设置")
            Case "btnToggleSidebar"
                Return If(en, "AI Chat", "智能聊天")
            Case "btnImageOcr"
                Return If(en, "Capture to sheet", "识图入表")
            Case "btnAiModel"
                Return If(en, "AI Model", "AI 模型")
            Case "btnOcrSettings"
                Return If(en, "OCR Model", "OCR 模型")
            Case "btnMobileControl"
                Return If(en, "Mobile", "手机遥控")
            Case Else
                Return control.Id
        End Select
    End Function

    Public Overridable Function GetControlSupertip(ByVal control As IRibbonControl) As String
        Dim en As Boolean = HostRibbonLocale.IsEnglish()
        Select Case control.Id
            Case "btnToggleSidebar"
                Return If(en, "Show or hide the SmartPilot chat pane", "显示或隐藏 SmartPilot 智航聊天面板")
            Case "btnImageOcr"
                Return If(en, "Recognize screenshots, tables, forms, and documents into the current sheet", "识别截图、表格、表单和票据，写入当前工作表")
            Case "btnAiModel"
                Return If(en, "AI model settings — provider, API, and model", "AI 模型配置 - 提供商、API、模型选择")
            Case "btnOcrSettings"
                Return If(en, "OCR engine settings — RapidOCR / PaddleOCR / cloud", "OCR 识别引擎配置 - RapidOCR/PaddleOCR/云端")
            Case "btnMobileControl"
                Return If(en, "Mobile remote OCR — scan, capture, upload", "手机遥控 OCR - 扫码拍照上传识别")
            Case Else
                Return ""
        End Select
    End Function

    ' ------------------------------------------------------------------
    ' 统一的功能区回调（onAction）
    ' ------------------------------------------------------------------
    Public Overridable Sub OnToggleSidebar(ByVal control As IRibbonControl)
        TogglePanel("main", "index.html", HostRibbonLocale.Text("SmartPilot", "SmartPilot 智航"))
    End Sub

    Public Overridable Sub OnToggleAiModel(ByVal control As IRibbonControl)
        TogglePanel("aisettings", "ai-settings-panel.html?tab=ai", HostRibbonLocale.Text("AI Model Settings", "AI 模型设置"))
    End Sub

    Public Overridable Sub OnToggleOcrPanel(ByVal control As IRibbonControl)
        TogglePanel("ocr", "ocr-panel.html", HostRibbonLocale.Text("Capture to sheet", "识图入表"))
    End Sub

    Public Overridable Sub OnToggleOcrSettings(ByVal control As IRibbonControl)
        TogglePanel("ocrsettings", "ocr-settings-panel.html", HostRibbonLocale.Text("OCR Model Settings", "OCR 模型设置"))
    End Sub

    Public Overridable Sub OnToggleMobilePanel(ByVal control As IRibbonControl)
        TogglePanel("mobile", "mobile-control-panel.html", HostRibbonLocale.Text("Mobile Control", "手机遥控"))
    End Sub

    ' ------------------------------------------------------------------
    ' 统一的图标回调（getImage）：程序化绘制，不依赖内嵌图片资源
    ' ------------------------------------------------------------------
    Public Overridable Function GetChatImage(ByVal control As IRibbonControl) As stdole.IPictureDisp
        Try
            Return HostRibbonIcons.CreateChat()
        Catch
            Return Nothing
        End Try
    End Function

    Public Overridable Function GetSettingImage(ByVal control As IRibbonControl) As stdole.IPictureDisp
        Try
            Return HostRibbonIcons.CreateAiModel()
        Catch
            Return Nothing
        End Try
    End Function

    Public Overridable Function GetOcrImage(ByVal control As IRibbonControl) As stdole.IPictureDisp
        Try
            Return HostRibbonIcons.CreateOcr()
        Catch
            Return Nothing
        End Try
    End Function

    Public Overridable Function GetMobileControlImage(ByVal control As IRibbonControl) As stdole.IPictureDisp
        Try
            Return HostRibbonIcons.CreateMobileControl()
        Catch
            Return Nothing
        End Try
    End Function
End Class

' ============================================================================
' 功能区图标工厂：用 GDI+ 直接绘制 32x32 图标并转成 IPictureDisp。
' 好处：不依赖任何嵌入图片资源（此前因资源名不匹配/缺失导致功能区按钮无图标）。
' ============================================================================
Friend NotInheritable Class HostRibbonIcons
    Private Sub New()
    End Sub

    Private Shared ReadOnly Accent As Color = Color.FromArgb(37, 99, 235)

    Public Shared Function CreateChat() As stdole.IPictureDisp
        Dim bmp As New Bitmap(32, 32)
        Using g As Graphics = Graphics.FromImage(bmp)
            g.SmoothingMode = SmoothingMode.AntiAlias
            g.Clear(Color.Transparent)

            Dim body As New Rectangle(3, 4, 26, 18)
            Using path As New GraphicsPath()
                AddRoundedRect(path, body, 7)
                path.AddPolygon(New Point() {
                    New Point(8, body.Bottom - 1),
                    New Point(15, body.Bottom - 1),
                    New Point(9, body.Bottom + 6)})
                Using brush As New SolidBrush(Accent)
                    g.FillPath(brush, path)
                End Using
            End Using

            Using dot As New SolidBrush(Color.White)
                g.FillEllipse(dot, 9, 10, 5, 5)
                g.FillEllipse(dot, 14, 10, 5, 5)
                g.FillEllipse(dot, 19, 10, 5, 5)
            End Using
        End Using
        Return ToPictureDisp(bmp)
    End Function

    Public Shared Function CreateAiModel() As stdole.IPictureDisp
        Dim bmp As New Bitmap(32, 32)
        Using g As Graphics = Graphics.FromImage(bmp)
            g.SmoothingMode = SmoothingMode.AntiAlias
            g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
            g.Clear(Color.Transparent)

            Dim rect As New Rectangle(3, 3, 26, 26)
            Using path As New GraphicsPath()
                AddRoundedRect(path, rect, 8)
                Using fill As New SolidBrush(Color.White)
                    g.FillPath(fill, path)
                End Using
                Using pen As New Pen(Accent, 2.0F)
                    g.DrawPath(pen, path)
                End Using
            End Using

            Using font As New Font("Segoe UI", 13.0F, FontStyle.Bold, GraphicsUnit.Pixel)
                Using sf As New StringFormat()
                    sf.Alignment = StringAlignment.Center
                    sf.LineAlignment = StringAlignment.Center
                    Using brush As New SolidBrush(Accent)
                        g.DrawString("AI", font, brush, New RectangleF(0, 0, 32, 32), sf)
                    End Using
                End Using
            End Using
        End Using
        Return ToPictureDisp(bmp)
    End Function

    Public Shared Function CreateOcr() As stdole.IPictureDisp
        Dim bmp As New Bitmap(32, 32)
        Using g As Graphics = Graphics.FromImage(bmp)
            g.SmoothingMode = SmoothingMode.AntiAlias
            g.Clear(Color.Transparent)

            ' 外框：天蓝圆角矩形（图片/相机底）
            Dim outer As New Rectangle(4, 6, 24, 20)
            Using p As New GraphicsPath()
                AddRoundedRect(p, outer, 5)
                Using brush As New SolidBrush(Accent)
                    g.FillPath(brush, p)
                End Using
            End Using

            ' 内白：图片内容区
            Dim inner As New Rectangle(7, 9, 18, 14)
            Using p2 As New GraphicsPath()
                AddRoundedRect(p2, inner, 3)
                Using white As New SolidBrush(Color.White)
                    g.FillPath(white, p2)
                End Using
            End Using

            ' 太阳
            Using sun As New SolidBrush(Accent)
                g.FillEllipse(sun, 10, 12, 4, 4)
            End Using

            ' 山（三角形）
            Using mountain As New SolidBrush(Accent)
                g.FillPolygon(mountain, New Point() {
                    New Point(9, 23), New Point(15, 15), New Point(20, 23)})
            End Using
        End Using
        Return ToPictureDisp(bmp)
    End Function

    Public Shared Function CreateMobileControl() As stdole.IPictureDisp
        Dim bmp As New Bitmap(32, 32)
        Using g As Graphics = Graphics.FromImage(bmp)
            g.SmoothingMode = SmoothingMode.AntiAlias
            g.Clear(Color.Transparent)

            ' 手机机身
            Dim body As New Rectangle(9, 3, 14, 26)
            Using p As New GraphicsPath()
                AddRoundedRect(p, body, 4)
                Using fill As New SolidBrush(Color.White)
                    g.FillPath(fill, p)
                End Using
                Using pen As New Pen(Accent, 2.0F)
                    g.DrawPath(pen, p)
                End Using
            End Using

            ' 摄像头点
            Using dot As New SolidBrush(Accent)
                g.FillEllipse(dot, 14, 6, 4, 2)
            End Using

            ' Home 指示条
            Using bar As New SolidBrush(Accent)
                g.FillRectangle(bar, 13, 25, 6, 2)
            End Using
        End Using
        Return ToPictureDisp(bmp)
    End Function

    Private Shared Sub AddRoundedRect(path As GraphicsPath, r As Rectangle, radius As Integer)
        Dim d As Integer = radius * 2
        path.AddArc(r.X, r.Y, d, d, 180, 90)
        path.AddArc(r.Right - d, r.Y, d, d, 270, 90)
        path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90)
        path.AddArc(r.X, r.Bottom - d, d, d, 90, 90)
        path.CloseFigure()
    End Sub

    Private Shared Function ToPictureDisp(img As Image) As stdole.IPictureDisp
        Return DirectCast(HostPictureConverter.Convert(img), stdole.IPictureDisp)
    End Function
End Class

' AxHost 子类：借其保护成员 GetIPictureDispFromPicture 把 Image 转成 IPictureDisp
Friend Class HostPictureConverter
    Inherits System.Windows.Forms.AxHost

    Public Sub New()
        MyBase.New("{00000000-0000-0000-0000-000000000000}")
    End Sub

    Public Shared Function Convert(img As Image) As Object
        Return GetIPictureDispFromPicture(img)
    End Function
End Class

''' <summary>
''' 功能区 / 任务窗格标题用的界面语言（与 WebView localStorage smartpilot_ui_locale 同步）。
''' 文件：%LocalAppData%\SmartPilot\ui-locale.txt （内容为目录短码，目前 zh / en）
''' 追加语言：在 LocaleCatalog 加一行，并在 UI/js/i18n.js 的 LOCALE_CATALOG 加同一短码。
''' 功能区按钮文案目前只有中/英；非中文短码先走英文，直到 GetControlLabel 扩成多语。
''' </summary>
Friend Module HostRibbonLocale
    Public Active As HostRibbonBase

    ''' <summary>每行第一个是短码，后面是可识别别名。顺序与 i18n.js LOCALE_CATALOG 保持一致。</summary>
    Private Function LocaleCatalog() As String()()
        Return New String()() {
            New String() {"zh", "zh-cn", "zh-hans"},
            New String() {"en", "en-us", "en-gb"}
        }
    End Function

    Private Function MatchLocale(locale As String) As String
        Dim s As String = If(locale, "").Trim().ToLowerInvariant().Replace("_"c, "-"c)
        If s.Length = 0 Then Return Nothing
        Dim table As String()() = LocaleCatalog()
        Dim row As String()
        Dim aliasName As String
        For Each row In table
            If row Is Nothing OrElse row.Length = 0 Then Continue For
            For Each aliasName In row
                If s = aliasName Then Return row(0)
            Next
        Next
        Dim best As String = Nothing
        Dim bestLen As Integer = -1
        For Each row In table
            If row Is Nothing OrElse row.Length = 0 Then Continue For
            For Each aliasName In row
                If s.Length > aliasName.Length AndAlso s.StartsWith(aliasName & "-", StringComparison.Ordinal) Then
                    If aliasName.Length > bestLen Then
                        best = row(0)
                        bestLen = aliasName.Length
                    End If
                End If
            Next
        Next
        Return best
    End Function

    ''' <summary>已知标签收成短码；无法识别时回退到目录第一项（zh）。</summary>
    Public Function Normalize(locale As String) As String
        Dim known As String = MatchLocale(locale)
        If known IsNot Nothing Then Return known
        Return LocaleCatalog()(0)(0)
    End Function

    Public Function LocalePath() As String
        Return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "SmartPilot", "ui-locale.txt")
    End Function

    Public Function GetLocale() As String
        Try
            Dim p As String = LocalePath()
            If File.Exists(p) Then
                Dim known As String = MatchLocale(File.ReadAllText(p, Encoding.UTF8))
                If known IsNot Nothing Then Return known
            End If
        Catch
        End Try
        Return LocaleCatalog()(0)(0)
    End Function

    ''' <summary>功能区只有中/英文案。已登记的非中文语言先用英文，避免新语言按钮掉回中文。</summary>
    Public Function IsEnglish() As Boolean
        Return GetLocale() <> "zh"
    End Function

    Public Function Text(en As String, zh As String) As String
        Return If(IsEnglish(), en, zh)
    End Function

    Public Sub SaveLocale(locale As String)
        Dim loc As String = Normalize(locale)
        Dim dir As String = Path.GetDirectoryName(LocalePath())
        If Not String.IsNullOrEmpty(dir) Then Directory.CreateDirectory(dir)
        File.WriteAllText(LocalePath(), loc, Encoding.UTF8)
    End Sub

    ''' <summary>无语言文件时按系统 UI 语言落盘，避免 getLabel 每次回落到不稳定的默认值。</summary>
    Public Sub EnsureLocaleFile()
        Try
            If File.Exists(LocalePath()) Then Return
            Dim loc As String = LocaleCatalog()(0)(0)
            Try
                Dim known As String = MatchLocale(CultureInfo.CurrentUICulture.Name)
                If known IsNot Nothing Then
                    loc = known
                ElseIf CultureInfo.CurrentUICulture.TwoLetterISOLanguageName <> "zh" Then
                    Dim en As String = MatchLocale("en")
                    If en IsNot Nothing Then loc = en
                End If
            Catch
            End Try
            SaveLocale(loc)
        Catch
        End Try
    End Sub

    Public Sub Invalidate()
        Try
            If Active IsNot Nothing Then Active.InvalidateLabels()
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' 让已打开的各个 WebView 跟上语言。PPT 主题模板等窗格和主聊天不是同一个 WebView，
    ''' postMessage / localStorage 过不去，只能由宿主执行这段脚本。
    ''' </summary>
    Public Function ApplyLocaleScript() As String
        Dim loc As String = GetLocale()
        Return "(function(){try{var a=window.App&&window.App.i18n;if(a&&typeof a.applyExternalLocale==='function')a.applyExternalLocale('" & loc & "');}catch(e){}})();"
    End Function

    Public Function PptThemePaneTitle() As String
        Return Text("PPT theme templates", "PPT 主题模板")
    End Function
End Module
