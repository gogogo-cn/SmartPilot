Option Strict On
Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports System.Threading.Tasks
Imports Microsoft.Web.WebView2.WinForms

Public Class ScreenshotService
        Implements IDisposable

        Private _logger As Utils.ChatLogger
        Private _ocrWebView As WebView2
        Private _isActive As Boolean = False
        Private _isSelecting As Boolean = False
        Private _selectionStart As Point
        Private _selectionEnd As Point
        Private _selectionRect As Rectangle
        Private _screenshotBitmap As Bitmap
        Private _screenshotForm As Form
        Private _tipForm As Form
        Private _pollThread As Threading.Thread
        Private _pollCts As Threading.CancellationTokenSource
        Private _isCapturing As Boolean = False
        Private _shots As New List(Of String)
        Private _pushedCount As Integer = 0
        Private _thumbImages As New List(Of Image)
        Private _thumbHost As FlowLayoutPanel
        Private _tipPreview As PictureBox
        Private _ctrlHeld As Boolean = False
        Private _enterHeld As Boolean = False
        Private _escHeld As Boolean = False

        Private Const VK_CONTROL As Integer = &H11
        Private Const VK_RETURN As Integer = &HD
        Private Const VK_ESCAPE As Integer = &H1B

        <DllImport("user32.dll")>
        Private Shared Function GetAsyncKeyState(ByVal vKey As Integer) As Short
        End Function

        <DllImport("user32.dll")>
        Private Shared Function GetCursorPos(ByRef lpPoint As Point) As Boolean
        End Function

        <DllImport("user32.dll")>
        Private Shared Function ShowWindow(ByVal hWnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
        End Function

        <DllImport("user32.dll")>
        Private Shared Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
        End Function

        Private Const SW_RESTORE As Integer = 9

        Public Sub New(logger As Utils.ChatLogger)
            _logger = logger
        End Sub

        ''' <summary>
        ''' Use OCR WebView as UI thread dispatcher
        ''' </summary>
        Private Sub RunOnUiThread(action As Action)
            If _ocrWebView IsNot Nothing Then
                If _ocrWebView.InvokeRequired Then
                    _ocrWebView.Invoke(action)
                Else
                    action()
                End If
            Else
                Try
                    ThisAddIn.RunOnUIThread(action)
                Catch ex As Exception  ' 
                    action()
                End Try
            End If
        End Sub

        Private _preCreatedTip As Form
        Private _preCreatedLabel As Label

        Private Sub EnsurePreCreatedTip()
            If _preCreatedTip IsNot Nothing Then Return

            _preCreatedTip = New Form() With {
                .FormBorderStyle = FormBorderStyle.FixedToolWindow,
                .StartPosition = FormStartPosition.Manual,
                .TopMost = True,
                .ShowInTaskbar = False,
                .Text = "连续截图",
                .Width = 400,
                .Height = 332,
                .BackColor = Color.FromArgb(244, 246, 248),
                .Padding = New Padding(8, 8, 8, 6)
            }

            _preCreatedLabel = New Label() With {
                .Text = TipText(),
                .ForeColor = Color.FromArgb(55, 65, 81),
                .BackColor = Color.White,
                .Font = New Font("Microsoft YaHei UI", 9),
                .TextAlign = ContentAlignment.MiddleLeft,
                .Dock = DockStyle.Bottom,
                .Height = 46,
                .Padding = New Padding(10, 0, 8, 0)
            }

            _thumbHost = New FlowLayoutPanel() With {
                .Height = 70,
                .Dock = DockStyle.Bottom,
                .FlowDirection = FlowDirection.LeftToRight,
                .WrapContents = False,
                .AutoScroll = True,
                .BackColor = Color.FromArgb(244, 246, 248),
                .Padding = New Padding(0, 4, 0, 0)
            }

            _tipPreview = New PictureBox() With {
                .Dock = DockStyle.Fill,
                .SizeMode = PictureBoxSizeMode.Zoom,
                .BackColor = Color.White
            }

            ' 后加入的控件先停靠：提示条贴底，缩略图在它上面，预览占满剩下的区域。
            _preCreatedTip.Controls.Add(_tipPreview)
            _preCreatedTip.Controls.Add(_thumbHost)
            _preCreatedTip.Controls.Add(_preCreatedLabel)
            PlaceTip()
        End Sub

        Private Sub PlaceTip()
            If _preCreatedTip Is Nothing OrElse _preCreatedTip.IsDisposed Then Return
            Dim area = Screen.PrimaryScreen.WorkingArea
            _preCreatedTip.Location = New Point(
                Math.Max(area.Left, area.Right - _preCreatedTip.Width - 16),
                Math.Max(area.Top, area.Bottom - _preCreatedTip.Height - 16))
        End Sub

        Public Sub SetOcrWebView(webView As WebView2)
            _ocrWebView = webView
        End Sub

        Public Sub StartOcr()
            Try
                If _isActive Then
                    _logger?.LogInfo("SCREENSHOT", "Screenshot already active, ignoring request")
                    Return
                End If

                _isActive = True
                _shots.Clear()
                _pushedCount = 0
                ClearTipThumbs()
                _ctrlHeld = False
                _enterHeld = False
                _escHeld = False
                _logger?.LogInfo("SCREENSHOT", "Starting continuous screenshot")

                QueueScript("window.beginScreenshotBatch && window.beginScreenshotBatch();")
                StartPollingOnUi()
                ShowTip()

            Catch ex As Exception
                _isActive = False
                HideStartupLoading()
                _logger?.LogError("SCREENSHOT", "Failed to start screenshot", ex)
                MessageBox.Show("Failed to start screenshot: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private _startupLoadingForm As Form
        Private _startupLoadingBar As ProgressBar
        Private _startupLoadingLabel As Label

        Private Sub ShowStartupLoading()
            Try
                _startupLoadingForm = New Form() With {
                    .FormBorderStyle = FormBorderStyle.None,
                    .StartPosition = FormStartPosition.CenterScreen,
                    .TopMost = True,
                    .Width = 320,
                    .Height = 120,
                    .BackColor = Color.FromArgb(30, 30, 30),
                    .ShowInTaskbar = False
                }

                _startupLoadingLabel = New Label() With {
                    .Text = "Initializing hotkey monitor...",
                    .ForeColor = Color.White,
                    .Font = New Font("Segoe UI", 10),
                    .TextAlign = ContentAlignment.MiddleCenter,
                    .Dock = DockStyle.Top,
                    .Height = 35
                }

                _startupLoadingBar = New ProgressBar() With {
                    .Dock = DockStyle.Fill,
                    .Style = ProgressBarStyle.Marquee,
                    .MarqueeAnimationSpeed = 30,
                    .Height = 40,
                    .BackColor = Color.FromArgb(60, 60, 60)
                }

                _startupLoadingForm.Controls.Add(_startupLoadingBar)
                _startupLoadingForm.Controls.Add(_startupLoadingLabel)
                _startupLoadingForm.Show()
            Catch ex As Exception  ' 
            End Try
        End Sub

        Private Sub UpdateStartupLoading(text As String)
            Try
                If _startupLoadingForm Is Nothing OrElse _startupLoadingForm.IsDisposed Then Return
                If _startupLoadingForm.InvokeRequired Then
                    _startupLoadingForm.Invoke(New Action(Sub() UpdateStartupLoading(text)))
                    Return
                End If
                If _startupLoadingLabel IsNot Nothing Then _startupLoadingLabel.Text = text
            Catch ex As Exception  ' 
            End Try
        End Sub

        Private Sub HideStartupLoading()
            Try
                If _startupLoadingForm IsNot Nothing AndAlso Not _startupLoadingForm.IsDisposed Then
                    _startupLoadingForm.Close()
                    _startupLoadingForm.Dispose()
                    _startupLoadingForm = Nothing
                End If
            Catch ex As Exception  ' 
            End Try
        End Sub

        Private Function TipText() As String
            Return "已截 " & _shots.Count.ToString() & " 张" & vbCrLf &
                   "Ctrl  下一张      Enter  识别全部      Esc  取消"
        End Function

        Private Sub UpdateTip()
            Try
                If _preCreatedLabel Is Nothing Then Return
                _preCreatedLabel.Text = TipText()
            Catch ex As Exception
            End Try
        End Sub

        Private Sub StartPollingOnUi()
            If _pollCts IsNot Nothing Then
                Try
                    _pollCts.Cancel()
                    _pollCts.Dispose()
                Catch ex As Exception
                End Try
            End If
            _pollCts = New Threading.CancellationTokenSource()
            _pollThread = New Threading.Thread(Sub() PollLoop(_pollCts.Token))
            _pollThread.IsBackground = True
            _pollThread.Priority = Threading.ThreadPriority.AboveNormal
            _pollThread.Start()
        End Sub

        Private Sub PollLoop(ct As Threading.CancellationToken)
            While Not ct.IsCancellationRequested
                If Not _isActive Then Exit While

                Dim ctrlDown = (GetAsyncKeyState(VK_CONTROL) And &H8000) <> 0
                Dim enterDown = (GetAsyncKeyState(VK_RETURN) And &H8000) <> 0
                Dim escDown = (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0

                If ctrlDown AndAlso Not _ctrlHeld AndAlso Not _isCapturing Then
                    _isCapturing = True
                    _ctrlHeld = True
                    InvokeCapture()
                    Exit While
                End If
                _ctrlHeld = ctrlDown

                If enterDown AndAlso Not _enterHeld AndAlso _shots.Count > 0 AndAlso Not _isCapturing Then
                    _enterHeld = True
                    RunOnUiThread(New Action(AddressOf FinishAndRecognize))
                    Exit While
                End If
                _enterHeld = enterDown

                If escDown AndAlso Not _escHeld Then
                    _escHeld = True
                    RunOnUiThread(New Action(AddressOf CancelSession))
                    Exit While
                End If
                _escHeld = escDown

                Threading.Thread.Sleep(15)
            End While
        End Sub

        Private Sub InvokeCapture()
            Try
                RunOnUiThread(New Action(AddressOf DoCapture))
            Catch ex As Exception  ' 
                DoCapture()
            End Try
        End Sub

        Private Sub DoCapture()
            HideTip()
            HideStartupLoading()
            ShowSelector()
        End Sub

        ''' <summary>
        ''' 先抄屏幕，再弹出选区。选区窗是半透明黑底，若先显示再 CopyFromScreen，裁出来的图就是这层黑底。
        ''' </summary>
        Private Function CaptureCleanScreen() As Boolean
            Try
                Dim bounds = Screen.PrimaryScreen.Bounds
                Dim bmp As New Bitmap(bounds.Width, bounds.Height)
                Using g = Graphics.FromImage(bmp)
                    g.CopyFromScreen(bounds.Location, Point.Empty, bounds.Size)
                End Using
                If _screenshotBitmap IsNot Nothing Then
                    _screenshotBitmap.Dispose()
                End If
                _screenshotBitmap = bmp
                _logger?.LogInfo("SCREENSHOT", $"Screen captured: {bounds.Width}x{bounds.Height}")
                Return True
            Catch ex As Exception
                _logger?.LogError("SCREENSHOT", "Failed to capture screen", ex)
                Return False
            End Try
        End Function

        Private Sub ShowTip()
            EnsurePreCreatedTip()
            UpdateTip()
            PlaceTip()
            _preCreatedTip.Show()
        End Sub

        Private Sub HideTip()
            Try
                If _preCreatedTip IsNot Nothing AndAlso Not _preCreatedTip.IsDisposed Then
                    _preCreatedTip.Hide()
                End If
            Catch ex As Exception  ' 
            End Try
        End Sub

        Private Sub ShowSelector()
            If Not CaptureCleanScreen() Then
                _isCapturing = False
                RestartWait()
                Return
            End If

            _screenshotForm = New Form() With {
                .FormBorderStyle = FormBorderStyle.None,
                .WindowState = FormWindowState.Maximized,
                .BackColor = Color.Black,
                .Cursor = Cursors.Cross,
                .TopMost = True,
                .Opacity = 1.0,
                .ShowInTaskbar = False,
                .KeyPreview = True
            }

            Dim overlay = New PictureBox() With {
                .Dock = DockStyle.Fill,
                .BackColor = Color.Black,
                .SizeMode = PictureBoxSizeMode.Normal,
                .Image = _screenshotBitmap,
                .Cursor = Cursors.Cross
            }

            AddHandler overlay.MouseDown,
                Sub(s, e)
                    If e.Button = MouseButtons.Left Then
                        _isSelecting = True
                        _selectionStart = e.Location
                        _selectionEnd = e.Location
                    End If
                End Sub

            AddHandler overlay.MouseMove,
                Sub(s, e)
                    If _isSelecting Then
                        _selectionEnd = e.Location
                        overlay.Invalidate()
                    End If
                End Sub

            AddHandler overlay.MouseUp,
                Sub(s, e)
                    If e.Button = MouseButtons.Left AndAlso _isSelecting Then
                        _isSelecting = False
                        _selectionEnd = e.Location
                        CalculateSelectionRect()

                        If _selectionRect.Width > 5 AndAlso _selectionRect.Height > 5 AndAlso _screenshotBitmap IsNot Nothing Then
                            Dim thumb As Image = Nothing
                            Dim shot = CropSelectionBase64(thumb)
                            If shot <> "" Then
                                _shots.Add(shot)
                                If thumb IsNot Nothing Then AddTipThumb(thumb)
                            ElseIf thumb IsNot Nothing Then
                                thumb.Dispose()
                            End If
                            _screenshotForm.DialogResult = DialogResult.Retry
                            _screenshotForm.Close()
                        End If
                    End If
                End Sub

            AddHandler overlay.Paint,
                Sub(s, e)
                    Dim area = overlay.ClientRectangle
                    Dim hole = Rectangle.Empty
                    If _isSelecting Then
                        Dim rect = GetCurrentSelectionRect()
                        If rect.Width > 0 AndAlso rect.Height > 0 Then hole = rect
                    End If
                    Using dimBrush As New SolidBrush(Color.FromArgb(110, 0, 0, 0))
                        If hole.IsEmpty Then
                            e.Graphics.FillRectangle(dimBrush, area)
                        Else
                            e.Graphics.FillRectangle(dimBrush, New Rectangle(0, 0, area.Width, Math.Max(0, hole.Top)))
                            e.Graphics.FillRectangle(dimBrush, New Rectangle(0, hole.Bottom, area.Width, Math.Max(0, area.Height - hole.Bottom)))
                            e.Graphics.FillRectangle(dimBrush, New Rectangle(0, hole.Top, Math.Max(0, hole.Left), hole.Height))
                            e.Graphics.FillRectangle(dimBrush, New Rectangle(hole.Right, hole.Top, Math.Max(0, area.Width - hole.Right), hole.Height))
                        End If
                    End Using
                    If Not hole.IsEmpty Then
                        Using pen As New Pen(Color.Lime, 2)
                            e.Graphics.DrawRectangle(pen, hole)
                        End Using
                        Dim sizeText = hole.Width.ToString() & " x " & hole.Height.ToString()
                        Dim textRect = New Rectangle(hole.X, Math.Max(0, hole.Y - 24), 120, 22)
                        Using textBrush As New SolidBrush(Color.FromArgb(200, 26, 115, 232))
                            e.Graphics.FillRectangle(textBrush, textRect)
                        End Using
                        Using font As New Font("Arial", 11, FontStyle.Bold)
                            e.Graphics.DrawString(sizeText, font, Brushes.White, textRect.X + 4, textRect.Y + 2)
                        End Using
                    End If
                End Sub

            AddHandler _screenshotForm.Shown,
                Sub(s, e)
                    Try
                        _screenshotForm.BringToFront()
                        _screenshotForm.Activate()
                    Catch ex As Exception  ' 
                    End Try
                End Sub

            AddHandler _screenshotForm.KeyDown,
                Sub(s, e)
                    If e.KeyCode = Keys.Escape Then
                        _screenshotForm.DialogResult = If(_shots.Count > 0, DialogResult.Retry, DialogResult.Cancel)
                        _screenshotForm.Close()
                    ElseIf e.KeyCode = Keys.Enter AndAlso _shots.Count > 0 Then
                        _screenshotForm.DialogResult = DialogResult.OK
                        _screenshotForm.Close()
                    End If
                End Sub

            _screenshotForm.Controls.Add(overlay)
            _screenshotForm.ShowDialog()

            Dim result = DialogResult.Cancel
            Try
                If _screenshotForm IsNot Nothing AndAlso Not _screenshotForm.IsDisposed Then
                    result = _screenshotForm.DialogResult
                End If
            Catch ex As Exception
            End Try
            ReleaseClosedSelector()

            If result = DialogResult.OK Then
                FinishAndRecognize()
            ElseIf result = DialogResult.Retry Then
                _isCapturing = False
                RestartWait()
            Else
                CancelSession()
            End If
        End Sub

        Private Sub ReleaseClosedSelector()
            Dim form = _screenshotForm
            _screenshotForm = Nothing
            If form Is Nothing Then Return
            Try
                If Not form.IsDisposed Then form.Dispose()
            Catch ex As Exception
            End Try
        End Sub

        Private Function GetMousePositionOnScreen() As Point
            Dim pt As New Point()
            GetCursorPos(pt)
            Return pt
        End Function

        Private Function GetCurrentSelectionRect() As Rectangle
            Dim x = Math.Min(_selectionStart.X, _selectionEnd.X)
            Dim y = Math.Min(_selectionStart.Y, _selectionEnd.Y)
            Dim width = Math.Abs(_selectionEnd.X - _selectionStart.X)
            Dim height = Math.Abs(_selectionEnd.Y - _selectionStart.Y)
            Return New Rectangle(x, y, width, height)
        End Function

        Private Sub CalculateSelectionRect()
            _selectionRect = GetCurrentSelectionRect()
        End Sub

        Private Function CropSelectionBase64(ByRef thumb As Image) As String
            thumb = Nothing
            Dim rect = _selectionRect
            Dim sourceBmp = _screenshotBitmap
            If sourceBmp Is Nothing OrElse rect.Width <= 0 OrElse rect.Height <= 0 Then Return ""
            Dim cropped As Bitmap = Nothing
            Dim ms As System.IO.MemoryStream = Nothing
            Try
                cropped = New Bitmap(rect.Width, rect.Height)
                Using g = Graphics.FromImage(cropped)
                    g.DrawImage(sourceBmp,
                        New Rectangle(0, 0, cropped.Width, cropped.Height),
                        rect,
                        GraphicsUnit.Pixel)
                End Using
                thumb = MakeThumb(cropped)
                ms = New System.IO.MemoryStream()
                cropped.Save(ms, Imaging.ImageFormat.Png)
                Return Convert.ToBase64String(ms.ToArray())
            Finally
                If ms IsNot Nothing Then ms.Dispose()
                If cropped IsNot Nothing Then cropped.Dispose()
            End Try
        End Function

        Private Function MakeThumb(source As Bitmap) As Bitmap
            Dim maxEdge = 280.0R
            Dim scale = Math.Min(maxEdge / source.Width, maxEdge / source.Height)
            If scale > 1 Then scale = 1
            Dim w = Math.Max(1, CInt(Math.Round(source.Width * scale)))
            Dim h = Math.Max(1, CInt(Math.Round(source.Height * scale)))
            Dim thumb As New Bitmap(w, h)
            Using g = Graphics.FromImage(thumb)
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
                g.DrawImage(source, New Rectangle(0, 0, w, h))
            End Using
            Return thumb
        End Function

        Private Sub AddTipThumb(thumb As Image)
            _thumbImages.Add(thumb)
            If _tipPreview IsNot Nothing Then _tipPreview.Image = thumb
            If _thumbHost Is Nothing Then Return
            For Each c As Control In _thumbHost.Controls
                c.BackColor = Color.FromArgb(209, 213, 219)
            Next
            Dim frame As New Panel() With {
                .Width = 76,
                .Height = 56,
                .Margin = New Padding(6, 4, 0, 4),
                .Padding = New Padding(2),
                .BackColor = Color.FromArgb(26, 115, 232)
            }
            Dim box As New PictureBox() With {
                .Dock = DockStyle.Fill,
                .Image = thumb,
                .SizeMode = PictureBoxSizeMode.Zoom,
                .BackColor = Color.White
            }
            frame.Controls.Add(box)
            _thumbHost.Controls.Add(frame)
            _thumbHost.ScrollControlIntoView(frame)
        End Sub

        Private Sub ReleaseThumbImages(parent As Control)
            For Each c As Control In parent.Controls
                Dim pb = TryCast(c, PictureBox)
                If pb IsNot Nothing Then pb.Image = Nothing
                If c.HasChildren Then ReleaseThumbImages(c)
            Next
        End Sub

        Private Sub ClearTipThumbs()
            Try
                If _tipPreview IsNot Nothing Then _tipPreview.Image = Nothing
                If _thumbHost IsNot Nothing Then
                    ReleaseThumbImages(_thumbHost)
                    _thumbHost.Controls.Clear()
                End If
                For Each img In _thumbImages
                    img.Dispose()
                Next
                _thumbImages.Clear()
            Catch ex As Exception
            End Try
        End Sub

        Private Function QueueScript(code As String) As Boolean
            Try
                If _ocrWebView Is Nothing OrElse _ocrWebView.CoreWebView2 Is Nothing Then Return False
                Dim task = _ocrWebView.CoreWebView2.ExecuteScriptAsync(code)
                Return True
            Catch ex As Exception
                _logger?.LogError("SCREENSHOT", "Queue script failed", ex)
                Return False
            End Try
        End Function

        Private Sub PushPendingShots()
            While _pushedCount < _shots.Count
                Dim safe = _shots(_pushedCount).Replace("'", "\'")
                If Not QueueScript("window.appendScreenshot('" & safe & "');") Then Exit While
                _pushedCount += 1
            End While
        End Sub

        Private Sub RestartWait()
            _isCapturing = False
            PushPendingShots()
            UpdateTip()
            ShowTip()
            ActivateExcelWindow()
            StartPollingOnUi()
        End Sub

        Private Sub FinishAndRecognize()
            HideTip()
            Dim shots = New List(Of String)(_shots)
            Dim pushed = _pushedCount
            _shots.Clear()
            _pushedCount = 0
            _isActive = False
            _isCapturing = False
            ClearTipThumbs()
            ActivateExcelWindow()
            FinishDelivery(shots, pushed)
            Cleanup()
        End Sub

        Private Sub CancelSession()
            _shots.Clear()
            _pushedCount = 0
            _isActive = False
            _isCapturing = False
            HideTip()
            ClearTipThumbs()
            QueueScript("window.clearScreenshotStrip && window.clearScreenshotStrip();")
            Cleanup()
        End Sub

        Private Async Sub FinishDelivery(shots As List(Of String), alreadyPushed As Integer)
            Try
                If shots Is Nothing OrElse shots.Count = 0 Then Return
                If _ocrWebView Is Nothing OrElse _ocrWebView.CoreWebView2 Is Nothing Then
                    _logger?.LogError("SCREENSHOT", "OCR WebView not ready for batch")
                    Return
                End If
                If alreadyPushed <= 0 Then
                    Await _ocrWebView.CoreWebView2.ExecuteScriptAsync("window.beginScreenshotBatch && window.beginScreenshotBatch();")
                    alreadyPushed = 0
                End If
                For i As Integer = alreadyPushed To shots.Count - 1
                    Dim safe = shots(i).Replace("'", "\'")
                    Await _ocrWebView.CoreWebView2.ExecuteScriptAsync("window.appendScreenshot('" & safe & "');")
                Next
                Await _ocrWebView.CoreWebView2.ExecuteScriptAsync("window.recognizeScreenshotBatch && window.recognizeScreenshotBatch();")
                _logger?.LogInfo("SCREENSHOT", "Recognizing " & shots.Count.ToString() & " screenshots")
            Catch ex As Exception
                _logger?.LogError("SCREENSHOT", "Deliver batch failed", ex)
            End Try
        End Sub

        Private Sub ProcessSelection()
            Try
                If _selectionRect.Width = 0 OrElse _selectionRect.Height = 0 Then
                    _logger?.LogWarning("SCREENSHOT", "Invalid selection area")
                    Return
                End If

                _logger?.LogInfo("SCREENSHOT", $"Selection: {_selectionRect.Width} x {_selectionRect.Height}")

                ShowLoadingOverlay()

                Dim rect = _selectionRect
                Dim sourceBmp = _screenshotBitmap

                Task.Run(Sub()
                    Try
                        UpdateLoadingProgress(10, "Cropping screenshot...")
                        Dim croppedBitmap = New Bitmap(rect.Width, rect.Height)
                        Using g = Graphics.FromImage(croppedBitmap)
                            g.DrawImage(sourceBmp,
                                New Rectangle(0, 0, croppedBitmap.Width, croppedBitmap.Height),
                                rect,
                                GraphicsUnit.Pixel)
                        End Using

                        UpdateLoadingProgress(40, "Converting format...")
                        Dim ms = New System.IO.MemoryStream()
                        croppedBitmap.Save(ms, Imaging.ImageFormat.Png)
                        Dim imageBytes = ms.ToArray()
                        Dim base64String = Convert.ToBase64String(imageBytes)

                        croppedBitmap.Dispose()
                        ms.Dispose()

                        _logger?.LogInfo("SCREENSHOT", $"Cropped, base64 len={base64String.Length}")

                        UpdateLoadingProgress(70, "Sending to OCR panel...")

                        RunOnUiThread(New Action(Sub()
                            Try
                                ActivateExcelWindow()
                                SendToOcr(base64String)
                            Catch ex2 As Exception
                                _logger?.LogError("SCREENSHOT", "UI thread send failed: " & ex2.Message, ex2)
                            Finally
                                _isActive = False
                                _isCapturing = False
                                HideLoadingOverlay()
                                Cleanup()
                            End Try
                        End Sub))

                    Catch ex As Exception
                        _logger?.LogError("SCREENSHOT", "Background crop failed", ex)
                        RunOnUiThread(New Action(Sub()
                            HideLoadingOverlay()
                            MessageBox.Show("Failed to process screenshot: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                            _isActive = False
                            _isCapturing = False
                            Cleanup()
                        End Sub))
                    End Try
                End Sub)

            Catch ex As Exception
                _logger?.LogError("SCREENSHOT", "Process selection failed", ex)
                HideLoadingOverlay()
                MessageBox.Show("Failed to process screenshot: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                _isActive = False
                _isCapturing = False
                Cleanup()
            End Try
        End Sub

        Private _loadingForm As Form
        Private _loadingBar As ProgressBar
        Private _loadingLabel As Label

        Private Sub ShowLoadingOverlay()
            Try
                _loadingForm = New Form() With {
                    .FormBorderStyle = FormBorderStyle.None,
                    .StartPosition = FormStartPosition.CenterScreen,
                    .TopMost = True,
                    .Width = 320,
                    .Height = 120,
                    .BackColor = Color.FromArgb(30, 30, 30),
                    .ShowInTaskbar = False
                }

                _loadingLabel = New Label() With {
                    .Text = "Processing screenshot...",
                    .ForeColor = Color.White,
                    .Font = New Font("Segoe UI", 10),
                    .TextAlign = ContentAlignment.MiddleCenter,
                    .Dock = DockStyle.Top,
                    .Height = 35
                }

                _loadingBar = New ProgressBar() With {
                    .Dock = DockStyle.Fill,
                    .Style = ProgressBarStyle.Continuous,
                    .Minimum = 0,
                    .Maximum = 100,
                    .Value = 0,
                    .Height = 40,
                    .ForeColor = Color.FromArgb(26, 115, 232),
                    .BackColor = Color.FromArgb(60, 60, 60)
                }

                _loadingForm.Controls.Add(_loadingBar)
                _loadingForm.Controls.Add(_loadingLabel)
                _loadingForm.Show()
            Catch ex As Exception  ' 
            End Try
        End Sub

        Private Sub UpdateLoadingProgress(value As Integer, text As String)
            Try
                If _loadingForm Is Nothing OrElse _loadingForm.IsDisposed Then Return
                If _loadingForm.InvokeRequired Then
                    _loadingForm.Invoke(New Action(Sub() UpdateLoadingProgress(value, text)))
                    Return
                End If
                If _loadingBar IsNot Nothing Then _loadingBar.Value = Math.Min(100, Math.Max(0, value))
                If _loadingLabel IsNot Nothing Then _loadingLabel.Text = text
            Catch ex As Exception  ' 
            End Try
        End Sub

        Private Sub HideLoadingOverlay()
            Try
                If _loadingForm IsNot Nothing AndAlso Not _loadingForm.IsDisposed Then
                    _loadingForm.Close()
                    _loadingForm.Dispose()
                    _loadingForm = Nothing
                End If
            Catch ex As Exception  ' 
            End Try
        End Sub

        Private Sub ActivateExcelWindow()
            Try
                If Globals.ThisAddIn IsNot Nothing AndAlso Globals.ThisAddIn.Application IsNot Nothing Then
                    Dim hwnd As IntPtr = New IntPtr(Globals.ThisAddIn.Application.Hwnd)
                    If hwnd <> IntPtr.Zero Then
                        ShowWindow(hwnd, SW_RESTORE)
                        SetForegroundWindow(hwnd)
                        _logger?.LogInfo("SCREENSHOT", "Excel window activated")
                    End If
                End If
            Catch ex As Exception  ' 
            End Try
        End Sub

        Private Sub SendToOcr(base64Image As String)
            Try
                _logger?.LogInfo("SCREENSHOT", $"SendToOcr: _ocrWebView={If(_ocrWebView Is Nothing, "Nothing", "OK")}, CoreWebView2={If(_ocrWebView?.CoreWebView2 Is Nothing, "Nothing", "OK")}")

                If _ocrWebView Is Nothing OrElse _ocrWebView.CoreWebView2 Is Nothing Then
                    Dim live = TaskPaneManager.Current?.OcrWebView
                    If live IsNot Nothing AndAlso live.CoreWebView2 IsNot Nothing Then
                        _ocrWebView = live
                    End If
                End If

                If _ocrWebView Is Nothing OrElse _ocrWebView.CoreWebView2 Is Nothing Then
                    _logger?.LogInfo("SCREENSHOT", "OCR WebView not ready, creating panel")
                    Try
                        Globals.ThisAddIn.ShowOcrPanel()
                    Catch ex As Exception  ' 
                    End Try

                    Dim capturedBase64 = base64Image
                    Task.Run(Sub()
                        Try
                            Dim retryCount = 0
                            While _ocrWebView Is Nothing AndAlso retryCount < 30
                                System.Threading.Thread.Sleep(100)
                                retryCount += 1
                                _logger?.LogInfo("SCREENSHOT", $"Waiting for WebView2... ({retryCount}/30)")
                            End While

                            RunOnUiThread(New Action(Sub()
                                If _ocrWebView IsNot Nothing AndAlso _ocrWebView.CoreWebView2 IsNot Nothing Then
                                    ExecuteScreenshotJs(capturedBase64)
                                Else
                                    _logger?.LogError("SCREENSHOT", "WebView2 creation timeout")
                                    MessageBox.Show("OCR panel initialization timeout", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                End If
                            End Sub))
                        Catch ex As Exception
                            _logger?.LogError("SCREENSHOT", "Background wait for WebView2 failed", ex)
                        End Try
                    End Sub)
                    Return
                End If

                ExecuteScreenshotJs(base64Image)

            Catch ex As Exception
                _logger?.LogError("SCREENSHOT", "Send to OCR failed", ex)
                MessageBox.Show("Send to OCR failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub ExecuteScreenshotJs(base64Image As String)
            Try
                If _ocrWebView Is Nothing OrElse _ocrWebView.CoreWebView2 Is Nothing Then
                    _logger?.LogError("SCREENSHOT", "OCR WebView still not initialized")
                    MessageBox.Show("OCR panel not ready, please open OCR panel first", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If

                Dim safeBase64 = base64Image.Replace("'", "\'")
                Dim jsCode = "window.loadScreenshotFromBase64('" & safeBase64 & "');"
                _logger?.LogInfo("SCREENSHOT", $"Executing JS, base64 len={base64Image.Length}")

                Dim task = _ocrWebView.CoreWebView2.ExecuteScriptAsync(jsCode)
                task.ContinueWith(Sub(t)
                    Try
                        If t.IsFaulted Then
                            _logger?.LogError("SCREENSHOT", "ExecuteScript failed: " & t.Exception?.Message, t.Exception)
                        Else
                            _logger?.LogInfo("SCREENSHOT", $"ExecuteScript result: {t.Result}")
                        End If
                    Catch ex As Exception
                        _logger?.LogError("SCREENSHOT", "ContinueWith exception: " & ex.Message, ex)
                    End Try
                End Sub)
            Catch ex As Exception
                _logger?.LogError("SCREENSHOT", "ExecuteScreenshotJs failed", ex)
            End Try
        End Sub

        Private Sub Cleanup()
            Try
                If _pollCts IsNot Nothing Then
                    _pollCts.Cancel()
                    _pollCts.Dispose()
                    _pollCts = Nothing
                End If

                If _pollThread IsNot Nothing Then
                    _pollThread = Nothing
                End If

                If _preCreatedTip IsNot Nothing AndAlso Not _preCreatedTip.IsDisposed Then
                    _preCreatedTip.Hide()
                End If

                If _screenshotBitmap IsNot Nothing Then
                    _screenshotBitmap.Dispose()
                    _screenshotBitmap = Nothing
                End If

                If _screenshotForm IsNot Nothing Then
                    Try
                        _screenshotForm.Dispose()
                    Catch ex As Exception  ' 
                    End Try
                    _screenshotForm = Nothing
                End If

                _isSelecting = False
                _selectionRect = Rectangle.Empty

                _logger?.LogInfo("SCREENSHOT", "Screenshot resources cleaned up")

            Catch ex As Exception
                _logger?.LogError("SCREENSHOT", "Failed to clean up screenshot resources", ex)
            End Try
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            _isActive = False
            _isCapturing = False
            Try
                If _pollCts IsNot Nothing Then
                    _pollCts.Cancel()
                    _pollCts.Dispose()
                    _pollCts = Nothing
                End If
                If _preCreatedTip IsNot Nothing Then
                    _preCreatedTip.Close()
                    _preCreatedTip.Dispose()
                    _preCreatedTip = Nothing
                End If
            Catch ex As Exception  ' 
            End Try
            Cleanup()
        End Sub

    End Class