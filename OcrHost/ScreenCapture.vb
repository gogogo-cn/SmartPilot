Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows.Forms

''' <summary>
''' 浏览器里的截图走系统共享框，和 Excel 插件差太远。这里用同一套本机框选：Ctrl 截一张，Enter 识别，Esc 取消。
''' </summary>
Public Class ScreenCaptureSession

    Public Class CaptureResult
        Public Property cancelled As Boolean
        Public Property images As List(Of String)
    End Class

    Private Shared _running As Integer
    Private Shared ReadOnly _liveLock As New Object()
    Private Shared _live As New List(Of String)

    Public Shared Sub ResetLive()
        SyncLock _liveLock
            _live = New List(Of String)()
        End SyncLock
    End Sub

    Public Shared Sub NoteShot(image As String)
        If String.IsNullOrEmpty(image) Then Return
        SyncLock _liveLock
            _live.Add(image)
        End SyncLock
    End Sub

    Public Shared Function ShotsAfter(after As Integer) As List(Of String)
        SyncLock _liveLock
            Dim list As New List(Of String)
            Dim start = Math.Max(0, after)
            For i = start To _live.Count - 1
                list.Add(_live(i))
            Next
            Return list
        End SyncLock
    End Function

    Public Shared Function LiveCount() As Integer
        SyncLock _liveLock
            Return _live.Count
        End SyncLock
    End Function

    Public Shared Function CaptureScreenshots() As CaptureResult
        If Interlocked.CompareExchange(_running, 1, 0) <> 0 Then
            Return New CaptureResult With {.cancelled = True, .images = New List(Of String)()}
        End If
        ResetLive()
        Dim box As New ResultBox()
        Dim worker As New Thread(Sub() box.Value = RunOnSta())
        worker.SetApartmentState(ApartmentState.STA)
        worker.IsBackground = True
        worker.Start()
        worker.Join()
        Interlocked.Exchange(_running, 0)
        If box.Value Is Nothing Then
            Return New CaptureResult With {.cancelled = True, .images = New List(Of String)()}
        End If
        Return box.Value
    End Function

    Private Class ResultBox
        Public Value As CaptureResult
    End Class

    Private Shared Function RunOnSta() As CaptureResult
        Dim session As New ShotSession()
        Application.Run(session)
        Return session.Outcome
    End Function

    Private Class ShotSession
        Inherits ApplicationContext

        Private Const VK_CONTROL As Integer = &H11
        Private Const VK_RETURN As Integer = &HD
        Private Const VK_ESCAPE As Integer = &H1B

        <DllImport("user32.dll")>
        Private Shared Function GetAsyncKeyState(vKey As Integer) As Short
        End Function

        Public Outcome As CaptureResult
        Private ReadOnly _shots As New List(Of String)
        Private ReadOnly _tip As TipForm
        Private ReadOnly _timer As System.Windows.Forms.Timer
        Private _ctrl As Boolean
        Private _enter As Boolean
        Private _esc As Boolean
        Private _busy As Boolean
        Private _done As Boolean

        Public Sub New()
            Outcome = New CaptureResult With {.cancelled = True, .images = New List(Of String)()}
            _tip = New TipForm()
            _tip.UpdateCount(0)
            MainForm = _tip
            AddHandler _tip.FormClosed, AddressOf OnTipClosed
            _ctrl = KeyDown(VK_CONTROL)
            _enter = KeyDown(VK_RETURN)
            _esc = KeyDown(VK_ESCAPE)
            _timer = New System.Windows.Forms.Timer() With {.Interval = 30}
            AddHandler _timer.Tick, AddressOf OnTick
            _timer.Start()
        End Sub

        Private Sub OnTipClosed(sender As Object, e As EventArgs)
            If _done Then Return
            _done = True
            _timer.Stop()
            Outcome.cancelled = True
            Outcome.images = New List(Of String)()
        End Sub

        Private Sub OnTick(sender As Object, e As EventArgs)
            If _busy OrElse _done Then Return
            Dim ctrl = KeyDown(VK_CONTROL)
            Dim enter = KeyDown(VK_RETURN)
            Dim esc = KeyDown(VK_ESCAPE)

            If ctrl AndAlso Not _ctrl Then
                _ctrl = True
                _busy = True
                _timer.Stop()
                _tip.Hide()
                Application.DoEvents()
                Thread.Sleep(40)
                Dim result = DialogResult.Cancel
                Try
                    result = ShowSelector()
                Catch
                    result = DialogResult.Retry
                Finally
                    _busy = False
                End Try
                If _done Then Return
                If result = DialogResult.OK Then
                    Finish(False)
                ElseIf result = DialogResult.Cancel AndAlso _shots.Count = 0 Then
                    Finish(True)
                Else
                    _tip.UpdateCount(_shots.Count)
                    _tip.Show()
                    _ctrl = KeyDown(VK_CONTROL)
                    _timer.Start()
                End If
                Return
            End If
            _ctrl = ctrl

            If enter AndAlso Not _enter AndAlso _shots.Count > 0 Then
                Finish(False)
                Return
            End If
            _enter = enter

            If esc AndAlso Not _esc Then
                Finish(True)
                Return
            End If
            _esc = esc
        End Sub

        Private Function ShowSelector() As DialogResult
            Dim bounds = Screen.PrimaryScreen.Bounds
            Dim bmp As Bitmap = Nothing
            Try
                bmp = New Bitmap(bounds.Width, bounds.Height)
                Using g = Graphics.FromImage(bmp)
                    g.CopyFromScreen(bounds.Location, Point.Empty, bounds.Size)
                End Using
            Catch
                If bmp IsNot Nothing Then bmp.Dispose()
                Return DialogResult.Retry
            End Try
            Using form As New SelectorForm(bmp, _shots)
                Return form.ShowDialog()
            End Using
        End Function

        Private Sub Finish(cancelled As Boolean)
            If _done Then Return
            _done = True
            _timer.Stop()
            Outcome.cancelled = cancelled
            Outcome.images = If(cancelled, New List(Of String)(), New List(Of String)(_shots))
            Try
                _tip.Close()
            Catch
            End Try
        End Sub

        Private Shared Function KeyDown(vKey As Integer) As Boolean
            Return (GetAsyncKeyState(vKey) And &H8000) <> 0
        End Function
    End Class

    Private Class TipForm
        Inherits Form

        Private ReadOnly _label As Label

        Public Sub New()
            FormBorderStyle = FormBorderStyle.None
            StartPosition = FormStartPosition.Manual
            ShowInTaskbar = False
            TopMost = True
            BackColor = Color.FromArgb(32, 32, 32)
            ForeColor = Color.White
            Width = 280
            Height = 92
            _label = New Label() With {
                .Dock = DockStyle.Fill,
                .ForeColor = Color.White,
                .Font = New Font("Microsoft YaHei UI", 10),
                .TextAlign = ContentAlignment.MiddleLeft,
                .Padding = New Padding(14, 8, 14, 8)
            }
            Controls.Add(_label)
            Place()
        End Sub

        Public Sub UpdateCount(count As Integer)
            _label.Text = "已截 " & count.ToString() & " 张" & vbCrLf &
                "先切到目标窗口" & vbCrLf &
                "Ctrl 框选    Enter 识别    Esc 取消"
        End Sub

        Private Sub Place()
            Dim area = Screen.PrimaryScreen.WorkingArea
            Location = New Point(Math.Max(area.Left, area.Right - Width - 16), Math.Max(area.Top, area.Bottom - Height - 16))
        End Sub

        Protected Overrides ReadOnly Property ShowWithoutActivation As Boolean
            Get
                Return True
            End Get
        End Property

        Protected Overrides ReadOnly Property CreateParams As CreateParams
            Get
                Dim cp = MyBase.CreateParams
                cp.ExStyle = cp.ExStyle Or &H8000000
                Return cp
            End Get
        End Property
    End Class

    Private Class SelectorForm
        Inherits Form

        Private ReadOnly _bitmap As Bitmap
        Private ReadOnly _shots As List(Of String)
        Private _selecting As Boolean
        Private _start As Point
        Private _end As Point

        Public Sub New(bitmap As Bitmap, shots As List(Of String))
            _bitmap = bitmap
            _shots = shots
            FormBorderStyle = FormBorderStyle.None
            StartPosition = FormStartPosition.Manual
            ShowInTaskbar = False
            TopMost = True
            KeyPreview = True
            Cursor = Cursors.Cross
            BackColor = Color.Black
            Dim bounds = Screen.PrimaryScreen.Bounds
            Location = bounds.Location
            ClientSize = New Size(bitmap.Width, bitmap.Height)
            SetStyle(ControlStyles.AllPaintingInWmPaint Or ControlStyles.UserPaint Or ControlStyles.OptimizedDoubleBuffer, True)
            UpdateStyles()
        End Sub

        Protected Overrides Sub OnMouseDown(e As MouseEventArgs)
            MyBase.OnMouseDown(e)
            If e.Button <> MouseButtons.Left Then Return
            _selecting = True
            _start = e.Location
            _end = e.Location
            Invalidate()
        End Sub

        Protected Overrides Sub OnMouseMove(e As MouseEventArgs)
            MyBase.OnMouseMove(e)
            If Not _selecting Then Return
            _end = e.Location
            Invalidate()
        End Sub

        Protected Overrides Sub OnMouseUp(e As MouseEventArgs)
            MyBase.OnMouseUp(e)
            If e.Button <> MouseButtons.Left OrElse Not _selecting Then Return
            _selecting = False
            _end = e.Location
            Dim shot = Crop(SelectionRect())
            If shot Is Nothing Then
                Invalidate()
                Return
            End If
            _shots.Add(shot)
            ScreenCaptureSession.NoteShot(shot)
            DialogResult = DialogResult.Retry
            Close()
        End Sub

        Protected Overrides Sub OnKeyDown(e As KeyEventArgs)
            MyBase.OnKeyDown(e)
            If e.KeyCode = Keys.Escape Then
                DialogResult = If(_shots.Count > 0, DialogResult.Retry, DialogResult.Cancel)
                Close()
            ElseIf e.KeyCode = Keys.Enter AndAlso _shots.Count > 0 Then
                DialogResult = DialogResult.OK
                Close()
            End If
        End Sub

        Protected Overrides Sub OnPaint(e As PaintEventArgs)
            Dim g = e.Graphics
            g.DrawImageUnscaled(_bitmap, 0, 0)
            Dim area = ClientRectangle
            Dim hole = Rectangle.Empty
            If _selecting Then
                Dim rect = SelectionRect()
                If rect.Width > 0 AndAlso rect.Height > 0 Then hole = rect
            End If
            Using dimBrush As New SolidBrush(Color.FromArgb(110, 0, 0, 0))
                If hole.IsEmpty Then
                    g.FillRectangle(dimBrush, area)
                Else
                    g.FillRectangle(dimBrush, New Rectangle(0, 0, area.Width, Math.Max(0, hole.Top)))
                    g.FillRectangle(dimBrush, New Rectangle(0, hole.Bottom, area.Width, Math.Max(0, area.Height - hole.Bottom)))
                    g.FillRectangle(dimBrush, New Rectangle(0, hole.Top, Math.Max(0, hole.Left), hole.Height))
                    g.FillRectangle(dimBrush, New Rectangle(hole.Right, hole.Top, Math.Max(0, area.Width - hole.Right), hole.Height))
                End If
            End Using
            If Not hole.IsEmpty Then
                Using pen As New Pen(Color.Lime, 2)
                    g.DrawRectangle(pen, hole)
                End Using
                Dim sizeText = hole.Width.ToString() & " x " & hole.Height.ToString()
                Dim textRect = New Rectangle(hole.X, Math.Max(0, hole.Y - 24), 120, 22)
                Using textBrush As New SolidBrush(Color.FromArgb(200, 26, 115, 232))
                    g.FillRectangle(textBrush, textRect)
                End Using
                Using font As New Font("Arial", 11, FontStyle.Bold)
                    g.DrawString(sizeText, font, Brushes.White, textRect.X + 4, textRect.Y + 2)
                End Using
            End If
        End Sub

        Private Function SelectionRect() As Rectangle
            Dim x = Math.Min(_start.X, _end.X)
            Dim y = Math.Min(_start.Y, _end.Y)
            Return New Rectangle(x, y, Math.Abs(_end.X - _start.X), Math.Abs(_end.Y - _start.Y))
        End Function

        Private Function Crop(rect As Rectangle) As String
            Dim bounds = New Rectangle(0, 0, _bitmap.Width, _bitmap.Height)
            Dim hit = Rectangle.Intersect(rect, bounds)
            If hit.Width < 6 OrElse hit.Height < 6 Then Return Nothing
            Using piece As New Bitmap(hit.Width, hit.Height)
                Using g = Graphics.FromImage(piece)
                    g.DrawImage(_bitmap, New Rectangle(0, 0, hit.Width, hit.Height), hit, GraphicsUnit.Pixel)
                End Using
                Using ms As New System.IO.MemoryStream()
                    piece.Save(ms, ImageFormat.Png)
                    Return Convert.ToBase64String(ms.ToArray())
                End Using
            End Using
        End Function

        Protected Overrides Sub Dispose(disposing As Boolean)
            If disposing AndAlso _bitmap IsNot Nothing Then _bitmap.Dispose()
            MyBase.Dispose(disposing)
        End Sub
    End Class

End Class
