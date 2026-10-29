Imports System.IO
Imports System.Drawing
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports ExcelAIVSTOAddIn.Services

Namespace Forms

    ''' <summary>
    ''' 图片OCR识别对话框 - 重新设计版
    ''' </summary>
    Public Class ImageOcrDialog
        Inherits Form

        ' 控件声明
        Private WithEvents btnSelectImage As Button
        Private WithEvents btnRecognize As Button
        Private WithEvents btnClose As Button
        Private WithEvents btnCopyResult As Button
        Private WithEvents btnInsertToExcel As Button
        Private WithEvents btnClearResult As Button
        Private WithEvents picPreview As PictureBox
        Private WithEvents cbEngine As ComboBox
        Private txtResult As TextBox
        Private lblStatus As Label
        Private progressBar As ProgressBar
        Private lblImageInfo As Label
        Private pnlEngineStatus As Panel
        Private lblEngineStatus As Label
        Private lblDragHint As Label
        Private pnlPreviewOverlay As Panel

        ' 服务
        Private _ocrManager As OcrServiceManager

        ' 当前图片路径
        Private _currentImagePath As String
        Private _recognitionResult As String

        ' Excel应用程序引用
        Private _excelApp As Microsoft.Office.Interop.Excel.Application

        ''' <summary>
        ''' 引擎列表项
        ''' </summary>
        Private Class EngineItem
            Public Property Engine As OcrEngineType
            Public Property Display As String
            Public Overrides Function ToString() As String
                Return Display
            End Function
        End Class

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        Public Sub New(excelApp As Microsoft.Office.Interop.Excel.Application)
            InitializeComponent()
            _excelApp = excelApp
            _ocrManager = OcrServiceManager.Instance

            PopulateEngineList()
            SyncEngineComboBox()
            UpdateStatus("就绪", StatusType.Ready)

            DetectEnginesAsync()
        End Sub

        ''' <summary>
        ''' 后台异步检测各引擎可用性
        ''' </summary>
        Private Async Sub DetectEnginesAsync()
            Try
                Await Task.Run(Sub()
                    System.Threading.Thread.Sleep(100)
                End Sub)

                PopulateEngineList()
                AutoSelectBestEngine()
                CheckOcrStatus()

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("OCR引擎后台检测异常: " & ex.Message)
            End Try
        End Sub

#Region "界面初始化"

        Private Sub InitializeComponent()
            Me.Text = "图片OCR识别"
            Me.Size = New System.Drawing.Size(900, 650)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.FormBorderStyle = FormBorderStyle.Sizable
            Me.MinimumSize = New System.Drawing.Size(700, 500)
            Me.BackColor = System.Drawing.Color.FromArgb(250, 250, 250)
            Me.Font = New System.Drawing.Font("微软雅黑", 9)

            ' 主布局面板
            Dim mainPanel As New TableLayoutPanel()
            mainPanel.Dock = DockStyle.Fill
            mainPanel.RowCount = 4
            mainPanel.ColumnCount = 1
            mainPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 52))
            mainPanel.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
            mainPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 36))
            mainPanel.RowStyles.Add(New RowStyle(SizeType.Absolute, 52))
            mainPanel.Padding = New Padding(12, 8, 12, 8)
            mainPanel.BackColor = Me.BackColor

            ' ===== 第1行：顶部工具栏 =====
            Dim topBar = CreateTopToolbar()
            mainPanel.Controls.Add(topBar, 0, 0)

            ' ===== 第2行：中间工作区（左右分栏） =====
            Dim workArea = CreateWorkArea()
            mainPanel.Controls.Add(workArea, 0, 1)

            ' ===== 第3行：进度与状态栏 =====
            Dim statusBar = CreateStatusBar()
            mainPanel.Controls.Add(statusBar, 0, 2)

            ' ===== 第4行：底部操作栏 =====
            Dim actionBar = CreateActionBar()
            mainPanel.Controls.Add(actionBar, 0, 3)

            Me.Controls.Add(mainPanel)
        End Sub

        ''' <summary>
        ''' 创建顶部工具栏
        ''' </summary>
        Private Function CreateTopToolbar() As Panel
            Dim panel As New Panel()
            panel.Dock = DockStyle.Fill
            panel.BackColor = System.Drawing.Color.FromArgb(240, 240, 240)
            panel.Padding = New Padding(10, 6, 10, 6)

            ' 左侧：选择图片按钮
            btnSelectImage = New Button()
            btnSelectImage.Text = "选择图片"
            btnSelectImage.Size = New System.Drawing.Size(90, 32)
            btnSelectImage.Location = New System.Drawing.Point(10, 6)
            btnSelectImage.BackColor = System.Drawing.Color.FromArgb(51, 122, 183)
            btnSelectImage.ForeColor = System.Drawing.Color.White
            btnSelectImage.FlatStyle = FlatStyle.Flat
            btnSelectImage.FlatAppearance.BorderSize = 0
            btnSelectImage.Font = New System.Drawing.Font("微软雅黑", 9, System.Drawing.FontStyle.Regular)
            btnSelectImage.Cursor = Cursors.Hand
            panel.Controls.Add(btnSelectImage)

            ' 分隔线1
            Dim sep1 As New Label()
            sep1.AutoSize = False
            sep1.Size = New System.Drawing.Size(1, 24)
            sep1.Location = New System.Drawing.Point(110, 10)
            sep1.BackColor = System.Drawing.Color.FromArgb(200, 200, 200)
            panel.Controls.Add(sep1)

            ' 引擎标签
            Dim lblEngine As New Label()
            lblEngine.Text = "识别引擎:"
            lblEngine.AutoSize = True
            lblEngine.Location = New System.Drawing.Point(122, 12)
            lblEngine.Font = New System.Drawing.Font("微软雅黑", 9)
            lblEngine.ForeColor = System.Drawing.Color.FromArgb(80, 80, 80)
            panel.Controls.Add(lblEngine)

            ' 引擎下拉框
            cbEngine = New ComboBox()
            cbEngine.Size = New System.Drawing.Size(180, 24)
            cbEngine.Location = New System.Drawing.Point(185, 9)
            cbEngine.DropDownStyle = ComboBoxStyle.DropDownList
            cbEngine.Font = New System.Drawing.Font("微软雅黑", 9)
            cbEngine.FlatStyle = FlatStyle.Flat
            cbEngine.BackColor = System.Drawing.Color.White
            panel.Controls.Add(cbEngine)

            ' 引擎状态指示灯面板
            pnlEngineStatus = New Panel()
            pnlEngineStatus.Size = New System.Drawing.Size(12, 12)
            pnlEngineStatus.Location = New System.Drawing.Point(370, 14)
            pnlEngineStatus.BackColor = System.Drawing.Color.FromArgb(200, 200, 200)
            pnlEngineStatus.BorderStyle = BorderStyle.None
            ' 做成圆形
            AddHandler pnlEngineStatus.Paint, Sub(s, e)
                e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                Using brush As New SolidBrush(pnlEngineStatus.BackColor)
                    e.Graphics.FillEllipse(brush, 0, 0, pnlEngineStatus.Width - 1, pnlEngineStatus.Height - 1)
                End Using
            End Sub
            panel.Controls.Add(pnlEngineStatus)

            lblEngineStatus = New Label()
            lblEngineStatus.Text = "检测中"
            lblEngineStatus.AutoSize = True
            lblEngineStatus.Location = New System.Drawing.Point(386, 12)
            lblEngineStatus.Font = New System.Drawing.Font("微软雅黑", 9)
            lblEngineStatus.ForeColor = System.Drawing.Color.FromArgb(120, 120, 120)
            panel.Controls.Add(lblEngineStatus)

            ' 分隔线2
            Dim sep2 As New Label()
            sep2.AutoSize = False
            sep2.Size = New System.Drawing.Size(1, 24)
            sep2.Location = New System.Drawing.Point(470, 10)
            sep2.BackColor = System.Drawing.Color.FromArgb(200, 200, 200)
            panel.Controls.Add(sep2)

            ' 开始识别按钮
            btnRecognize = New Button()
            btnRecognize.Text = "开始识别"
            btnRecognize.Size = New System.Drawing.Size(90, 32)
            btnRecognize.Location = New System.Drawing.Point(480, 6)
            btnRecognize.BackColor = System.Drawing.Color.FromArgb(92, 184, 92)
            btnRecognize.ForeColor = System.Drawing.Color.White
            btnRecognize.FlatStyle = FlatStyle.Flat
            btnRecognize.FlatAppearance.BorderSize = 0
            btnRecognize.Font = New System.Drawing.Font("微软雅黑", 9, System.Drawing.FontStyle.Regular)
            btnRecognize.Cursor = Cursors.Hand
            btnRecognize.Enabled = False
            panel.Controls.Add(btnRecognize)

            Return panel
        End Function

        ''' <summary>
        ''' 创建工作区（左右分栏）
        ''' </summary>
        Private Function CreateWorkArea() As TableLayoutPanel
            Dim panel As New TableLayoutPanel()
            panel.Dock = DockStyle.Fill
            panel.RowCount = 1
            panel.ColumnCount = 2
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
            panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 50))
            panel.Padding = New Padding(0, 8, 0, 8)
            panel.BackColor = Me.BackColor

            ' 左侧：图片预览
            panel.Controls.Add(CreatePreviewPanel(), 0, 0)

            ' 右侧：识别结果
            panel.Controls.Add(CreateResultPanel(), 1, 0)

            Return panel
        End Function

        ''' <summary>
        ''' 创建图片预览面板
        ''' </summary>
        Private Function CreatePreviewPanel() As Panel
            Dim container As New Panel()
            container.Dock = DockStyle.Fill
            container.Padding = New Padding(0, 0, 6, 0)
            container.BackColor = Me.BackColor

            ' 标题栏
            Dim titleBar As New Panel()
            titleBar.Dock = DockStyle.Top
            titleBar.Height = 28
            titleBar.BackColor = Me.BackColor

            Dim lblTitle As New Label()
            lblTitle.Text = "图片预览"
            lblTitle.Dock = DockStyle.Left
            lblTitle.AutoSize = True
            lblTitle.Font = New System.Drawing.Font("微软雅黑", 10, System.Drawing.FontStyle.Bold)
            lblTitle.ForeColor = System.Drawing.Color.FromArgb(60, 60, 60)
            lblTitle.TextAlign = ContentAlignment.MiddleLeft
            titleBar.Controls.Add(lblTitle)

            container.Controls.Add(titleBar)

            ' 预览区域容器
            Dim previewContainer As New Panel()
            previewContainer.Dock = DockStyle.Fill
            previewContainer.BackColor = System.Drawing.Color.FromArgb(245, 245, 245)
            previewContainer.BorderStyle = BorderStyle.FixedSingle
            previewContainer.Padding = New Padding(4)

            ' PictureBox
            picPreview = New PictureBox()
            picPreview.Dock = DockStyle.Fill
            picPreview.SizeMode = PictureBoxSizeMode.Zoom
            picPreview.BackColor = System.Drawing.Color.FromArgb(245, 245, 245)
            previewContainer.Controls.Add(picPreview)

            ' 拖拽提示覆盖层（当没有图片时显示）
            pnlPreviewOverlay = New Panel()
            pnlPreviewOverlay.Dock = DockStyle.Fill
            pnlPreviewOverlay.BackColor = System.Drawing.Color.FromArgb(245, 245, 245)

            ' 拖拽图标（用Label模拟）
            Dim lblIcon As New Label()
            lblIcon.Text = "+"
            lblIcon.Font = New System.Drawing.Font("微软雅黑", 48, System.Drawing.FontStyle.Regular)
            lblIcon.ForeColor = System.Drawing.Color.FromArgb(180, 180, 180)
            lblIcon.AutoSize = True
            lblIcon.Location = New System.Drawing.Point(0, 0)
            AddHandler lblIcon.SizeChanged, Sub(s, e)
                CenterControl(lblIcon, pnlPreviewOverlay)
            End Sub
            AddHandler pnlPreviewOverlay.SizeChanged, Sub(s, e)
                CenterControl(lblIcon, pnlPreviewOverlay)
                CenterControl(lblDragHint, pnlPreviewOverlay, 50)
            End Sub
            pnlPreviewOverlay.Controls.Add(lblIcon)

            ' 拖拽提示文字
            lblDragHint = New Label()
            lblDragHint.Text = "拖拽图片到此处" & Environment.NewLine & "或点击""选择图片""按钮"
            lblDragHint.AutoSize = True
            lblDragHint.Font = New System.Drawing.Font("微软雅黑", 10)
            lblDragHint.ForeColor = System.Drawing.Color.FromArgb(150, 150, 150)
            lblDragHint.TextAlign = ContentAlignment.MiddleCenter
            pnlPreviewOverlay.Controls.Add(lblDragHint)

            previewContainer.Controls.Add(pnlPreviewOverlay)

            container.Controls.Add(previewContainer)

            ' 图片信息标签
            lblImageInfo = New Label()
            lblImageInfo.Text = "未选择图片"
            lblImageInfo.Dock = DockStyle.Bottom
            lblImageInfo.Height = 24
            lblImageInfo.Font = New System.Drawing.Font("微软雅黑", 9)
            lblImageInfo.ForeColor = System.Drawing.Color.FromArgb(130, 130, 130)
            lblImageInfo.TextAlign = ContentAlignment.MiddleLeft
            lblImageInfo.Padding = New Padding(0, 4, 0, 0)
            container.Controls.Add(lblImageInfo)

            Return container
        End Function

        ''' <summary>
        ''' 创建识别结果面板
        ''' </summary>
        Private Function CreateResultPanel() As Panel
            Dim container As New Panel()
            container.Dock = DockStyle.Fill
            container.Padding = New Padding(6, 0, 0, 0)
            container.BackColor = Me.BackColor

            ' 标题栏（带操作按钮）
            Dim titleBar As New Panel()
            titleBar.Dock = DockStyle.Top
            titleBar.Height = 28
            titleBar.BackColor = Me.BackColor

            Dim lblTitle As New Label()
            lblTitle.Text = "识别结果"
            lblTitle.Dock = DockStyle.Left
            lblTitle.AutoSize = True
            lblTitle.Font = New System.Drawing.Font("微软雅黑", 10, System.Drawing.FontStyle.Bold)
            lblTitle.ForeColor = System.Drawing.Color.FromArgb(60, 60, 60)
            lblTitle.TextAlign = ContentAlignment.MiddleLeft
            titleBar.Controls.Add(lblTitle)

            ' 清空按钮（小）
            btnClearResult = New Button()
            btnClearResult.Text = "清空"
            btnClearResult.Size = New System.Drawing.Size(50, 24)
            btnClearResult.Dock = DockStyle.Right
            btnClearResult.FlatStyle = FlatStyle.Flat
            btnClearResult.FlatAppearance.BorderSize = 0
            btnClearResult.BackColor = System.Drawing.Color.FromArgb(230, 230, 230)
            btnClearResult.ForeColor = System.Drawing.Color.FromArgb(80, 80, 80)
            btnClearResult.Font = New System.Drawing.Font("微软雅黑", 8)
            btnClearResult.Cursor = Cursors.Hand
            btnClearResult.Enabled = False
            titleBar.Controls.Add(btnClearResult)

            ' 复制按钮（小）
            btnCopyResult = New Button()
            btnCopyResult.Text = "复制"
            btnCopyResult.Size = New System.Drawing.Size(50, 24)
            btnCopyResult.Dock = DockStyle.Right
            btnCopyResult.FlatStyle = FlatStyle.Flat
            btnCopyResult.FlatAppearance.BorderSize = 0
            btnCopyResult.BackColor = System.Drawing.Color.FromArgb(230, 230, 230)
            btnCopyResult.ForeColor = System.Drawing.Color.FromArgb(80, 80, 80)
            btnCopyResult.Font = New System.Drawing.Font("微软雅黑", 8)
            btnCopyResult.Cursor = Cursors.Hand
            btnCopyResult.Enabled = False
            btnCopyResult.Margin = New Padding(0, 0, 4, 0)
            titleBar.Controls.Add(btnCopyResult)

            container.Controls.Add(titleBar)

            ' 结果文本框
            txtResult = New TextBox()
            txtResult.Dock = DockStyle.Fill
            txtResult.Multiline = True
            txtResult.ScrollBars = ScrollBars.Both
            txtResult.Font = New System.Drawing.Font("Consolas", 10)
            txtResult.ReadOnly = True
            txtResult.BackColor = System.Drawing.Color.White
            txtResult.BorderStyle = BorderStyle.FixedSingle
            txtResult.ForeColor = System.Drawing.Color.FromArgb(50, 50, 50)
            container.Controls.Add(txtResult)

            Return container
        End Function

        ''' <summary>
        ''' 创建状态栏
        ''' </summary>
        Private Function CreateStatusBar() As Panel
            Dim panel As New Panel()
            panel.Dock = DockStyle.Fill
            panel.BackColor = Me.BackColor
            panel.Padding = New Padding(0, 4, 0, 4)

            progressBar = New ProgressBar()
            progressBar.Dock = DockStyle.Fill
            progressBar.Height = 20
            progressBar.Style = ProgressBarStyle.Marquee
            progressBar.Visible = False
            panel.Controls.Add(progressBar)

            Return panel
        End Function

        ''' <summary>
        ''' 创建底部操作栏
        ''' </summary>
        Private Function CreateActionBar() As Panel
            Dim panel As New Panel()
            panel.Dock = DockStyle.Fill
            panel.BackColor = System.Drawing.Color.FromArgb(240, 240, 240)
            panel.Padding = New Padding(10, 8, 10, 8)

            ' 左侧状态文字
            lblStatus = New Label()
            lblStatus.Text = "就绪"
            lblStatus.AutoSize = True
            lblStatus.Location = New System.Drawing.Point(10, 14)
            lblStatus.Font = New System.Drawing.Font("微软雅黑", 9)
            lblStatus.ForeColor = System.Drawing.Color.FromArgb(100, 100, 100)
            panel.Controls.Add(lblStatus)

            ' 右侧按钮容器
            Dim buttonPanel As New FlowLayoutPanel()
            buttonPanel.Dock = DockStyle.Right
            buttonPanel.FlowDirection = FlowDirection.RightToLeft
            buttonPanel.WrapContents = False
            buttonPanel.AutoSize = True
            buttonPanel.BackColor = Color.Transparent

            ' 关闭按钮
            btnClose = New Button()
            btnClose.Text = "关闭"
            btnClose.Size = New System.Drawing.Size(80, 32)
            btnClose.BackColor = System.Drawing.Color.FromArgb(180, 180, 180)
            btnClose.ForeColor = System.Drawing.Color.White
            btnClose.FlatStyle = FlatStyle.Flat
            btnClose.FlatAppearance.BorderSize = 0
            btnClose.Font = New System.Drawing.Font("微软雅黑", 9)
            btnClose.Cursor = Cursors.Hand
            buttonPanel.Controls.Add(btnClose)

            ' 插入到Excel按钮
            btnInsertToExcel = New Button()
            btnInsertToExcel.Text = "插入到Excel"
            btnInsertToExcel.Size = New System.Drawing.Size(110, 32)
            btnInsertToExcel.BackColor = System.Drawing.Color.FromArgb(92, 184, 92)
            btnInsertToExcel.ForeColor = System.Drawing.Color.White
            btnInsertToExcel.FlatStyle = FlatStyle.Flat
            btnInsertToExcel.FlatAppearance.BorderSize = 0
            btnInsertToExcel.Font = New System.Drawing.Font("微软雅黑", 9)
            btnInsertToExcel.Cursor = Cursors.Hand
            btnInsertToExcel.Enabled = False
            btnInsertToExcel.Margin = New Padding(0, 0, 8, 0)
            buttonPanel.Controls.Add(btnInsertToExcel)

            panel.Controls.Add(buttonPanel)

            Return panel
        End Function

        ''' <summary>
        ''' 将控件居中到容器
        ''' </summary>
        Private Sub CenterControl(ctrl As Control, container As Control, Optional offsetY As Integer = 0)
            If container.Width > ctrl.Width AndAlso container.Height > ctrl.Height Then
                ctrl.Left = (container.Width - ctrl.Width) \ 2
                ctrl.Top = (container.Height - ctrl.Height) \ 2 - offsetY
            End If
        End Sub

#End Region

#Region "状态管理"

        Private Enum StatusType
            Ready
            Info
            Success
            Warning
            [Error]
        End Enum

        Private Sub UpdateStatus(message As String, type As StatusType)
            lblStatus.Text = message
            Select Case type
                Case StatusType.Ready
                    lblStatus.ForeColor = System.Drawing.Color.FromArgb(100, 100, 100)
                Case StatusType.Info
                    lblStatus.ForeColor = System.Drawing.Color.FromArgb(51, 122, 183)
                Case StatusType.Success
                    lblStatus.ForeColor = System.Drawing.Color.FromArgb(92, 184, 92)
                Case StatusType.Warning
                    lblStatus.ForeColor = System.Drawing.Color.FromArgb(240, 173, 78)
                Case StatusType.Error
                    lblStatus.ForeColor = System.Drawing.Color.FromArgb(217, 83, 79)
            End Select
        End Sub

        Private Sub UpdateEngineStatus(isAvailable As Boolean, message As String)
            If isAvailable Then
                pnlEngineStatus.BackColor = System.Drawing.Color.FromArgb(92, 184, 92)
                lblEngineStatus.Text = message
                lblEngineStatus.ForeColor = System.Drawing.Color.FromArgb(92, 184, 92)
            Else
                pnlEngineStatus.BackColor = System.Drawing.Color.FromArgb(217, 83, 79)
                lblEngineStatus.Text = message
                lblEngineStatus.ForeColor = System.Drawing.Color.FromArgb(217, 83, 79)
            End If
            pnlEngineStatus.Invalidate()
        End Sub

#End Region

#Region "引擎管理"

        ''' <summary>
        ''' 填充引擎选择列表
        ''' </summary>
        Private Sub PopulateEngineList()
            cbEngine.Items.Clear()
            Dim engines = _ocrManager.GetAvailableEngines()
            For Each kvp In engines
                cbEngine.Items.Add(New EngineItem With {.Engine = kvp.Key, .Display = kvp.Value})
            Next
        End Sub

        ''' <summary>
        ''' 自动选择最佳可用引擎
        ''' </summary>
        Private Sub AutoSelectBestEngine()
            Dim engines = _ocrManager.GetAvailableEngines()
            If engines.Count = 0 Then Return

            Dim priorityOrder As New List(Of OcrEngineType) From {
                OcrEngineType.PaddleOCR,
                OcrEngineType.CloudOCR
            }

            Dim selected As OcrEngineType? = Nothing
            For Each engine In priorityOrder
                If engines.ContainsKey(engine) AndAlso _ocrManager.IsEngineAvailable(engine) Then
                    selected = engine
                    Exit For
                End If
            Next

            If selected.HasValue Then
                _ocrManager.CurrentEngine = selected.Value
            ElseIf engines.ContainsKey(OcrEngineType.PaddleOCR) Then
                _ocrManager.CurrentEngine = OcrEngineType.PaddleOCR
            Else
                _ocrManager.CurrentEngine = engines.Keys(0)
            End If

            SyncEngineComboBox()
        End Sub

        ''' <summary>
        ''' 同步下拉框与当前引擎设置
        ''' </summary>
        Private Sub SyncEngineComboBox()
            For i As Integer = 0 To cbEngine.Items.Count - 1
                Dim item = CType(cbEngine.Items(i), EngineItem)
                If item.Engine = _ocrManager.CurrentEngine Then
                    cbEngine.SelectedIndex = i
                    Return
                End If
            Next
        End Sub

        ''' <summary>
        ''' 检查OCR引擎状态
        ''' </summary>
        Private Sub CheckOcrStatus()
            Dim currentEngine = _ocrManager.CurrentEngine
            Dim engineName = OcrServiceManager.GetEngineDisplayName(currentEngine)
            Dim isAvailable = _ocrManager.IsEngineAvailable(currentEngine)

            If isAvailable Then
                UpdateEngineStatus(True, $"{engineName} 就绪")
            Else
                UpdateEngineStatus(False, $"{engineName} 不可用")
            End If
        End Sub

        ''' <summary>
        ''' 引擎选择变更事件
        ''' </summary>
        Private Sub cbEngine_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbEngine.SelectedIndexChanged
            If cbEngine.SelectedItem IsNot Nothing AndAlso TypeOf cbEngine.SelectedItem Is EngineItem Then
                Dim item = CType(cbEngine.SelectedItem, EngineItem)
                If _ocrManager IsNot Nothing Then
                    _ocrManager.CurrentEngine = item.Engine
                    CheckOcrStatus()
                End If
            End If
        End Sub

#End Region

#Region "图片处理"

        ''' <summary>
        ''' 选择图片按钮点击事件
        ''' </summary>
        Private Sub btnSelectImage_Click(sender As Object, e As EventArgs) Handles btnSelectImage.Click
            Using openFileDialog As New OpenFileDialog()
                openFileDialog.Filter = "图片文件|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff;*.webp|所有文件|*.*"
                openFileDialog.Title = "选择要识别的图片"
                openFileDialog.Multiselect = False

                If openFileDialog.ShowDialog() = DialogResult.OK Then
                    LoadImage(openFileDialog.FileName)
                End If
            End Using
        End Sub

        ''' <summary>
        ''' 加载图片
        ''' </summary>
        Private Sub LoadImage(imagePath As String)
            Try
                _currentImagePath = imagePath

                Using fs As New FileStream(imagePath, FileMode.Open, FileAccess.Read)
                    picPreview.Image = Image.FromStream(fs)
                End Using

                ' 隐藏拖拽提示
                pnlPreviewOverlay.Visible = False

                ' 更新图片信息
                Dim fileInfo As New FileInfo(imagePath)
                Dim img = picPreview.Image
                lblImageInfo.Text = $"{fileInfo.Name}   |   {img.Width} x {img.Height}   |   {fileInfo.Length / 1024:N1} KB"

                ' 启用识别按钮
                btnRecognize.Enabled = True
                txtResult.Text = ""
                _recognitionResult = ""
                btnCopyResult.Enabled = False
                btnInsertToExcel.Enabled = False
                btnClearResult.Enabled = False

                UpdateStatus("图片已加载，点击""开始识别""", StatusType.Info)

            Catch ex As Exception
                MessageBox.Show("加载图片失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>
        ''' 拖拽图片到预览区
        ''' </summary>
        Private Sub picPreview_DragEnter(sender As Object, e As DragEventArgs) Handles picPreview.DragEnter
            If e.Data.GetDataPresent(DataFormats.FileDrop) Then
                Dim files = CType(e.Data.GetData(DataFormats.FileDrop), String())
                If files.Length > 0 Then
                    Dim ext = Path.GetExtension(files(0)).ToLower()
                    If {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".webp"}.Contains(ext) Then
                        e.Effect = DragDropEffects.Copy
                    End If
                End If
            End If
        End Sub

        ''' <summary>
        ''' 拖拽放下图片
        ''' </summary>
        Private Sub picPreview_DragDrop(sender As Object, e As DragEventArgs) Handles picPreview.DragDrop
            Dim files = CType(e.Data.GetData(DataFormats.FileDrop), String())
            If files.Length > 0 Then
                LoadImage(files(0))
            End If
        End Sub

#End Region

#Region "OCR识别"

        ''' <summary>
        ''' 开始识别按钮点击事件
        ''' </summary>
        Private Async Sub btnRecognize_Click(sender As Object, e As EventArgs) Handles btnRecognize.Click
            If String.IsNullOrEmpty(_currentImagePath) OrElse Not File.Exists(_currentImagePath) Then
                MessageBox.Show("请先选择图片", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' 当前引擎不可用时，自动切换到可用引擎
            If Not _ocrManager.IsEngineAvailable(_ocrManager.CurrentEngine) Then
                Dim oldEngine = OcrServiceManager.GetEngineDisplayName(_ocrManager.CurrentEngine)
                AutoSelectBestEngine()
                Dim newEngine = OcrServiceManager.GetEngineDisplayName(_ocrManager.CurrentEngine)
                UpdateStatus($"{oldEngine} 不可用，已自动切换至 {newEngine}", StatusType.Warning)
            End If

            ' 开始识别
            btnRecognize.Enabled = False
            btnSelectImage.Enabled = False
            cbEngine.Enabled = False
            progressBar.Visible = True
            UpdateStatus($"正在识别... ({OcrServiceManager.GetEngineDisplayName(_ocrManager.CurrentEngine)})", StatusType.Info)
            txtResult.Text = "识别中，请稍候..."
            Me.Cursor = Cursors.WaitCursor

            Try
                _recognitionResult = Await _ocrManager.RecognizeText(_currentImagePath)

                txtResult.Text = _recognitionResult
                txtResult.ReadOnly = False

                btnCopyResult.Enabled = True
                btnInsertToExcel.Enabled = True
                btnClearResult.Enabled = True

                UpdateStatus($"识别完成   |   共 {_recognitionResult.Length} 字符", StatusType.Success)

            Catch ex As Exception
                txtResult.Text = "识别失败: " & ex.Message
                UpdateStatus("识别失败", StatusType.Error)
            Finally
                btnRecognize.Enabled = True
                btnSelectImage.Enabled = True
                cbEngine.Enabled = True
                progressBar.Visible = False
                Me.Cursor = Cursors.Default
            End Try
        End Sub

#End Region

#Region "结果操作"

        ''' <summary>
        ''' 复制结果按钮点击事件
        ''' </summary>
        Private Sub btnCopyResult_Click(sender As Object, e As EventArgs) Handles btnCopyResult.Click
            If Not String.IsNullOrEmpty(txtResult.Text) Then
                Clipboard.SetText(txtResult.Text)
                UpdateStatus("已复制到剪贴板", StatusType.Success)

                Task.Delay(2000).ContinueWith(Sub()
                    Me.Invoke(Sub()
                        If lblStatus.Text = "已复制到剪贴板" Then
                            UpdateStatus("识别完成", StatusType.Success)
                        End If
                    End Sub)
                End Sub)
            End If
        End Sub

        ''' <summary>
        ''' 清空结果按钮点击事件
        ''' </summary>
        Private Sub btnClearResult_Click(sender As Object, e As EventArgs) Handles btnClearResult.Click
            txtResult.Text = ""
            _recognitionResult = ""
            btnCopyResult.Enabled = False
            btnInsertToExcel.Enabled = False
            btnClearResult.Enabled = False
            UpdateStatus("结果已清空", StatusType.Info)
        End Sub

        ''' <summary>
        ''' 插入到Excel按钮点击事件
        ''' </summary>
        Private Sub btnInsertToExcel_Click(sender As Object, e As EventArgs) Handles btnInsertToExcel.Click
            If String.IsNullOrEmpty(txtResult.Text) Then
                Return
            End If

            Try
                If _excelApp Is Nothing Then
                    MessageBox.Show("无法访问Excel应用程序", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Return
                End If

                Dim selection = TryCast(_excelApp.Selection, Microsoft.Office.Interop.Excel.Range)
                If selection Is Nothing Then
                    MessageBox.Show("请先选择要插入的单元格", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                Dim lines = txtResult.Text.Split({Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

                If lines.Length = 1 Then
                    selection.Value = txtResult.Text
                Else
                    For i As Integer = 0 To lines.Length - 1
                        selection.Offset(i, 0).Value = lines(i)
                    Next
                End If

                UpdateStatus("已插入到Excel", StatusType.Success)
                Me.DialogResult = DialogResult.OK
                Me.Close()

            Catch ex As Exception
                MessageBox.Show("插入到Excel失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>
        ''' 关闭按钮点击事件
        ''' </summary>
        Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
            Me.DialogResult = DialogResult.Cancel
            Me.Close()
        End Sub

#End Region

        ''' <summary>
        ''' 窗体加载事件
        ''' </summary>
        Private Sub ImageOcrDialog_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            picPreview.AllowDrop = True
        End Sub

    End Class

End Namespace
