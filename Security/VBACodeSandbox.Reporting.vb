'【注意】VBACodeSandbox 是 VBA 安全沙箱，通过 COM IDispatch 与 Excel 动态交互，
'  本质依赖后期绑定。此处单独关闭 Option Strict，而核心服务仍保持严格类型检查。
Option Strict Off
Option Explicit On
Option Infer On
Imports System.Text.RegularExpressions
Imports System.Security.Cryptography
Imports System.Text
Imports System.Linq
Imports System.Threading
Imports System.Windows.Forms
Imports System.Drawing
Imports ExcelAIVSToAddIn.Utils

Namespace Security

    ' VBACodeSandbox.Reporting.vb — 安全检查报告 / Markdown 清理 / 白名单确认对话框
    ' （类级 XML 文档只在主文件 VBACodeSandbox.vb 保留一份，避免 BC42314 分部重复）
    Partial Public Class VBACodeSandbox

        Public Shared Function GetFormattedReport(result As SecurityCheckResult) As String
            Dim report As New System.Text.StringBuilder()

            report.AppendLine("=== VBA Code Security Check Report ===")
            report.AppendLine(String.Format("Security Status: {0}", If(result.IsSafe, "Safe", "Unsafe")))
            report.AppendLine(String.Format("Risk Level: {0}", result.RiskLevel))
            report.AppendLine()

            If result.Errors.Count > 0 Then
                report.AppendLine("[Errors]")
                For Each err As String In result.Errors
                    report.AppendLine(String.Format("  - {0}", err))
                Next
                report.AppendLine()
            End If

            If result.Warnings.Count > 0 Then
                report.AppendLine("[Warnings]")
                For Each warning As String In result.Warnings
                    report.AppendLine(String.Format("  - {0}", warning))
                Next
                report.AppendLine()
            End If

            If result.Errors.Count = 0 AndAlso result.Warnings.Count = 0 Then
                report.AppendLine("No security issues found")
            End If

            Return report.ToString()
        End Function

        ''' <summary>
        ''' 清理 Markdown 代码块标记（LLM 可能返回 ```vba ... ``` 格式）
        ''' </summary>
        Private Shared Function CleanMarkdownCodeBlock(vbaCode As String) As String
            If String.IsNullOrWhiteSpace(vbaCode) Then Return vbaCode

            Dim result = vbaCode.Trim()

            ' 移除开头的 ```vba 或 ```vb 或 ```
            Dim codeBlockPattern As String = "^```(?:vba|vb)?\s*"
            result = System.Text.RegularExpressions.Regex.Replace(result, codeBlockPattern, "", System.Text.RegularExpressions.RegexOptions.IgnoreCase)

            ' 移除末尾的 ```
            If result.EndsWith("```") Then
                result = result.Substring(0, result.Length - 3).TrimEnd()
            End If

            Return result
        End Function

        ''' <summary>
        ''' 显示带"加入白名单"选项的安全确认对话框
        ''' </summary>
        ''' <param name="message">警告/确认消息</param>
        ''' <param name="unknownOps">不在白名单中的操作列表</param>
        ''' <param name="isCommonExcelOp">是否为常见 Excel 操作</param>
        ''' <returns>DialogResult: Yes=执行, Retry=加入白名单并执行, No=取消</returns>
        Private Shared Function ShowWhitelistConfirmDialog(message As String, unknownOps As List(Of String), isCommonExcelOp As Boolean) As DialogResult
            Dim result As DialogResult = DialogResult.No

            Dim form As New Form()
            form.Text = "安全确认"
            form.FormBorderStyle = FormBorderStyle.FixedDialog
            form.MaximizeBox = False
            form.MinimizeBox = False
            form.StartPosition = FormStartPosition.CenterScreen
            form.Font = New Font("Microsoft YaHei", 9.0F)
            form.ClientSize = New Size(440, 340)
            form.BackColor = Color.White

            ' ── 顶部区域：图标 + 标题 + 警告信息 ──
            ' 警告图标
            Dim iconBox As New PictureBox()
            iconBox.Location = New Point(16, 16)
            iconBox.Size = New Size(32, 32)
            iconBox.SizeMode = PictureBoxSizeMode.AutoSize
            iconBox.Image = SystemIcons.Warning.ToBitmap()
            form.Controls.Add(iconBox)

            ' 警告信息（自动换行，固定宽度）
            Dim messageLabel As New Label()
            messageLabel.Location = New Point(56, 18)
            messageLabel.Size = New Size(368, 56)
            messageLabel.AutoSize = False
            messageLabel.Text = message
            messageLabel.AutoEllipsis = True
            messageLabel.ForeColor = Color.FromArgb(51, 51, 51)
            form.Controls.Add(messageLabel)

            ' ── 分隔线 ──
            Dim line1 As New Panel()
            line1.Location = New Point(16, 80)
            line1.Size = New Size(408, 1)
            line1.BackColor = Color.FromArgb(224, 224, 224)
            form.Controls.Add(line1)

            ' ── 操作列表区域 ──
            ' 列表标题
            Dim opsLabel As New Label()
            opsLabel.Location = New Point(16, 88)
            opsLabel.Size = New Size(408, 18)
            opsLabel.Text = String.Format("不在安全白名单内的操作：", unknownOps.Count)
            opsLabel.Font = New Font("Microsoft YaHei", 9.0F, FontStyle.Bold)
            opsLabel.ForeColor = Color.FromArgb(51, 51, 51)
            form.Controls.Add(opsLabel)

            ' 操作列表容器（浅灰色背景）
            Dim opsPanel As New Panel()
            opsPanel.Location = New Point(16, 110)
            opsPanel.Size = New Size(408, 112)
            opsPanel.BackColor = Color.FromArgb(250, 250, 250)
            opsPanel.BorderStyle = BorderStyle.FixedSingle
            form.Controls.Add(opsPanel)

            ' 操作列表内容（每行一项，带 • 前缀）
            Dim opsText As String = String.Join(Environment.NewLine, unknownOps.Select(Function(op) "• " & op))
            Dim opsTextBox As New TextBox()
            opsTextBox.Location = New Point(8, 6)
            opsTextBox.Size = New Size(392, 100)
            opsTextBox.Multiline = True
            opsTextBox.ReadOnly = True
            opsTextBox.ScrollBars = ScrollBars.Vertical
            opsTextBox.BackColor = Color.FromArgb(250, 250, 250)
            opsTextBox.BorderStyle = BorderStyle.None
            opsTextBox.Font = New Font("Consolas", 9.0F)
            opsTextBox.Text = opsText
            opsPanel.Controls.Add(opsTextBox)

            ' ── 提示文字 ──
            Dim tipLabel As New Label()
            tipLabel.Location = New Point(16, 228)
            tipLabel.Size = New Size(408, 32)
            tipLabel.Text = "点击「加入白名单」可将上述操作永久加入安全白名单，下次执行同类代码无需确认。"
            tipLabel.ForeColor = Color.FromArgb(120, 120, 120)
            tipLabel.Font = New Font("Microsoft YaHei", 8.5F)
            tipLabel.AutoEllipsis = True
            form.Controls.Add(tipLabel)

            ' ── 分隔线 ──
            Dim line2 As New Panel()
            line2.Location = New Point(16, 268)
            line2.Size = New Size(408, 1)
            line2.BackColor = Color.FromArgb(224, 224, 224)
            form.Controls.Add(line2)

            ' ── 按钮区域 ──
            ' 按钮：否(N) - 左
            Dim btnNo As New Button()
            btnNo.Text = "否(&N)"
            btnNo.DialogResult = DialogResult.No
            btnNo.Location = New Point(16, 284)
            btnNo.Size = New Size(110, 32)
            btnNo.FlatStyle = FlatStyle.Flat
            btnNo.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200)
            btnNo.BackColor = Color.White
            form.Controls.Add(btnNo)

            ' 按钮：加入白名单并执行(W) - 中
            Dim btnWhitelist As New Button()
            btnWhitelist.Text = "加入白名单(&W)"
            btnWhitelist.DialogResult = DialogResult.Retry
            btnWhitelist.Location = New Point(165, 284)
            btnWhitelist.Size = New Size(120, 32)
            btnWhitelist.FlatStyle = FlatStyle.Flat
            btnWhitelist.FlatAppearance.BorderColor = Color.FromArgb(198, 239, 206)
            btnWhitelist.BackColor = Color.FromArgb(198, 239, 206)
            form.Controls.Add(btnWhitelist)

            ' 按钮：是(Y) - 右
            Dim btnYes As New Button()
            btnYes.Text = "是(&Y)"
            btnYes.DialogResult = DialogResult.Yes
            btnYes.Location = New Point(314, 284)
            btnYes.Size = New Size(110, 32)
            btnYes.FlatStyle = FlatStyle.Flat
            btnYes.FlatAppearance.BorderColor = Color.FromArgb(189, 215, 238)
            btnYes.BackColor = Color.FromArgb(189, 215, 238)
            form.Controls.Add(btnYes)

            ' 设置默认按钮和取消按钮
            form.AcceptButton = btnYes
            form.CancelButton = btnNo

            result = form.ShowDialog()
            form.Dispose()

            Return result
        End Function

        ' 【历史备注】Execute VBA code (in sandbox environment) —— 该方法定义在其他分部文件中

    End Class

End Namespace
