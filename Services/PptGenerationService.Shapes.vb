Option Explicit On
Option Infer On
Option Strict Off

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Linq
Imports System.Text
Imports System.Xml.Linq
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json.Serialization
Imports ExcelAIVSTOAddIn.Utils
Imports ExcelAIVSTOAddIn.Bridges

Namespace Services

Partial Public Class PptGenerationService

        Private Shared Sub AddBadge(slide As Object, toRelease As List(Of Object), text As String,
                                    left As Integer, top As Integer, width As Integer, height As Integer,
                                    fillColor As Long, textColor As Long,
                                    Optional fontSize As Integer = 12,
                                    Optional fontName As String = "微软雅黑")
            AddRoundedRect(slide, toRelease, left, top, width, height, fillColor, 50)
            AddCenteredTextBox(slide, toRelease, text, left, top, width, height, fontSize, True, textColor, fontName)
        End Sub
        Private Shared Sub AddDecorativeCircle(slide As Object, toRelease As List(Of Object),
                                               left As Integer, top As Integer, diameter As Integer,
                                               color As Long, transparency As Single)
            Dim shape As Object = Nothing
            Try
                shape = slide.Shapes.AddShape(9, SX(left), SY(top), SW(diameter), SH(diameter)) ' msoShapeOval = 9
                shape.Fill.ForeColor.RGB = color
                Try
                    shape.Fill.Transparency = transparency
                Catch
                End Try
                Try
                    shape.Line.Visible = False
                Catch
                End Try
                toRelease.Add(shape)
            Catch ex As Exception
                ChatLogger.Instance.LogWarning("PPT_GEN", "AddDecorativeCircle 失败: " & ex.Message)
                Try
                    If shape IsNot Nothing Then ComReleaseHelper.Release(shape)
                Catch
                End Try
            End Try
        End Sub
        Private Shared Sub AddFooter(slide As Object, toRelease As List(Of Object), pageNo As Integer, color As Long,
                                     Optional fontName As String = "微软雅黑")
            Try
                Dim ft As Object = slide.Shapes.AddTextbox(1, SX(40), SY(700), SW(300), SH(20))
                ft.TextFrame.TextRange.Text = "第 " & pageNo.ToString() & " 页"
                ft.TextFrame.TextRange.Font.Size = 10
                ft.TextFrame.TextRange.Font.Color.RGB = color
                ft.TextFrame.TextRange.Font.Name = fontName
                toRelease.Add(ft)
            Catch
            End Try
        End Sub
        Private Shared Sub AddCenteredTextBox(slide As Object, toRelease As List(Of Object), text As String,
                                              left As Integer, top As Integer, width As Integer, height As Integer,
                                              fontSize As Integer, bold As Boolean, color As Long,
                                              Optional fontName As String = "微软雅黑")
            Dim tb As Object = Nothing
            Try
                tb = slide.Shapes.AddTextbox(1, SX(left), SY(top), SW(width), SH(height))
                tb.TextFrame.TextRange.Text = text
                tb.TextFrame.TextRange.Font.Name = fontName
                tb.TextFrame.TextRange.Font.Size = fontSize
                If bold Then tb.TextFrame.TextRange.Font.Bold = True
                tb.TextFrame.TextRange.Font.Color.RGB = color
                Try
                    tb.TextFrame.VerticalAnchor = 3 ' msoAnchorMiddle
                Catch
                End Try
                toRelease.Add(tb)
            Catch
                Try
                    If tb IsNot Nothing Then ComReleaseHelper.Release(tb)
                Catch
                End Try
            End Try
        End Sub
        Private Shared Sub AddGradientRect(slide As Object, toRelease As List(Of Object),
                                           left As Integer, top As Integer, width As Integer, height As Integer,
                                           color1 As Long, color2 As Long,
                                           Optional radius As Integer = 10,
                                           Optional angle As Integer = 1)
            Dim shape As Object = Nothing
            Try
                shape = slide.Shapes.AddShape(5, SX(left), SY(top), SW(width), SH(height)) ' msoShapeRoundedRectangle = 5
                Try
                    shape.Fill.TwoColorGradient(angle, 1)
                    shape.Fill.ForeColor.RGB = color1
                    shape.Fill.BackColor.RGB = color2
                Catch
                    shape.Fill.ForeColor.RGB = color1
                End Try
                Try
                    shape.Adjustments(1) = Math.Min(0.5, radius / 100.0)
                Catch
                End Try
                Try
                    shape.Line.Visible = False
                Catch
                End Try
                toRelease.Add(shape)
            Catch ex As Exception
                ChatLogger.Instance.LogWarning("PPT_GEN", "AddGradientRect 失败: " & ex.Message)
                Try
                    If shape IsNot Nothing Then ComReleaseHelper.Release(shape)
                Catch
                End Try
            End Try
        End Sub
        Private Shared Sub AddTextBox(slide As Object, toRelease As List(Of Object), text As String,
                                      left As Integer, top As Integer, width As Integer, height As Integer,
                                      fontSize As Integer, bold As Boolean, color As Long,
                                      Optional fontName As String = "微软雅黑")
            Dim tb As Object = Nothing
            Try
                tb = slide.Shapes.AddTextbox(1, SX(left), SY(top), SW(width), SH(height))
                tb.TextFrame.TextRange.Text = text
                tb.TextFrame.TextRange.Font.Name = fontName
                ' 自动缩小字号以适应文本框（防止长文本溢出覆盖其他元素）
                Dim effSize As Integer = fontSize
                If height > 0 AndAlso Not String.IsNullOrEmpty(text) Then
                    Dim totalLines As Integer = 0
                    For Each ln As String In text.Split(New String() {vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
                        Dim charsPerLine As Integer = Math.Max(3, CInt(width / (fontSize * 0.6)))
                        totalLines += Math.Max(1, CInt(Math.Ceiling(Len(ln) / charsPerLine)))
                    Next
                    Dim neededH As Integer = CInt(totalLines * fontSize * 1.35)
                    If neededH > height Then
                        effSize = Math.Max(8, CInt(fontSize * height / neededH))
                    End If
                End If
                tb.TextFrame.TextRange.Font.Size = effSize
                If bold Then tb.TextFrame.TextRange.Font.Bold = True
                tb.TextFrame.TextRange.Font.Color.RGB = color
                toRelease.Add(tb)
            Catch
                Try
                    If tb IsNot Nothing Then ComReleaseHelper.Release(tb)
                Catch
                End Try
            End Try
        End Sub
        ''' <summary>
        ''' 设置幻灯片背景：优先套用母版背景（纯色/渐变/背景图），未提取到时用主题色。
        ''' 随后在内容之下绘制母版背景图形（logo/装饰）与版式装饰形状。
        ''' </summary>
        Private Shared Sub SetSlideBackground(slide As Object, toRelease As List(Of Object), color As Long)
            ' PowerPoint 背景必须先 Solid() 再设色，否则部分版本抛 0x800A03EC
            Dim bg As PptBackgroundProfile = If(_activeLayout IsNot Nothing, _activeLayout.Background, Nothing)
            If bg IsNot Nothing Then
                Try
                    ApplyMasterBackground(slide, bg)
                Catch
                    Try
                        slide.Background.Fill.Solid()
                        slide.Background.Fill.ForeColor.RGB = color
                    Catch
                    End Try
                End Try
            Else
                Try
                    slide.Background.Fill.Solid()
                    slide.Background.Fill.ForeColor.RGB = color
                Catch
                End Try
            End If

            ' 母版背景图形（每页复现，先绘制避免遮挡正文）
            If _activeLayout IsNot Nothing AndAlso _activeLayout.MasterShapes IsNot Nothing Then
                For Each sp As PptShapeProfile In _activeLayout.MasterShapes
                    DrawProfileShape(slide, toRelease, sp)
                Next
            End If

            ' 套用版式画像中的装饰形状（在内容之前绘制，避免遮挡正文）
            If _activeLayout IsNot Nothing AndAlso _activeLayout.Shapes IsNot Nothing Then
                For Each sp As PptShapeProfile In _activeLayout.Shapes
                    DrawProfileShape(slide, toRelease, sp)
                Next
            End If
        End Sub

        ''' <summary>
        ''' 套用母版背景：纯色/线性渐变（角度近似 + GradientAngle 精确兜底）/背景图片
        ''' </summary>
        Private Shared Sub ApplyMasterBackground(slide As Object, bg As PptBackgroundProfile)
            Select Case bg.Kind
                Case "color"
                    slide.Background.Fill.Solid()
                    slide.Background.Fill.ForeColor.RGB = RGB((bg.Color >> 16) And &HFF, (bg.Color >> 8) And &HFF, bg.Color And &HFF)
                Case "gradient"
                    ' TwoColorGradient(MsoGradientStyle)：1=水平 2=垂直 3=对角上 4=对角下，按角度近似
                    Dim angle As Integer = bg.GradientAngle Mod 360
                    If angle < 0 Then angle += 360
                    Dim style As Integer = 1
                    If angle >= 45 AndAlso angle < 135 Then
                        style = 2
                    ElseIf angle >= 135 AndAlso angle < 225 Then
                        style = 1
                    ElseIf angle >= 225 AndAlso angle < 315 Then
                        style = 2
                    End If
                    If angle >= 30 AndAlso angle < 60 Then style = 4
                    If angle >= 120 AndAlso angle < 150 Then style = 3
                    If angle >= 210 AndAlso angle < 240 Then style = 4
                    If angle >= 300 AndAlso angle < 330 Then style = 3
                    slide.Background.Fill.TwoColorGradient(style, 1)
                    slide.Background.Fill.ForeColor.RGB = RGB((bg.GradientColor1 >> 16) And &HFF, (bg.GradientColor1 >> 8) And &HFF, bg.GradientColor1 And &HFF)
                    slide.Background.Fill.BackColor.RGB = RGB((bg.GradientColor2 >> 16) And &HFF, (bg.GradientColor2 >> 8) And &HFF, bg.GradientColor2 And &HFF)
                    Try
                        ' Office 2010+：精确角度
                        slide.Background.Fill.GradientAngle = angle
                    Catch
                    End Try
                Case "picture"
                    Dim fullPath As String = Path.Combine(PathHelper.ConfigDirectory, bg.ImagePath)
                    If File.Exists(fullPath) Then
                        slide.Background.Fill.UserPicture(fullPath)
                    End If
            End Select
        End Sub

        ''' <summary>
        ''' 绘制版式/母版画像中的单个形状：图片（logo）优先，其次纯线条（分隔线），最后填充形状
        ''' </summary>
        Private Shared Sub DrawProfileShape(slide As Object, toRelease As List(Of Object), sp As PptShapeProfile)
            If sp Is Nothing Then Return
            ' 图片形状（logo 等）
            If Not String.IsNullOrEmpty(sp.ImagePath) Then
                Dim fullPath As String = Path.Combine(PathHelper.ConfigDirectory, sp.ImagePath)
                If File.Exists(fullPath) Then
                    Dim pic As Object = Nothing
                    Try
                        pic = slide.Shapes.AddPicture(fullPath, 0, -1, SX(sp.Left), SY(sp.Top), SW(sp.Width), SH(sp.Height))
                        Try
                            pic.LockAspectRatio = 0
                        Catch
                        End Try
                        toRelease.Add(pic)
                    Catch
                        Try
                            If pic IsNot Nothing Then ComReleaseHelper.Release(pic)
                        Catch
                        End Try
                    End Try
                End If
                Return
            End If
            ' 纯线条形状（分隔线/装饰线）
            If sp.Fill < 0 AndAlso sp.LineColor >= 0 Then
                Dim line As Object = Nothing
                Try
                    line = slide.Shapes.AddLine(SX(sp.Left), SY(sp.Top), SX(sp.Left + sp.Width), SY(sp.Top + sp.Height))
                    line.Line.ForeColor.RGB = RGB((sp.LineColor >> 16) And &HFF, (sp.LineColor >> 8) And &HFF, sp.LineColor And &HFF)
                    Try
                        line.Line.Weight = 1.5
                    Catch
                    End Try
                    toRelease.Add(line)
                Catch
                    Try
                        If line IsNot Nothing Then ComReleaseHelper.Release(line)
                    Catch
                    End Try
                End Try
                Return
            End If
            ' 常规填充形状
            If sp.Fill < 0 Then Return
            ' 画像填充色为 RGB 整数（如 0x1F4E79），需转为 PowerPoint 的 BGR 格式
            Dim fill As Long = RGB((sp.Fill >> 16) And &HFF, (sp.Fill >> 8) And &HFF, sp.Fill And &HFF)
            Select Case sp.Kind.ToLowerInvariant()
                Case "roundrect"
                    AddRoundedRect(slide, toRelease, sp.Left, sp.Top, sp.Width, sp.Height, fill, 10)
                Case "oval", "ellipse"
                    Dim shape As Object = Nothing
                    Try
                        shape = slide.Shapes.AddShape(9, SX(sp.Left), SY(sp.Top), SW(sp.Width), SH(sp.Height))
                        shape.Fill.ForeColor.RGB = fill
                        Try
                            shape.Line.Visible = False
                        Catch
                        End Try
                        toRelease.Add(shape)
                    Catch
                        Try
                            If shape IsNot Nothing Then ComReleaseHelper.Release(shape)
                        Catch
                        End Try
                    End Try
                Case Else ' rect 及其他类型回退为矩形
                    AddRibbon(slide, toRelease, fill, sp.Left, sp.Top, sp.Width, sp.Height)
            End Select
        End Sub
        Private Shared Function AddRoundedRect(slide As Object, toRelease As List(Of Object),
                                               left As Integer, top As Integer, width As Integer, height As Integer,
                                               fillColor As Long,
                                               Optional radius As Integer = 10,
                                               Optional lineColor As Long = 0,
                                               Optional lineVisible As Boolean = False) As Object
            Dim shape As Object = Nothing
            Try
                shape = slide.Shapes.AddShape(5, SX(left), SY(top), SW(width), SH(height)) ' msoShapeRoundedRectangle = 5
                shape.Fill.ForeColor.RGB = fillColor
                Try
                    shape.Adjustments(1) = Math.Min(0.5, radius / 100.0)
                Catch
                End Try
                Try
                    shape.Line.Visible = lineVisible
                    If lineVisible Then shape.Line.ForeColor.RGB = lineColor
                Catch
                End Try
                toRelease.Add(shape)
            Catch ex As Exception
                ChatLogger.Instance.LogWarning("PPT_GEN", "AddRoundedRect 失败: " & ex.Message)
                Try
                    If shape IsNot Nothing Then ComReleaseHelper.Release(shape)
                Catch
                End Try
            End Try
            Return shape
        End Function
        Private Shared Sub AddRibbon(slide As Object, toRelease As List(Of Object), color As Long,
                                     left As Integer, top As Integer, width As Integer, height As Integer)
            Dim shape As Object = Nothing
            Try
                shape = slide.Shapes.AddShape(1, SX(left), SY(top), SW(width), SH(height)) ' msoShapeRectangle = 1
                shape.Fill.ForeColor.RGB = color
                Try
                    shape.Line.Visible = False
                Catch
                End Try
                toRelease.Add(shape)
            Catch ex As Exception
                ChatLogger.Instance.LogWarning("PPT_GEN", "AddRibbon 失败: " & ex.Message)
                Try
                    If shape IsNot Nothing Then ComReleaseHelper.Release(shape)
                Catch
                End Try
            End Try
        End Sub
End Class
End Namespace
