' ============================================================
' PptImagePaletteExtractor.vb — 从参考图提取品牌调色板
'
' 目标：让 AI 把用户给的品牌图/参考图变成一套可用 PPT 主题配色。
' 算法（确定性，可单测）：
'   1. 加载图片（本地路径 / http(s) URL 下载到临时目录）
'   2. 降采样到 64x64（跳过 alpha<128 的像素）
'   3. RGB 5bit/通道量化聚类 → 按出现频次排序
'   4. 选色：最高频簇=背景 bg；次高频非 bg 簇=主色 primary；
'      第三频=辅助 secondary；饱和度最高=强调 accent；
'      light=bg 提亮；text 按 bg 明暗取深/浅；gray=中性灰
'   5. 输出 ImagePalette（含统计），由调用方组装主题 JSON 并走
'      PptThemeDesignValidator（WCAG + 令牌）兜底"合理性"
' ============================================================
Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Net
Imports Newtonsoft.Json.Linq

Namespace Services

    Public NotInheritable Class PptImagePaletteExtractor

        Private Sub New()
        End Sub

        ' 缩略图最长边（提色采样尺寸）
        Private Const THUMB_MAX As Integer = 64
        ' 候选簇最小占比（剔除噪点）
        Private Const MIN_CLUSTER_RATIO As Double = 0.02
        ' 饱和度门槛：只有 S >= 该值的簇才有资格当选强调色
        Private Const ACCENT_MIN_SATURATION As Double = 0.3

        ''' <summary>提色结果（RGB [r,g,b]）</summary>
        Public Class ImagePalette
            Public Property Bg As Integer()
            Public Property Primary As Integer()
            Public Property Secondary As Integer()
            Public Property Accent As Integer()
            Public Property Light As Integer()
            Public Property Text As Integer()
            Public Property Gray As Integer()
            Public Property ValidPixels As Integer = 0
            Public Property ClusterCount As Integer = 0
        End Class

        ''' <summary>
        ''' 从图片提取主色调色板。图片不存在/无法解码时返回 Nothing。
        ''' 支持本地路径与 http(s) URL（URL 先下载到临时目录）。
        ''' </summary>
        Public Shared Function Extract(imagePath As String) As ImagePalette
            If String.IsNullOrWhiteSpace(imagePath) Then Return Nothing
            Dim localPath As String = ResolveImagePath(imagePath)
            If String.IsNullOrWhiteSpace(localPath) OrElse Not IO.File.Exists(localPath) Then Return Nothing

            Dim bitmap As Bitmap = Nothing
            Try
                ' 解码原图（FromFile 会锁定文件，用完即 Dispose）
                Using source As New Bitmap(localPath)
                    ' 降采样到 THUMB_MAX 长边
                    Dim scale As Double = Math.Min(1.0, THUMB_MAX / Math.Max(source.Width, source.Height))
                    Dim tw As Integer = Math.Max(1, CInt(Math.Round(source.Width * scale)))
                    Dim th As Integer = Math.Max(1, CInt(Math.Round(source.Height * scale)))
                    bitmap = New Bitmap(tw, th)
                    Using g As Graphics = Graphics.FromImage(bitmap)
                        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBilinear
                        g.DrawImage(source, 0, 0, tw, th)
                    End Using
                End Using
            Catch
                Return Nothing
            End Try

            Try
                Using bitmap
                    Return BuildPalette(bitmap)
                End Using
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 从位图像素聚类构建调色板。
        ''' 公开以便单测直接喂合成位图（不依赖磁盘文件）。
        ''' </summary>
        Public Shared Function BuildPalette(bitmap As Bitmap) As ImagePalette
            ' 1) 像素量化：RGB 5bit/通道 → 桶计数（key/计数用 Long，避免 VB 溢出检查）
            Dim buckets As New Dictionary(Of Long, Long)()
            Dim total As Long = 0
            For y As Integer = 0 To bitmap.Height - 1
                For x As Integer = 0 To bitmap.Width - 1
                    Dim c As Color = bitmap.GetPixel(x, y)
                    If c.A < 128 Then Continue For
                    Dim r5 As Long = CLng(c.R) >> 3
                    Dim g5 As Long = CLng(c.G) >> 3
                    Dim b5 As Long = CLng(c.B) >> 3
                    Dim key As Long = (r5 << 10) Or (g5 << 5) Or b5
                    Dim n As Long
                    If buckets.TryGetValue(key, n) Then
                        buckets(key) = n + 1
                    Else
                        buckets(key) = 1
                    End If
                    total += 1
                Next
            Next

            Dim palette As New ImagePalette()
            palette.ValidPixels = CInt(Math.Min(Integer.MaxValue, total))
            palette.ClusterCount = buckets.Count
            If total = 0 OrElse buckets.Count = 0 Then
                Return FallbackPalette()
            End If

            ' 2) 排序：频次降序；同频按饱和度降序（更"有颜色"的簇优先）
            Dim ordered As New List(Of KeyValuePair(Of Long, Long))()
            For Each kv As KeyValuePair(Of Long, Long) In buckets
                ordered.Add(kv)
            Next
            ordered.Sort(Function(a, b)
                             Dim d As Integer = b.Value.CompareTo(a.Value)
                             If d <> 0 Then Return d
                             Dim ca As Color = ClusterToColor(a.Key)
                             Dim cb As Color = ClusterToColor(b.Key)
                             Return Saturation(cb).CompareTo(Saturation(ca))
                         End Function)

            ' 3) 候选簇（占比 >= MIN_CLUSTER_RATIO），不足则回退最高频
            Dim candidates As New List(Of KeyValuePair(Of Long, Long))()
            For Each kv As KeyValuePair(Of Long, Long) In ordered
                If kv.Value * 1.0 / total >= MIN_CLUSTER_RATIO Then
                    candidates.Add(kv)
                End If
                If candidates.Count >= 8 Then Exit For
            Next
            If candidates.Count = 0 Then
                candidates.Add(ordered(0))
            End If

            ' 4) 选色
            ' bg：最高频簇（白底图 → 白为最高频；深底图 → 深为最高频）
            Dim bg As Color = ClusterToColor(candidates(0).Key)
            ' primary：次高频且与 bg 可区分（色距 >= 24 或亮度差 >= 0.05）；没有则取次高频
            Dim primary As Color = bg
            For i As Integer = 1 To candidates.Count - 1
                Dim c As Color = ClusterToColor(candidates(i).Key)
                If ColorDistinct(c, bg) Then
                    primary = c
                    Exit For
                End If
            Next
            If primary.Equals(bg) Then primary = ClusterToColor(candidates(1).Key)

            ' secondary：第三频（与 bg、primary 均区分）
            Dim secondary As Color = primary
            For i As Integer = 1 To candidates.Count - 1
                Dim c As Color = ClusterToColor(candidates(i).Key)
                If ColorDistinct(c, bg) AndAlso ColorDistinct(c, primary) Then
                    secondary = c
                    Exit For
                End If
            Next
            If secondary.Equals(primary) Then secondary = Mix(bg, primary, 0.5)

            ' accent：非 bg 簇中饱和度最高者（灰色簇不参与）
            Dim accent As Color = primary
            Dim bestS As Double = -1.0
            For i As Integer = 0 To candidates.Count - 1
                Dim c As Color = ClusterToColor(candidates(i).Key)
                Dim s As Double = Saturation(c)
                If ColorDistinct(c, bg) AndAlso s >= ACCENT_MIN_SATURATION AndAlso s > bestS Then
                    accent = c
                    bestS = s
                End If
            Next
            If bestS < 0 Then accent = ShiftHue(primary, 30)

            ' light：bg 向白走 20%；text 按 bg 明暗取深/浅；gray 中性灰
            Dim light As Color = Lift(bg, 0.2)
            Dim text As Color = If(RelativeLuminance(bg.R, bg.G, bg.B) >= 0.45, Color.FromArgb(33, 37, 41), Color.FromArgb(240, 244, 250))
            Dim gray As Color = Color.FromArgb(150, 150, 150)

            palette.Bg = ToRgb(bg)
            palette.Primary = ToRgb(primary)
            palette.Secondary = ToRgb(secondary)
            palette.Accent = ToRgb(accent)
            palette.Light = ToRgb(light)
            palette.Text = ToRgb(text)
            palette.Gray = ToRgb(gray)
            Return palette
        End Function

        ' ------------------------------------------------------------
        ' 内部工具
        ' ------------------------------------------------------------

        Private Shared Function FallbackPalette() As ImagePalette
            Dim p As New ImagePalette()
            p.Bg = New Integer() {245, 247, 250}
            p.Primary = New Integer() {31, 78, 121}
            p.Secondary = New Integer() {190, 167, 91}
            p.Accent = New Integer() {255, 165, 0}
            p.Light = New Integer() {239, 243, 248}
            p.Text = New Integer() {33, 37, 41}
            p.Gray = New Integer() {150, 150, 150}
            Return p
        End Function

        ''' <summary>5bit 桶 key → 颜色（取桶中值恢复）</summary>
        Private Shared Function ClusterToColor(key As Long) As Color
            Dim r As Integer = CInt((key >> 10) And 31) << 3 Or 4
            Dim g As Integer = CInt((key >> 5) And 31) << 3 Or 4
            Dim b As Integer = CInt(key And 31) << 3 Or 4
            Return Color.FromArgb(r, g, b)
        End Function

        ''' <summary>HSV 饱和度</summary>
        Private Shared Function Saturation(c As Color) As Double
            Dim max As Double = Math.Max(c.R, Math.Max(c.G, c.B)) / 255.0
            Dim min As Double = Math.Min(c.R, Math.Min(c.G, c.B)) / 255.0
            If max <= 0 Then Return 0.0
            Return (max - min) / max
        End Function

        ''' <summary>颜色是否足够区分（欧氏色距或亮度差）</summary>
        Private Shared Function ColorDistinct(a As Color, b As Color) As Boolean
            ' 注意：VB 算术不提升窄类型，Byte-Byte 会按 Byte 运算并溢出，须显式 CInt
            Dim dr As Integer = CInt(a.R) - CInt(b.R)
            Dim dg As Integer = CInt(a.G) - CInt(b.G)
            Dim db As Integer = CInt(a.B) - CInt(b.B)
            Dim dist As Long = CLng(dr) * dr + CLng(dg) * dg + CLng(db) * db
            If dist >= 24L * 24L Then Return True
            Dim la As Double = RelativeLuminance(a.R, a.G, a.B)
            Dim lb As Double = RelativeLuminance(b.R, b.G, b.B)
            Return Math.Abs(la - lb) >= 0.05
        End Function

        ''' <summary>两色混合 ratio=a 的占比</summary>
        Private Shared Function Mix(a As Color, b As Color, ratio As Double) As Color
            Dim r As Integer = CInt(Math.Min(255, a.R * ratio + b.R * (1 - ratio)))
            Dim g As Integer = CInt(Math.Min(255, a.G * ratio + b.G * (1 - ratio)))
            Dim bl As Integer = CInt(Math.Min(255, a.B * ratio + b.B * (1 - ratio)))
            Return Color.FromArgb(r, g, bl)
        End Function

        ''' <summary>向白提升 ratio（0-1）</summary>
        Private Shared Function Lift(c As Color, ratio As Double) As Color
            Dim r As Integer = CInt(c.R + (255 - CInt(c.R)) * ratio)
            Dim g As Integer = CInt(c.G + (255 - CInt(c.G)) * ratio)
            Dim b As Integer = CInt(c.B + (255 - CInt(c.B)) * ratio)
            Return Color.FromArgb(Math.Min(255, r), Math.Min(255, g), Math.Min(255, b))
        End Function

        ''' <summary>主色色调偏移（无高饱和簇时生成强调色）</summary>
        Private Shared Function ShiftHue(c As Color, degrees As Integer) As Color
            Dim h As Single
            Dim s As Single
            Dim v As Single
            ColorToHsv(c, h, s, v)
            h = (h + degrees) Mod 360.0F
            Return HsvToColor(h, s, v)
        End Function

        Private Shared Sub ColorToHsv(c As Color, ByRef h As Single, ByRef s As Single, ByRef v As Single)
            Dim max As Single = Math.Max(c.R, Math.Max(c.G, c.B)) / 255.0F
            Dim min As Single = Math.Min(c.R, Math.Min(c.G, c.B)) / 255.0F
            v = max
            Dim d As Single = max - min
            s = If(max = 0.0F, 0.0F, d / max)
            If d = 0.0F Then
                h = 0.0F
            Else
                Dim rn As Single = (max - c.R / 255.0F) / d
                Dim gn As Single = (max - c.G / 255.0F) / d
                Dim bn As Single = (max - c.B / 255.0F) / d
                If c.R = CInt(max * 255) Then
                    h = bn - gn
                ElseIf c.G = CInt(max * 255) Then
                    h = 2.0F + rn - bn
                Else
                    h = 4.0F + gn - rn
                End If
                h *= 60.0F
                If h < 0.0F Then h += 360.0F
            End If
        End Sub

        Private Shared Function HsvToColor(h As Single, s As Single, v As Single) As Color
            Dim c As Single = v * s
            Dim x As Single = c * (1.0F - Math.Abs((h / 60.0F) Mod 2.0F - 1.0F))
            Dim m As Single = v - c
            Dim r As Single = 0
            Dim g As Single = 0
            Dim b As Single = 0
            If h < 60.0F Then
                r = c : g = x
            ElseIf h < 120.0F Then
                r = x : g = c
            ElseIf h < 180.0F Then
                g = c : b = x
            ElseIf h < 240.0F Then
                g = x : b = c
            ElseIf h < 300.0F Then
                r = x : b = c
            Else
                r = c : b = x
            End If
            Return Color.FromArgb(CInt((r + m) * 255), CInt((g + m) * 255), CInt((b + m) * 255))
        End Function

        ''' <summary>sRGB 相对亮度（与 PptThemeDesignValidator 同公式）</summary>
        Public Shared Function RelativeLuminance(r As Integer, g As Integer, b As Integer) As Double
            Return 0.2126 * ChannelLinear(r) + 0.7152 * ChannelLinear(g) + 0.0722 * ChannelLinear(b)
        End Function

        Private Shared Function ChannelLinear(c As Integer) As Double
            Dim v As Double = Math.Max(0, Math.Min(255, c)) / 255.0
            If v <= 0.03928 Then Return v / 12.92
            Return Math.Pow((v + 0.055) / 1.055, 2.4)
        End Function

        Private Shared Function ToRgb(c As Color) As Integer()
            Return New Integer() {c.R, c.G, c.B}
        End Function

        ''' <summary>
        ''' 图片路径解析：本地路径直接用；http(s) URL 先下载到临时目录再返回本地路径。
        ''' 与 PptGenerationService.Media.ResolveImagePath 同契约（独立实现，避免跨文件依赖）。
        ''' </summary>
        Public Shared Function ResolveImagePath(src As String) As String
            If String.IsNullOrWhiteSpace(src) Then Return ""
            Dim trimmed As String = src.Trim()
            If Not (trimmed.StartsWith("http://", StringComparison.OrdinalIgnoreCase) OrElse
                    trimmed.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) Then
                Return trimmed
            End If
            Dim uri As Uri = Nothing
            If Not Uri.TryCreate(trimmed, UriKind.Absolute, uri) Then Return ""
            Try
                Dim ext As String = ""
                Dim urlPath As String = uri.AbsolutePath
                If Not String.IsNullOrWhiteSpace(urlPath) Then ext = IO.Path.GetExtension(urlPath)
                If String.IsNullOrWhiteSpace(ext) Then ext = ".jpg"
                Dim tmp As String = IO.Path.Combine(IO.Path.GetTempPath(), "ppt_theme_img_" & Guid.NewGuid().ToString("N") & ext)
                Using wc As New WebClient()
                    wc.DownloadFile(uri, tmp)
                End Using
                If IO.File.Exists(tmp) Then Return tmp
            Catch
                Return ""
            End Try
            Return ""
        End Function

        ''' <summary>
        ''' 组装主题 JSON：{id,name,font,styleDirection?,designTokens?,colors:{提取色}}。
        ''' styleDirection 由校验器展开（显式 colors 覆盖 preset 配色，保留 preset 版式）。
        ''' </summary>
        Public Shared Function BuildThemeJson(palette As ImagePalette,
                                              themeId As String,
                                              name As String,
                                              font As String,
                                              Optional styleDirection As String = "",
                                              Optional designTokensJson As String = "") As String
            If palette Is Nothing Then palette = FallbackPalette()
            Dim theme As New JObject()
            theme("id") = If(String.IsNullOrWhiteSpace(themeId), "img_" & Guid.NewGuid().ToString("N").Substring(0, 8), themeId)
            theme("name") = If(String.IsNullOrWhiteSpace(name), Convert.ToString(theme("id")), name)
            theme("font") = If(String.IsNullOrWhiteSpace(font), "微软雅黑", font)
            If Not String.IsNullOrWhiteSpace(styleDirection) Then theme("styleDirection") = styleDirection.Trim().ToLowerInvariant()
            If Not String.IsNullOrWhiteSpace(designTokensJson) Then
                Try
                    Dim dt As JObject = JObject.Parse(designTokensJson)
                    If dt IsNot Nothing AndAlso dt.Count > 0 Then theme("designTokens") = dt
                Catch
                End Try
            End If
            theme("colors") = New JObject(
                New JProperty("primary", New JArray(palette.Primary(0), palette.Primary(1), palette.Primary(2))),
                New JProperty("secondary", New JArray(palette.Secondary(0), palette.Secondary(1), palette.Secondary(2))),
                New JProperty("accent", New JArray(palette.Accent(0), palette.Accent(1), palette.Accent(2))),
                New JProperty("bg", New JArray(palette.Bg(0), palette.Bg(1), palette.Bg(2))),
                New JProperty("light", New JArray(palette.Light(0), palette.Light(1), palette.Light(2))),
                New JProperty("text", New JArray(palette.Text(0), palette.Text(1), palette.Text(2))),
                New JProperty("gray", New JArray(palette.Gray(0), palette.Gray(1), palette.Gray(2))))
            Return theme.ToString(Newtonsoft.Json.Formatting.None)
        End Function

    End Class

End Namespace
