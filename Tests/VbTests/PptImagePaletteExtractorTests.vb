' ============================================================
' PptImagePaletteExtractorTests.vb — 图片提色引擎单元测试
'
' 被测对象：Services/PptImagePaletteExtractor.vb（System.Drawing 提色聚类）
' 覆盖：亮底图（白底+蓝色主体→bg≈白、primary≈蓝）、
'       暗底图（深底+亮绿→bg 暗、text 浅）、
'       URL 路径解析、BuildThemeJson 组装、合成图全流程
' ============================================================
Imports System
Imports System.Drawing
Imports ExcelAIVSTOAddIn.Services
Imports Newtonsoft.Json.Linq

Module PptImagePaletteExtractorTests

    Private _failures As Integer = 0
    Private _total As Integer = 0

    Private Sub Check(condition As Boolean, label As String)
        _total += 1
        If condition Then
            Console.WriteLine("  ok " & label)
        Else
            _failures += 1
            Console.WriteLine("  FAIL " & label)
        End If
    End Sub

    ''' <summary>合成测试位图：背景填充 + 若干色块（模拟品牌图构图）</summary>
    Private Function MakeBitmap(width As Integer, height As Integer,
                                bg As Color, ParamArray blocks As (Color, Integer, Integer, Integer, Integer)()) As Bitmap
        Dim bmp As New Bitmap(width, height)
        Using g As Graphics = Graphics.FromImage(bmp)
            g.Clear(bg)
            For Each blk In blocks
                Using br As New SolidBrush(blk.Item1)
                    g.FillRectangle(br, blk.Item2, blk.Item3, blk.Item4, blk.Item5)
                End Using
            Next
        End Using
        Return bmp
    End Function

    Public Function Run() As (Total As Integer, Failures As Integer)
        Console.WriteLine("== PptImagePaletteExtractor 图片提色单元测试 ==")

        TestLightBackgroundImage()
        TestDarkBackgroundImage()
        TestBuildThemeJson()
        TestBadPathReturnsNothing()
        TestResolveImagePath()

        Console.WriteLine()
        If _failures = 0 Then
            Console.WriteLine("[PASS] PptImagePaletteExtractor " & _total.ToString() & " 项断言通过")
        Else
            Console.WriteLine("[FAIL] PptImagePaletteExtractor " & _failures.ToString() & " / " & _total.ToString() & " 项断言失败")
        End If
        Return (_total, _failures)
    End Function

    ' ------------------------------------------------------------
    ' 1) 亮底图：白底(70%) + 蓝色块(25%) + 小红条(5%)
    '    → bg≈白、primary≈蓝、accent 存在、text 深
    ' ------------------------------------------------------------
    Private Sub TestLightBackgroundImage()
        Console.WriteLine("[Light] 白底蓝主体提色")
        Dim blocks As (Color, Integer, Integer, Integer, Integer)() = {
            (Color.FromArgb(0, 90, 156), 20, 20, 80, 80),   ' 蓝色块
            (Color.FromArgb(230, 30, 30), 100, 100, 20, 20)  ' 小红条
        }
        Using bmp = MakeBitmap(160, 160, Color.White, blocks)
            Dim p = PptImagePaletteExtractor.BuildPalette(bmp)
            Check(p IsNot Nothing, "提取成功")
            Check(p.Bg(0) >= 240 AndAlso p.Bg(1) >= 240 AndAlso p.Bg(2) >= 240, "bg 接近白（R=" & p.Bg(0).ToString() & "）")
            ' 蓝色块占 25% → 应为主色候选；验证 blue 通道明显大于红
            Check(p.Primary(2) >= 100 AndAlso p.Primary(2) > p.Primary(0), "primary 为蓝色系（B=" & p.Primary(2).ToString() & " R=" & p.Primary(0).ToString() & "）")
            Check(p.Text(0) <= 80, "亮底 text 为深色（R=" & p.Text(0).ToString() & "）")
            Check(p.Accent IsNot Nothing AndAlso p.Accent.Length = 3, "accent 已生成")
        End Using
    End Sub

    ' ------------------------------------------------------------
    ' 2) 暗底图：深蓝底(80%) + 亮绿块(15%)
    '    → bg 暗、text 浅
    ' ------------------------------------------------------------
    Private Sub TestDarkBackgroundImage()
        Console.WriteLine("[Dark] 深底亮绿提色")
        Dim blocks As (Color, Integer, Integer, Integer, Integer)() = {
            (Color.FromArgb(0, 200, 120), 20, 20, 80, 80)   ' 亮绿块
        }
        Using bmp = MakeBitmap(160, 160, Color.FromArgb(16, 24, 48), blocks)
            Dim p = PptImagePaletteExtractor.BuildPalette(bmp)
            Check(p IsNot Nothing, "提取成功")
            Check(p.Bg(0) <= 60 AndAlso p.Bg(2) <= 90, "bg 为深色（R=" & p.Bg(0).ToString() & " B=" & p.Bg(2).ToString() & "）")
            Check(p.Text(0) >= 200, "暗底 text 为浅色（R=" & p.Text(0).ToString() & "）")
            Check(p.Primary(1) >= 100, "primary 含绿色分量（G=" & p.Primary(1).ToString() & "）")
        End Using
    End Sub

    ' ------------------------------------------------------------
    ' 3) BuildThemeJson：组装合法主题 JSON（可直接进校验器）
    ' ------------------------------------------------------------
    Private Sub TestBuildThemeJson()
        Console.WriteLine("[Json] BuildThemeJson 组装")
        Dim p As New PptImagePaletteExtractor.ImagePalette()
        p.Bg = New Integer() {245, 247, 250}
        p.Primary = New Integer() {31, 78, 121}
        p.Secondary = New Integer() {190, 167, 91}
        p.Accent = New Integer() {255, 165, 0}
        p.Light = New Integer() {239, 243, 248}
        p.Text = New Integer() {33, 37, 41}
        p.Gray = New Integer() {150, 150, 150}

        Dim json = PptImagePaletteExtractor.BuildThemeJson(p, "img_test", "测试主题", "微软雅黑", "tech", "{""density"":""compact""}")
        Dim theme = JObject.Parse(json)
        Check(CStr(theme("id")) = "img_test", "id 保留")
        Check(CStr(theme("styleDirection")) = "tech", "styleDirection 透传")
        Check(theme("designTokens") IsNot Nothing AndAlso CStr(theme("designTokens")("density")) = "compact", "designTokens 透传")
        Dim colors = TryCast(theme("colors"), JObject)
        Check(colors IsNot Nothing, "colors 存在")
        Check(CInt(colors("primary")(0)) = 31 AndAlso CInt(colors("primary")(1)) = 78 AndAlso CInt(colors("primary")(2)) = 121, "primary 为 [31,78,121]")
        ' 组装结果可直接通过校验器（含 tech preset 展开 + WCAG）
        Dim validation = PptThemeDesignValidator.ValidateAndFix(json)
        Check(CBool(validation("success")), "组装主题通过校验器")
    End Sub

    ' ------------------------------------------------------------
    ' 4) 无效路径 → Nothing
    ' ------------------------------------------------------------
    Private Sub TestBadPathReturnsNothing()
        Console.WriteLine("[Path] 无效图片返回 Nothing")
        Check(PptImagePaletteExtractor.Extract("Z:\definitely\missing\img.png") Is Nothing, "不存在的本地路径返回 Nothing")
        Check(PptImagePaletteExtractor.Extract("") Is Nothing, "空路径返回 Nothing")
    End Sub

    ' ------------------------------------------------------------
    ' 5) ResolveImagePath：本地路径直接返回；URL 走下载逻辑
    ' ------------------------------------------------------------
    Private Sub TestResolveImagePath()
        Console.WriteLine("[Path] 路径解析")
        Check(PptImagePaletteExtractor.ResolveImagePath("C:\temp\a.png") = "C:\temp\a.png", "本地路径原样返回")
        Check(PptImagePaletteExtractor.ResolveImagePath("") = "", "空串返回空")
    End Sub

End Module
