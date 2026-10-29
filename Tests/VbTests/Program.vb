Imports System
Imports ExcelAIVSTOAddIn.Services.Mcp

Module Program

    Private _failures As Integer = 0
    Private _total As Integer = 0

    Sub Check(condition As Boolean, label As String)
        _total += 1
        If condition Then
            Console.WriteLine("  ok " & label)
        Else
            _failures += 1
            Console.WriteLine("  FAIL " & label)
        End If
    End Sub

    Sub Main()
        Console.WriteLine("== ExcelMcpCoreUtils 单元测试 ==")
        Console.WriteLine()

        ' --- BuildChunkName：{original}/{i}/{index}/{start}/{end} 占位符 ---
        Console.WriteLine("[BuildChunkName]")
        Check(ExcelMcpCoreUtils.BuildChunkName("{index}", "销售", 3, 10, 20) = "3", "{index} -> 3")
        Check(ExcelMcpCoreUtils.BuildChunkName("{i}", "销售", 3, 10, 20) = "3", "{i} -> 3")
        Check(ExcelMcpCoreUtils.BuildChunkName("{original}", "销售", 3, 10, 20) = "销售", "{original} -> 销售")
        Check(ExcelMcpCoreUtils.BuildChunkName("{original}_{start}_{end}", "销售", 3, 10, 20) = "销售_10_20", "组合占位符")
        Check(ExcelMcpCoreUtils.BuildChunkName("分块{index}", "销售", 7, 1, 100) = "分块7", "前缀+占位符")
        Check(ExcelMcpCoreUtils.BuildChunkName("", "销售", 3, 10, 20) = "", "空模板")
        Check(ExcelMcpCoreUtils.BuildChunkName("无占位符名称", "销售", 3, 10, 20) = "无占位符名称", "无占位符保持不变")

        ' --- ColumnLetterToNumber：A=1, Z=26, AA=27, 不区分大小写 ---
        Console.WriteLine("[ColumnLetterToNumber]")
        Check(ExcelMcpCoreUtils.ColumnLetterToNumber("A") = 1, "A -> 1")
        Check(ExcelMcpCoreUtils.ColumnLetterToNumber("Z") = 26, "Z -> 26")
        Check(ExcelMcpCoreUtils.ColumnLetterToNumber("AA") = 27, "AA -> 27")
        Check(ExcelMcpCoreUtils.ColumnLetterToNumber("AZ") = 52, "AZ -> 52")
        Check(ExcelMcpCoreUtils.ColumnLetterToNumber("a") = 1, "小写 a -> 1")
        Check(ExcelMcpCoreUtils.ColumnLetterToNumber("az") = 52, "小写 az -> 52")

        ' --- NumberToColumnLetter：1=A, 26=Z, 27=AA, 边界回环 ---
        Console.WriteLine("[NumberToColumnLetter]")
        Check(ExcelMcpCoreUtils.NumberToColumnLetter(1) = "A", "1 -> A")
        Check(ExcelMcpCoreUtils.NumberToColumnLetter(26) = "Z", "26 -> Z")
        Check(ExcelMcpCoreUtils.NumberToColumnLetter(27) = "AA", "27 -> AA")
        Check(ExcelMcpCoreUtils.NumberToColumnLetter(52) = "AZ", "52 -> AZ")
        Check(ExcelMcpCoreUtils.NumberToColumnLetter(702) = "ZZ", "702 -> ZZ")
        Check(ExcelMcpCoreUtils.NumberToColumnLetter(703) = "AAA", "703 -> AAA")

        ' --- 字母/数字互逆（回环一致） ---
        Console.WriteLine("[RoundTrip]")
        Dim roundTripOk = True
        For i As Integer = 1 To 2000
            Dim letter = ExcelMcpCoreUtils.NumberToColumnLetter(i)
            If ExcelMcpCoreUtils.ColumnLetterToNumber(letter) <> i Then roundTripOk = False
        Next
        Check(roundTripOk, "1..2000 字母<->数字互逆")

        Console.WriteLine()
        ' 运行 PptDesignEngine 决策链测试，合并失败计数
        Dim pivot = PivotLogicTests.Run()
        _total += pivot.Total
        _failures += pivot.Failures
        Console.WriteLine()

        Dim ppt = PptDesignEngineTests.Run()
        _total += ppt.Total
        _failures += ppt.Failures
        Console.WriteLine()

        Dim theme = PptThemeDesignValidatorTests.Run()
        _total += theme.Total
        _failures += theme.Failures
        Console.WriteLine()

        Dim palette = PptImagePaletteExtractorTests.Run()
        _total += palette.Total
        _failures += palette.Failures

        Console.WriteLine()
        If _failures = 0 Then
            Console.WriteLine("[PASS] 全部 " & _total.ToString() & " 项断言通过")
            Environment.Exit(0)
        Else
            Console.WriteLine("[FAIL] " & _failures.ToString() & " / " & _total.ToString() & " 项断言失败")
            Environment.Exit(1)
        End If
    End Sub

End Module