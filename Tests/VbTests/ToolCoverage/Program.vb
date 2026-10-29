Imports System

Module Program
    Sub Main(args() As String)
        Console.WriteLine("== PPT/Word 原子工具覆盖测试 ==")
        Console.WriteLine("运行目录: " & AppContext.BaseDirectory)
        Try
            ToolCoverage.RunAll()
            Console.WriteLine()
            Console.WriteLine(">>>> ALL TESTS PASSED <<<<")
        Catch ex As Exception
            Console.WriteLine()
            Console.WriteLine(">>>> TEST FAILED: " & ex.Message & " <<<<")
            Environment.Exit(1)
        End Try
    End Sub
End Module
