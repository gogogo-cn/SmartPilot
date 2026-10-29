' COM-free 纯逻辑工具集：从 ExcelMcpTools partial 类中抽取的可测共享方法。
' 该文件同时被 add-in 工程与 Tests\VbTests 测试工程编译（共享源文件模式），
' 因此不得引用任何 COM / Excel Interop 类型，且保持 Option Strict On 可编译。
Option Strict On
Option Explicit On
Option Infer On

Imports System

Namespace Services.Mcp

    ''' <summary>
    ''' 纯逻辑工具集（无 Excel COM 依赖），便于在独立 VB 测试工程中单元测试。
    ''' 抽取来源：ExcelMcpTools.DataOps 等 partial 文件的 Private 辅助方法（零行为变更搬移）。
    ''' </summary>
    Public NotInheritable Class ExcelMcpCoreUtils

        Private Sub New()
        End Sub

        ''' <summary>
        ''' 按下发命名模板生成分块文件名/表名，支持 {original}/{i}/{index}/{start}/{end} 占位符。
        ''' </summary>
        Public Shared Function BuildChunkName(pattern As String, original As String, chunk As Integer, startRow As Integer, endRow As Integer) As String
            Dim name = pattern _
                .Replace("{original}", original) _
                .Replace("{i}", chunk.ToString()) _
                .Replace("{index}", chunk.ToString()) _
                .Replace("{start}", startRow.ToString()) _
                .Replace("{end}", endRow.ToString())
            Return name
        End Function

        ''' <summary>
        ''' 列字母转列号：A=1, Z=26, AA=27（不区分大小写）。
        ''' </summary>
        Public Shared Function ColumnLetterToNumber(letter As String) As Integer
            Dim result = 0
            For Each c As Char In letter.ToUpper()
                result = result * 26 + (AscW(c) - AscW("A"c) + 1)
            Next
            Return result
        End Function

        ''' <summary>
        ''' 列号转列字母：1=A, 26=Z, 27=AA。
        ''' </summary>
        Public Shared Function NumberToColumnLetter(num As Integer) As String
            Dim result = ""
            While num > 0
                num -= 1
                result = ChrW(AscW("A"c) + (num Mod 26)) & result
                num \= 26
            End While
            Return result
        End Function

    End Class

End Namespace