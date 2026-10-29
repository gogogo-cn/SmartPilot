Imports System.Collections.Generic
Imports System.IO

Namespace Services

    ''' <summary>
    ''' 定位随插件分发的 OCR Python 脚本。独立服务的工作目录可能和插件目录不同。
    ''' </summary>
    Friend Module OcrScriptPaths

        Friend Function Resolve(fileName As String) As String
            Dim candidates As New List(Of String)
            Dim plugin = Environment.GetEnvironmentVariable("SMARTPILOT_PLUGIN_DIR")
            If Not String.IsNullOrEmpty(plugin) Then
                candidates.Add(Path.Combine(plugin, "Services", fileName))
                candidates.Add(Path.Combine(plugin, fileName))
            End If
            candidates.Add(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Services", fileName))
            candidates.Add(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName))
            Try
                Dim asmDir = Path.GetDirectoryName(GetType(OcrServiceManager).Assembly.Location)
                If Not String.IsNullOrEmpty(asmDir) Then
                    candidates.Add(Path.Combine(asmDir, "Services", fileName))
                    candidates.Add(Path.Combine(asmDir, fileName))
                End If
            Catch
            End Try

            For Each candidate In candidates
                If File.Exists(candidate) Then Return candidate
            Next
            Return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Services", fileName)
        End Function

    End Module

End Namespace
