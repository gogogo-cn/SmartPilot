' ============================================================================
' HostConfig - 宿主无关的配置存储（AI 模型配置等）
'
' 存储位置：%APPDATA%\SmartPilotOffice\ai_config.json
' 首次读取时若本文件不存在，会尝试从 Excel 版加载项导出的 llm_config.json
' 导入（字段结构一致：provider / model / apiKey / apiUrl / baseUrl）。
' ============================================================================
Option Strict Off
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Text

Public Module HostConfig

    Private ReadOnly RootDir As String =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SmartPilotOffice")
    Private ReadOnly ConfigPath As String = Path.Combine(RootDir, "ai_config.json")

    Private Const EMPTY_CONFIG As String = "{""provider"":"""",""model"":"""",""apiKey"":"""",""apiUrl"":"""",""baseUrl"":""""}"

    ''' <summary>读取 AI 配置（JSON 字符串）</summary>
    Public Function GetAIConfigJson() As String
        Try
            If File.Exists(ConfigPath) Then
                Return File.ReadAllText(ConfigPath)
            End If
            Dim imported = TryImportFromExcelConfig()
            If imported IsNot Nothing Then
                Directory.CreateDirectory(RootDir)
                File.WriteAllText(ConfigPath, imported)
                Return imported
            End If
            Return EMPTY_CONFIG
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("[HostConfig] 读取失败: " & ex.Message)
            Return EMPTY_CONFIG
        End Try
    End Function

    ''' <summary>
    ''' 清掉本机 API 密钥：写出空的 ai_config.json（避免下次从 Excel 配置再导入），
    ''' 并抹掉 llm_config.json 与 Excel config.ini 里的密钥。不删除授权会话。
    ''' </summary>
    Public Function ClearStoredApiKey() As String
        Try
            Dim provider = ""
            Dim model = ""
            Dim apiUrl = ""
            If File.Exists(ConfigPath) Then
                Dim raw = File.ReadAllText(ConfigPath)
                provider = ExtractJsonString(raw, "provider")
                model = ExtractJsonString(raw, "model")
                apiUrl = FirstNonEmpty(ExtractJsonString(raw, "apiUrl"), ExtractJsonString(raw, "baseUrl"))
            End If
            Directory.CreateDirectory(RootDir)
            Dim written = "{" &
                """provider"":""" & Escape(provider) & """," &
                """apiType"":""" & Escape(provider) & """," &
                """model"":""" & Escape(model) & """," &
                """apiUrl"":""" & Escape(apiUrl) & """," &
                """baseUrl"":""" & Escape(apiUrl) & """," &
                """apiKey"":""""" &
                "}"
            File.WriteAllText(ConfigPath, written)
            Dim left = ExtractJsonString(File.ReadAllText(ConfigPath), "apiKey")
            If left <> "" Then
                Return "{""success"":false,""error"":""apiKey 仍在""}"
            End If
            If Not BlankExportedKeyFiles() Then
                Return "{""success"":false,""error"":""导出的密钥文件没有清干净""}"
            End If
            Return "{""success"":true}"
        Catch ex As Exception
            Return "{""success"":false,""error"":""" & Escape(ex.Message) & """}"
        End Try
    End Function

    Private Function BlankExportedKeyFiles() As Boolean
        Dim ok = True
        For Each candidate As String In ExcelLlmConfigCandidates()
            If Not File.Exists(candidate) Then Continue For
            Dim raw = File.ReadAllText(candidate)
            Dim cleared = BlankApiKeyJson(raw)
            File.WriteAllText(candidate, cleared)
            If ExtractJsonString(cleared, "apiKey") <> "" OrElse ExtractJsonString(cleared, "api_key") <> "" Then
                ok = False
            End If
        Next
        Dim ini = ExcelSettingsIniPath()
        If File.Exists(ini) Then
            Dim kept As New List(Of String)()
            For Each iniLine As String In File.ReadAllLines(ini)
                Dim sep = iniLine.IndexOf("="c)
                Dim key = If(sep > 0, iniLine.Substring(0, sep).Trim(), "")
                If IsKeyMaterial(key) Then Continue For
                kept.Add(iniLine)
            Next
            File.WriteAllLines(ini, kept)
        End If
        Return ok
    End Function

    Private Function BlankApiKeyJson(raw As String) As String
        Dim rx = System.Text.RegularExpressions.RegexOptions.IgnoreCase
        Dim cleared = System.Text.RegularExpressions.Regex.Replace(raw, """apiKey""\s*:\s*""(?:\\.|[^""])*""", """apiKey"":""""", rx)
        cleared = System.Text.RegularExpressions.Regex.Replace(cleared, """apiKey""\s*:\s*null", """apiKey"":""""", rx)
        cleared = System.Text.RegularExpressions.Regex.Replace(cleared, """api_key""\s*:\s*""(?:\\.|[^""])*""", """api_key"":""""", rx)
        cleared = System.Text.RegularExpressions.Regex.Replace(cleared, """api_key""\s*:\s*null", """api_key"":""""", rx)
        Return cleared
    End Function

    Private Function IsKeyMaterial(key As String) As Boolean
        If String.IsNullOrEmpty(key) Then Return False
        Dim lower = key.ToLowerInvariant()
        If lower.Contains("token") Then Return False
        Return lower.Contains("apikey") OrElse lower.Contains("api_key") OrElse
            lower.Contains("secret") OrElse lower.Contains("password") OrElse lower.Contains("privatekey")
    End Function

    Private Function ExcelLlmConfigCandidates() As String()
        Dim appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        Dim localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
        Return New String() {
            Path.Combine(appData, "ExcelAIVSTOAddIn", "Config", "llm_config.json"),
            Path.Combine(localAppData, "ExcelAIVSTOAddIn", "Config", "llm_config.json"),
            Path.Combine(appData, "ExcelAIVSTOAddIn", "llm_config.json")
        }
    End Function

    Private Function ExcelSettingsIniPath() As String
        Dim appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
        Return Path.Combine(appData, "ExcelAIVSTOAddIn", "config.ini")
    End Function

    ''' <summary>保存 AI 配置（前端传入完整 JSON 字符串）</summary>
    Public Function SaveAIConfigJson(json As String) As String
        Try
            If String.IsNullOrWhiteSpace(json) Then
                Return "{""success"":false,""error"":""配置为空""}"
            End If
            Directory.CreateDirectory(RootDir)
            File.WriteAllText(ConfigPath, json)
            Return "{""success"":true}"
        Catch ex As Exception
            Return "{""success"":false,""error"":""" & Escape(ex.Message) & """}"
        End Try
    End Function

    ''' <summary>从 Excel 版加载项已导出的 llm_config.json 导入（一次性）</summary>
    Private Function TryImportFromExcelConfig() As String
        Try
            Dim appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            Dim localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
            Dim candidates As String() = {
                Path.Combine(appData, "ExcelAIVSTOAddIn", "Config", "llm_config.json"),
                Path.Combine(localAppData, "ExcelAIVSTOAddIn", "Config", "llm_config.json"),
                Path.Combine(appData, "ExcelAIVSTOAddIn", "llm_config.json")
            }
            For Each p As String In candidates
                If File.Exists(p) Then
                    Dim raw = File.ReadAllText(p)
                    Dim provider = ExtractJsonString(raw, "provider")
                    Dim model = ExtractJsonString(raw, "model")
                    Dim apiKey = ExtractJsonString(raw, "apiKey")
                    Dim apiUrl = FirstNonEmpty(ExtractJsonString(raw, "apiUrl"), ExtractJsonString(raw, "baseUrl"))
                    If String.IsNullOrEmpty(provider) AndAlso String.IsNullOrEmpty(apiKey) Then Continue For
                    Return "{" &
                        """provider"":""" & Escape(provider) & """," &
                        """apiType"":""" & Escape(provider) & """," &
                        """model"":""" & Escape(model) & """," &
                        """apiUrl"":""" & Escape(apiUrl) & """," &
                        """baseUrl"":""" & Escape(apiUrl) & """," &
                        """apiKey"":""" & Escape(apiKey) & """" &
                        "}"
                End If
            Next
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("[HostConfig] 导入 Excel 配置失败: " & ex.Message)
        End Try
        Return Nothing
    End Function

    ' —— 极简 JSON 取值（避免引入第三方 JSON 依赖）——
    Private Function ExtractJsonString(json As String, key As String) As String
        Try
            Dim mark = """" & key & """"
            Dim idx = json.IndexOf(mark, StringComparison.OrdinalIgnoreCase)
            If idx < 0 Then Return ""
            Dim colon = json.IndexOf(":"c, idx + mark.Length)
            If colon < 0 Then Return ""
            Dim p = colon + 1
            While p < json.Length AndAlso Char.IsWhiteSpace(json(p))
                p += 1
            End While
            If p >= json.Length OrElse json(p) <> """"c Then Return ""
            p += 1
            Dim sb As New StringBuilder()
            While p < json.Length
                Dim c = json(p)
                If c = "\"c AndAlso p + 1 < json.Length Then
                    sb.Append(json(p + 1))
                    p += 2
                    Continue While
                End If
                If c = """"c Then Exit While
                sb.Append(c)
                p += 1
            End While
            Return sb.ToString()
        Catch
            Return ""
        End Try
    End Function

    Private Function FirstNonEmpty(ParamArray values As String()) As String
        For Each v As String In values
            If Not String.IsNullOrEmpty(v) Then Return v
        Next
        Return ""
    End Function

    Private Function Escape(s As String) As String
        If s Is Nothing Then Return ""
        Return s.Replace("\", "\\").Replace("""", "\""").Replace(vbCrLf, " ").Replace(vbLf, " ").Replace(vbTab, " ")
    End Function
End Module
