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

    ''' <summary>
    ''' VBA Code Sandbox - Provides secure VBA code execution environment
    ''' Includes code review, security checks, execution isolation and other features
    ''' </summary>
    Partial Public Class VBACodeSandbox

        ' —— 沙箱执行策略常量 ——
        Private Const MAX_CODE_LENGTH As Integer = 50000        ' 单次执行代码最大字符数
        Private Const MAX_VBA_LINE_COUNT As Integer = 500       ' VBA 代码最大行数
        Private Const MAX_NESTING_DEPTH As Integer = 10         ' 最大括号嵌套深度
        Private Const MAX_OBFUSCATED_LINE_LENGTH As Integer = 500 ' 混淆检测：单行最大长度
        Private Const OBFUSCATION_SPECIAL_CHAR_RATIO As Double = 0.3 ' 混淆检测：特殊字符比例阈值
        Private Const SANDBOX_TIMEOUT_SECONDS As Integer = 30   ' 沙箱执行超时（秒）
        Private Const SANDBOX_MAX_COPY_ROWS As Long = 1000      ' 沙箱回写最大行数
        Private Const SANDBOX_MAX_COPY_COLS As Integer = 50     ' 沙箱回写最大列数
        Private Const CHR_OBFUSCATION_THRESHOLD As Integer = 3  ' Chr() 混淆检测最小调用次数

        ' Dangerous keywords - These are direct VBA statement/function calls
        Private Shared ReadOnly DangerousKeywords As String() = New String() {
            "Kill ", "Kill(",
            "DeleteFile",
            "RmDir ", "RmDir(", "Rmdir ", "Rmdir(",
            "FileCopy ", "FileCopy(",
            "SetAttr ", "SetAttr(",
            "Shell ", "Shell(",
            "AppActivate ", "AppActivate(",
            "SendKeys ", "SendKeys(",
            "Environ(",
            "RegWrite", "RegDelete", "RegRead",
            "ChDir ", "ChDir(", "ChDrive ", "ChDrive(",
            "MkDir ", "MkDir(",
            "Name ", "Name(",
            "SetAttr ", "SetAttr(",
            "DDEExecute",
            "DDEInitiate",
            "DDETerminate",
            "ExecuteExcel4Macro",
            "Declare Function",
            "Declare Sub",
            "Lib """,
            "TerminateProcess",
            "CreateProcess",
            "WinExec",
            "LoadLibrary",
            "GetProcAddress",
            "VirtualAlloc",
            "WriteProcessMemory",
            "CreateRemoteThread",
            "OpenProcess",
            "CloseHandle",
            "ShellExecute",
            "ShellExecuteA",
            "ShellExecuteW",
            "URLDownloadToFile",
            "URLDownloadToFileA",
            "URLDownloadToFileW",
            "InetExecute",
            "WinHttp.WinHttpRequest",
            "WinHttp.WinHttpRequest.5.1",
            "MSXML2.ServerXMLHTTP",
            "Microsoft.XMLHTTP",
            "ADODB.Stream",
            "Scriptlet.TypeLib",
            "HTMLFile",
            "HTMLDocument",
            "RunDll32",
            "mshta.exe",
            "cmd.exe",
            "powershell.exe",
            "pwsh.exe",
            "wscript.exe",
            "cscript.exe",
            "regsvr32",
            "msiexec",
            "certutil",
            "bitsadmin",
            "wmic ",
            "Open """,
            "Open ",
            "Put #",
            "Get #",
            "Print #",
            "Write #",
            "Input #",
            "Close #",
            "KillProcess",
            "EndProcess",
            "Process.Kill",
            "Stop-Process"
        }

        ' Critical keywords - Reject execution immediately if found
        ' Note: Scripting.FileSystemObject is NOT in this list — it is downgraded
        ' to a warning (Medium risk). Destructive FSO methods (.DeleteFile,
        ' .DeleteFolder, .CopyFile, .MoveFile, .CreateTextFile, .OpenTextFile,
        ' .OpenAsTextStream) are still blocked by CheckObfuscationBypasses.
        ' 【安全策略调整】Critical = 无条件阻止（仅保留真正可执行任意代码的危险项）
        ' 以下已移至 DangerousKeywords(High) 级别，因为它们是 LLM 生成 Excel VBA 的常见合法操作：
        '   Shell(, CreateObject("WScript.Shell"), CreateObject("WScript.Network"), CreateObject("Shell.Application")
        '   这些操作仍会被检测为 High 风险，但在 requireConfirmation:=False 模式下（自定义Skill/工作流）可直接执行
        Private Shared ReadOnly CriticalKeywords As String() = New String() {
            "CreateObject(""MSScriptControl.ScriptControl"")",
            "CreateObject(""InternetExplorer.Application"")",
            "CallByName",
            "InvokeMember",
            "GetRef",
            "ADODB.Connection",
            "MSScriptControl.ScriptControl",
            "ExecuteExcel4Macro(",
            "ExecuteExcel4Macro ",
            "Declare Function ",
            "Declare Sub ",
            "XlmRegister",
            "REGISTER ",
            "REGISTER(",
            "CALLBACK",
            "EnumWindows",
            "FindWindow",
            "FindWindowA",
            "FindWindowW",
            "SendMessage",
            "PostMessage",
            "SetWindowsHookEx",
            "WH_KEYBOARD_LL",
            "WH_MOUSE_LL",
            "GetAsyncKeyState",
            "CreateRemoteThread",
            "WriteProcessMemory",
            "NtUnmapViewOfSection",
            "QueueUserAPC",
            "SetThreadContext"
        }

        ' Obfuscation-related critical patterns (detected via regex/semantic analysis)
        ' These match attempts to construct dangerous strings via concatenation or Chr()
        Private Shared ReadOnly ObfuscationBypassPatterns As String() = New String() {
            """wscript.shell""",
            """wscript.network""",
            """scripting.filesystemobject""",
            """shell.application""",
            """internetexplorer.application""",
            """msxml2.xmlhttp""",
            """winhttp.winhttprequest""",
            """adodb.stream""",
            """adodb.connection""",
            """microsoft.xmlhttp""",
            "scriptcontrol",
            "vbscript.regexp",
            "msscriptcontrol.scriptcontrol"
        }

        ' Additional bypass signatures checked via normalized code
        Private Shared ReadOnly HighRiskFunctions As String() = New String() {
            "Eval(",
            "Execute(",
            "ExecuteGlobal(",
            "AddressOf ",
            "Application.Run(",
            "VBProject",
            "VBE.",
            "CommandBars(",
            "GetObject(",
            "DDEExecute",
            "DDEInitiate",
            "DDETerminate",
            "Application.DDEExecute",
            "ExecuteExcel4Macro(",
            "Run(""",
            "Run """,
            ".Run(""",
            ".Run """,
            "ThisWorkbook.VBProject",
            "ActiveWorkbook.VBProject",
            "VBComponents",
            "CodeModule",
            "AddFromFile",
            "AddFromString",
            "Import(",
            "Export(",
            "ReplaceLine(",
            "InsertLines(",
            "DeleteLines(",
            "VBE.ActiveVBProject",
            "TrustAccessToVBProject",
            "OnTime ",
            "OnTime(",
            "Wait(",
            "Application.Wait",
            "SendKeys(",
            "DoEvent",
            "AppActivate ",
            "EnableCancelKey",
            "Calculation = xlCalculationAutomatic",
            "Calculation = xlCalculationManual",
            "EnableEvents = False",
            "EnableEvents = True",
            "DisplayAlerts = False",
            "DisplayAlerts = True",
            "ScreenUpdating = False",
            "ScreenUpdating = True",
            "Visible = False",
            "WindowState = xlMinimized",
            "WindowState = xlHidden"
        }

        ' —— 白名单配置键名 ——
        Private Const WHITELIST_CONFIG_KEY As String = "VBA.CustomWhitelist"

        ' 内置安全操作白名单（只读，不允许运行时修改）
        Private Shared ReadOnly BuiltInSafeOperations As String() = New String() {
            "Range(",
            "Cells(",
            "Worksheets(",
            "ActiveSheet",
            "ActiveWorkbook",
            "Selection",
            "MsgBox",
            "InputBox",
            "Debug.Print",
            "Format(",
            "Left(",
            "Right(",
            "Mid(",
            "Len(",
            "Trim(",
            "UCase(",
            "LCase(",
            "Replace(",
            "Split(",
            "Join(",
            "CStr(",
            "CInt(",
            "CLng(",
            "CDbl(",
            "CDate(",
            "Now",
            "Date",
            "Time",
            "Year(",
            "Month(",
            "Day(",
            "Hour(",
            "Minute(",
            "Second("
        }

        ' 运行时白名单（内置 + 用户自定义），支持动态添加
        Private Shared _runtimeWhitelist As List(Of String) = Nothing
        Private Shared _whitelistLock As New Object()

        ''' <summary>
        ''' 获取当前完整的安全操作白名单（内置 + 用户自定义持久化项）
        ''' </summary>
        Private Shared ReadOnly Property AllowedSafeOperations As List(Of String)
            Get
                If _runtimeWhitelist Is Nothing Then
                    SyncLock _whitelistLock
                        EnsureWhitelistInitialized()
                    End SyncLock
                End If
                Return _runtimeWhitelist
            End Get
        End Property

        ''' <summary>
        ''' 从配置文件加载用户自定义白名单
        ''' </summary>
        Private Shared Sub LoadUserWhitelist()
            Try
                Dim saved As String = Utils.ConfigurationManager.Instance.GetSetting(WHITELIST_CONFIG_KEY, "")
                If Not String.IsNullOrWhiteSpace(saved) Then
                    Dim items() As String = saved.Split(New Char() {","c}, StringSplitOptions.RemoveEmptyEntries)
                    For Each item In items
                        Dim trimmed As String = item.Trim()
                        If Not String.IsNullOrEmpty(trimmed) AndAlso Not _runtimeWhitelist.Contains(trimmed) Then
                            _runtimeWhitelist.Add(trimmed)
                        End If
                    Next
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"加载用户白名单失败: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' 持久化用户自定义白名单到配置文件
        ''' </summary>
        Private Shared Sub SaveUserWhitelist()
            Try
                ' 只保存非内置的自定义项
                Dim userItems As New List(Of String)()
                For Each op In _runtimeWhitelist
                    If Not BuiltInSafeOperations.Contains(op) Then
                        userItems.Add(op)
                    End If
                Next
                Dim serialized As String = String.Join(",", userItems)
                Utils.ConfigurationManager.Instance.SetSetting(WHITELIST_CONFIG_KEY, serialized)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"保存用户白名单失败: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' 确保运行时白名单已初始化（内部方法，需在锁内调用）
        ''' </summary>
        Private Shared Sub EnsureWhitelistInitialized()
            If _runtimeWhitelist Is Nothing Then
                _runtimeWhitelist = New List(Of String)(BuiltInSafeOperations)
                LoadUserWhitelist()
            End If
        End Sub

        ''' <summary>
        ''' 向白名单添加一个或多个操作（带持久化）
        ''' </summary>
        ''' <param name="operations">要添加的操作名称数组</param>
        ''' <returns>实际新增的操作数量</returns>
        Public Shared Function AddToWhitelist(ParamArray operations() As String) As Integer
            If operations Is Nothing OrElse operations.Length = 0 Then Return 0

            SyncLock _whitelistLock
                EnsureWhitelistInitialized()
                Dim added As Integer = 0
                For Each op In operations
                    If String.IsNullOrWhiteSpace(op) Then Continue For
                    Dim clean As String = op.Trim()
                    If Not _runtimeWhitelist.Contains(clean) Then
                        _runtimeWhitelist.Add(clean)
                        added += 1
                    End If
                Next
                If added > 0 Then
                    SaveUserWhitelist()
                End If
                Return added
            End SyncLock
        End Function

        ''' <summary>
        ''' 获取当前白名单中所有操作的副本
        ''' </summary>
        Public Shared Function GetWhitelistOperations() As List(Of String)
            SyncLock _whitelistLock
                EnsureWhitelistInitialized()
                Return New List(Of String)(_runtimeWhitelist)
            End SyncLock
        End Function

        ''' <summary>
        ''' 从白名单移除指定操作（仅限用户自定义项，内置项不可移除）
        ''' </summary>
        ''' <returns>是否实际移除了某项</returns>
        Public Shared Function RemoveFromWhitelist(opName As String) As Boolean
            If String.IsNullOrWhiteSpace(opName) Then Return False

            SyncLock _whitelistLock
                EnsureWhitelistInitialized()
                Dim clean As String = opName.Trim()
                If BuiltInSafeOperations.Contains(clean) Then
                    ' 内置项不允许移除
                    Return False
                End If
                Dim removed As Boolean = _runtimeWhitelist.Remove(clean)
                If removed Then
                    SaveUserWhitelist()
                End If
                Return removed
            End SyncLock
        End Function

        ''' <summary>
        ''' Code security check result
        ''' </summary>
        Public Class SecurityCheckResult
            Public Property IsSafe As Boolean
            Public Property Warnings As List(Of String)
            Public Property Errors As List(Of String)
            Public Property RiskLevel As RiskLevel
            Public Property CodeHash As String
            Public Property CheckTime As DateTime
            ' 存储不在白名单中的操作（用于"加入白名单"功能）
            Public Property UnknownWhitelistOps As List(Of String)

            Public Sub New()
                Warnings = New List(Of String)()
                Errors = New List(Of String)()
                UnknownWhitelistOps = New List(Of String)()
                RiskLevel = RiskLevel.Safe
                CheckTime = DateTime.Now
            End Sub

            ''' <summary>
            ''' Get summary of check result
            ''' </summary>
            Public Function GetSummary() As String
                Dim summary As New StringBuilder()
                summary.AppendLine($"Check Time: {CheckTime:yyyy-MM-dd HH:mm:ss}")
                summary.AppendLine($"Code Hash: {CodeHash}")
                summary.AppendLine($"Security Status: {If(IsSafe, "Safe", "Unsafe")}")
                summary.AppendLine($"Risk Level: {RiskLevel}")
                summary.AppendLine($"Warning Count: {Warnings.Count}")
                summary.AppendLine($"Error Count: {Errors.Count}")
                Return summary.ToString()
            End Function
        End Class

        Public Enum RiskLevel
            Safe = 0
            Low = 1
            Medium = 2
            High = 3
            Critical = 4
        End Enum

        ''' <summary>
        ''' Common Excel operations detected in VBA code
        ''' </summary>
        Public Enum CommonExcelOperation
            None
            ClearSheet
            SplitWorkbook
            MergeSheets
            RemoveDuplicates
            CreateChart
        End Enum
    End Class

End Namespace
