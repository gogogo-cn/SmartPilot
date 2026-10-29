Option Strict On
Option Infer On

Imports System
Imports System.IO
Imports System.Reflection
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports ExcelAIVSTOAddIn.Skills.Testing

Partial Public NotInheritable Class ThisAddIn

#Region "启动逻辑"

    Private Sub ThisAddIn_Startup(sender As Object, e As EventArgs) Handles Me.Startup
        Try
            Try
                Dim loc As String = System.Reflection.Assembly.GetExecutingAssembly().Location
                Dim ft As String = ""
                Try
                    ft = New System.IO.FileInfo(loc).LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                Catch
                End Try
                System.IO.File.WriteAllText(System.IO.Path.Combine(System.IO.Path.GetTempPath(), "addin_path.log"), _
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " | Location=" & loc & " | FileTime=" & ft & vbCrLf)
            Catch
            End Try
            
            _chatLogger = Utils.ChatLogger.Instance
            _chatLogger.LogInfo("ADDIN", "SmartPilot add-in starting")

            _ribbon = New Ribbon()

            InitializeConfigHotReload()

            Try
                Skills.SkillTriggerEngine.Instance.Initialize(Me.Application)
                _chatLogger?.LogInfo("ADDIN", "Skill trigger engine initialized")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[ThisAddIn] Trigger engine init skipped: " & ex.Message)
            End Try

            ' 启用 Skill 热重载：监控系统/自定义技能目录，外部修改 .md 后自动重新挂载
            Try
                Skills.SkillManager.Instance.HotReloadEnabled = True
                _chatLogger?.LogInfo("ADDIN", "Skill hot reload enabled (system + custom skills dirs)")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[ThisAddIn] Skill hot reload enable failed: " & ex.Message)
            End Try

            Try
                AddHandler Me.Application.SheetSelectionChange, AddressOf Application_SheetSelectionChange
                _chatLogger?.LogInfo("ADDIN", "SheetSelectionChange event subscribed")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[ThisAddIn] SheetSelectionChange event subscription failed: " & ex.Message)
            End Try

            Try
                AddHandler Bridges.MobileControlBridge.Instance.OcrCompleted, AddressOf OnMobileOcrCompleted
                AddHandler Bridges.MobileControlBridge.Instance.UploadReceived, AddressOf OnMobileUploadReceived
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[ThisAddIn] MobileControlBridge event subscription failed: " & ex.Message)
            End Try

            ' 后台拉起独立 OCR 服务。识别不在 Excel 进程里跑，Word/PPT 也可以连同一个进程。
            Try
                Task.Run(Sub()
                             Try
                                 Services.OcrHostClient.EnsureStarted()
                             Catch ex As Exception
                                 System.Diagnostics.Debug.WriteLine("[ThisAddIn] OcrHost start skipped: " & ex.Message)
                             End Try
                         End Sub)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[ThisAddIn] OcrHost schedule failed: " & ex.Message)
            End Try

            Try
                _excelBridge = ExcelAIVSTOAddIn.ExcelBridge.Instance
                Services.ThinkingEngine.Instance.Initialize(_excelBridge)
                _chatLogger?.LogInfo("ADDIN", "ThinkingEngine initialized, memory service connected")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[ThisAddIn] ThinkingEngine init failed: " & ex.Message)
            End Try

            ' 初始化任务调度器
            Try
                Bridges.TaskSchedulerBridge.Instance.Initialize()
                _chatLogger?.LogInfo("ADDIN", "TaskSchedulerBridge initialized")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[ThisAddIn] TaskSchedulerBridge init failed: " & ex.Message)
            End Try

            ' 9528 / 9527 不在打开 Excel 时启动。手机遥控在 ShowMobilePanel 里拉起 9528；
            ' 独立网页版在用户打开对应页面时调用 EnsureLocalBridge 拉起 9527。

            ' ======================================================================
            '  [方案A接入] 启动后3秒 后台跑：MCP工具体检 + Skill全量回归
            '  - 完全异步 SafeFork，不阻塞UI；报告写到 TestReports\ 目录
            '  - 开关：环境变量 EXCEL_AI_ENABLE_STARTUP_REGRESSION = 1 强制开启；
            '          未设置时，DEBUG 编译默认开启，RELEASE 默认关闭
            ' ======================================================================
            Try
                Dim enableRegression As Boolean = False
                Dim envSwitch As String = Environment.GetEnvironmentVariable("EXCEL_AI_ENABLE_STARTUP_REGRESSION")
                If Not String.IsNullOrWhiteSpace(envSwitch) Then
                    enableRegression = envSwitch.Trim() = "1"
                Else
#If DEBUG Then
                    enableRegression = True
#Else
                    enableRegression = False
#End If
                End If

                If enableRegression Then
                    Dim __discard = Task.Run(Async Function()
                                                 Await Task.Delay(millisecondsDelay:=3000)
                                                 Try
                                                     ' 自动定位项目根（bin\Debug\XXX → 上 3 层到 csproj 所在目录）
                                                     Dim asmDir As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                                                     Dim projectRoot As String = asmDir
                                                     For _i As Integer = 1 To 3
                                                         If Directory.Exists(Path.Combine(projectRoot, "Skills")) Then Exit For
                                                         projectRoot = Directory.GetParent(projectRoot)?.FullName
                                                         If String.IsNullOrWhiteSpace(projectRoot) Then Exit For
                                                     Next
                                                     If Not Directory.Exists(Path.Combine(projectRoot, "Skills")) Then projectRoot = asmDir
                                                     Dim reportDir As String = Path.Combine(If(asmDir, projectRoot), "TestReports")
                                                     Directory.CreateDirectory(reportDir)

                                                     _chatLogger?.LogInfo("TEST", "后台开始：MCP工具完善度体检 + Skill全量回归 → " & reportDir)

                                                     ' RegressionProgress 是 POCO Class（不用ValueTuple，避免Option Strict命名成员访问失败）
                                                     Dim prog As IProgress(Of McpAndSkillTestRunner.RegressionProgress) =
                                                         New Progress(Of McpAndSkillTestRunner.RegressionProgress)(
                                                             Sub(t)
                                                                 ' t 是 Class 属性，无后期绑定，Option Strict绝对安全
                                                                 System.Diagnostics.Debug.WriteLine(String.Format("[REGRESSION] {0,3}%  [{1}] {2}",
                                                                     CInt(t.Percent), CStr(t.Stage), CStr(t.Message)))
                                                             End Sub)
                                                     ' 用 Option Infer 推断 fullReport 的真实类型，避免写显式长类型名报错
                                                     Dim fullReport = Await McpAndSkillTestRunner.RunAllAsync(
                                                         projectRoot,
                                                         reportDir,
                                                         SkillTestLevel.DefinitionValidation,
                                                         SkillTestLevel.EndToEndIntegration,
                                                         prog)

                                                     _chatLogger?.LogInfo("TEST", "完成！" & fullReport.SummaryLine)
                                                     _chatLogger?.LogInfo("TEST", "报告：" & fullReport.CombinedReportMarkdownPath)
                                                 Catch rEx As Exception
                                                     System.Diagnostics.Debug.WriteLine("[REGRESSION] 失败: " & rEx.ToString())
                                                     _chatLogger?.LogError("TEST", "后台回归运行异常: " & rEx.Message, rEx)
                                                 End Try
                                                 ' 修复 BC42105：Async Function 必须所有路径有Return值
                                                 Return Task.CompletedTask
                                             End Function)
                End If
            Catch startupTestInitEx As Exception
                System.Diagnostics.Debug.WriteLine("[ThisAddIn] Mcp+Skill回归启动器初始化失败(不影响AddIn其它功能): " & startupTestInitEx.Message)
            End Try

        Catch ex As Exception
            Dim errorMsg As String = $"Unhandled exception during add-in startup.{vbCrLf}{vbCrLf}" &
                $"Error: {ex.Message}{vbCrLf}" &
                $"Type: {ex.GetType().FullName}{vbCrLf}" &
                $"Stack: {ex.StackTrace}"

            Try
                System.Diagnostics.Debug.WriteLine(errorMsg)
                System.Diagnostics.EventLog.WriteEntry("ExcelAIVSTOAddIn", errorMsg,
                    System.Diagnostics.EventLogEntryType.Error, 1000)
            Catch
            End Try

            Try
                MessageBox.Show(errorMsg, "SmartPilot - Startup Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error)
            Catch
            End Try
        End Try
    End Sub

#End Region

End Class