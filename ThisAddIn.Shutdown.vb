Option Strict On
Option Infer On

Imports System
Imports ExcelAIVSTOAddIn.Utils

Partial Public NotInheritable Class ThisAddIn

#Region "关闭逻辑"

    Private Sub ThisAddIn_Shutdown(sender As Object, e As EventArgs) Handles Me.Shutdown
        Dim sw As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
        Dim shutdownMon As Utils.ShutdownMonitor = Utils.ShutdownMonitor.Instance
        shutdownMon.StartMonitoring()
        Try
            _chatLogger?.LogInfo("ADDIN", "SmartPilot add-in starting shutdown...")

            shutdownMon.BeginStep("Stop mobile remote service")
            Try
                ' 提前停 9528（监听 + 断开活跃连接）：浏览器独立页后续请求立即失败，
                ' 避免关闭流程期间产生新的 HandleClient 线程持有 COM 引用，导致 GC 死锁卡顿。
                ' （BridgeRegistry.ShutdownAll 仍会兜底再次 Dispose）
                Try : Bridges.MobileControlBridge.Instance.StopServer() : Catch : End Try
                shutdownMon.EndStep("Success")
            Catch ex As Exception
                shutdownMon.EndStep("Failed", ex.Message)
            End Try

            shutdownMon.BeginStep("Stop local bridge server (py-excel-bridge)")
            Try
                ' 仅终止本会话由 EnsureStarted 拉起的子进程；若端口已被外部桥占用（未由本会话启动），
                ' _process 为 Nothing → StopServer 直接返回，不会误杀用户手动启动的桥。
                Services.BridgeServerLauncher.StopServer()
                shutdownMon.EndStep("Success")
            Catch ex As Exception
                shutdownMon.EndStep("Failed", ex.Message)
            End Try

            shutdownMon.BeginStep("Stop skill trigger")
            Try
                Skills.SkillTriggerEngine.Instance.Shutdown()
                shutdownMon.EndStep("Success")
            Catch ex As Exception
                shutdownMon.EndStep("Failed", ex.Message)
            End Try

            shutdownMon.BeginStep("Stop task scheduler")
            Try
                Bridges.TaskSchedulerBridge.Instance.Dispose()
                shutdownMon.EndStep("Success")
            Catch ex As Exception
                shutdownMon.EndStep("Failed", ex.Message)
            End Try

            shutdownMon.BeginStep("Stop OCR service")
            Try
                ' 统一生命周期：由 BridgeRegistry.ShutdownAll 释放 OcrBridge(→OcrServiceManager)
                ' 下方「Stop OCR engine process」步骤仍直接释放 OcrServiceManager 作为兜底
                shutdownMon.EndStep("Skipped（由 BridgeRegistry 统一释放）")
            Catch ex As Exception
                shutdownMon.EndStep("Failed", ex.Message)
            End Try

            shutdownMon.BeginStep("Stop config monitor")
            If _configWatcher IsNot Nothing Then
                Try
                    RemoveHandler _configWatcher.ConfigFileChanged, AddressOf OnConfigFileChanged
                    _configWatcher.Dispose()
                    _configWatcher = Nothing
                    shutdownMon.EndStep("Success")
                Catch ex As Exception
                    shutdownMon.EndStep("Failed", ex.Message)
                End Try
            Else
                shutdownMon.EndStep("Skipped", "Config monitor not initialized")
            End If

            shutdownMon.BeginStep("Stop task pane")
            If _taskPaneManager IsNot Nothing Then
                Try
                    ' Dispose 已调用 StopAllWebViews + RemoveWebViewControlsBeforeClose，并异步释放 WebView2
                    _taskPaneManager.Dispose()
                    _taskPaneManager = Nothing
                    shutdownMon.EndStep("Success")
                Catch ex As Exception
                    shutdownMon.EndStep("Failed", ex.Message)
                End Try
            Else
                shutdownMon.EndStep("Skipped", "Task pane manager not initialized")
            End If

            shutdownMon.BeginStep("Stop WebView2")
            If _webViewManager IsNot Nothing Then
                Try
                    ' Dispose 已调用 StopWebView，并异步释放 WebView2
                    _webViewManager.Dispose()
                    _webViewManager = Nothing
                    shutdownMon.EndStep("Success")
                Catch ex As Exception
                    shutdownMon.EndStep("Failed", ex.Message)
                End Try
            Else
                shutdownMon.EndStep("Skipped", "WebView2 manager not initialized")
            End If

            shutdownMon.BeginStep("Stop screenshot service")
            If _screenshotService IsNot Nothing Then
                Try
                    _screenshotService.Dispose()
                    _screenshotService = Nothing
                    shutdownMon.EndStep("Success")
                Catch ex As Exception
                    shutdownMon.EndStep("Failed", ex.Message)
                End Try
            Else
                shutdownMon.EndStep("Skipped", "Screenshot service not initialized")
            End If

            shutdownMon.BeginStep("Release bridges (BridgeRegistry)")
            Try
                ' [Critical fix] 统一生命周期：McpBridge/OcrBridge/RagBridge/MobileControlBridge 等
                ' 资源型子 Bridge 已在各自构造函数中自助注册进 BridgeRegistry，
                ' 由 ShutdownAll 统一释放其持有的 TcpListener/MCP 客户端/Python 进程/HttpClient，
                ' 避免 Excel 关闭时挂起或资源残留。无论 _excelBridge 是否已初始化都应执行。
                Bridges.BridgeRegistry.Instance.ShutdownAll()
                If _excelBridge IsNot Nothing Then _excelBridge = Nothing
                shutdownMon.EndStep("Success")
            Catch ex As Exception
                shutdownMon.EndStep("Failed", ex.Message)
            End Try

            ' 识别引擎在独立进程 OcrHost 里。这里的 Dispose 在插件进程中是空操作，
            ' 不会关掉 Excel/Word/PPT 共用的那份服务。
            shutdownMon.BeginStep("Stop OCR engine process")
            Try
                Services.OcrServiceManager.Instance.Dispose()
                shutdownMon.EndStep("Success")
            Catch ex As Exception
                shutdownMon.EndStep("Failed", ex.Message)
            End Try

            shutdownMon.BeginStep("Release RAG knowledge base")
            Try
                Services.RAG.KnowledgeBaseManager.Instance.Dispose()
                shutdownMon.EndStep("Success")
            Catch ex As Exception
                shutdownMon.EndStep("Failed", ex.Message)
            End Try

            shutdownMon.BeginStep("Persist conversation memory")
            Try
                Services.ConversationMemoryService.Instance.PersistCurrentConversation()
                Services.ConversationMemoryService.Instance.SavePreferences()
                Services.ConversationMemoryService.Instance.Dispose()
                shutdownMon.EndStep("Success")
            Catch ex As Exception
                shutdownMon.EndStep("Failed", ex.Message)
            End Try

            shutdownMon.BeginStep("Stop log service")
            Try
                Utils.LogServiceManager.Instance.StopBackgroundWriter()
                shutdownMon.EndStep("Success")
            Catch ex As Exception
                shutdownMon.EndStep("Failed", ex.Message)
            End Try

            ' [Final cleanup] 在不阻塞关闭的情况下释放 COM 引用。
            ' 注意：GC.WaitForPendingFinalizers 会等待 Finalizer 线程；若 9528 的请求线程
            ' （HandleClient）持有 COM 对象，其 Finalizer 需回 STA 主线程才能释放，
            ' 而主线程正等在这里 → 死锁（"打开独立页后关闭 Excel 卡"）。
            ' 因此放到后台线程执行，主线程不等待。
            Try
                Dim gcThread As New Threading.Thread(Sub()
                    Try
                        GC.Collect()
                        GC.WaitForPendingFinalizers()
                    Catch
                    End Try
                End Sub)
                gcThread.IsBackground = True
                gcThread.Priority = Threading.ThreadPriority.BelowNormal
                gcThread.Start()
            Catch
            End Try

            sw.Stop()
            _chatLogger?.LogInfo("ADDIN", $"Add-in shutdown complete, elapsed: {sw.ElapsedMilliseconds.ToString()}ms")

        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "Exception during shutdown", ex)
        Finally
            shutdownMon.EndMonitoring()
        End Try
    End Sub

#End Region

End Class