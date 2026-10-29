Imports System.IO
Imports System.Linq
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json

Namespace Services

    ''' <summary>
    ''' OCR 批量处理器
    ''' 管理批量识别任务队列，支持并发控制、暂停/恢复/取消、进度报告
    ''' </summary>
    Public Class OcrBatchProcessor
        Implements IDisposable

        Private _taskQueue As New List(Of OcrBatchTask)()
        Private _syncLock As New Object()
        Private _cancellationTokenSource As CancellationTokenSource
        Private _isProcessing As Boolean = False
        Private _isPaused As Boolean = False
        Private _maxConcurrency As Integer = 2
        Private _ocrManager As OcrServiceManager
        Private _disposed As Boolean = False

        ''' <summary>当前处理状态</summary>
        Public ReadOnly Property IsProcessing As Boolean
            Get
                Return _isProcessing
            End Get
        End Property

        ''' <summary>当前是否暂停</summary>
        Public ReadOnly Property IsPaused As Boolean
            Get
                Return _isPaused
            End Get
        End Property

        ''' <summary>任务总数</summary>
        Public ReadOnly Property TotalTasks As Integer
            Get
                SyncLock _syncLock
                    Return _taskQueue.Count
                End SyncLock
            End Get
        End Property

        ''' <summary>已完成任务数</summary>
        Public ReadOnly Property CompletedCount As Integer
            Get
                SyncLock _syncLock
                    Return _taskQueue.Where(Function(t) t.Status = BatchTaskStatus.Completed).Count()
                End SyncLock
            End Get
        End Property

        ''' <summary>失败任务数</summary>
        Public ReadOnly Property FailedCount As Integer
            Get
                SyncLock _syncLock
                    Return _taskQueue.Where(Function(t) t.Status = BatchTaskStatus.Failed).Count()
                End SyncLock
            End Get
        End Property

        ''' <summary>最大并发数</summary>
        Public Property MaxConcurrency As Integer
            Get
                Return _maxConcurrency
            End Get
            Set(value As Integer)
                _maxConcurrency = Math.Max(1, Math.Min(value, 8))
            End Set
        End Property

        Public Sub New(Optional ocrManager As OcrServiceManager = Nothing)
            _ocrManager = If(ocrManager, OcrServiceManager.Instance)
            _cancellationTokenSource = New CancellationTokenSource()
        End Sub

        ''' <summary>
        ''' 从文件夹中扫描支持的图片文件
        ''' </summary>
        Public Function ScanFolder(folderPath As String, Optional recursive As Boolean = False) As List(Of String)
            Dim supportedExtensions As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
                ".png", ".jpg", ".jpeg", ".bmp", ".tiff", ".tif", ".gif"
            }

            Dim files As New List(Of String)()
            Try
                Dim searchOpt As SearchOption = If(recursive, SearchOption.AllDirectories, SearchOption.TopDirectoryOnly)
                For Each ext In supportedExtensions
                    files.AddRange(Directory.GetFiles(folderPath, "*" & ext, searchOpt))
                Next
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("扫描文件夹失败: " & ex.Message)
            End Try

            Return files
        End Function

        ''' <summary>
        ''' 添加文件到任务队列
        ''' </summary>
        Public Sub AddFiles(filePaths As IEnumerable(Of String))
            SyncLock _syncLock
                For Each fp In filePaths
                    If File.Exists(fp) Then
                        _taskQueue.Add(New OcrBatchTask() With {
                            .TaskId = Guid.NewGuid().ToString("N"),
                            .FileName = Path.GetFileName(fp),
                            .FilePath = fp,
                            .Status = BatchTaskStatus.Pending
                        })
                    End If
                Next
            End SyncLock
        End Sub

        ''' <summary>
        ''' 添加单个文件
        ''' </summary>
        Public Sub AddFile(filePath As String)
            If File.Exists(filePath) Then
                AddFiles({filePath})
            End If
        End Sub

        ''' <summary>
        ''' 获取当前任务列表（副本）
        ''' </summary>
        Public Function GetTasks() As List(Of OcrBatchTask)
            SyncLock _syncLock
                Return New List(Of OcrBatchTask)(_taskQueue)
            End SyncLock
        End Function

        ''' <summary>
        ''' 获取进度报告
        ''' </summary>
        Public Function GetProgress() As OcrBatchProgress
            SyncLock _syncLock
                Dim completed = _taskQueue.Where(Function(t) t.Status = BatchTaskStatus.Completed).Count()
                Dim failed = _taskQueue.Where(Function(t) t.Status = BatchTaskStatus.Failed).Count()
                Dim processing = _taskQueue.Where(Function(t) t.Status = BatchTaskStatus.Processing).Count()
                Dim currentFile = _taskQueue.AsEnumerable().FirstOrDefault(Function(t) t.Status = BatchTaskStatus.Processing)

                Return New OcrBatchProgress() With {
                    .TotalTasks = _taskQueue.Count,
                    .CompletedTasks = completed,
                    .FailedTasks = failed,
                    .ProcessingTasks = processing,
                    .CurrentFileName = If(currentFile IsNot Nothing, currentFile.FileName, ""),
                    .StatusMessage = $"已完成 {completed + failed}/{_taskQueue.Count} 个"
                }
            End SyncLock
        End Function

        ''' <summary>
        ''' 启动批量处理
        ''' </summary>
        Public Async Function StartAsync(Optional progressCallback As Action(Of OcrBatchProgress) = Nothing) As Task(Of List(Of OcrBatchTask))
            If _isProcessing Then
                System.Diagnostics.Debug.WriteLine("批量处理已在进行中")
                Return GetTasks()
            End If

            _isProcessing = True
            _isPaused = False
            _cancellationTokenSource = New CancellationTokenSource()
            Dim token = _cancellationTokenSource.Token

            Try
                ' 获取待处理任务
                Dim pendingTasks As List(Of OcrBatchTask) = Nothing
                SyncLock _syncLock
                    pendingTasks = _taskQueue.Where(Function(t) t.Status = BatchTaskStatus.Pending).ToList()
                End SyncLock

                If pendingTasks.Count = 0 Then
                    System.Diagnostics.Debug.WriteLine("没有待处理的任务")
                    Return GetTasks()
                End If

                ' 使用 SemaphoreSlim 控制并发
                Using semaphore As New SemaphoreSlim(_maxConcurrency)
                    Dim tasks = pendingTasks.Select(Function(t)
                        Return Task.Run(Async Function()
                            Await semaphore.WaitAsync(token)

                            Try
                                Await ProcessSingleTaskAsync(t, token, progressCallback)
                            Catch ex As OperationCanceledException
                                SyncLock _syncLock
                                    t.Status = BatchTaskStatus.Cancelled
                                End SyncLock
                            Catch ex As Exception
                                SyncLock _syncLock
                                    t.Status = BatchTaskStatus.Failed
                                    t.ErrorMessage = ex.Message
                                    t.CompletedAt = DateTime.Now
                                End SyncLock
                            Finally
                                semaphore.Release()
                            End Try
                        End Function, token)
                    End Function).ToArray()

                    Await Task.WhenAll(tasks)
                End Using

            Catch ex As OperationCanceledException
                System.Diagnostics.Debug.WriteLine("批量处理被取消")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("批量处理异常: " & ex.Message)
            Finally
                _isProcessing = False
                _isPaused = False
            End Try

            Return GetTasks()
        End Function

        ''' <summary>
        ''' 处理单个任务
        ''' </summary>
        Private Async Function ProcessSingleTaskAsync(currentTask As OcrBatchTask, token As CancellationToken, progressCallback As Action(Of OcrBatchProgress)) As Task
            token.ThrowIfCancellationRequested()

            ' 暂停检查
            While _isPaused
                Await Task.Delay(200, token)
                token.ThrowIfCancellationRequested()
            End While

            SyncLock _syncLock
                currentTask.Status = BatchTaskStatus.Processing
                currentTask.StartedAt = DateTime.Now
            End SyncLock

            ' 报告进度
            progressCallback?.Invoke(GetProgress())

            Try
                Dim sw As New System.Diagnostics.Stopwatch()
                sw.Start()

                Dim doc = Await _ocrManager.RecognizeDocumentAsync(currentTask.FilePath, currentTask.FileName)

                sw.Stop()

                token.ThrowIfCancellationRequested()

                SyncLock _syncLock
                    currentTask.Document = doc
                    currentTask.Status = BatchTaskStatus.Completed
                    currentTask.CompletedAt = DateTime.Now
                End SyncLock

            Catch ex As Exception When Not TypeOf ex Is OperationCanceledException
                SyncLock _syncLock
                    currentTask.Status = BatchTaskStatus.Failed
                    currentTask.ErrorMessage = ex.Message
                    currentTask.CompletedAt = DateTime.Now
                End SyncLock
            End Try

            ' 报告进度
            progressCallback?.Invoke(GetProgress())
        End Function

        ''' <summary>
        ''' 暂停处理
        ''' </summary>
        Public Sub Pause()
            _isPaused = True
        End Sub

        ''' <summary>
        ''' 恢复处理
        ''' </summary>
        Public Sub ResumeProcessing()
            _isPaused = False
        End Sub

        ''' <summary>
        ''' 取消所有处理
        ''' </summary>
        Public Sub Cancel()
            Try
                _cancellationTokenSource?.Cancel()
            Catch ex As Exception  ' 
            End Try

            SyncLock _syncLock
                For Each task In _taskQueue.Where(Function(t) t.Status = BatchTaskStatus.Pending)
                    task.Status = BatchTaskStatus.Cancelled
                Next
            End SyncLock
        End Sub

        ''' <summary>
        ''' 清空任务队列
        ''' </summary>
        Public Sub ClearTasks()
            SyncLock _syncLock
                _taskQueue.Clear()
            End SyncLock
        End Sub

        ''' <summary>
        ''' 将批量结果序列化为 JSON（供前端使用）
        ''' </summary>
        Public Function SerializeResultsToJson() As String
            SyncLock _syncLock
                Dim results = _taskQueue.Where(Function(t) t.Document IsNot Nothing).Select(Function(t)
                    Return New With {
                        Key .file = t.FileName,
                        Key .status = t.Status.ToString(),
                        Key .text = If(t.Document IsNot Nothing, t.Document.RawText, ""),
                        Key .error = t.ErrorMessage,
                        Key .elapsed = If(t.Document IsNot Nothing, t.Document.ElapsedSeconds, 0.0)
                    }
                End Function).ToList()

                Return JsonConvert.SerializeObject(New With {
                    Key .total = _taskQueue.Count,
                    Key .completed = CompletedCount,
                    Key .failed = FailedCount,
                    Key .results = results
                })
            End SyncLock
        End Function

#Region "IDisposable"

        Public Sub Dispose() Implements IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            Cancel()
            ClearTasks()
        End Sub

#End Region

    End Class

End Namespace
