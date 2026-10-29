Imports System.Collections.Generic
Imports System.IO
Imports Newtonsoft.Json

Namespace Context
    ''' <summary>
    ''' 操作记录
    ''' </summary>
    Public Class ActionRecord
        Public Property ActionType As String
        Public Property Result As Object
        Public Property Timestamp As DateTime
    End Class
    
    ''' <summary>
    ''' 用户偏好设置
    ''' </summary>
    Public Class UserPreferences
        Public Property DefaultChartType As String = "柱状图"
        Public Property PreferredLanguage As String = "zh-CN"
        Public Property AutoSave As Boolean = True
    End Class
    
    ''' <summary>
    ''' 列映射记录（用于记忆字段映射）
    ''' </summary>
    Public Class ColumnMapping
        Public Property Headers As List(Of String)
        Public Property Mapping As Dictionary(Of String, String)
        Public Property UsageCount As Integer
        Public Property LastUsed As DateTime
    End Class
    
    ''' <summary>
    ''' 短期记忆
    ''' </summary>
    Public Class ShortTermMemory
        Public Property CurrentSelection As String
        Public Property ActiveSheet As String
        Public Property ActiveWorkbook As String
        Public Property RecentActions As List(Of ActionRecord)
        Public Property LastAction As ActionRecord
        
        Public Sub New()
            RecentActions = New List(Of ActionRecord)()
        End Sub
    End Class
    
    ''' <summary>
    ''' 长期记忆
    ''' </summary>
    Public Class LongTermMemory
        Public Property UserPreferences As UserPreferences
        Public Property SkillUsage As Dictionary(Of String, Integer)
        Public Property ColumnMappings As List(Of ColumnMapping)
        
        Public Sub New()
            UserPreferences = New UserPreferences()
            SkillUsage = New Dictionary(Of String, Integer)()
            ColumnMappings = New List(Of ColumnMapping)()
        End Sub
    End Class
    
    ''' <summary>
    ''' 上下文信息
    ''' </summary>
    Public Class ContextInfo
        Public Property ShortTerm As ShortTermMemory
        Public Property LongTerm As LongTermMemory
    End Class
    
    ''' <summary>
    ''' 上下文存储管理类（单例模式）
    ''' </summary>
    Public Class ContextStore
        Private Shared _instance As ContextStore
        Private Shared ReadOnly _lock As New Object()
        Private _shortTermMemory As ShortTermMemory
        Private _longTermMemory As LongTermMemory
        Private _memoryFilePath As String

        Public Shared Function GetInstance() As ContextStore
            If _instance Is Nothing Then
                SyncLock _lock
                    If _instance Is Nothing Then
                        _instance = New ContextStore()
                        System.Threading.Thread.MemoryBarrier()
                    End If
                End SyncLock
            End If
            Return _instance
        End Function
        
        Private Sub New()
            _shortTermMemory = New ShortTermMemory()
            _memoryFilePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ExcelAIVSTOAddIn", "Memory.json"
            )
            LoadLongTermMemory()
        End Sub
        
        ''' <summary>
        ''' 获取完整上下文
        ''' </summary>
        Public Function GetContext() As ContextInfo
            Return New ContextInfo With {
                .ShortTerm = _shortTermMemory,
                .LongTerm = _longTermMemory
            }
        End Function
        
        ''' <summary>
        ''' 记录操作到历史
        ''' </summary>
        Public Sub RecordAction(actionType As String, result As Object)
            Dim record As New ActionRecord With {
                .ActionType = actionType,
                .Result = result,
                .Timestamp = DateTime.Now
            }
            
            _shortTermMemory.RecentActions.Add(record)
            _shortTermMemory.LastAction = record
            
            ' 只保留最近20条
            If _shortTermMemory.RecentActions.Count > 20 Then
                _shortTermMemory.RecentActions.RemoveAt(0)
            End If
            
            ' 更新Skill使用统计
            If _longTermMemory.SkillUsage.ContainsKey(actionType) Then
                _longTermMemory.SkillUsage(actionType) += 1
            Else
                _longTermMemory.SkillUsage(actionType) = 1
            End If
            
            SaveLongTermMemory()
        End Sub
        
        ''' <summary>
        ''' 更新当前选中区域信息
        ''' </summary>
        Public Sub UpdateCurrentSelection(range As String, sheetName As String)
            Try
                _shortTermMemory.CurrentSelection = range
                _shortTermMemory.ActiveSheet = sheetName
                
                ' 获取工作簿名
                Dim excelApp = Globals.ThisAddIn.GetExcelApp()
                If excelApp.ActiveWorkbook IsNot Nothing Then
                    _shortTermMemory.ActiveWorkbook = excelApp.ActiveWorkbook.Name
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"记录选择变化失败: {ex.Message}")
            End Try
        End Sub
        
        ''' <summary>
        ''' 加载长期记忆
        ''' </summary>
        Private Sub LoadLongTermMemory()
            Try
                If File.Exists(_memoryFilePath) Then
                    Dim json = File.ReadAllText(_memoryFilePath)
                    _longTermMemory = JsonConvert.DeserializeObject(Of LongTermMemory)(json)
                Else
                    _longTermMemory = New LongTermMemory()
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"加载长期记忆失败: {ex.Message}")
                _longTermMemory = New LongTermMemory()
            End Try
        End Sub
        
        ''' <summary>
        ''' 保存长期记忆
        ''' </summary>
        Public Sub SaveLongTermMemory()
            Try
                Dim dir = Path.GetDirectoryName(_memoryFilePath)
                If Not Directory.Exists(dir) Then
                    Directory.CreateDirectory(dir)
                End If
                
                Dim json = JsonConvert.SerializeObject(_longTermMemory, Formatting.Indented)
                File.WriteAllText(_memoryFilePath, json)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"保存长期记忆失败: {ex.Message}")
            End Try
        End Sub
        
        ''' <summary>
        ''' 保存列映射模板
        ''' </summary>
        Public Sub SaveColumnMapping(headers As List(Of String), mapping As Dictionary(Of String, String))
            Try
                ' 查找是否已存在相同的映射
                Dim templateKey = String.Join("|", headers.OrderBy(Function(h) h))
                Dim existing = _longTermMemory.ColumnMappings.FirstOrDefault(Function(cm)
                    Dim existingHeaders = If(cm.Headers Is Nothing, New List(Of String)(), cm.Headers)
                    Dim existingKey = String.Join("|", existingHeaders.OrderBy(Function(h) h))
                    Return existingKey = templateKey
                End Function)
                
                If existing IsNot Nothing Then
                    existing.Mapping = mapping
                    existing.UsageCount += 1
                    existing.LastUsed = DateTime.Now
                Else
                    _longTermMemory.ColumnMappings.Add(New ColumnMapping With {
                        .Headers = headers,
                        .Mapping = mapping,
                        .UsageCount = 1,
                        .LastUsed = DateTime.Now
                    })
                End If
                
                SaveLongTermMemory()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"保存列映射失败: {ex.Message}")
            End Try
        End Sub
        
        ''' <summary>
        ''' 查找匹配的列映射
        ''' </summary>
        Public Function FindColumnMapping(headers As List(Of String)) As ColumnMapping
            Try
                ' 精确匹配
                Dim templateKey = String.Join("|", headers.OrderBy(Function(h) h))
                Dim exactMatch = _longTermMemory.ColumnMappings.FirstOrDefault(Function(cm)
                    Dim existingHeaders = If(cm.Headers Is Nothing, New List(Of String)(), cm.Headers)
                    Dim existingKey = String.Join("|", existingHeaders.OrderBy(Function(h) h))
                    Return existingKey = templateKey
                End Function)
                
                If exactMatch IsNot Nothing Then
                    Return exactMatch
                End If
                
                ' 模糊匹配：找使用最多的
                If _longTermMemory.ColumnMappings.Count > 0 Then
                    Return _longTermMemory.ColumnMappings.OrderByDescending(Function(cm) cm.UsageCount).First()
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"查找列映射失败: {ex.Message}")
            End Try
            
            Return Nothing
        End Function
    End Class
End Namespace
