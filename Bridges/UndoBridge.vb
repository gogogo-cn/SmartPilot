Option Explicit On
Option Infer On

Imports System.Runtime.InteropServices
Imports System.Diagnostics
Imports Newtonsoft.Json
Imports ExcelAIVSTOAddIn.Utils
Imports Microsoft.Office.Interop.Excel

Namespace Bridges

    ''' <summary>
    ''' UndoBridge - 撤销功能管理桥梁类
    ''' 提供 COM 可见的方法供前端 JavaScript 调用，用于撤销操作管理
    ''' </summary>
    <ComVisible(True)>
    <ClassInterface(ClassInterfaceType.AutoDispatch)>
    <ProgId("ExcelAIVSTOAddIn.UndoBridge")>
    Public Class UndoBridge
        Implements IBridge, IUndoBridge

        ' 单例实例 - 使用 Lazy(Of T) 确保线程安全
        Private Shared ReadOnly _lazyInstance As New Lazy(Of UndoBridge)(Function() New UndoBridge(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        ' 撤销栈相关
        Private _undoStack As New List(Of UndoAction)()
        Private _redoStack As New List(Of UndoAction)()
        Private _undoStackLock As New Object()
        Private Const MAX_UNDO_STACK_SIZE As Integer = 50
        ' 防止撤销/重做过程中因 Excel 事件回调（如 SheetChange）导致同线程重入递归/死锁
        Private _isProcessing As Boolean = False

#Region "单例模式"

        ''' <summary>
        ''' 获取 UndoBridge 单例实例（线程安全 Lazy 初始化）
        ''' </summary>
        Public Shared ReadOnly Property Instance As UndoBridge
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        ''' <summary>
        ''' 私有构造函数（单例模式）
        ''' </summary>
        Private Sub New()
            ' 注册到 BridgeRegistry：MCP 工具层（ExcelMcpTools/ActionEngine/ExcelDataBridge 等）
            ' 均通过 Resolve(Of IUndoBridge)() 获取实例做撤销登记；若不注册，Resolve 返回 Nothing，
            ' 所有 AddUndoAction 会抛 NRE 并被调用方的 Try/Catch 静默吞掉 → 撤销栈永远为空。
            Try
                BridgeRegistry.Instance.Register(Me)
            Catch ex As Exception
                Debug.WriteLine($"[UndoBridge] 注册到 BridgeRegistry 失败: {ex.Message}")
            End Try
        End Sub

#End Region

#Region "撤销功能"

        ''' <summary>
        ''' 获取撤销栈信息
        ''' </summary>
        <ComVisible(True)>
        Public Function GetUndoStack() As String Implements IUndoBridge.GetUndoStack
            Try
                SyncLock _undoStackLock
                    Return JsonConvert.SerializeObject(New With {
                        .count = _undoStack.Count,
                        .actions = _undoStack.Select(Function(a) New With {
                            .actionType = a.ActionType,
                            .address = a.Address,
                            .sheetName = a.SheetName,
                            .timestamp = a.Timestamp
                        }).ToList()
                    })
                End SyncLock
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.count = 0, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 执行撤销操作
        ''' </summary>
        <ComVisible(True)>
        Public Function UndoLastAction() As String Implements IUndoBridge.UndoLastAction
            Try
                SyncLock _undoStackLock
                    If _isProcessing Then
                        Return JsonConvert.SerializeObject(New With {
                            .success = False,
                            .message = "撤销操作正在进行中，请稍候"
                        })
                    End If
                    _isProcessing = True
                    Try
                        If _undoStack.Count = 0 Then
                            Return JsonConvert.SerializeObject(New With {
                                .success = False,
                                .message = "没有可撤销的操作"
                            })
                        End If

                    Dim action = _undoStack(_undoStack.Count - 1)
                    _undoStack.RemoveAt(_undoStack.Count - 1)

                    ' 将撤销的动作推入重做栈
                    _redoStack.Add(action)
                    While _redoStack.Count > MAX_UNDO_STACK_SIZE
                        _redoStack.RemoveAt(0)
                    End While

                    Select Case action.ActionType.ToLower()
                        ' ---- 单元格/格式/行列增删：统一走"整表备份恢复"，不存在局部差异 ----
                        Case "sheet_backup", "setcell", "fill_data", "format", "clear_format",
                             "vba_execute", "delete_rows", "insert_rows", "delete_columns", "insert_columns"
                            Return UndoSheetBackup(action)

                        ' ---- 工作表结构类：保留原有专用还原 ----
                        Case "create_sheet"
                            Return UndoCreateSheet(action)
                        Case "merge_sheets"
                            Return UndoMergeSheets(action)
                        Case "split_workbook"
                            Return UndoSplitWorkbook(action)
                        Case "rename_sheet"
                            Return UndoRenameSheet(action)
                        Case "delete_sheet"
                            Return UndoDeleteSheet(action)
                        Case Else
                            Return JsonConvert.SerializeObject(New With {
                                .success = False,
                                .message = $"不支持的撤销类型: {action.ActionType}"
                            })
                    End Select
                    Finally
                        _isProcessing = False
                    End Try
                End SyncLock
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .message = $"撤销操作异常: {ex.Message}"
                })
            End Try
        End Function

        ''' <summary>
        ''' 按名称查找工作表（优化：直接索引而非遍历）
        ''' </summary>
        Private Function FindWorksheetByName(sheetName As String) As Worksheet
            Dim excelApp As Application = Globals.ThisAddIn.Application
            Dim wb As Workbook = excelApp.ActiveWorkbook
            If wb Is Nothing Then Return Nothing
            Try
                For Each ws As Worksheet In wb.Worksheets
                    If ws.Name = sheetName Then
                        Return ws
                    End If
                Next
            Catch ex As Exception  ' 
            End Try
            Return Nothing
        End Function


        ''' <summary>
        ''' 撤销创建工作表：删除该工作表
        ''' </summary>
        Private Function UndoCreateSheet(action As UndoAction) As String
            Dim wb As Workbook = Nothing
            Dim worksheets As Sheets = Nothing
            Try
                Dim excelApp As Application = Globals.ThisAddIn.Application
                wb = excelApp.ActiveWorkbook
                If wb Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .message = "当前没有打开的工作簿"})
                End If

                Dim sheetName = action.SheetName
                worksheets = wb.Worksheets
                Dim sheetCount As Integer = worksheets.Count

                For i As Integer = 1 To sheetCount
                    Dim ws As Worksheet = Nothing
                    Try
                        ws = CType(worksheets(i), Worksheet)
                        If ws.Name = sheetName Then
                            ' Excel 至少保留一个可见工作表
                            If sheetCount <= 1 Then
                                Return JsonConvert.SerializeObject(New With {.success = False, .message = "无法删除唯一的工作表"})
                            End If
                            ws.Delete()
                            Return JsonConvert.SerializeObject(New With {
                                .success = True,
                                .message = $"已撤销创建工作表: {sheetName}",
                                .actionType = action.ActionType,
                                .sheetName = sheetName
                            })
                        End If
                    Finally
                        Utils.ComReleaseHelper.Release(ws)
                    End Try
                Next

                Return JsonConvert.SerializeObject(New With {.success = False, .message = $"找不到工作表: {sheetName}"})
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .message = $"撤销失败: {ex.Message}"})
            Finally
                Utils.ComReleaseHelper.Release(worksheets)
                Utils.ComReleaseHelper.Release(wb)
            End Try
        End Function


        ''' <summary>
        ''' 撤销合并工作表：删除目标工作表
        ''' </summary>
        Private Function UndoMergeSheets(action As UndoAction) As String
            Return UndoCreateSheet(action)
        End Function

        ''' <summary>
        ''' 撤销拆分工作簿：删除生成的文件
        ''' Payload 存储为 JSON：{ "files": ["path1", "path2", ...] }
        ''' </summary>
        Private Function UndoSplitWorkbook(action As UndoAction) As String
            Try
                If String.IsNullOrWhiteSpace(action.Payload) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .message = "缺少撤销数据"})
                End If

                Dim payload = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(action.Payload)
                Dim files = JsonConvert.DeserializeObject(Of List(Of String))(payload("files").ToString())
                Dim deletedCount = 0
                For Each file In files
                    If System.IO.File.Exists(file) Then
                        System.IO.File.Delete(file)
                        deletedCount += 1
                    End If
                Next

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .message = $"已撤销拆分工作簿，删除 {deletedCount} 个文件",
                    .actionType = action.ActionType,
                    .deletedCount = deletedCount
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .message = $"撤销失败: {ex.Message}"})
            End Try
        End Function

#Region "行/列插入删除撤销（内容操作）"

        ''' <summary>
        ''' 解析行/列增删撤销动作的 Payload：{ "index": 行号/列号, "count": 数量, "data": 被删内容(可选) }
        ''' </summary>
        Private Function ParseRowColPayload(action As UndoAction, ByRef index As Integer, ByRef count As Integer, ByRef data As List(Of List(Of Object))) As Boolean
            If String.IsNullOrWhiteSpace(action.Payload) Then Return False
            Try
                Dim payload = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(action.Payload)
                If Not payload.ContainsKey("index") OrElse Not payload.ContainsKey("count") Then Return False
                index = Convert.ToInt32(payload("index"))
                count = Convert.ToInt32(payload("count"))
                If payload.ContainsKey("data") AndAlso payload("data") IsNot Nothing Then
                    data = JsonConvert.DeserializeObject(Of List(Of List(Of Object)))(payload("data").ToString())
                End If
                Return True
            Catch
                Return False
            End Try
        End Function









#End Region

#Region "格式与工作表撤销（内容操作）"




        ' 以下公共方法供 MCP / ActionEngine 在写操作前调用，生成撤销快照
        <ComVisible(True)>
        Public Function CaptureSheetSnapshot(ws As Worksheet) As String
            Dim used As Range = Nothing
            Try
                used = ws.UsedRange
                Dim addr = used.Address(False, False)
                Dim vals As Object()() = Nothing
                Dim v = used.Value
                If v IsNot Nothing Then
                    Dim arr = DirectCast(v, Object(,))
                    Dim lb1 = arr.GetLowerBound(0), ub1 = arr.GetUpperBound(0)
                    Dim lb2 = arr.GetLowerBound(1), ub2 = arr.GetUpperBound(1)
                    Dim rc = ub1 - lb1 + 1, cc = ub2 - lb2 + 1
                    vals = New Object(rc - 1)() {}
                    For r = 0 To rc - 1
                        Dim rowArr(cc - 1) As Object
                        For c = 0 To cc - 1
                            rowArr(c) = arr(lb1 + r, lb2 + c)
                        Next
                        vals(r) = rowArr
                    Next
                End If
                Dim nfObj As Object = Nothing
                Try
                    nfObj = used.NumberFormat
                Catch
                End Try
                Dim nf As Object = If(TypeOf nfObj Is String, nfObj, "")
                If Not (TypeOf nfObj Is String) AndAlso nfObj IsNot Nothing Then
                    Try
                        Dim arr = DirectCast(nfObj, Object(,))
                        Dim lb1 = arr.GetLowerBound(0), ub1 = arr.GetUpperBound(0)
                        Dim lb2 = arr.GetLowerBound(1), ub2 = arr.GetUpperBound(1)
                        Dim rc = ub1 - lb1 + 1, cc = ub2 - lb2 + 1
                        Dim nfArr(rc - 1)() As Object
                        For r = 0 To rc - 1
                            Dim rowArr(cc - 1) As Object
                            For c = 0 To cc - 1
                                rowArr(c) = arr(lb1 + r, lb2 + c)
                            Next
                            nfArr(r) = rowArr
                        Next
                        nf = nfArr
                    Catch
                        nf = ""
                    End Try
                End If
                Dim colWidths As New List(Of Object)()
                Try
                    For c = 1 To used.Columns.Count
                        colWidths.Add(DirectCast(used.Columns(c), Range).ColumnWidth)
                    Next
                Catch
                End Try
                Return JsonConvert.SerializeObject(New With {
                    .index = ws.Index, .name = ws.Name, .addr = addr,
                    .values = vals, .nf = nf, .colWidths = colWidths
                })
            Catch
                Return "{}"
            Finally
                If used IsNot Nothing Then Utils.ComReleaseHelper.Release(used)
            End Try
        End Function

#Region "VBA 执行撤销（整表快照）"

        Private Const XlNoneIdx As Integer = -4142
        Private Const XlHAlignGeneralIdx As Integer = -4109
        Private Const XlVAlignBottomIdx As Integer = -4107
        Private Const MAX_FULL_SNAPSHOT_CELLS As Integer = 100000

        ''' <summary>最近一次 CaptureSheetFullSnapshot 失败原因（用于前端诊断提示）</summary>
        Public Shared Property LastSnapshotError As String = ""

        ''' <summary>锯齿数组安全取值（数组或元素可能为 Nothing）</summary>
        Private Shared Function J(v As Object()(), r As Integer, c As Integer) As Object
            If v Is Nothing Then Return Nothing
            Dim row As Object() = v(r)
            If row Is Nothing Then Return Nothing
            Return row(c)
        End Function

        Private Shared Function FmtBool(v As Object()(), r As Integer, c As Integer) As Boolean
            Dim o = J(v, r, c)
            If o Is Nothing OrElse TypeOf o Is DBNull Then Return False
            Return CBool(o)
        End Function

        Private Shared Function FmtDbl(v As Object()(), r As Integer, c As Integer) As Double
            Dim o = J(v, r, c)
            If o Is Nothing OrElse TypeOf o Is DBNull Then Return 0
            Return Convert.ToDouble(o)
        End Function

        Private Shared Function FmtInt(v As Object()(), r As Integer, c As Integer, defVal As Integer) As Integer
            Dim o = J(v, r, c)
            If o Is Nothing OrElse TypeOf o Is DBNull Then Return defVal
            Return Convert.ToInt32(o)
        End Function

        ''' <summary>
        ''' 把 COM 属性值（标量 / 二维数组 / Nothing）统一转为锯齿数组
        ''' </summary>
        Private Function PropToJagged(o As Object) As Object()()
            If o Is Nothing OrElse TypeOf o Is DBNull Then Return Nothing
            If TypeOf o Is Object(,) Then
                Dim arr = DirectCast(o, Object(,))
                Dim lb1 = arr.GetLowerBound(0), ub1 = arr.GetUpperBound(0)
                Dim lb2 = arr.GetLowerBound(1), ub2 = arr.GetUpperBound(1)
                Dim rc = ub1 - lb1 + 1, cc = ub2 - lb2 + 1
                Dim res(rc - 1)() As Object
                For r = 0 To rc - 1
                    Dim row(cc - 1) As Object
                    For c = 0 To cc - 1
                        row(c) = arr(lb1 + r, lb2 + c)
                    Next
                    res(r) = row
                Next
                Return res
            End If
            Dim singleArr(0)() As Object
            Dim singleRow(0) As Object
            singleRow(0) = o
            singleArr(0) = singleRow
            Return singleArr
        End Function

        ''' <summary>
        ''' 捕获工作表整表快照（值 + 每格格式 + 数字格式 + 列宽），供 VBA 执行（如表格美化）撤销用
        ''' </summary>
        <ComVisible(True)>
        Public Function CaptureSheetFullSnapshot(ws As Worksheet) As String
            Dim used As Range = Nothing
            Dim fnt As Font = Nothing
            Dim intr As Interior = Nothing
            Dim brd As Borders = Nothing
            Try
                used = ws.UsedRange
                Dim addr = used.Address(False, False)
                Dim rc = used.Rows.Count
                Dim cc = used.Columns.Count
                Dim cellCount = rc * cc

                ' 用 FormulaLocal 而非 Value：保留公式（否则撤销会把公式变成死值，表就"乱"了）
                Dim vals = PropToJagged(used.FormulaLocal)

                Dim colWidths As New List(Of Object)()
                For c = 1 To cc
                    Dim colRng = DirectCast(used.Columns(c), Range)
                    colWidths.Add(colRng.ColumnWidth)
                    Utils.ComReleaseHelper.Release(colRng)
                Next

                Dim payload As New Dictionary(Of String, Object)
                payload("addr") = addr
                payload("rows") = rc
                payload("cols") = cc
                payload("values") = vals
                payload("colWidths") = colWidths

                ' 数字格式：先尝试整体读取，混合格式（返回 Null）时按列读取
                Dim nfAll As Object = Nothing
                Try
                    nfAll = used.NumberFormat
                Catch
                    nfAll = Nothing
                End Try
                If TypeOf nfAll Is DBNull Then nfAll = Nothing
                If nfAll IsNot Nothing Then
                    payload("nfAll") = nfAll
                Else
                    Dim nfCols As New List(Of Object)()
                    For c = 1 To cc
                        Dim colRng = DirectCast(used.Columns(c), Range)
                        Dim colNf As Object = Nothing
                        Try
                            colNf = colRng.NumberFormat
                        Catch
                        End Try
                        If TypeOf colNf Is DBNull Then colNf = Nothing
                        nfCols.Add(colNf)
                        Utils.ComReleaseHelper.Release(colRng)
                    Next
                    payload("nfCols") = nfCols
                End If

                ' 合并单元格：本期未捕获（Range.Cells 两参写法在本 PIA 下为后期绑定，需改用 A1 地址转换，留待后续），撤销后合并状态不保证还原

                ' 每格格式（仅记录非默认项；超大表降级为只恢复值与列宽）
                If cellCount <= MAX_FULL_SNAPSHOT_CELLS Then
                    fnt = used.Font
                    intr = used.Interior
                    brd = used.Borders
                    Try
                        ' 整块读属性：若区域为混合格式，Excel 返回标量 DBNull（不是逐格数组），
                        ' 此时 PropToJagged 得 Nothing —— 必须改用逐格读取，否则真实格式会被当默认丢弃
                        Dim fBold = PropToJagged(fnt.Bold)
                        Dim fItalic = PropToJagged(fnt.Italic)
                        Dim fColor = PropToJagged(fnt.Color)
                        Dim fName = PropToJagged(fnt.Name)
                        Dim fSize = PropToJagged(fnt.Size)
                        Dim iIdx = PropToJagged(intr.ColorIndex)
                        Dim iColor = PropToJagged(intr.Color)
                        Dim hAlign = PropToJagged(used.HorizontalAlignment)
                        Dim vAlign = PropToJagged(used.VerticalAlignment)
                        Dim bStyle = PropToJagged(brd.LineStyle)

                        ' 混合格式需逐格读（避免把真实格式当成默认丢失）；超大表逐格读过慢则降级跳过
                        Dim usePerCell As Boolean = (fBold Is Nothing OrElse fItalic Is Nothing OrElse fColor Is Nothing _
                            OrElse fName Is Nothing OrElse fSize Is Nothing OrElse iIdx Is Nothing OrElse iColor Is Nothing _
                            OrElse hAlign Is Nothing OrElse vAlign Is Nothing OrElse bStyle Is Nothing) _
                            AndAlso cellCount <= 20000

                        If usePerCell OrElse fBold IsNot Nothing Then
                            Dim fmt(rc - 1)() As Object
                            For r = 0 To rc - 1
                                Dim rowFmt(cc - 1) As Object
                                For c = 0 To cc - 1
                                    Dim cellRng As Range = Nothing
                                    If usePerCell Then cellRng = DirectCast(used.Cells(r + 1, c + 1), Range)
                                    Dim o As New Dictionary(Of String, Object)
                                    Dim hasFmt = False
                                    Dim bb As Boolean = If(usePerCell, CBool(cellRng.Font.Bold), FmtBool(fBold, r, c))
                                    If bb Then o("b") = True : hasFmt = True
                                    Dim it As Boolean = If(usePerCell, CBool(cellRng.Font.Italic), FmtBool(fItalic, r, c))
                                    If it Then o("i") = True : hasFmt = True
                                    Dim fc As Double = If(usePerCell, CDbl(cellRng.Font.Color), FmtDbl(fColor, r, c))
                                    If fc <> 0 Then o("fc") = fc : hasFmt = True
                                    Dim iIdxVal As Integer = If(usePerCell, Convert.ToInt32(cellRng.Interior.ColorIndex), FmtInt(iIdx, r, c, XlNoneIdx))
                                    If iIdxVal <> XlNoneIdx Then
                                        o("icc") = If(usePerCell, CDbl(cellRng.Interior.Color), FmtDbl(iColor, r, c))
                                        hasFmt = True
                                    End If
                                    Dim ha As Integer = If(usePerCell, Convert.ToInt32(cellRng.HorizontalAlignment), FmtInt(hAlign, r, c, XlHAlignGeneralIdx))
                                    If ha <> XlHAlignGeneralIdx Then o("ha") = ha : hasFmt = True
                                    Dim va As Integer = If(usePerCell, Convert.ToInt32(cellRng.VerticalAlignment), FmtInt(vAlign, r, c, XlVAlignBottomIdx))
                                    If va <> XlVAlignBottomIdx Then o("va") = va : hasFmt = True
                                    Dim bs As Integer = If(usePerCell, Convert.ToInt32(cellRng.Borders.LineStyle), FmtInt(bStyle, r, c, XlNoneIdx))
                                    If bs <> XlNoneIdx Then o("bs") = bs : hasFmt = True
                                    Dim fn As String = If(usePerCell, If(Convert.ToString(cellRng.Font.Name), ""), If(Convert.ToString(J(fName, r, c)), ""))
                                    If fn <> "宋体" AndAlso fn <> "等线" Then o("fn") = fn : hasFmt = True
                                    Dim fs As Double = If(usePerCell, CDbl(cellRng.Font.Size), FmtDbl(fSize, r, c))
                                    If fs <> 11 Then o("fs") = fs : hasFmt = True
                                    If usePerCell Then Utils.ComReleaseHelper.Release(cellRng)
                                    rowFmt(c) = If(hasFmt, CObj(o), Nothing)
                                Next
                                fmt(r) = rowFmt
                            Next
                            payload("fmt") = fmt
                        End If
                    Catch fmtEx As Exception
                        ' 格式捕获失败时降级：仍保留值/数字格式/列宽，撤销可还原数据本体
                        Debug.WriteLine($"[CaptureSheetFullSnapshot] 格式捕获降级: {fmtEx.Message}")
                    End Try
                End If

                ' 合并单元格：整表还原必须保留，否则撤销后合并状态丢失
                Dim merged As New List(Of String)()
                Try
                    CollectMergedAreas(ws, used, merged)
                Catch
                End Try
                payload("merged") = merged

                ' 行高
                Dim rowHeights As New List(Of Object)()
                Try
                    For r = 1 To rc
                        Dim rw = DirectCast(used.Rows(r), Range)
                        rowHeights.Add(CDbl(rw.RowHeight))
                        Utils.ComReleaseHelper.Release(rw)
                    Next
                Catch
                End Try
                payload("rowHeights") = rowHeights

                Return JsonConvert.SerializeObject(payload)
            Catch ex As Exception
                LastSnapshotError = ex.Message
                Debug.WriteLine($"[CaptureSheetFullSnapshot] 快照失败: {ex.Message}")
                Return "{}"
            Finally
                If fnt IsNot Nothing Then Utils.ComReleaseHelper.Release(fnt)
                If intr IsNot Nothing Then Utils.ComReleaseHelper.Release(intr)
                If brd IsNot Nothing Then Utils.ComReleaseHelper.Release(brd)
                If used IsNot Nothing Then Utils.ComReleaseHelper.Release(used)
            End Try
        End Function

        ''' <summary>
        ''' 把单格格式快照应用回单元格（仅设置快照中记录的属性）
        ''' </summary>
        Private Sub ApplyCellFmtSnapshot(cell As Range, f As Dictionary(Of String, Object))
            If f.ContainsKey("b") Then cell.Font.Bold = True
            If f.ContainsKey("i") Then cell.Font.Italic = True
            If f.ContainsKey("fc") Then cell.Font.Color = Convert.ToDouble(f("fc"))
            If f.ContainsKey("icc") Then cell.Interior.Color = Convert.ToDouble(f("icc"))
            If f.ContainsKey("ha") Then cell.HorizontalAlignment = Convert.ToInt32(f("ha"))
            If f.ContainsKey("va") Then cell.VerticalAlignment = Convert.ToInt32(f("va"))
            If f.ContainsKey("bs") Then cell.Borders.LineStyle = Convert.ToInt32(f("bs"))
            If f.ContainsKey("fn") Then cell.Font.Name = f("fn").ToString()
            If f.ContainsKey("fs") Then cell.Font.Size = Convert.ToDouble(f("fs"))
        End Sub

        ''' <summary>
        ''' 还原后清理美化可能扩展到快照 addr 之外、且原本为空的单元格上的残留格式
        ''' （如边框/底色落在原 used range 之外的空白单元格）。只清无值单元格的格式，
        ''' 不会误伤 addr 内已还原的内容，也不会清除 addr 外用户自己的数据。
        ''' </summary>
        Private Sub ClearStrayFormatOutside(ws As Worksheet, addr As String)
            Try
                Dim snapRng = ws.Range(addr)
                Dim aR1 = snapRng.Row, aC1 = snapRng.Column
                Dim aR2 = aR1 + snapRng.Rows.Count - 1, aC2 = aC1 + snapRng.Columns.Count - 1
                Utils.ComReleaseHelper.Release(snapRng)
                Dim big = ws.UsedRange
                If big Is Nothing Then Return
                If big.Cells.Count > 500000 Then
                    Utils.ComReleaseHelper.Release(big)
                    Return
                End If
                For Each cc As Range In big.Cells
                    Dim rr = cc.Row, cl = cc.Column
                    If rr < aR1 OrElse rr > aR2 OrElse cl < aC1 OrElse cl > aC2 Then
                        If cc.Value Is Nothing Then cc.ClearFormats()
                    End If
                Next
                Utils.ComReleaseHelper.Release(big)
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' 收集 used range 内的合并单元格区域地址（去重），用于整表快照的合并状态还原
        ''' </summary>
        Private Sub CollectMergedAreas(ws As Worksheet, used As Range, merged As List(Of String))
            Dim seen As New HashSet(Of String)()
            For Each c As Range In used.Cells
                Try
                    If CBool(c.MergeCells) Then
                        Dim ma = c.MergeArea
                        Dim a = ma.Address(False, False)
                        If Not seen.Contains(a) Then
                            seen.Add(a)
                            merged.Add(a)
                        End If
                        Utils.ComReleaseHelper.Release(ma)
                    End If
                Catch
                End Try
            Next
        End Sub

        ''' <summary>
        ''' 整表备份：在"增删改"操作执行【之前】调用，把受影响工作表（used range 全部内容+格式+
        ''' 公式+合并单元格+列宽+行高）拍成快照压入撤销栈。撤销时整张还原，无需关心操作细节。
        ''' 批量操作期间（UndoSuppressed=True）由首个操作统一备份，子操作不重复登记。
        ''' </summary>
        Public Sub AddSheetBackup(sheetName As String) Implements IUndoBridge.AddSheetBackup
            If UndoSuppressed Then Return
            SyncLock _undoStackLock
                Dim ws = FindWorksheetByName(sheetName)
                If ws Is Nothing Then Return
                Dim snap = CaptureSheetFullSnapshot(ws)
                Utils.ComReleaseHelper.Release(ws)
                If snap = "{}" Then Return
                _undoStack.Add(New UndoAction("sheet_backup", sheetName, Nothing, Nothing, sheetName, snap))
                While _undoStack.Count > MAX_UNDO_STACK_SIZE
                    _undoStack.RemoveAt(0)
                End While
                _redoStack.Clear()
            End SyncLock
        End Sub

        ' 撤销整表备份：清空当前 used range 与快照区域的并集后，按快照整体还原
        ' （值/公式、数字格式、列宽、行高、每格格式、合并单元格）。所有 cell 类操作统一走这里。
        ''' <summary>
        ''' 整表撤销：先在恢复前补拍一张"当前(改后)"快照挂到 action.NewValue 供重做，再整体还原到改前。
        ''' 重做快照只在用户真正按下撤销时才捕获，正常写操作零额外开销。
        ''' </summary>
        Private Function UndoSheetBackup(action As UndoAction) As String
            Try
                Dim wsSnap = FindWorksheetByName(action.SheetName)
                If wsSnap IsNot Nothing Then
                    Dim post = CaptureSheetFullSnapshot(wsSnap)
                    If post IsNot Nothing AndAlso post <> "{}" Then action.NewValue = post
                    Utils.ComReleaseHelper.Release(wsSnap)
                End If
            Catch
            End Try
            Return RestoreSheetSnapshot(action, action.Payload, True)
        End Function

        ''' <summary>
        ''' 整表重做：还原到撤销那一刻捕获的"改后"快照
        ''' </summary>
        Private Function RedoSheetBackup(action As UndoAction) As String
            If action.NewValue Is Nothing OrElse action.NewValue.ToString() = "{}" Then
                Return JsonConvert.SerializeObject(New With {.success = False, .message = "没有可重做的快照数据"})
            End If
            Return RestoreSheetSnapshot(action, action.NewValue.ToString(), False)
        End Function

        ''' <summary>
        ''' 按整表快照还原工作表（撤销/重做共用）
        ''' </summary>
        Private Function RestoreSheetSnapshot(action As UndoAction, snapJson As String, isUndo As Boolean) As String
            Dim ws As Worksheet = Nothing
            Dim rng As Range = Nothing
            Try
                ws = FindWorksheetByName(action.SheetName)
                If ws Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .message = $"找不到工作表: {action.SheetName}"})
                End If
                If String.IsNullOrWhiteSpace(snapJson) OrElse snapJson = "{}" Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .message = "快照不可用，无法还原"})
                End If
                Dim p = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(snapJson)
                If Not p.ContainsKey("addr") Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .message = "快照不可用，无法还原"})
                End If
                Dim addr = p("addr").ToString()
                Dim rc = Convert.ToInt32(p("rows"))
                Dim cc = Convert.ToInt32(p("cols"))

                ' 整表还原：清空"快照区域 ∪ 当前 used range"的整个并集，确保插入/删除行列后也能完整回退
                Dim backupRng = ws.Range(addr)
                Dim uR1 = backupRng.Row, uC1 = backupRng.Column
                Dim uR2 = uR1 + backupRng.Rows.Count - 1, uC2 = uC1 + backupRng.Columns.Count - 1
                Utils.ComReleaseHelper.Release(backupRng)
                Dim cur = ws.UsedRange
                If cur IsNot Nothing Then
                    uR1 = Math.Min(uR1, cur.Row)
                    uC1 = Math.Min(uC1, cur.Column)
                    uR2 = Math.Max(uR2, cur.Row + cur.Rows.Count - 1)
                    uC2 = Math.Max(uC2, cur.Column + cur.Columns.Count - 1)
                    Utils.ComReleaseHelper.Release(cur)
                End If
                Dim unionRng = ws.Range(ws.Cells(uR1, uC1), ws.Cells(uR2, uC2))
                Try : unionRng.UnMerge() : Catch : End Try
                Try : unionRng.Clear() : Catch : End Try
                Utils.ComReleaseHelper.Release(unionRng)

                ' 回填值（用 FormulaLocal 保留公式，避免撤销后公式变死值）
                If p.ContainsKey("values") AndAlso p("values") IsNot Nothing Then
                    Dim data = JsonConvert.DeserializeObject(Of List(Of List(Of Object)))(p("values").ToString())
                    If data.Count > 0 AndAlso data(0).Count > 0 Then
                        rng = ws.Range(addr)
                        Dim vals(rc - 1, cc - 1) As Object
                        For r = 0 To rc - 1
                            For c = 0 To cc - 1
                                vals(r, c) = data(r)(c)
                            Next
                        Next
                        rng.FormulaLocal = vals
                    End If
                End If

                If rng Is Nothing Then rng = ws.Range(addr)

                ' 回填数字格式
                If p.ContainsKey("nfAll") AndAlso p("nfAll") IsNot Nothing Then
                    rng.NumberFormat = p("nfAll").ToString()
                ElseIf p.ContainsKey("nfCols") AndAlso p("nfCols") IsNot Nothing Then
                    Dim nfCols = JsonConvert.DeserializeObject(Of List(Of Object))(p("nfCols").ToString())
                    For c = 0 To nfCols.Count - 1
                        If nfCols(c) IsNot Nothing Then
                            Dim colRng = DirectCast(rng.Columns(c + 1), Range)
                            colRng.NumberFormat = nfCols(c).ToString()
                            Utils.ComReleaseHelper.Release(colRng)
                        End If
                    Next
                End If

                ' 回填列宽
                If p.ContainsKey("colWidths") AndAlso p("colWidths") IsNot Nothing Then
                    Dim cw = JsonConvert.DeserializeObject(Of List(Of Object))(p("colWidths").ToString())
                    For c = 0 To cw.Count - 1
                        If cw(c) IsNot Nothing Then
                            Dim colRng = DirectCast(rng.Columns(c + 1), Range)
                            colRng.ColumnWidth = Convert.ToDouble(cw(c))
                            Utils.ComReleaseHelper.Release(colRng)
                        End If
                    Next
                End If

                ' 回填行高
                If p.ContainsKey("rowHeights") AndAlso p("rowHeights") IsNot Nothing Then
                    Dim rh = JsonConvert.DeserializeObject(Of List(Of Object))(p("rowHeights").ToString())
                    For r = 0 To rh.Count - 1
                        If rh(r) IsNot Nothing Then
                            Try
                                Dim rw = DirectCast(ws.Rows(r + 1), Range)
                                rw.RowHeight = Convert.ToDouble(rh(r))
                                Utils.ComReleaseHelper.Release(rw)
                            Catch
                            End Try
                        End If
                    Next
                End If

                ' 回填每格格式
                If p.ContainsKey("fmt") AndAlso p("fmt") IsNot Nothing Then
                    Dim fmt = JsonConvert.DeserializeObject(Of List(Of List(Of Dictionary(Of String, Object))))(p("fmt").ToString())
                    For r = 0 To fmt.Count - 1
                        For c = 0 To fmt(r).Count - 1
                            Dim f = fmt(r)(c)
                            If f IsNot Nothing Then
                                Dim cell = DirectCast(rng.Cells(r + 1, c + 1), Range)
                                ApplyCellFmtSnapshot(cell, f)
                                Utils.ComReleaseHelper.Release(cell)
                            End If
                        Next
                    Next
                End If

                ' 清理美化扩展出快照区域外的残留格式（原为空白单元格上的边框/底色等），
                ' 否则这些无边框单元格撤销后仍残留美化边框
                ClearStrayFormatOutside(ws, addr)

                ' 还原合并单元格
                If p.ContainsKey("merged") AndAlso p("merged") IsNot Nothing Then
                    Dim merged = JsonConvert.DeserializeObject(Of List(Of String))(p("merged").ToString())
                    For Each m In merged
                        Try
                            ws.Range(m).Merge()
                        Catch
                        End Try
                    Next
                End If

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .message = If(isUndo,
                        $"已撤销操作，工作表 {action.SheetName} 已整体恢复到操作前状态",
                        $"已重做操作，工作表 {action.SheetName} 已整体恢复到操作后状态"),
                    .actionType = action.ActionType,
                    .sheetName = action.SheetName})
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .message = $"撤销失败: {ex.Message}"})
            Finally
                If rng IsNot Nothing Then Utils.ComReleaseHelper.Release(rng)
                If ws IsNot Nothing Then Utils.ComReleaseHelper.Release(ws)
            End Try
        End Function

#End Region





        Private Function UndoRenameSheet(action As UndoAction) As String
            Dim wb As Workbook = Nothing
            Dim ws As Worksheet = Nothing
            Try
                Dim excelApp = Globals.ThisAddIn.Application
                wb = excelApp.ActiveWorkbook
                If wb Is Nothing Then Return JsonConvert.SerializeObject(New With {.success = False, .message = "当前没有打开的工作簿"})
                Dim p = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(action.Payload)
                Dim oldName = p("oldName").ToString()
                ws = CType(wb.Worksheets(action.SheetName), Worksheet)
                ws.Name = oldName
                Return JsonConvert.SerializeObject(New With {.success = True, .message = $"已撤销重命名: {action.SheetName} -> {oldName}", .actionType = action.ActionType, .sheetName = oldName})
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .message = $"撤销失败: {ex.Message}"})
            Finally
                If ws IsNot Nothing Then Utils.ComReleaseHelper.Release(ws)
                If wb IsNot Nothing Then Utils.ComReleaseHelper.Release(wb)
            End Try
        End Function

        Private Function RedoRenameSheet(action As UndoAction) As String
            Dim wb As Workbook = Nothing
            Dim ws As Worksheet = Nothing
            Try
                Dim excelApp = Globals.ThisAddIn.Application
                wb = excelApp.ActiveWorkbook
                If wb Is Nothing Then Return JsonConvert.SerializeObject(New With {.success = False, .message = "当前没有打开的工作簿"})
                Dim p = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(action.Payload)
                Dim newName = p("newName").ToString()
                ' 撤销后工作表名已变回 oldName，重做须用 oldName 定位
                ws = CType(wb.Worksheets(p("oldName").ToString()), Worksheet)
                ws.Name = newName
                Return JsonConvert.SerializeObject(New With {.success = True, .message = $"已重做重命名: {action.SheetName} -> {newName}", .actionType = action.ActionType, .sheetName = newName})
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .message = $"重做失败: {ex.Message}"})
            Finally
                If ws IsNot Nothing Then Utils.ComReleaseHelper.Release(ws)
                If wb IsNot Nothing Then Utils.ComReleaseHelper.Release(wb)
            End Try
        End Function

        Private Function UndoDeleteSheet(action As UndoAction) As String
            Dim wb As Workbook = Nothing
            Dim newWs As Worksheet = Nothing
            Dim rng As Range = Nothing
            Try
                Dim excelApp = Globals.ThisAddIn.Application
                wb = excelApp.ActiveWorkbook
                If wb Is Nothing Then Return JsonConvert.SerializeObject(New With {.success = False, .message = "当前没有打开的工作簿"})
                Dim p = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(action.Payload)
                Dim sheetName = p("name").ToString()
                Dim addr = If(p.ContainsKey("addr"), p("addr").ToString(), "")
                newWs = CType(wb.Worksheets.Add(), Worksheet)
                newWs.Name = sheetName
                Try
                    Dim idx = Convert.ToInt32(p("index"))
                    If idx >= 1 AndAlso idx <= wb.Worksheets.Count Then
                        newWs.Move(Before:=wb.Worksheets(idx))
                    End If
                Catch
                End Try
                If Not String.IsNullOrWhiteSpace(addr) AndAlso p.ContainsKey("values") AndAlso p("values") IsNot Nothing Then
                    rng = newWs.Range(addr)
                    Dim data = JsonConvert.DeserializeObject(Of List(Of List(Of Object)))(p("values").ToString())
                    If data.Count > 0 Then
                        Dim cc = data(0).Count
                        Dim vals(data.Count - 1, cc - 1) As Object
                        For r = 0 To data.Count - 1
                            For c = 0 To cc - 1
                                vals(r, c) = data(r)(c)
                            Next
                        Next
                        rng.FormulaLocal = vals
                    End If
                    If p.ContainsKey("nf") AndAlso p("nf") IsNot Nothing Then
                        Dim nf = p("nf")
                        If TypeOf nf Is String Then
                            Try
                                rng.NumberFormat = nf.ToString()
                            Catch
                            End Try
                        Else
                            Try
                                Dim nfArr = JsonConvert.DeserializeObject(Of List(Of List(Of Object)))(nf.ToString())
                                For r = 0 To nfArr.Count - 1
                                    For c = 0 To nfArr(r).Count - 1
                                        If nfArr(r)(c) IsNot Nothing AndAlso nfArr(r)(c).ToString() <> "" Then
                                            CType(rng.Cells(r + 1, c + 1), Range).NumberFormat = nfArr(r)(c).ToString()
                                        End If
                                    Next
                                Next
                            Catch
                            End Try
                        End If
                    End If
                    If p.ContainsKey("colWidths") AndAlso p("colWidths") IsNot Nothing Then
                        Try
                            Dim cw = JsonConvert.DeserializeObject(Of List(Of Object))(p("colWidths").ToString())
                            For c = 0 To cw.Count - 1
                                If cw(c) IsNot Nothing Then DirectCast(newWs.Columns(c + 1), Range).ColumnWidth = Convert.ToDouble(cw(c))
                            Next
                        Catch
                        End Try
                    End If
                End If
                Return JsonConvert.SerializeObject(New With {.success = True, .message = $"已撤销删除工作表: {sheetName}", .actionType = action.ActionType, .sheetName = sheetName})
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .message = $"撤销失败: {ex.Message}"})
            Finally
                If rng IsNot Nothing Then Utils.ComReleaseHelper.Release(rng)
                If newWs IsNot Nothing Then Utils.ComReleaseHelper.Release(newWs)
                If wb IsNot Nothing Then Utils.ComReleaseHelper.Release(wb)
            End Try
        End Function

        Private Function RedoDeleteSheet(action As UndoAction) As String
            Dim wb As Workbook = Nothing
            Dim ws As Worksheet = Nothing
            Try
                Dim excelApp = Globals.ThisAddIn.Application
                wb = excelApp.ActiveWorkbook
                If wb Is Nothing Then Return JsonConvert.SerializeObject(New With {.success = False, .message = "当前没有打开的工作簿"})
                Dim sheetName = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(action.Payload)("name").ToString()
                ws = CType(wb.Worksheets(sheetName), Worksheet)
                If wb.Worksheets.Count <= 1 Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .message = "无法删除唯一的工作表"})
                End If
                excelApp.DisplayAlerts = False
                ws.Delete()
                excelApp.DisplayAlerts = True
                Return JsonConvert.SerializeObject(New With {.success = True, .message = $"已重做删除工作表: {sheetName}", .actionType = action.ActionType, .sheetName = sheetName})
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .message = $"重做失败: {ex.Message}"})
            Finally
                If ws IsNot Nothing Then Utils.ComReleaseHelper.Release(ws)
                If wb IsNot Nothing Then Utils.ComReleaseHelper.Release(wb)
            End Try
        End Function

#End Region

        ''' <summary>
        ''' 清空撤销栈
        ''' </summary>
        <ComVisible(True)>
        Public Sub ClearUndoStack() Implements IUndoBridge.ClearUndoStack
            SyncLock _undoStackLock
                _undoStack.Clear()
            End SyncLock
        End Sub

        ''' <summary>
        ''' 获取重做栈信息
        ''' </summary>
        <ComVisible(True)>
        Public Function GetRedoStack() As String Implements IUndoBridge.GetRedoStack
            Try
                SyncLock _undoStackLock
                    Return JsonConvert.SerializeObject(New With {
                        .count = _redoStack.Count,
                        .actions = _redoStack.Select(Function(a) New With {
                            .actionType = a.ActionType,
                            .address = a.Address,
                            .sheetName = a.SheetName,
                            .timestamp = a.Timestamp
                        }).ToList()
                    })
                End SyncLock
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.count = 0, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 执行重做操作
        ''' </summary>
        <ComVisible(True)>
        Public Function RedoLastAction() As String Implements IUndoBridge.RedoLastAction
            Try
                SyncLock _undoStackLock
                    If _isProcessing Then
                        Return JsonConvert.SerializeObject(New With {
                            .success = False,
                            .message = "重做操作正在进行中，请稍候"
                        })
                    End If
                    _isProcessing = True
                    Try
                        If _redoStack.Count = 0 Then
                            Return JsonConvert.SerializeObject(New With {
                                .success = False,
                                .message = "没有可重做的操作"
                            })
                        End If

                    Dim action = _redoStack(_redoStack.Count - 1)
                    _redoStack.RemoveAt(_redoStack.Count - 1)

                    ' 将重做的动作推回撤销栈
                    _undoStack.Add(action)
                    While _undoStack.Count > MAX_UNDO_STACK_SIZE
                        _undoStack.RemoveAt(0)
                    End While

                    Select Case action.ActionType.ToLower()
                        ' ---- 单元格/格式/行列增删：整表快照重做 ----
                        Case "sheet_backup", "setcell", "fill_data", "format", "clear_format",
                             "vba_execute", "delete_rows", "insert_rows", "delete_columns", "insert_columns"
                            Return RedoSheetBackup(action)

                        ' ---- 工作表结构类：保留原有专用重做 ----
                        Case "create_sheet"
                            Return RedoCreateSheet(action)
                        Case "merge_sheets"
                            Return RedoMergeSheets(action)
                        Case "split_workbook"
                            Return RedoSplitWorkbook(action)
                        Case "rename_sheet"
                            Return RedoRenameSheet(action)
                        Case "delete_sheet"
                            Return RedoDeleteSheet(action)
                        Case Else
                            Return JsonConvert.SerializeObject(New With {
                                .success = False,
                                .message = $"不支持的重做类型: {action.ActionType}"
                            })
                    End Select
                    Finally
                        _isProcessing = False
                    End Try
                End SyncLock
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .message = $"重做操作异常: {ex.Message}"
                })
            End Try
        End Function

        ''' <summary>
        ''' 清空重做栈
        ''' </summary>
        <ComVisible(True)>
        Public Sub ClearRedoStack() Implements IUndoBridge.ClearRedoStack
            SyncLock _undoStackLock
                _redoStack.Clear()
            End SyncLock
        End Sub


        ''' <summary>
        ''' 重做创建工作表：重新创建该工作表
        ''' </summary>
        Private Function RedoCreateSheet(action As UndoAction) As String
            Dim wb As Workbook = Nothing
            Dim worksheets As Sheets = Nothing
            Try
                Dim excelApp As Application = Globals.ThisAddIn.Application
                wb = excelApp.ActiveWorkbook
                If wb Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .message = "当前没有打开的工作簿"})
                End If

                worksheets = wb.Worksheets
                Dim sheetName = action.SheetName

                ' 检查是否已存在同名工作表
                For i As Integer = 1 To worksheets.Count
                    Dim ws As Worksheet = Nothing
                    Try
                        ws = CType(worksheets(i), Worksheet)
                        If ws.Name = sheetName Then
                            Return JsonConvert.SerializeObject(New With {
                                .success = True,
                                .message = $"工作表已存在: {sheetName}"
                            })
                        End If
                    Finally
                        Utils.ComReleaseHelper.Release(ws)
                    End Try
                Next

                Dim newSheet As Worksheet = CType(worksheets.Add(After:=worksheets(worksheets.Count)), Worksheet)
                newSheet.Name = sheetName

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .message = $"已重做创建工作表: {sheetName}",
                    .actionType = action.ActionType,
                    .sheetName = sheetName
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .message = $"重做失败: {ex.Message}"})
            Finally
                Utils.ComReleaseHelper.Release(worksheets)
                Utils.ComReleaseHelper.Release(wb)
            End Try
        End Function


        ''' <summary>
        ''' 重做合并工作表：重新创建工作表并恢复数据
        ''' Payload 格式：{ "sheetName": "...", "data": [[...], [...]] }
        ''' </summary>
        Private Function RedoMergeSheets(action As UndoAction) As String
            Dim wb As Workbook = Nothing
            Dim worksheets As Sheets = Nothing
            Dim newSheet As Worksheet = Nothing
            Try
                If String.IsNullOrWhiteSpace(action.Payload) Then
                    ' 兼容旧版本：无 Payload 时降级为仅创建空表
                    Return RedoCreateSheet(action)
                End If

                Dim excelApp As Application = Globals.ThisAddIn.Application
                wb = excelApp.ActiveWorkbook
                If wb Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .message = "当前没有打开的工作簿"})
                End If

                Dim payload = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(action.Payload)
                Dim sheetName = payload("sheetName")?.ToString()
                If String.IsNullOrWhiteSpace(sheetName) Then sheetName = action.SheetName
                Dim dataToken = TryCast(payload("data"), Newtonsoft.Json.Linq.JToken)
                Dim dataList = If(dataToken IsNot Nothing,
                    JsonConvert.DeserializeObject(Of List(Of List(Of Object)))(dataToken.ToString()),
                    Nothing)

                worksheets = wb.Worksheets

                ' 检查是否已存在同名工作表
                For i As Integer = 1 To worksheets.Count
                    Dim ws As Worksheet = Nothing
                    Try
                        ws = CType(worksheets(i), Worksheet)
                        If ws.Name = sheetName Then
                            ' 已存在，清空后恢复数据
                            CType(ws.Cells, Range).Clear()
                            newSheet = ws
                            Exit For
                        End If
                    Finally
                        If ws IsNot Nothing AndAlso Not Object.ReferenceEquals(ws, newSheet) Then
                            Utils.ComReleaseHelper.Release(ws)
                        End If
                    End Try
                Next

                ' 如果不存在，创建新工作表
                If newSheet Is Nothing Then
                    newSheet = CType(worksheets.Add(After:=worksheets(worksheets.Count)), Worksheet)
                    newSheet.Name = sheetName
                End If

                ' 写入数据（如果有）
                If dataList IsNot Nothing AndAlso dataList.Count > 0 Then
                    Dim rowCount = dataList.Count
                    Dim colCount = If(rowCount > 0, dataList(0).Count, 0)
                    If rowCount > 0 AndAlso colCount > 0 Then
                        Dim values(rowCount - 1, colCount - 1) As Object
                        For r = 0 To rowCount - 1
                            For c = 0 To colCount - 1
                                values(r, c) = If(dataList(r).Count > c, dataList(r)(c), Nothing)
                            Next
                        Next
                        Dim topLeft As Range = CType(newSheet.Cells(1, 1), Range)
                        Dim bottomRight As Range = CType(newSheet.Cells(rowCount, colCount), Range)
                        Dim rng As Range = CType(newSheet.Range(topLeft, bottomRight), Range)
                        rng.Value = values
                        Utils.ComReleaseHelper.Release(topLeft)
                        Utils.ComReleaseHelper.Release(bottomRight)
                        Utils.ComReleaseHelper.Release(rng)
                    End If
                End If

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .message = $"已重做合并工作表: {sheetName}",
                    .actionType = action.ActionType,
                    .sheetName = sheetName
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .message = $"重做失败: {ex.Message}"})
            Finally
                Utils.ComReleaseHelper.Release(worksheets)
                Utils.ComReleaseHelper.Release(newSheet)
                Utils.ComReleaseHelper.Release(wb)
            End Try
        End Function

        ''' <summary>
        ''' 重做拆分工作簿：重新生成文件
        ''' Payload 格式：{ "files": [...], "sourceWorkbook": "...", "outputFolder": "...", "splitMethod": "...", "sheetCount": N }
        ''' </summary>
        Private Function RedoSplitWorkbook(action As UndoAction) As String
            Dim sourceWb As Workbook = Nothing
            Dim app As Application = Nothing
            Try
                If String.IsNullOrWhiteSpace(action.Payload) Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .message = "缺少重做数据"
                    })
                End If

                Dim payload = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(action.Payload)
                Dim files = If(payload.ContainsKey("files"),
                    JsonConvert.DeserializeObject(Of List(Of String))(payload("files").ToString()),
                    New List(Of String)())
                Dim sourceWorkbookPath = payload("sourceWorkbook")?.ToString()
                Dim outputFolder = payload("outputFolder")?.ToString()
                Dim splitMethod = If(payload.ContainsKey("splitMethod"), payload("splitMethod")?.ToString(), "by_sheet")

                ' 如果源工作簿路径缺失，尝试使用当前活动工作簿
                app = Globals.ThisAddIn.Application
                Dim needOpenSource As Boolean = False
                If Not String.IsNullOrWhiteSpace(sourceWorkbookPath) AndAlso System.IO.File.Exists(sourceWorkbookPath) Then
                    ' 检查该工作簿是否已打开
                    For Each wbLoop As Workbook In app.Workbooks
                        If String.Equals(wbLoop.FullName, sourceWorkbookPath, StringComparison.OrdinalIgnoreCase) Then
                            sourceWb = wbLoop
                            Exit For
                        End If
                    Next
                    If sourceWb Is Nothing Then
                        needOpenSource = True
                        sourceWb = app.Workbooks.Open(sourceWorkbookPath)
                    End If
                Else
                    sourceWb = CType(app.ActiveWorkbook, Workbook)
                End If

                If sourceWb Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .message = "无法获取源工作簿"
                    })
                End If

                ' 确定输出目录
                If String.IsNullOrWhiteSpace(outputFolder) OrElse Not System.IO.Directory.Exists(outputFolder) Then
                    Try
                        outputFolder = System.IO.Path.GetDirectoryName(sourceWb.FullName)
                    Catch ex As Exception  ' 
                        outputFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                    End Try
                End If

                Dim generatedFiles As New List(Of String)()
                Dim baseName As String = System.IO.Path.GetFileNameWithoutExtension(sourceWb.Name)
                Dim sheetCount As Integer = CInt(sourceWb.Worksheets.Count)

                ' 先删除已存在的同名文件（避免冲突）
                For Each existingFile In files
                    If System.IO.File.Exists(existingFile) Then
                        Try
                            System.IO.File.Delete(existingFile)
                        Catch ex As Exception  ' 
                        End Try
                    End If
                Next

                ' 按工作表重新拆分
                For i As Integer = sheetCount To 1 Step -1
                    Dim srcWs As Worksheet = Nothing
                    Dim newWb As Workbook = Nothing
                    Try
                        srcWs = CType(sourceWb.Worksheets(i), Worksheet)
                        Dim sheetName As String = srcWs.Name
                        srcWs.Copy()
                        newWb = CType(app.ActiveWorkbook, Workbook)
                        Dim fileName As String = $"{baseName}_{sheetName}.xlsx"
                        Dim filePath As String = System.IO.Path.Combine(outputFolder, fileName)
                        ' 处理路径冲突
                        If System.IO.File.Exists(filePath) Then
                            Dim ts As String = DateTime.Now.ToString("HHmmss")
                            fileName = $"{baseName}_{sheetName}_{ts}.xlsx"
                            filePath = System.IO.Path.Combine(outputFolder, fileName)
                        End If
                        newWb.SaveAs(filePath)
                        newWb.Close(SaveChanges:=False)
                        generatedFiles.Add(filePath)
                    Finally
                        Utils.ComReleaseHelper.Release(srcWs)
                        Utils.ComReleaseHelper.Release(newWb)
                    End Try
                Next

                ' 如果是我们打开的源工作簿，关闭它（不保存）
                If needOpenSource AndAlso sourceWb IsNot Nothing Then
                    sourceWb.Close(SaveChanges:=False)
                End If

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .message = $"已重做拆分工作簿，生成 {generatedFiles.Count} 个文件",
                    .actionType = action.ActionType,
                    .generatedCount = generatedFiles.Count,
                    .files = generatedFiles
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .message = $"重做拆分失败: {ex.Message}"
                })
            Finally
                ' 注意：不要随意关闭 sourceWb，可能是用户正在使用的
            End Try
        End Function

        ''' <summary>批量操作（如整次美化）期间临时抑制子操作登记，使整次操作合并为单一撤销项</summary>
        Public Shared Property UndoSuppressed As Boolean = False

        ''' <summary>
        ''' 添加撤销动作到栈
        ''' </summary>
        Public Sub AddUndoAction(action As UndoAction) Implements IUndoBridge.AddUndoAction
            If UndoSuppressed Then Return
            SyncLock _undoStackLock
                _undoStack.Add(action)

                ' 限制栈大小
                While _undoStack.Count > MAX_UNDO_STACK_SIZE
                    _undoStack.RemoveAt(0)
                End While

                ' 有新操作加入时，清空重做栈
                _redoStack.Clear()
            End SyncLock
        End Sub

#End Region

#Region "IBridge 生命周期"

        Public ReadOnly Property BridgeName As String Implements IBridge.BridgeName
            Get
                Return "UndoBridge"
            End Get
        End Property

        Private _disposed As Boolean = False
        Public Sub Dispose() Implements IBridge.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not _disposed Then
                _disposed = True
            End If
        End Sub

#End Region

    End Class

End Namespace
