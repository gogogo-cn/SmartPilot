Option Strict On
Option Infer On
Imports System.IO
Imports Microsoft.Office.Interop.Excel
Imports ExcelAIVSTOAddIn.Utils

Namespace Services

    ''' <summary>
    ''' 填入模式枚举
    ''' </summary>
    Public Enum FillMode
        Auto = 0
        CreateNew = 1
        SmartMatch = 2
    End Enum

    ''' <summary>
    ''' Excel操作服务 - 支持多种类型数据填入
    ''' </summary>
    Public Class ExcelService
        Private _excelApp As Application
        Private Shared ReadOnly _lazy As New Lazy(Of ExcelService)(Function() New ExcelService())

        ''' <summary>
        ''' 共享单例（推荐使用）——无状态工具类，只需一个实例
        ''' </summary>
        Public Shared ReadOnly Property Instance As ExcelService
            Get
                Return _lazy.Value
            End Get
        End Property

        Public Sub New()
            _excelApp = Globals.ThisAddIn.Application
        End Sub

        ''' <summary>
        ''' 根据数据类型获取字段映射
        ''' </summary>
        Private Function GetFieldMapping(dataType As String) As List(Of KeyValuePair(Of String, String))
            Select Case dataType
                ' [已弃用 - 识别类Skill已移除] 身份证/名片 Case 块已弃用
                'Case "身份证"
                '    Return New List(Of KeyValuePair(Of String, String)) From {
                '        New KeyValuePair(Of String, String)("name", "姓名"),
                '        New KeyValuePair(Of String, String)("gender", "性别"),
                '        New KeyValuePair(Of String, String)("ethnicity", "民族"),
                '        New KeyValuePair(Of String, String)("birthDate", "出生日期"),
                '        New KeyValuePair(Of String, String)("address", "住址"),
                '        New KeyValuePair(Of String, String)("idNumber", "公民身份号码"),
                '        New KeyValuePair(Of String, String)("issueAuthority", "签发机关"),
                '        New KeyValuePair(Of String, String)("validDate", "有效期限")
                '    }
                'Case "名片"
                '    Return New List(Of KeyValuePair(Of String, String)) From {
                '        New KeyValuePair(Of String, String)("name", "姓名"),
                '        New KeyValuePair(Of String, String)("company", "公司"),
                '        New KeyValuePair(Of String, String)("position", "职位"),
                '        New KeyValuePair(Of String, String)("phone", "电话"),
                '        New KeyValuePair(Of String, String)("mobile", "手机"),
                '        New KeyValuePair(Of String, String)("email", "邮箱"),
                '        New KeyValuePair(Of String, String)("address", "地址"),
                '        New KeyValuePair(Of String, String)("website", "网址")
                '    }
                Case Else ' 发票及其他类型
                    Return New List(Of KeyValuePair(Of String, String)) From {
                        New KeyValuePair(Of String, String)("invoiceCode", "发票代码"),
                        New KeyValuePair(Of String, String)("invoiceNumber", "发票号码"),
                        New KeyValuePair(Of String, String)("invoiceDate", "开票日期"),
                        New KeyValuePair(Of String, String)("amount", "金额"),
                        New KeyValuePair(Of String, String)("totalAmount", "价税合计"),
                        New KeyValuePair(Of String, String)("taxRate", "税率"),
                        New KeyValuePair(Of String, String)("taxAmount", "税额"),
                        New KeyValuePair(Of String, String)("sellerName", "销售方名称"),
                        New KeyValuePair(Of String, String)("sellerTaxId", "销售方税号"),
                        New KeyValuePair(Of String, String)("sellerAddressPhone", "销售方地址电话"),
                        New KeyValuePair(Of String, String)("sellerBankAccount", "销售方开户行"),
                        New KeyValuePair(Of String, String)("buyerName", "购买方名称"),
                        New KeyValuePair(Of String, String)("buyerTaxId", "购买方税号"),
                        New KeyValuePair(Of String, String)("buyerAddressPhone", "购买方地址电话"),
                        New KeyValuePair(Of String, String)("buyerBankAccount", "购买方开户行"),
                        New KeyValuePair(Of String, String)("itemName", "货物名称"),
                        New KeyValuePair(Of String, String)("specModel", "规格型号"),
                        New KeyValuePair(Of String, String)("unit", "单位"),
                        New KeyValuePair(Of String, String)("quantity", "数量"),
                        New KeyValuePair(Of String, String)("unitPrice", "单价"),
                        New KeyValuePair(Of String, String)("payee", "收款人"),
                        New KeyValuePair(Of String, String)("reviewer", "复核"),
                        New KeyValuePair(Of String, String)("drawer", "开票人"),
                        New KeyValuePair(Of String, String)("verifyCode", "校验码"),
                        New KeyValuePair(Of String, String)("machineCode", "机器编号"),
                        New KeyValuePair(Of String, String)("remarks", "备注")
                    }
            End Select
        End Function

        ''' <summary>
        ''' 获取需要文本格式的字段名
        ''' </summary>
        Private Function GetTextFields(dataType As String) As List(Of String)
            Select Case dataType
                Case "身份证"
                    Return New List(Of String) From {"idNumber"}
                Case "名片"
                    Return New List(Of String) From {"phone", "mobile"}
                Case Else ' 发票
                    Return New List(Of String) From {"invoiceCode", "invoiceNumber", "sellerTaxId", "buyerTaxId", "verifyCode", "machineCode"}
            End Select
        End Function

        ''' <summary>
        ''' 检测表头类型
        ''' </summary>
        Private Function DetectHeaderType(sheet As Worksheet) As String
            Dim headers As New List(Of String)()
            Dim col As Integer = 1
            
            ' 读取第一行的表头
            Do While True
                Dim headerCell As Range = Nothing
                Dim headerValue As Object = Nothing
                Try
                    headerCell = CType(sheet.Cells(1, col), Range)
                    headerValue = headerCell.Value
                Finally
                    ComReleaseHelper.Release(headerCell)
                End Try
                If headerValue Is Nothing OrElse String.IsNullOrEmpty(headerValue.ToString().Trim()) Then
                    Exit Do
                End If
                headers.Add(headerValue.ToString().Trim().ToLower())
                col += 1
                If col > 50 Then Exit Do
            Loop
            
            ' 定义各类数据的特征表头（按优先级排序）
            Dim invoiceKeywords = New String() {"发票代码", "发票号码", "开票日期", "价税合计", "销售方名称", "购买方名称", "纳税人识别号"}
            Dim idCardKeywords = New String() {"姓名", "性别", "民族", "出生日期", "住址", "公民身份号码", "签发机关", "有效期限"}
            Dim businessCardKeywords = New String() {"姓名", "公司", "职位", "电话", "手机", "邮箱", "网址"}
            
            ' 统计匹配数量
            Dim invoiceMatch As Integer = 0
            Dim idCardMatch As Integer = 0
            Dim businessCardMatch As Integer = 0
            
            For Each header In headers
                ' 检查发票关键词
                For Each keyword In invoiceKeywords
                    If header.Contains(keyword.ToLower()) Then
                        invoiceMatch += 1
                        Exit For
                    End If
                Next
                
                ' 检查身份证关键词
                For Each keyword In idCardKeywords
                    If header.Contains(keyword.ToLower()) Then
                        idCardMatch += 1
                        Exit For
                    End If
                Next
                
                ' 检查名片关键词
                For Each keyword In businessCardKeywords
                    If header.Contains(keyword.ToLower()) Then
                        businessCardMatch += 1
                        Exit For
                    End If
                Next
            Next
            
            System.Diagnostics.Debug.WriteLine($"表头检测 - 发票匹配: {invoiceMatch}, 身份证匹配: {idCardMatch}, 名片匹配: {businessCardMatch}")
            
            ' 返回匹配最多的类型（降低阈值，提高检测灵敏度）
            Dim maxMatch = Math.Max(invoiceMatch, Math.Max(idCardMatch, businessCardMatch))
            If maxMatch = 0 Then Return "" ' 无法确定类型
            
            ' 如果有明确匹配，返回对应类型
            If invoiceMatch >= 2 Then Return "发票"
            If idCardMatch >= 2 Then Return "身份证"
            If businessCardMatch >= 2 Then Return "名片"
            
            ' 如果只有一个匹配，检查是否是强特征字段
            If invoiceMatch = 1 AndAlso headers.Any(Function(h) h.Contains("发票")) Then Return "发票"
            If idCardMatch = 1 AndAlso headers.Any(Function(h) h.Contains("民族") Or h.Contains("公民身份号码")) Then Return "身份证"
            If businessCardMatch = 1 AndAlso headers.Any(Function(h) h.Contains("公司") Or h.Contains("职位")) Then Return "名片"
            
            Return "" ' 无法确定类型
        End Function

        ''' <summary>
        ''' 获取默认工作表名称
        ''' </summary>
        Private Function fGetDefaultWorksheetName(dataType As String) As String
            Select Case dataType
                Case "身份证"
                    Return "身份证信息"
                Case "名片"
                    Return "联系人"
                Case "发票"
                    Return "发票记录"
                Case "表格"
                    Return "表格数据"
                Case Else
                    Return dataType & "数据"
            End Select
        End Function

        ''' <summary>
        ''' 获取字段别名映射（用于支持多种字段名写法）
        ''' </summary>
        Private Function GetFieldAliases(dataType As String) As Dictionary(Of String, List(Of String))
            Dim aliases As New Dictionary(Of String, List(Of String))
            
            Select Case dataType
                Case "身份证"
                    aliases.Add("idNumber", New List(Of String) From {"idNumber", "身份证号码", "公民身份号码", "身份证号", "号码"})
                    aliases.Add("validDate", New List(Of String) From {"validDate", "有效期限", "有效期", "有效日期"})
                    aliases.Add("issueAuthority", New List(Of String) From {"issueAuthority", "签发机关", "签发单位"})
                    aliases.Add("name", New List(Of String) From {"name", "姓名", "名字"})
                    aliases.Add("gender", New List(Of String) From {"gender", "性别", "性"})
                    aliases.Add("ethnicity", New List(Of String) From {"ethnicity", "民族"})
                    aliases.Add("birthDate", New List(Of String) From {"birthDate", "出生日期", "出生"})
                    aliases.Add("address", New List(Of String) From {"address", "住址", "地址"})
                    
                Case "名片"
                    aliases.Add("name", New List(Of String) From {"name", "姓名", "名字"})
                    aliases.Add("company", New List(Of String) From {"company", "公司", "单位", "公司名称", "单位名称"})
                    aliases.Add("position", New List(Of String) From {"position", "职位", "职务", "岗位"})
                    aliases.Add("phone", New List(Of String) From {"phone", "电话", "固定电话", "座机"})
                    aliases.Add("mobile", New List(Of String) From {"mobile", "手机", "移动电话", "手机号码"})
                    aliases.Add("email", New List(Of String) From {"email", "邮箱", "电子邮件", "E-mail", "e-mail"})
                    aliases.Add("address", New List(Of String) From {"address", "地址", "住址", "单位地址", "公司地址"})
                    aliases.Add("website", New List(Of String) From {"website", "网址", "网站", "主页"})
                    
                Case Else ' 发票及其他类型
                    aliases.Add("invoiceCode", New List(Of String) From {"invoiceCode", "发票代码", "代码"})
                    aliases.Add("invoiceNumber", New List(Of String) From {"invoiceNumber", "发票号码", "号码", "发票号码"})
                    aliases.Add("invoiceDate", New List(Of String) From {"invoiceDate", "开票日期", "日期"})
                    aliases.Add("amount", New List(Of String) From {"amount", "金额", "不含税金额"})
                    aliases.Add("totalAmount", New List(Of String) From {"totalAmount", "价税合计", "合计金额", "总金额"})
                    aliases.Add("taxRate", New List(Of String) From {"taxRate", "税率"})
                    aliases.Add("taxAmount", New List(Of String) From {"taxAmount", "税额", "税金"})
                    aliases.Add("sellerName", New List(Of String) From {"sellerName", "销售方名称", "销售方", "卖方", "销售单位"})
                    aliases.Add("sellerTaxId", New List(Of String) From {"sellerTaxId", "销售方税号", "销售方纳税人识别号", "卖方税号"})
                    aliases.Add("sellerAddressPhone", New List(Of String) From {"sellerAddressPhone", "销售方地址电话", "销售方地址", "销售方电话"})
                    aliases.Add("sellerBankAccount", New List(Of String) From {"sellerBankAccount", "销售方开户行", "销售方银行", "卖方开户行"})
                    aliases.Add("buyerName", New List(Of String) From {"buyerName", "购买方名称", "购买方", "买方", "购买单位", "购方"})
                    aliases.Add("buyerTaxId", New List(Of String) From {"buyerTaxId", "购买方税号", "购买方纳税人识别号", "买方税号"})
                    aliases.Add("buyerAddressPhone", New List(Of String) From {"buyerAddressPhone", "购买方地址电话", "购买方地址", "购买方电话"})
                    aliases.Add("buyerBankAccount", New List(Of String) From {"buyerBankAccount", "购买方开户行", "购买方银行", "买方开户行"})
                    aliases.Add("itemName", New List(Of String) From {"itemName", "货物名称", "项目名称", "商品名称", "货物或应税劳务名称", "通行费", "费用名称"})
                    aliases.Add("specModel", New List(Of String) From {"specModel", "规格型号", "规格", "型号"})
                    aliases.Add("unit", New List(Of String) From {"unit", "单位", "计量单位"})
                    aliases.Add("quantity", New List(Of String) From {"quantity", "数量", "件数"})
                    aliases.Add("unitPrice", New List(Of String) From {"unitPrice", "单价", "价格"})
                    aliases.Add("payee", New List(Of String) From {"payee", "收款人", "收款", "收款方"})
                    aliases.Add("reviewer", New List(Of String) From {"reviewer", "复核", "复核人", "审核"})
                    aliases.Add("drawer", New List(Of String) From {"drawer", "开票人", "开票", "开票员"})
                    aliases.Add("verifyCode", New List(Of String) From {"verifyCode", "校验码", "验证码"})
                    aliases.Add("machineCode", New List(Of String) From {"machineCode", "机器编号", "机器码"})
                    aliases.Add("remarks", New List(Of String) From {"remarks", "备注", "附注", "说明"})
            End Select
            
            Return aliases
        End Function

        ''' <summary>
        ''' 通用数据填入Excel方法（支持发票、身份证、名片）
        ''' </summary>
        Public Sub FillDataToExcel(data As Dictionary(Of String, Object), dataType As String, imagePath As String)
            Dim workbook As Workbook = Nothing
            Dim activeCell As Range = Nothing
            Dim targetRange As Range = Nothing
            Try
                System.Diagnostics.Debug.WriteLine("===== 开始填入Excel（类型: " & dataType & "）=====")
                
                If _excelApp Is Nothing Then
                    Throw New Exception("Excel应用不可用")
                End If

                workbook = _excelApp.ActiveWorkbook
                If workbook Is Nothing Then
                    workbook = _excelApp.Workbooks.Add()
                End If

                Dim fields = GetFieldMapping(dataType)
                Dim colCount = fields.Count
                
                If colCount = 0 Then
                    System.Diagnostics.Debug.WriteLine("===== 没有字段映射，跳过填入 =====")
                    Return
                End If

                activeCell = CType(_excelApp.ActiveCell, Range)
                Dim startRow As Integer = CInt(activeCell.Row)
                Dim startCol As Integer = CInt(activeCell.Column)

                Dim dataArray(colCount - 1, 1) As Object

                For colIndex = 0 To colCount - 1
                    dataArray(colIndex, 0) = fields(colIndex).Value
                    
                    Dim fieldName As String = fields(colIndex).Value.ToLower()
                    Dim value As String = ""
                    For Each dataKvp In data
                        If dataKvp.Key.ToLower() = fields(colIndex).Key.ToLower() OrElse dataKvp.Key.ToLower() = fieldName OrElse dataKvp.Key.Contains(fields(colIndex).Value) Then
                            value = If(dataKvp.Value?.ToString(), "")
                            Exit For
                        End If
                    Next
                    dataArray(colIndex, 1) = value
                Next

                targetRange = CType(_excelApp.Cells(startRow, startCol), Range).Resize(2, colCount)
                targetRange.Value2 = dataArray
                
                System.Diagnostics.Debug.WriteLine("===== 填入Excel完成 =====")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("FillDataToExcel 错误: " & ex.Message)
                Throw
            Finally
                ComReleaseHelper.Release(targetRange)
                ComReleaseHelper.Release(activeCell)
                ComReleaseHelper.Release(workbook)
            End Try
        End Sub

        ''' <summary>
        ''' 多维表格数据填入Excel（支持数组格式）
        ''' </summary>
        Public Sub FillDataToExcel(dataArray As List(Of Dictionary(Of String, Object)), dataType As String, imagePath As String)
            Dim workbook As Workbook = Nothing
            Dim activeCell As Range = Nothing
            Dim targetRange As Range = Nothing
            Try
                System.Diagnostics.Debug.WriteLine("===== 开始填入多维表格到Excel（" & dataArray.Count & "行）=====")
                
                If _excelApp Is Nothing OrElse dataArray.Count = 0 Then
                    Throw New Exception("Excel应用不可用或数据为空")
                End If

                workbook = _excelApp.ActiveWorkbook
                If workbook Is Nothing Then
                    workbook = _excelApp.Workbooks.Add()
                End If

                Dim columns = dataArray(0).Keys.ToList()

                Dim defaultColumnPatterns As New List(Of String) From {
                    "列\d+", "Column\d+", "Col\d+", "表头\d+", "Header\d+", "字段\d+", "Field\d+"
                }
                Dim hasDefaultHeaders As Boolean = False
                Dim defaultColumnRegex As New System.Text.RegularExpressions.Regex(
                    String.Join("|", defaultColumnPatterns),
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase
                )

                If columns.Count > 0 Then
                    hasDefaultHeaders = columns.All(Function(col) defaultColumnRegex.IsMatch(col.Trim()))
                End If

                activeCell = CType(_excelApp.ActiveCell, Range)
                Dim startRow As Integer = CInt(activeCell.Row)
                Dim startCol As Integer = CInt(activeCell.Column)

                Dim rowCount As Integer = dataArray.Count
                Dim colCount As Integer = columns.Count
                Dim includeHeader As Boolean = Not hasDefaultHeaders
                Dim totalRows As Integer = rowCount + If(includeHeader, 1, 0)

                Dim outputArray(totalRows - 1, colCount - 1) As Object

                If includeHeader Then
                    For colIndex = 0 To colCount - 1
                        outputArray(0, colIndex) = columns(colIndex)
                    Next
                End If

                For dataRowIndex = 0 To rowCount - 1
                    Dim row As Dictionary(Of String, Object) = dataArray(dataRowIndex)
                    Dim outputRowIndex As Integer = dataRowIndex + If(includeHeader, 1, 0)
                    For colIndex = 0 To colCount - 1
                        Dim value As String = ""
                        If row.ContainsKey(columns(colIndex)) AndAlso row(columns(colIndex)) IsNot Nothing Then
                            value = row(columns(colIndex)).ToString()
                        End If
                        outputArray(outputRowIndex, colIndex) = value
                    Next
                Next

                targetRange = CType(_excelApp.Cells(startRow, startCol), Range).Resize(totalRows, colCount)
                targetRange.Value2 = outputArray

                If hasDefaultHeaders Then
                    System.Diagnostics.Debug.WriteLine("⚠️ 检测到默认列名（列1、列2等），跳过表头行")
                End If
                
                System.Diagnostics.Debug.WriteLine("===== 多维表格填入完成 =====")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("FillDataToExcel(数组) 错误: " & ex.Message)
                Throw
            Finally
                ComReleaseHelper.Release(targetRange)
                ComReleaseHelper.Release(activeCell)
                ComReleaseHelper.Release(workbook)
            End Try
        End Sub

        ''' <summary>
        ''' 简单直接的发票数据填入Excel（兼容旧代码）
        ''' </summary>
        Public Sub FillInvoiceData(invoiceData As Dictionary(Of String, Object), imagePath As String)
            FillDataToExcel(invoiceData, "发票", imagePath)
        End Sub

        ''' <summary>
        ''' 兼容旧代码 - 添加发票记录（智能模式）
        ''' </summary>
        Public Sub AddInvoiceRecordSmart(sheetName As String, invoiceData As Dictionary(Of String, Object), imagePath As String, mode As FillMode)
            FillInvoiceData(invoiceData, imagePath)
        End Sub

        ''' <summary>
        ''' 兼容旧代码 - 添加通用记录（智能模式）
        ''' </summary>
        Public Sub AddGenericRecordSmart(sheetName As String, data As Dictionary(Of String, Object), mode As FillMode)
            FillInvoiceData(data, "")
        End Sub

        ''' <summary>
        ''' 兼容旧代码 - 添加发票记录
        ''' </summary>
        Public Sub AddInvoiceRecord(sheetName As String, invoiceInfo As Object, imagePath As String)
        End Sub

        ''' <summary>
        ''' 兼容旧代码 - 保存工作簿
        ''' </summary>
        Public Sub SaveWorkbook()
        End Sub

    End Class

End Namespace