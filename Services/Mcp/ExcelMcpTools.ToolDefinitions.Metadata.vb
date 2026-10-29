Option Strict On
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Reflection
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ExcelAIVSTOAddIn.Utils

Namespace Services.Mcp

    Partial Public Class ExcelMcpTools

        ''' <summary>工具元数据显式覆盖（GetHardcodedToolMetadata 的 Select Case）</summary>
        Private Shared Sub ApplyExplicitToolMetadata(meta As ToolMeta, toolName As String)
            Select Case toolName
                Case "excel_read_range"
                    meta.DisplayName = "读取Excel数据"
                    meta.Keywords = New List(Of String)({"读取", "查看", "获取数据", "显示内容", "看看", "查"})
                Case "excel_write_range"
                    meta.DisplayName = "写入Excel数据"
                    meta.Keywords = New List(Of String)({"写入", "填充", "输入数据", "填入"})
                Case "excel_get_sheets"
                    meta.DisplayName = "获取工作表列表"
                    meta.Keywords = New List(Of String)({"工作表列表", "有哪些sheet", "sheets"})
                Case "excel_create_sheet"
                    meta.DisplayName = "创建工作表"
                    meta.Keywords = New List(Of String)({"新建工作表", "创建sheet", "新建sheet"})
                Case "excel_get_workbook_info"
                    meta.DisplayName = "工作簿信息"
                    meta.Keywords = New List(Of String)({"工作簿信息", "文件信息"})
                Case "excel_get_open_workbooks"
                    meta.DisplayName = "打开的工作簿"
                    meta.Keywords = New List(Of String)({"打开的工作簿", "所有工作簿"})
                Case "excel_search"
                    meta.DisplayName = "搜索Excel"
                    meta.Keywords = New List(Of String)({"搜索", "查找", "搜索内容"})
                Case "excel_create_chart"
                    meta.DisplayName = "创建图表"
                    meta.Keywords = New List(Of String)({"图表", "画图", "作图", "柱状图", "折线图", "饼图", "可视化"})
                Case "excel_format_range"
                    meta.DisplayName = "格式化单元格"
                    meta.Keywords = New List(Of String)({"格式化", "设置格式", "字体", "颜色", "边框", "对齐"})
                Case "excel_apply_conditional_format"
                    meta.DisplayName = "应用条件格式"
                    meta.Keywords = New List(Of String)({"条件格式", "数据条", "色阶", "图标集", "热力图", "渐变", "一键格式化"})
                Case "excel_remove_duplicates"
                    meta.DisplayName = "删除重复行"
                    meta.Keywords = New List(Of String)({"去重", "删除重复", "重复项", "去重复", "删重", "去重处理"})
                Case "excel_clean_data"
                    meta.DisplayName = "智能数据清洗"
                    meta.Keywords = New List(Of String)({"清洗", "数据清洗", "去空格", "填充空值", "统一日期", "统一数字", "文本规范化", "删除空行", "数据整理", "整理数据"})
                Case "excel_clear_range"
                    meta.DisplayName = "清空单元格"
                    meta.Keywords = New List(Of String)({"清空", "清除内容", "清除数据", "删除内容", "清空选中", "清空数据"})
                Case "excel_clear_format"
                    meta.DisplayName = "清除格式"
                    meta.Keywords = New List(Of String)({"清空格式", "清除格式", "去掉格式", "取消格式", "清除", "格式化清除"})
                Case "excel_clear_sheet"
                    meta.DisplayName = "清空工作表"
                    meta.Keywords = New List(Of String)({"清空工作表", "清空整个sheet", "清除所有数据"})
                Case "excel_sort_range"
                    meta.DisplayName = "排序"
                    meta.Keywords = New List(Of String)({"排序", "升序", "降序", "按...排"})
                Case "excel_apply_filter"
                    meta.DisplayName = "应用筛选"
                    meta.Keywords = New List(Of String)({"筛选", "过滤", "只看"})
                Case "excel_clear_filter"
                    meta.DisplayName = "清除筛选"
                    meta.Keywords = New List(Of String)({"清除筛选", "取消筛选", "显示全部"})
                Case "excel_split_by_column"
                    meta.DisplayName = "按列拆分"
                    meta.Keywords = New List(Of String)({"按列拆分", "拆分数据", "按...分", "分类"})
                Case "excel_split_workbook"
                    meta.DisplayName = "拆分工作簿"
                    meta.Keywords = New List(Of String)({"拆分工作簿", "按sheet拆分"})
                Case "excel_merge_workbooks"
                    meta.DisplayName = "合并工作簿"
                    meta.Keywords = New List(Of String)({"合并工作簿", "合并文件", "汇总"})
                Case "excel_merge_sheets"
                    meta.DisplayName = "合并工作表"
                    meta.Keywords = New List(Of String)({"合并sheet", "合并工作表", "合并多表", "合到一张表", "合并所有sheet", "合并到一张"})
                Case "excel_insert_rows"
                    meta.DisplayName = "插入数据行"
                    meta.Keywords = New List(Of String)({"插入数据", "追加数据", "写入数据行", "插入表格", "insert rows", "逐行写入"})
                Case "excel_split_by_rows"
                    meta.DisplayName = "按行数拆分"
                    meta.Keywords = New List(Of String)({"按行拆分", "按行数拆", "拆成多份", "每1000行", "每500行", "拆文件"})
                Case "excel_join_tables"
                    meta.DisplayName = "跨表关联"
                    meta.Keywords = New List(Of String)({"跨表关联", "关联两个表", "表连接", "join", "关联表", "连接两张表"})
                Case "excel_aggregate"
                    meta.DisplayName = "分组聚合统计"
                    meta.Keywords = New List(Of String)({"聚合", "分组统计", "求和", "平均", "计数", "合计", "汇总统计", "按...分组"})
                Case "excel_date_group"
                    meta.DisplayName = "日期分组"
                    meta.Keywords = New List(Of String)({"按年", "按季度", "按月", "按周", "按日", "日期分组", "时间分组", "年度", "季度", "月度", "周度", "财年", "按月汇总", "按季汇总", "按年汇总"})
                Case "excel_import_csv"
                    meta.DisplayName = "导入CSV"
                    meta.Keywords = New List(Of String)({"导入CSV", "导入csv"})
                Case "excel_export_csv"
                    meta.DisplayName = "导出CSV"
                    meta.Keywords = New List(Of String)({"导出CSV", "导出csv", "另存为CSV"})
                Case "excel_create_pivot_table"
                    meta.DisplayName = "创建数据透视表"
                    meta.Keywords = New List(Of String)({"透视表", "数据透视", "交叉表"})
                    meta.Category = CapabilityCategory.Analyze
                    ' 会向工作表写入透视表对象：Analyze 分类的默认 isWrite=False 不适用，须显式声明
                    meta.IsWrite = True
                Case "excel_create_slicer"
                    meta.DisplayName = "创建原生切片器"
                    meta.Keywords = New List(Of String)({"切片器", "原生切片器", "联动筛选", "透视表筛选"})
                    meta.Category = CapabilityCategory.Analyze
                    ' 会向工作表写入切片器对象
                    meta.IsWrite = True
                Case "excel_add_shape"
                    meta.DisplayName = "绘制形状"
                    meta.Keywords = New List(Of String)({"形状", "卡片", "圆角", "色块", "横幅", "分隔条", "看板卡片", "图形"})
                    meta.Category = CapabilityCategory.Format
                Case "excel_display_html_in_browser"
                    meta.DisplayName = "浏览器展示"
                    meta.Keywords = New List(Of String)({"浏览器展示", "HTML展示", "浏览器查看", "HTML表格"})
                Case "excel_autofit"
                    meta.DisplayName = "自动调整列宽行高"
                    meta.Keywords = New List(Of String)({"自动调整", "列宽", "行高", "自适应"})
                Case "excel_merge_cells"
                    meta.DisplayName = "合并单元格"
                    meta.Keywords = New List(Of String)({"合并单元格"})
                Case "excel_unmerge_cells"
                    meta.DisplayName = "取消合并单元格"
                    meta.Keywords = New List(Of String)({"取消合并", "拆分单元格"})
                Case "excel_run_python_script"
                    meta.DisplayName = "执行Python脚本"
                    meta.Keywords = New List(Of String)({"Python", "python脚本", "数据分析脚本"})
                Case "excel_fill_data"
                    meta.DisplayName = "智能填充数据"
                    meta.Keywords = New List(Of String)({"填充", "智能填充", "自动填充", "填入"})
                Case "excel_ocr_recognize"
                    meta.DisplayName = "OCR识别"
                    meta.Keywords = New List(Of String)({"OCR", "识别图片", "图片文字"})
                Case "excel_read_file"
                    meta.DisplayName = "读取文件"
                    meta.Keywords = New List(Of String)({"读取文件", "打开文件", "查看文件"})
                Case "excel_list_files"
                    meta.DisplayName = "列出文件"
                    meta.Keywords = New List(Of String)({"列出文件", "文件列表", "目录"})
                Case "excel_browse_files"
                    meta.DisplayName = "选择文件"
                    meta.Keywords = New List(Of String)({"选择文件", "浏览文件", "打开文件"})
                Case "excel_take_screenshot"
                    meta.DisplayName = "截图"
                    meta.Keywords = New List(Of String)({"截图", "截屏", "预览"})
                Case "excel_save_workbook"
                    meta.DisplayName = "保存工作簿"
                    meta.Keywords = New List(Of String)({"保存", "另存为"})
                Case "excel_set_cell_value"
                    meta.DisplayName = "设置单元格值"
                    meta.Keywords = New List(Of String)({"设置单元格", "修改单元格"})
                Case "excel_get_cell_value"
                    meta.DisplayName = "获取单元格值"
                    meta.Keywords = New List(Of String)({"单元格值"})
                Case "excel_add_formula"
                    meta.DisplayName = "添加公式"
                    meta.Keywords = New List(Of String)({"公式", "函数", "SUM", "VLOOKUP", "IF"})
                Case "excel_insert_row"
                    meta.DisplayName = "插入行"
                    meta.Keywords = New List(Of String)({"插入行", "新增行"})
                Case "excel_insert_column"
                    meta.DisplayName = "插入列"
                    meta.Keywords = New List(Of String)({"插入列", "新增列"})
                Case "excel_delete_row"
                    meta.DisplayName = "删除行"
                    meta.Keywords = New List(Of String)({"删除行", "移除行"})
                Case "excel_delete_column"
                    meta.DisplayName = "删除列"
                    meta.Keywords = New List(Of String)({"删除列", "移除列"})
                Case "excel_copy_sheet"
                    meta.DisplayName = "复制工作表"
                    meta.Keywords = New List(Of String)({"复制工作表", "复制sheet", "备份工作表"})
                Case "excel_copy_range"
                    meta.DisplayName = "复制数据范围"
                    meta.Keywords = New List(Of String)({"复制数据", "复制范围", "复制区域", "拷贝数据"})
                    meta.Category = CapabilityCategory.Manage
                    meta.IsWrite = True
                Case "excel_rename_sheet"
                    meta.DisplayName = "重命名工作表"
                    meta.Keywords = New List(Of String)({"重命名", "改名"})
                Case "excel_delete_sheet"
                    meta.DisplayName = "删除工作表"
                    meta.Keywords = New List(Of String)({"删除工作表", "删除sheet"})
                Case "excel_protect_sheet"
                    meta.DisplayName = "保护工作表"
                    meta.Keywords = New List(Of String)({"保护", "锁定"})
                Case "excel_unprotect_sheet"
                    meta.DisplayName = "取消保护"
                    meta.Keywords = New List(Of String)({"取消保护", "解锁"})
                Case "excel_open_workbook"
                    meta.DisplayName = "打开工作簿"
                    meta.Keywords = New List(Of String)({"打开工作簿", "打开文件"})
                Case "excel_close_workbook"
                    meta.DisplayName = "关闭工作簿"
                    meta.Keywords = New List(Of String)({"关闭工作簿", "关闭文件"})
                Case "excel_get_used_range"
                    meta.DisplayName = "获取已用范围"
                    meta.Keywords = New List(Of String)({"已用范围", "数据范围", "usedrange"})
                Case "excel_check_errors"
                    meta.DisplayName = "检查表格错误"
                    meta.Keywords = New List(Of String)({"检查错误", "错误检查", "错误", "有没有错误", "包含错误", "#DIV/0", "#N/A", "#VALUE", "公式错误", "异常值", "校验"})
                Case "excel_execute_vba"
                    meta.DisplayName = "执行VBA代码"
                    meta.Keywords = New List(Of String)({"VBA", "宏", "vba代码"})
                Case "excel_read_vba_macros"
                    meta.DisplayName = "读取VBA宏"
                    meta.Keywords = New List(Of String)({"VBA宏", "宏代码", "查看宏"})
                Case "excel_clipboard_read"
                    meta.DisplayName = "读取剪贴板"
                    meta.Keywords = New List(Of String)({"剪贴板", "粘贴板"})
                Case "excel_clipboard_write"
                    meta.DisplayName = "写入剪贴板"
                    meta.Keywords = New List(Of String)({"复制到剪贴板", "写入剪贴板"})
                Case "excel_batch_execute"
                    meta.DisplayName = "批量操作"
                    meta.Keywords = New List(Of String)({"批量", "批量执行"})
                Case "excel_delete_chart"
                    meta.DisplayName = "删除图表"
                    meta.Keywords = New List(Of String)({"删除图表"})
                Case "excel_update_chart_data"
                    meta.DisplayName = "更新图表数据"
                    meta.Keywords = New List(Of String)({"更新图表", "修改图表数据"})
                Case "excel_set_chart_title"
                    meta.DisplayName = "设置图表标题"
                    meta.Keywords = New List(Of String)({"图表标题"})
                Case "excel_list_charts"
                    meta.DisplayName = "列出所有图表"
                    meta.Keywords = New List(Of String)({"图表列表", "查看图表", "有哪些图表", "图表名称"})
                Case "excel_set_chart_properties"
                    meta.DisplayName = "设置图表属性"
                    meta.Keywords = New List(Of String)({"图表属性", "改图表", "调整图表", "图表美化", "趋势线", "图例", "坐标轴标题", "数据标签"})
                Case "excel_export_chart_image"
                    meta.DisplayName = "导出图表图片"
                    meta.Keywords = New List(Of String)({"图表导出", "保存图片", "导出图片", "生成图片"})
                Case "excel_set_pivot_field_layout"
                    meta.DisplayName = "配置透视表字段"
                    meta.Keywords = New List(Of String)({"透视表字段", "调整透视表", "改透视表", "透视表布局", "换数据源", "透视表行字段", "透视表值字段"})
                    meta.Category = CapabilityCategory.Analyze
                    ' 修改透视表布局属于写操作
                    meta.IsWrite = True
                Case "excel_refresh_pivot_table"
                    meta.DisplayName = "刷新透视表"
                    meta.Keywords = New List(Of String)({"刷新透视表", "更新透视表"})
                    meta.Category = CapabilityCategory.Analyze
                Case "excel_get_layout_snapshot"
                    meta.DisplayName = "布局快照"
                    meta.Keywords = New List(Of String)({"布局", "快照", "看板布局", "已放什么", "还能放", "图表位置", "已占用区域", "layout", "snapshot"})
            End Select
        End Sub

    End Class

End Namespace