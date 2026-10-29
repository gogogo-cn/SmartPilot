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

        ''' <summary>Layer3 系统与扩展 工具定义</summary>
        Private Shared Sub BuildSystemTools(tools As List(Of McpTool))
            ' ========== <Layer>3 - 系统与扩展</Layer> ==========

            ' [L3] 41. excel_run_python_script - 执行 Python 脚本
            tools.Add(New McpTool("excel_run_python_script", "执行 Python 脚本进行数据分析。已预置 pandas(pd)、numpy(np)、matplotlib.pyplot(plt) 无需手动 import。调用 plt.show() 会自动保存图表为 PNG 并返回路径。脚本中可使用 input_data 变量获取输入数据，将结果赋值给 output 变量返回。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"script", New McpToolProperty("string", "Python 脚本内容")},
                    {"timeout", New McpToolProperty("integer", "执行超时时间（秒），默认 30")}
                },
                .Required = New List(Of String) From {"script"}
            }))

            ' [L3] 42. excel_take_screenshot - 截取 Excel 区域截图
            tools.Add(New McpTool("excel_take_screenshot", "截取 Excel 区域截图，用于生成操作预览或异常排查", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"range", New McpToolProperty("string", "要截图的范围，如 A1:D20（可选，不指定则截取当前可见区域）")},
                    {"outputPath", New McpToolProperty("string", "截图保存路径（可选，不指定则保存到临时目录）")}
                }
            }))

            ' [L3] 42b. excel_display_html_in_browser - 在系统默认浏览器中打开 HTML 内容
            tools.Add(New McpTool("excel_display_html_in_browser", "在系统默认浏览器中展示内容。【两种用法】1)不传htmlContent参数：自动读取当前选中区域（或最近read_range的选区）数据，生成美观的表格HTML并在浏览器中打开——用于'用HTML列出/展示/浏览器查看数据'等场景；2)传htmlContent参数：直接展示自定义HTML内容。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"htmlContent", New McpToolProperty("string", "自定义 HTML 内容（可选，不传则自动读取选区数据生成表格HTML）")},
                    {"title", New McpToolProperty("string", "页面标题，也用于生成文件名（可选，默认：选中数据报表）")},
                    {"range", New McpToolProperty("string", "数据范围地址，如A1:H10（可选，仅在不传htmlContent时生效；不传则自动使用上下文选区/实时Selection）")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，仅在不传htmlContent时生效，默认当前工作表）")}
                },
                .Required = New List(Of String)()
            }))

            ' [L3] 43. excel_clipboard_read - 读取剪贴板内容
            tools.Add(New McpTool("excel_clipboard_read", "读取剪贴板内容", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {}
            }))

            ' [L3] 44. excel_clipboard_write - 写入内容到剪贴板
            tools.Add(New McpTool("excel_clipboard_write", "写入内容到剪贴板", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"content", New McpToolProperty("string", "要写入剪贴板的内容")}
                },
                .Required = New List(Of String) From {"content"}
            }))

            ' [L3] 45. excel_aggregate - Excel 原生聚合引擎（分组统计，仅返回结果）
            tools.Add(New McpTool("excel_aggregate", "按指定维度列分组聚合数值列，仅返回聚合结果数据（不做画图、不写回工作表）。聚合方式支持 sum/avg/count/max/min。需画图请先调用本工具取得结果后再用图表类工具；需落盘请再用 excel_write_range。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"dimColumn", New McpToolProperty("string", "【可选】分组维度列名（如：部门、地区、产品名称）。留空时自动选择典型维度列（优先级：部门>地区>产品>渠道>分类）")},
                    {"valueColumn", New McpToolProperty("string", "聚合数值列名（如：业绩、销售额、数量）。为空时自动推断第一个数值列")},
                    {"aggFunc", New McpToolProperty("string", "聚合函数：sum（默认）、avg、count、max、min")},
                    {"sheet", New McpToolProperty("string", "源数据工作表名（可选，默认当前活动工作表）")},
                    {"range", New McpToolProperty("string", "源数据范围（可选，默认当前选中区域）")},
                    {"targetSheet", New McpToolProperty("string", "结果落地表名（可选）。传入时把聚合结果写入该表，不存在则新建；留空则仅返回数据，不写回 Excel")},
                    {"skipEmptyRows", New McpToolProperty("boolean", "是否跳过全空行（默认 false，true 时自动过滤全部单元格均为空的行）")}
                },
                .Required = New List(Of String) From {}
            }))

            ' [L3] 45b. excel_date_group - 日期时间分组（Compute + 可选落地）
            tools.Add(New McpTool("excel_date_group", "将日期列按指定粒度（year/quarter/month/week/day）分组计算，默认仅返回分组结果数据，不写回 Excel。传入 aggColumn 时按组对数值列求和并输出汇总表；不传则给每一行追加一个分组列。支持财年（fiscalYearStart）与周起始日（weekStart）自定义。需落地请传入 targetSheet。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"dateColumn", New McpToolProperty("string", "日期列的列名（如：日期、创建时间、订单日期）")},
                    {"groupBy", New McpToolProperty("string", "分组粒度：year（按年）、quarter（按季度）、month（按月）、week（按周）、day（按日）")},
                    {"weekStart", New McpToolProperty("string", "【week 粒度专用】周起始日：monday（周一，默认）或 sunday（周日）")},
                    {"fiscalYearStart", New McpToolProperty("integer", "财年起始月份（1-12），默认 1（自然年）。如财年从 4 月开始则传 4")},
                    {"fiscalQuarter", New McpToolProperty("boolean", "是否使用财年季度。true 时按 fiscalYearStart 定义的财年起始计算季度")},
                    {"aggColumn", New McpToolProperty("string", "【可选】要按组聚合的数值列名。提供时输出按时间粒度求和后的汇总表；不填则仅给每行追加分组列")},
                    {"outputColumn", New McpToolProperty("string", "分组结果列名（可选，留空默认'年份'/'季度'/'月份'/'周'/'日期'）")},
                    {"sheet", New McpToolProperty("string", "源数据工作表名（可选，默认当前活动工作表）")},
                    {"range", New McpToolProperty("string", "源数据范围（可选，默认当前选中区域）")},
                    {"targetSheet", New McpToolProperty("string", "结果落地表名（可选）。传入时把分组结果写入该表，不存在则新建；留空则仅返回数据，不写回 Excel")}
                },
                .Required = New List(Of String) From {"dateColumn", "groupBy"}
            }))

            ' [L3] 46. excel_remove_duplicates - 选区/指定范围 去重（仅返回去重后数据，不回写）
            tools.Add(New McpTool("excel_remove_duplicates",
                "读取范围并按指定列去重，仅返回去重后的唯一数据（二维数组，含表头），不写回 Excel、不高亮、不删除。columns 不传时默认按范围内【全部列】做完全重复匹配；按部分列时传 columns 数组，如 ['A','B']（列字母）、[1,2]（列序号）、或 ['客户','产品']（表头列名）——三种格式都接受，多列表示这些列值全部相同才视为重复。hasHeaderRow 不传则自动检测首行是否为表头。需真正落盘请再用 excel_write_range。",
                New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"columns", New McpToolProperty("array", "去重依据列数组（可多列）：传列字母=['A','B']、1基列序号=[1,2]、或表头列名=['客户','产品'] 皆可；多列表示这些列的值全部相同才视为重复；不传则默认【全部列】完全匹配")},
                    {"hasHeaderRow", New McpToolProperty("boolean", "范围首行是否为表头（True=视为表头不参与去重；False=纯数据无表头）。不传则按首行样式+内容自动判断。")},
                    {"sheet", New McpToolProperty("string", "工作表名（可选，默认当前活动工作表）")},
                    {"range", New McpToolProperty("string", "去重范围（可选，默认当前选中区域 / 最终回退 UsedRange）")},
                    {"targetSheet", New McpToolProperty("string", "结果落地表名（可选）。传入时把去重结果写入该表，不存在则新建；留空则仅返回数据，不写回 Excel")}
                },
                .Required = New List(Of String) From {}
            }))


            ' [L3] 47. excel_clean_data - 选区/指定范围 智能数据清洗（默认零副作用）
            tools.Add(New McpTool("excel_clean_data",
                "对选中区域或指定范围执行智能数据清洗（原生 Excel 引擎，无需 VBA）。" &
                "【默认行为】dryRun=true，仅返回清洗后的数据预览，不修改 Excel。" &
                "三种使用模式：" &
                "A) 纯预览（dryRun=true，默认）→ 返回 cleanedRows/removedRows 数量和清洗后的数据，Excel 不动；" &
                "B) 零副作用落盘（targetSheet='清洗结果'）→ 结果写入新表，源表完全不变；" &
                "C) 就地修改（writeBack=true + autoSave=true 显式开启）→ 源表就地改值/删空行，autoSave=true 时自动 Save；否则工作簿 Dirty 需手动保存。" &
                "operations 必填：trimSpaces(去首尾空格)/fillEmpty(填充空值)/normalizeDates(统一为yyyy-mm-dd)/normalizeNumbers(统一数字格式)/standardizeText(文本规范化)/removeEmptyRows(删全空行)/removeDuplicates(去重)。" &
                "fillMethod：zero/delete/ffill(前值填充，默认)/placeholder(填fillValue)。" &
                "range 不传取当前选区，hasHeaderRow 默认 true。",
                New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"range", New McpToolProperty("string", "数据范围（可选，默认当前选中区域 / 回退 UsedRange）")},
                    {"sheet", New McpToolProperty("string", "工作表名（可选，默认当前活动工作表）")},
                    {"hasHeaderRow", New McpToolProperty("boolean", "首行是否为表头（默认 true，表头不参与清洗）")},
                    {"operations", New McpToolProperty("array", "清洗操作列表（必填）：trimSpaces/fillEmpty/normalizeDates/normalizeNumbers/standardizeText/removeEmptyRows/removeDuplicates")},
                    {"fillMethod", New McpToolProperty("string", "空值填充方式：zero/delete/ffill/placeholder（默认 ffill）")},
                    {"fillValue", New McpToolProperty("string", "占位文本（fillMethod=placeholder 时，默认'待补充'）")},
                    {"columns", New McpToolProperty("array", "去重依据列（removeDuplicates 时有效）：列字母=['A','B']、1基列序号=[1,2]、或表头列名=['客户','产品']；不传则按全部列整行匹配")},
                    {"targetSheet", New McpToolProperty("string", "结果落地表名（可选）。传入时把清洗结果写入该表，不存在则新建；不影响源数据（零副作用落盘）")},
                    {"dryRun", New McpToolProperty("boolean", "仅预览不修改（默认 true）。关闭后才会对源表产生副作用，需配合 writeBack/autoSave 使用")},
                    {"writeBack", New McpToolProperty("boolean", "就地写回源表（默认 false）。仅在 dryRun=false 时生效。写入后工作簿被标记 Dirty")},
                    {"autoSave", New McpToolProperty("boolean", "写回后自动保存（默认 false）。仅在 writeBack=true 时生效")}
                },
                .Required = New List(Of String) From {"operations"}
            }))

        End Sub

    End Class

End Namespace