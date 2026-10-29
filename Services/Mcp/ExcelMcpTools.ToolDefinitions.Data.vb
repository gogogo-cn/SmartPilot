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

        ''' <summary>Layer2 数据操作 工具定义</summary>
        Private Shared Sub BuildDataTools(tools As List(Of McpTool))
            ' ========== <Layer>2 - 数据操作</Layer> ==========

            ' [L2] 6. excel_search
            tools.Add(New McpTool("excel_search", "在工作表中搜索指定内容", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"query", New McpToolProperty("string", "搜索关键词")},
                    {"sheet", New McpToolProperty("string", "要搜索的工作表名称（不指定则搜所有工作表）")},
                    {"range", New McpToolProperty("string", "可选：搜索范围，如 A1:C100")}
                },
                .Required = New List(Of String) From {"query"}
            }))

            ' [L3] 7. excel_create_chart
            tools.Add(New McpTool("excel_create_chart", "根据数据创建图表（支持 20 种常用+高级类型：含堆叠/100%/折线带点/圆环/雷达/统计/销售转化/现代层级。自动错开位置，支持样式/图例/轴/数据标签/趋势线/组合次轴）", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"chartType", New McpToolProperty("string", "【auto=自动选择（推荐：用户未明确图表类型时传 auto，引擎按数据特征自动选型）·中文对应英文枚举】最常用→xlColumnClustered=柱状图/柱形图/分组柱/对比图、xlPie=饼图/扇形图/占比图(≤6类)、xlLineMarkers=折线图/折线带点/趋势图、xlColumnStacked=堆叠柱/累积柱、xlColumnStacked100=100%堆叠柱/百分比柱、xlBarClustered=条形图/横向柱状图、xlDoughnut=圆环图/环形图(>6类)、xlArea=面积图、xlXYScatter=散点图、xlRadar=雷达图、xlBubble=气泡图、xlWaterfall=瀑布图、xlFunnel=漏斗图、xlBoxwhisker=箱线图、xlTreemap=树状图、xlSunburst=旭日图、xlHistogram=直方图。未知类型将返回错误，不再静默降级。", New List(Of String) From {
                        "xlColumnClustered", "xlColumnStacked", "xlColumnStacked100",
                        "xlBarClustered", "xlBarStacked", "xlBarStacked100",
                        "xlLine", "xlLineMarkers",
                        "xlArea",
                        "xlPie", "xlDoughnut",
                        "xlXYScatter",
                        "xlRadar", "xlRadarFilled",
                        "xlBubble",
                        "xlSunburst", "xlTreemap",
                        "xlHistogram", "xlBoxwhisker",
                        "xlFunnel", "xlWaterfall"
                    })},
                    {"dataRange", New McpToolProperty("string", "【必填·强烈建议显式写绝对范围】数据范围（推荐带表头首行），绝对引用或工作表引用都可，如 A1:H6、Sheet1!A1:B12、分类汇总表!A1:E20。时间序列必须包含日期列，漏斗/瀑布必须按步骤顺序排列。")},
                    {"title", New McpToolProperty("string", "【强烈建议显式传递】图表标题（用于报告/老板一眼识别图含义），建议包含分析维度与核心指标，如『2025 1-6月 各区域销售额对比』、『用户转化漏斗』。")},
                    {"sheet", New McpToolProperty("string", "【推荐传递】放置图表的工作表名称；不传则默认当前活动工作表。建议与 dataRange 所在工作表保持一致（避免跨 Sheet 引用）。")},
                    {"position", New McpToolProperty("string", "【可选】图表位置与大小，JSON：{left, top, width, height}；单位=磅（默认 Excel 标准）。也支持按单元格区域锚定：{anchorRange:""B5:G18""}，引擎自动换算磅坐标（看板布局推荐）。留空=自动在数据区下方或现有图表之后错开排列（避免多张图重叠），适合一次性批量建图。")},
                    {"seriesBy", New McpToolProperty("string", "【推荐传递】数据系列方向：columns（按列分组，每列=1系列，默认）/ rows（按行分组，每行=1系列）。多指标表头在行首时传 rows。")},
                    {"switchRowColumn", New McpToolProperty("boolean", "【可选】等价于 Excel『切换行/列』按钮：True 直接反转数据方向（覆盖 seriesBy）。用于图做出来后系列错位、分类/系列对调时快速修正。")},
                    {"hasLegend", New McpToolProperty("boolean", "【强烈建议显式传递】是否显示图例（默认 True）。饼图/圆环/漏斗/瀑布/树状/旭日：建议 False（用数据标签标注类别名更清楚）。")},
                    {"legendPosition", New McpToolProperty("string", "【推荐传递】图例位置：xlLegendPositionBottom（默认，分组柱/折线）/ xlLegendPositionRight（条/散点/气泡）/ xlLegendPositionTop / xlLegendPositionLeft / xlLegendPositionCorner")},
                    {"xAxisTitle", New McpToolProperty("string", "【强烈建议显式传递】横坐标轴（分类轴）标题，如『月份』、『产品类别』、『分数区间』。让看的人立刻知道 X 轴在表达什么。")},
                    {"yAxisTitle", New McpToolProperty("string", "【强烈建议显式传递】纵坐标轴（数值轴）标题，如『销售额(万元)』、『数量(个)』、『占比(%)』。组合图+次轴时，此处标题默认主轴；次轴标题请在 set_chart_properties 再补充。")},
                    {"showDataLabels", New McpToolProperty("boolean", "【强烈建议显式传递】是否显示数据标签：饼图/圆环/漏斗/瀑布/100%堆叠 强烈建议 True（否则数值丢失）；时间序列折线/散点/气泡 图点多时建议 False（避免视觉拥堵）。")},
                    {"dataLabelPosition", New McpToolProperty("string", "【推荐传递】数据标签位置：xlLabelPositionOutsideEnd（柱/条默认，防止遮挡条体）/ xlLabelPositionCenter（居中）/ xlLabelPositionInsideEnd（内靠尾端）/ xlLabelPositionInsideBase（内靠基线）/ xlLabelPositionBestFit（饼图/圆环/漏斗推荐，自动避免重叠）/ xlLabelPositionAbove（折线数据点上方）")},
                    {"chartStyle", New McpToolProperty("integer", "【强烈建议显式传递】Excel 内置图表主题样式 1~208：201=Office 2016 默认商务蓝(推荐) / 202=多色活泼 / 203=单色灰 / 204=单色橙 / 205=单色绿 / 206=简约浅灰。传了就直接应用，不做任何二次美化。")},
                    {"useSecondaryAxisLastSeries", New McpToolProperty("boolean", "【组合图专用】把最后一个系列放到**次坐标轴**，一步到位制作 柱(主)+折(次) 组合图。常见用法：2系列柱(销量+同比%) → 最后系列(同比%)放次轴（范围 0-100% 不再被销量压缩）。需要精细控制（指定任意系列到次轴/取消次轴）请在 set_chart_properties 用 secondaryAxisSeriesIndex。")},
                    {"dataValues", New McpToolProperty("string", "可选内存数组模式。传二维数组JSON字符串：首行为表头其余为明细。提供时作为dataRange的内存替代，直接内存绑定系列、不落地工作表，第1列视为分类其余数值列各为一个系列。与dataRange二选一，内存模式可省略dataRange。")}
                },
                .Required = New List(Of String) From {"chartType"}
            }))

            ' [L2] 8. excel_format_range
            tools.Add(New McpTool("excel_format_range", "格式化指定范围的单元格", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称")},
                    {"range", New McpToolProperty("string", "单元格范围，如 A1:D10")},
                    {"autoRange", New McpToolProperty("boolean", "【可选】未传 range 时，是否自动使用指定工作表（未指定则活动工作表）的已用区域（UsedRange）推断整表。默认 false：range 为空时回退当前界面选区，但选区所在表与 sheet 不一致时会弃用并明确报错，不会静默格式化到别的表")},
                    {"format", New McpToolProperty("string", "格式选项 JSON：{bold, italic, fontSize, fontName, fontColor, headerFontColor, bgColor, bgScope, numberFormat, borderStyle, borderColor, rowHeight, columnWidth}。bgScope=header 时底色只上表头。headerFontColor 只改表头字色，不要用 fontColor 把整段涂白。rowHeight/columnWidth 用于固定行高/列宽（看板 KPI 卡大字号标题撑高行、固定卡片列宽）")},
                    {"conditionalFormat", New McpToolProperty("string", "条件格式 JSON：{type:'dataBar'|'colorScale'|'iconSet', ...}")},
                    {"alignment", New McpToolProperty("string", "对齐 JSON：{horizontal:'center', vertical:'middle', wrapText:true}")}
                },
                .Required = New List(Of String) From {"range", "format"}
            }))

            ' [L1] 9. excel_get_workbook_info
            tools.Add(New McpTool("excel_get_workbook_info", "获取工作簿的基本信息（名称、路径、工作表数等）", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "可选：Excel 文件完整路径。不指定则使用当前活动工作簿")}
                }
            }))

            ' [L1] 9a. excel_get_open_workbooks
            tools.Add(New McpTool("excel_get_open_workbooks", "获取当前 Excel 中所有打开的工作簿列表（名称、路径、工作表数、哪个是活动工作簿）", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {}
            }))

            ' [L3] 10. excel_fill_data
            tools.Add(New McpTool("excel_fill_data", "智能填入数据到 Excel，自动识别数据类型并格式化", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"data", New McpToolProperty("string", "要填入的数据，JSON 数组格式")},
                    {"dataType", New McpToolProperty("string", "数据类型标识（可选）：invoice, form, table, document, general")}
                },
                .Required = New List(Of String) From {"data"}
            }))

            ' [L3] 11. excel_delete_sheet
            tools.Add(New McpTool("excel_delete_sheet", "删除指定工作表", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"name", New McpToolProperty("string", "要删除的工作表名称")}
                },
                .Required = New List(Of String) From {"name"}
            }))

            ' [L3] 12. excel_rename_sheet
            tools.Add(New McpTool("excel_rename_sheet", "重命名工作表", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"oldName", New McpToolProperty("string", "当前工作表名称")},
                    {"newName", New McpToolProperty("string", "新工作表名称")}
                },
                .Required = New List(Of String) From {"oldName", "newName"}
            }))

            ' [L3] 13. excel_ocr_recognize
            tools.Add(New McpTool("excel_ocr_recognize", "使用本地 OCR 引擎识别图片中的文字（支持 PaddleOCR / RapidOCR / CloudOCR）。需要提供图片路径或 Base64 数据。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"imagePath", New McpToolProperty("string", "图片文件路径（如 C:\temp\screenshot.png），与 imageData 二选一")},
                    {"imageData", New McpToolProperty("string", "Base64 编码的图片数据（可选，与 imagePath 二选一）")},
                    {"engine", New McpToolProperty("string", "指定 OCR 引擎：paddleocr / rapidocr / cloudocr（不指定则使用默认引擎）")},
                    {"language", New McpToolProperty("string", "识别语言：ch（中文，默认）、en（英文），仅对 PaddleOCR 生效")}
                },
                .Required = New List(Of String) From {"imagePath"}
            }))

            ' [L3] 14. excel_read_file
            tools.Add(New McpTool("excel_read_file", "读取本地文件内容。必须指定完整文件路径。支持 DOCX、PDF、TXT、XLSX、CSV、MD、JSON、XML 等格式。注意：只能读取单个文件，不能读取目录。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "文件完整绝对路径（如 C:\Desktop\WPS\文档.docx），必须包含文件名和扩展名")},
                    {"maxChars", New McpToolProperty("integer", "返回的最大字符数（默认 50000，超过则截断）")}
                },
                .Required = New List(Of String) From {"filePath"}
            }))

            ' [L3] 15. excel_list_files
            tools.Add(New McpTool("excel_list_files", "列出指定目录下的所有文件和子目录。当用户请求读取文件夹或目录时，应先调用此工具获取文件列表。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"directoryPath", New McpToolProperty("string", "目录完整绝对路径（如 C:\Desktop\WPS），不要包含文件名")}
                },
                .Required = New List(Of String) From {"directoryPath"}
            }))

            ' [L3] 16. excel_browse_files - 弹出文件选择对话框，让用户选择文件并返回路径
            tools.Add(New McpTool("excel_browse_files", "弹出文件选择对话框让用户选择文件，返回所选文件的完整路径。当需要文件路径但不知道具体路径时，应先调用此工具让用户选择文件，再用返回的路径调用 excel_read_file 读取内容。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filter", New McpToolProperty("string", "文件类型过滤器（可选）。常用值：all=所有文件, excel=Excel文件(*.xlsx;*.xls), doc=文档(*.docx;*.pdf), image=图片(*.jpg;*.png), csv=CSV文件。默认 all")}
                }
            }))

            ' [L3] 16a. excel_copy_sheet - 复制工作表
            tools.Add(New McpTool("excel_copy_sheet", "复制指定工作表", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"source", New McpToolProperty("string", "源工作表名称")},
                    {"target", New McpToolProperty("string", "新工作表名称（可选）")}
                },
                .Required = New List(Of String) From {"source"}
            }))

            ' [L3] 16a-2. excel_copy_range - 复制指定范围的数据到目标工作表
            tools.Add(New McpTool("excel_copy_range", "复制指定范围的数据到目标工作表（源数据保持不变）。用于将选中区域或指定范围的数据复制到另一张工作表。如果目标工作表已有数据，粘贴会覆盖对应区域。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sourceSheet", New McpToolProperty("string", "源工作表名称")},
                    {"sourceRange", New McpToolProperty("string", "源数据范围，如 A1:D100（可选，不传则默认使用源工作表的已用区域 UsedRange）")},
                    {"destinationSheet", New McpToolProperty("string", "目标工作表名称")},
                    {"destinationRange", New McpToolProperty("string", "目标区域左上角单元格，如 A1（可选，默认 A1）")}
                },
                .Required = New List(Of String) From {"sourceSheet", "destinationSheet"}
            }))

            ' [L2] 16b. excel_clear_range - 清除范围内容（默认清除当前选中区域）
            tools.Add(New McpTool("excel_clear_range", "清除指定范围内单元格的内容/数据。按用户意图设置清除粒度：默认仅清内容（clearContent=true, clearFormat=false）；要求'全部清除'（内容+格式）→传 clearContent=true, clearFormat=true。若用户只想清格式而不动内容，应使用excel_clear_format（本工具清格式只是辅助清数据的兜底，不鼓励单独清格式）。range 可选，不传则后端自动选选区（优先最近read_range选区，其次实时选中区域）。⚠️仅当用户明确说'清空整个工作表/清空整个Sheet'才用excel_clear_sheet。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")},
                    {"range", New McpToolProperty("string", "单元格范围，如 A1:D10（可选，不传则后端自动选择选区）")},
                    {"clearContent", New McpToolProperty("boolean", "是否清除单元格内容（默认 true；仅清除格式时传 false）")},
                    {"clearFormat", New McpToolProperty("boolean", "是否清除单元格格式（默认 false；要求内容+格式一并清除时传 true）")}
                },
                .Required = New List(Of String)()
            }))

            ' [L2] 16b-2. excel_clear_format - 仅清除单元格格式（独立于清数据，便于AI多步骤规划）
            tools.Add(New McpTool("excel_clear_format", "仅清除指定范围内的单元格格式（字体、颜色、边框、对齐、填充等），保留单元格内容不变。适合用户说'清空格式'/'去掉格式'/'清除格式'/'取消格式'时使用。若用户要清空的是内容/数据，应使用excel_clear_range；若两者都要，建议按步骤分别调用。range 可选，不传则后端自动选选区（优先最近read_range选区，其次实时选中区域）。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")},
                    {"range", New McpToolProperty("string", "单元格范围，如 A1:D10（可选，不传则后端自动选择选区）")}
                },
                .Required = New List(Of String)()
            }))

            ' [L2] 16c. excel_clear_sheet - 清空整个工作表
            tools.Add(New McpTool("excel_clear_sheet", "清空指定工作表的所有内容和格式（保留工作表本身，不删除工作表）。⚠️危险操作：此工具会清除整个工作表的所有数据！仅当用户明确说'清空整个工作表'/'清空整个Sheet'/'清除所有数据'时才使用。如果用户说'清空数据'/'清空选中'/'清除内容'，应该使用excel_clear_range。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")}
                },
                .Required = New List(Of String)()
            }))

            ' [L1] 17. excel_set_cell_value - 设置单个单元格的值
            tools.Add(New McpTool("excel_set_cell_value", "设置指定单元格的值", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")},
                    {"cell", New McpToolProperty("string", "单元格地址，如 A1")},
                    {"value", New McpToolProperty("string", "要设置的值")}
                },
                .Required = New List(Of String) From {"cell", "value"}
            }))

            ' [L1] 18. excel_get_cell_value - 获取单个单元格的值
            tools.Add(New McpTool("excel_get_cell_value", "获取指定单元格的值", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")},
                    {"cell", New McpToolProperty("string", "单元格地址，如 A1")}
                },
                .Required = New List(Of String) From {"cell"}
            }))

            ' [L2] 19. excel_add_formula - 在单元格中添加公式
            tools.Add(New McpTool("excel_add_formula", "在指定单元格中添加公式", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")},
                    {"cell", New McpToolProperty("string", "单元格地址，如 A1")},
                    {"formula", New McpToolProperty("string", "公式内容，如 =SUM(A1:A10)")}
                },
                .Required = New List(Of String) From {"cell", "formula"}
            }))

            ' [L2] 20. excel_insert_row - 在指定位置插入行
            tools.Add(New McpTool("excel_insert_row", "在指定位置插入指定数量的行", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")},
                    {"row", New McpToolProperty("integer", "起始行号（从1开始）")},
                    {"count", New McpToolProperty("integer", "要插入的行数，默认1")}
                },
                .Required = New List(Of String) From {"row"}
            }))

            ' [L2] 21. excel_insert_column - 在指定位置插入列
            tools.Add(New McpTool("excel_insert_column", "在指定位置插入指定数量的列", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")},
                    {"column", New McpToolProperty("integer", "起始列号（从1开始，1=A列）")},
                    {"count", New McpToolProperty("integer", "要插入的列数，默认1")}
                },
                .Required = New List(Of String) From {"column"}
            }))

            ' [L2] 22. excel_delete_row - 删除指定行
            tools.Add(New McpTool("excel_delete_row", "删除指定位置的行", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")},
                    {"row", New McpToolProperty("integer", "起始行号（从1开始）")},
                    {"count", New McpToolProperty("integer", "要删除的行数，默认1")}
                },
                .Required = New List(Of String) From {"row"}
            }))

            ' [L2] 23. excel_delete_column - 删除指定列
            tools.Add(New McpTool("excel_delete_column", "删除指定位置的列", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")},
                    {"column", New McpToolProperty("integer", "起始列号（从1开始，1=A列）")},
                    {"count", New McpToolProperty("integer", "要删除的列数，默认1")}
                },
                .Required = New List(Of String) From {"column"}
            }))

            ' [L2] 24. excel_merge_cells - 合并单元格
            tools.Add(New McpTool("excel_merge_cells", "合并指定范围的单元格。按用户意图设置合并粒度：只说'合并'→mergeType=merge（默认，仅合并）；说'合并后居中/合并居中'→mergeType=mergeAndCenter（合并后文字居中）。range 必填。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")},
                    {"range", New McpToolProperty("string", "要合并的单元格范围，如 A1:C3")},
                    {"mergeType", New McpToolProperty("string", "合并类型：merge（仅合并，默认）或 mergeAndCenter（合并后居中）")}
                },
                .Required = New List(Of String) From {"range"}
            }))

            ' [L2] 25. excel_unmerge_cells - 取消合并单元格
            tools.Add(New McpTool("excel_unmerge_cells", "取消合并指定范围的单元格。若用户希望取消合并后把每个原合并区域的值展开填充到整片区域（如'取消合并并填充/铺满/复制到所有单元格'），传 fillAfterUnmerge=true；默认 false 时仅取消合并、只保留左上角值。range 必填。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")},
                    {"range", New McpToolProperty("string", "要取消合并的单元格范围，如 A1:C3")},
                    {"fillAfterUnmerge", New McpToolProperty("boolean", "取消合并后是否把原合并区域左上角的值展开填充到整个区域（默认 false）")}
                },
                .Required = New List(Of String) From {"range"}
            }))

            ' [L2] 26. excel_autofit - 自动调整列宽或行高
            tools.Add(New McpTool("excel_autofit", "自动调整指定范围的列宽或行高", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")},
                    {"range", New McpToolProperty("string", "要调整的单元格范围，如 A1:C10（可选，不指定则调整已使用范围）")},
                    {"type", New McpToolProperty("string", "调整类型：column（列宽，默认）或 row（行高）")}
                }
            }))

            ' [L2] 27. excel_save_workbook - 保存工作簿
            tools.Add(New McpTool("excel_save_workbook", "保存当前工作簿。如果指定了 filePath 参数，则另存为到指定路径", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "可选：另存为的完整文件路径（如 C:\\Documents\\report.xlsx），不指定则保存到原路径")}
                }
            }))

            ' [L2] 28. excel_protect_sheet - 保护工作表
            tools.Add(New McpTool("excel_protect_sheet", "保护指定工作表，防止修改", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")},
                    {"password", New McpToolProperty("string", "可选：保护密码")}
                }
            }))

            ' [L2] 29. excel_unprotect_sheet - 取消保护工作表
            tools.Add(New McpTool("excel_unprotect_sheet", "取消指定工作表的保护", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")},
                    {"password", New McpToolProperty("string", "可选：保护密码（如果设置了密码）")}
                }
            }))

            ' [L3] 30. excel_split_workbook - 按工作表拆分工作簿（统一版）
            tools.Add(New McpTool("excel_split_workbook", "将工作簿按工作表拆分为多个独立的 Excel 文件。支持自定义输出目录、命名模板、文件格式。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"outputDir", New McpToolProperty("string", "输出文件夹路径，空则与源文件同目录")},
                    {"namingPattern", New McpToolProperty("string", "文件名模板：{original}=原文件名, {sheet}=工作表名, {index}=序号。默认'{original}_{sheet}'")},
                    {"fileFormat", New McpToolProperty("string", "输出格式：xlsx（默认）、xls、csv")},
                    {"includeHidden", New McpToolProperty("boolean", "是否包含隐藏工作表，默认 false")}
                }
            }))

            ' [L3] 30a. excel_open_workbook - 打开指定路径的 Excel 文件
            tools.Add(New McpTool("excel_open_workbook", "打开指定路径的 Excel 文件", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "Excel 文件完整路径")}
                },
                .Required = New List(Of String) From {"filePath"}
            }))

            ' [L3] 30b. excel_close_workbook - 关闭指定工作簿
            tools.Add(New McpTool("excel_close_workbook", "关闭指定工作簿", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "工作簿路径，空则关闭当前活动工作簿")},
                    {"saveChanges", New McpToolProperty("boolean", "是否保存更改，默认 false")}
                }
            }))

            ' [L1] 30c. excel_get_used_range - 获取工作表的已使用范围
            tools.Add(New McpTool("excel_get_used_range", "获取工作表的已使用范围地址和行列数", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")}
                }
            }))

            ' [L1] 30c+. excel_check_errors - 扫描工作表中的公式错误单元格
            tools.Add(New McpTool("excel_check_errors", "扫描工作表中的公式错误单元格（如 #DIV/0!、#N/A、#VALUE!、#REF! 等）。默认扫描当前活动工作表的已使用区域；可通过 sheet/range 参数指定范围。返回每个错误单元格的地址、错误类型和公式。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")},
                    {"range", New McpToolProperty("string", "【可选】要检查的范围，如 A1:F100；留空则扫描整张表的已使用区域")}
                }
            }))

            ' [L3] 30d. excel_batch_execute - 批量执行多个操作
            tools.Add(New McpTool("excel_batch_execute", "批量执行多个 Excel 操作，失败时自动回滚", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"operations", New McpToolProperty("string", "操作列表 JSON，每个含 toolName 和 arguments")}
                },
                .Required = New List(Of String) From {"operations"}
            }))

            ' [L2] 30e. excel_sort_range - 对指定范围排序（完整还原 Excel 原生「排序」对话框）
            tools.Add(New McpTool("excel_sort_range", "对指定范围排序，完整还原 Excel 原生「排序」对话框能力：① 多级排序依据/次要关键字（sortBy 数组按顺序生效）；② 排序依据支持「单元格值/单元格颜色/字体颜色/条件格式图标」（sortBy[].sortOn）；③ 次序升序/降序（sortBy[].order，颜色排序时 ascending=指定颜色排最前、descending=排最后）；④ 选项：区分大小写 matchCase、方向 orientation（按列/按行）、排序方法 sortMethod（拼音/笔划）。未指定 sortBy 时默认按第一列升序。按颜色排序必须同时给出 color（如 #FF0000）；按图标排序需给出 iconSet（图标集序号或名称如 3Arrows）与 icon（图标序号，从 1 开始）。排序成功后返回表头+前几行预览数据，可直接用于总结。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"range", New McpToolProperty("string", "要排序的范围，如 A1:D100")},
                    {"sortBy", New McpToolProperty("string", "排序列列表 JSON，如 [{""column"":""A"",""order"":""ascending""}]，支持多列（第一个=排序依据，其余=次要关键字），列可用字母(A/B/C)、数字(1/2/3，范围内相对列号)或中文列名(如""销售额"")。每项可带：sortOn=排序依据，values(默认，单元格值)|cellColor(单元格颜色)|fontColor(字体颜色)|icon(条件格式图标)；color=排序颜色，仅颜色排序需要，支持 #FF0000 / FF0000 / ""255,0,0"" / 中文颜色名（如""红色""）；iconSet=图标集（序号或名称如 3Arrows）与 icon=图标序号(1起)，仅图标排序需要。示例：[{""column"":""状态"",""sortOn"":""cellColor"",""color"":""#FF0000"",""order"":""ascending""}]（红色排最前）")},
                    {"hasHeader", New McpToolProperty("boolean", "是否有表头，默认 true")},
                    {"matchCase", New McpToolProperty("boolean", "【可选】区分大小写（原生对话框「选项」），默认 false")},
                    {"orientation", New McpToolProperty("string", "【可选】排序方向（原生对话框「选项」）：columns=按列排序（默认，沿垂直方向重排行）| rows=按行排序（沿水平方向重排列）")},
                    {"sortMethod", New McpToolProperty("string", "【可选】排序方法（原生对话框「选项」，仅对文本生效）：pinYin=按字母顺序（拼音，默认）| stroke=按笔划")},
                    {"previewRows", New McpToolProperty("integer", "【可选】排序成功后回传前几行预览数据（默认 5，最大 50），用于总结展示")}
                },
                .Required = New List(Of String) From {"range"}
            }))

            ' [L2] 30f. excel_apply_filter - 对指定范围应用自动筛选
            tools.Add(New McpTool("excel_apply_filter", "对指定范围应用自动筛选。条件键可用列字母(A/B)或表头名(订单金额/销售区域)，值支持比较运算符前缀(>3000 / <=100)。未传 range 时使用当前选区；传 autoRange=true 时使用整张已用区域。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为选定区域/当前活动工作表）")},
                    {"range", New McpToolProperty("string", "要筛选的范围，如 A1:D100（可选；留空则用当前选区，配合 autoRange 用整张已用区域）")},
                    {"autoRange", New McpToolProperty("boolean", "为 true 且未传 range 时，自动使用当前工作表的整张已用区域（UsedRange）")},
                    {"criteria", New McpToolProperty("string", "筛选条件 JSON。键可为列字母(A/B)或表头名(订单金额/区域)，值支持比较运算符前缀，如 {""订单金额"":"">3000"",""区域"":""华东""}")}
                },
                .Required = New List(Of String) From {"criteria"}
            }))

            ' [L2] 30g. excel_clear_filter - 清除指定工作表的筛选
            tools.Add(New McpTool("excel_clear_filter", "清除指定工作表的筛选", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")}
                }
            }))

            ' [L3] 31. excel_split_by_column - 按指定列的唯一值拆分数据
            tools.Add(New McpTool("excel_split_by_column", "按指定列的唯一值将数据拆分为多个工作表或工作簿", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"column", New McpToolProperty("string", "要拆分的列名或列号（如'A'或'部门'）")},
                    {"outputMode", New McpToolProperty("string", "输出方式：sheets（拆分到同一工作簿的新工作表）或 files（拆分到独立文件）")},
                    {"outputDir", New McpToolProperty("string", "files 模式时的输出目录")},
                    {"namingPattern", New McpToolProperty("string", "命名模板，{value}=列值, {original}=原文件名")}
                },
                .Required = New List(Of String) From {"column"}
            }))

            ' [L3] 32. excel_create_pivot_table - 创建数据透视表
            tools.Add(New McpTool("excel_create_pivot_table", "创建数据透视表", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sourceRange", New McpToolProperty("string", "源数据范围，如 Sheet1!A1:D100")},
                    {"targetSheet", New McpToolProperty("string", "放置透视表的工作表名称")},
                    {"targetCell", New McpToolProperty("string", "放置位置，如 A1")},
                    {"rowFields", New McpToolProperty("string", "行字段列表，JSON 数组")},
                    {"columnFields", New McpToolProperty("string", "列字段列表，JSON 数组")},
                    {"valueFields", New McpToolProperty("string", "值字段列表，JSON 数组")},
                    {"valueFunctions", New McpToolProperty("string", "聚合函数列表：Sum/Count/Average/Max/Min，JSON 数组")}
                },
                .Required = New List(Of String) From {"sourceRange"}
            }))

            ' [L3] 33. excel_merge_workbooks - 合并多个工作簿
            tools.Add(New McpTool("excel_merge_workbooks", "将多个 Excel 文件合并到一个工作簿，每个文件的工作表全部复制进来", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePaths", New McpToolProperty("string", "要合并的文件路径列表，JSON 数组")},
                    {"targetPath", New McpToolProperty("string", "合并后的保存路径，空则新建工作簿不保存")},
                    {"sheetPrefix", New McpToolProperty("string", "工作表名前缀，如'[文件1]'")}
                },
                .Required = New List(Of String) From {"filePaths"}
            }))

            ' [L3] 33b. excel_merge_sheets - 合并同一工作簿内多个工作表到一张表
            tools.Add(New McpTool("excel_merge_sheets", "将同一工作簿中多个工作表的数据按行合并到一起并返回合并结果。sourceSheets 不传时自动合并除目标表外的所有可见工作表；每个源表的表头只保留一次，数据行按源表顺序依次追加。默认仅返回合并数据（不写回 Excel）；传入 targetSheet 时才把结果写入该表", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sourceSheets", New McpToolProperty("string", "源工作表名列表，JSON 数组；为空则合并除目标表外的所有可见工作表")},
                    {"targetSheet", New McpToolProperty("string", "【可选】结果落地表名。传入时把合并后的数据写入该表（不存在则新建）；留空则仅返回合并数据，不写回 Excel")},
                    {"includeHeaders", New McpToolProperty("boolean", "是否保留表头，默认 true（仅保留第一张表的表头）")},
                    {"mergeAll", New McpToolProperty("boolean", "用户明确要求'合并所有Sheet'时为 true：忽略当前选中的工作表标签，强制合并全部可见表")}
                },
                .Required = New List(Of String) From {}
            }))

            ' [L3] 33b2. excel_insert_rows - 原子能力：向指定工作表插入数据行（供合并/追加等操作型任务分步调用）
            tools.Add(New McpTool("excel_insert_rows", "向指定工作表写入数据行（原子插入能力）。rows 为二维数组（JSON），写入从 startRow（1 基）开始；append=true 时自动定位到目标表末尾追加。返回实际起始行、写入行数与下一空行，便于分步连续插入", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "目标工作表名")},
                    {"rows", New McpToolProperty("string", "要写入的数据行二维数组，JSON，如 [[""a"",1],[""b"",2]]")},
                    {"startRow", New McpToolProperty("integer", "起始行号（1 基）；与 append 二选一，默认 1")},
                    {"append", New McpToolProperty("boolean", "是否追加到目标表末尾（自动取下一个空行），默认 true")},
                    {"createIfMissing", New McpToolProperty("boolean", "目标表不存在时是否自动新建，默认 true")}
                },
                .Required = New List(Of String) From {"sheet", "rows"}
            }))

            ' [L3] 33c. excel_split_by_rows - 按行数拆分
            tools.Add(New McpTool("excel_split_by_rows", "将当前工作表数据按指定行数拆分为多个工作表或文件。每份都会带上表头（includeHeaders=true 时）", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"rowsPerFile", New McpToolProperty("integer", "每个结果包含的数据行数（不含表头），默认 500")},
                    {"includeHeaders", New McpToolProperty("boolean", "每个结果是否都保留表头，默认 true")},
                    {"outputMode", New McpToolProperty("string", "输出方式：sheets（拆分到同一工作簿的新工作表，默认）或 files（拆分到独立文件）")},
                    {"outputDir", New McpToolProperty("string", "files 模式时的输出目录，默认源文件目录")},
                    {"namingPattern", New McpToolProperty("string", "命名模板，{i}=序号, {original}=原文件名，默认'{original}_第{i}份'")}
                },
                .Required = New List(Of String) From {"rowsPerFile"}
            }))

            ' [L3] 33d. excel_join_tables - 跨表关联（数据库风格 JOIN）
            tools.Add(New McpTool("excel_join_tables", "按关联键将两个工作表进行跨表关联（内/左/右/全连接，类似数据库 JOIN），仅返回关联结果数据，不回写工作表。左右表各需一行表头，关联键列可传表头名、列序号或列字母。需落盘请再用 excel_write_range。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"leftSheet", New McpToolProperty("string", "左表工作表名")},
                    {"rightSheet", New McpToolProperty("string", "右表工作表名")},
                    {"leftKeyColumns", New McpToolProperty("string", "左表关联键列（列名/列序号/列字母的 JSON 数组，如 [""ID""]）")},
                    {"rightKeyColumns", New McpToolProperty("string", "右表关联键列（列名/列序号/列字母的 JSON 数组，如 [""订单ID""]，与左表一一对应）")},
                    {"joinType", New McpToolProperty("string", "连接类型：inner（内连接，默认）/ left（左连接）/ right（右连接）/ full（全外连接）")},
                    {"includeHeaders", New McpToolProperty("boolean", "是否保留表头，默认 true")},
                    {"targetSheet", New McpToolProperty("string", "结果落地表名（可选）。传入时把关联结果写入该表，不存在则新建；留空则仅返回数据，不写回 Excel")}
                },
                .Required = New List(Of String) From {"leftSheet", "rightSheet", "leftKeyColumns", "rightKeyColumns"}
            }))

            ' [L3] 34. excel_import_csv - 导入 CSV 文件
            tools.Add(New McpTool("excel_import_csv", "导入 CSV 文件到当前工作簿", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "CSV 文件完整路径")},
                    {"delimiter", New McpToolProperty("string", "分隔符，默认逗号")},
                    {"encoding", New McpToolProperty("string", "编码：utf8（默认）/ gbk / ansi")},
                    {"targetSheet", New McpToolProperty("string", "目标工作表，空则新建")}
                },
                .Required = New List(Of String) From {"filePath"}
            }))

            ' [L3] 35. excel_export_csv - 导出为 CSV
            tools.Add(New McpTool("excel_export_csv", "将工作表导出为 CSV", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")},
                    {"filePath", New McpToolProperty("string", "【可选】CSV 保存路径。留空则自动保存到桌面，文件名为工作表名称+时间戳")},
                    {"delimiter", New McpToolProperty("string", "分隔符，默认逗号")},
                    {"includeHeader", New McpToolProperty("boolean", "是否包含表头，默认 true")}
                },
                .Required = New List(Of String) From {}
            }))

            ' [L3] 36. excel_delete_chart - 删除图表
            tools.Add(New McpTool("excel_delete_chart", "删除指定工作表上的图表", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")},
                    {"chartName", New McpToolProperty("string", "图表名称（可选，不指定则删除最后一个图表）")}
                }
            }))

            ' [L3] 37. excel_update_chart_data - 更新图表数据源
            tools.Add(New McpTool("excel_update_chart_data", "更新图表的数据源范围", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")},
                    {"chartName", New McpToolProperty("string", "图表名称（可选，不指定则更新最后一个图表）")},
                    {"newDataRange", New McpToolProperty("string", "新的数据范围，如 Sheet1!A1:B10")}
                },
                .Required = New List(Of String) From {"newDataRange"}
            }))

            ' [L3] 38. excel_set_chart_title - 设置图表标题
            tools.Add(New McpTool("excel_set_chart_title", "设置图表标题", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")},
                    {"chartName", New McpToolProperty("string", "图表名称（可选，不指定则设置最后一个图表的标题）")},
                    {"title", New McpToolProperty("string", "图表标题")}
                },
                .Required = New List(Of String) From {"title"}
            }))

            ' [L3] 38b. excel_list_charts - 列出工作表中所有图表（编辑/删除前先查名称，这是图表编辑的前提）
            tools.Add(New McpTool("excel_list_charts", "列出指定工作表上的所有图表（名称、类型、位置、大小、标题），修改或删除图表前必须先调用本工具取得准确的 chartName", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")}
                }
            }))

            ' [L3] 38c. excel_set_chart_properties - 统一设置图表属性（位置/大小/类型/样式/图例/轴/数据标签/切换行列/趋势线/组合次轴/分系列类型）
            tools.Add(New McpTool("excel_set_chart_properties", "统一设置已存在图表的属性（位置、大小、类型、样式、图例、坐标轴标题、数据标签、系列方向、趋势线、组合次轴、分系列类型），避免分别调用多个小工具。支持为任意系列单独指定类型+次轴以实现柱+折线的组合图。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")},
                    {"chartName", New McpToolProperty("string", "图表名称（可选，默认操作最后一个图表）；推荐先调用 excel_list_charts 获取")},
                    {"chartType", New McpToolProperty("string", "修改整个图表的类型：同 excel_create_chart（20 种）")},
                    {"seriesChartType", New McpToolProperty("string", "仅修改单个系列的自身类型（用于组合图）：同 excel_create_chart（20 种），例如把最后一个系列从柱状改为折线。需配合 seriesIndex 使用")},
                    {"seriesIndex", New McpToolProperty("integer", "seriesChartType 或 secondaryAxisSeriesIndex 作用的系列编号（从 1 开始）。留空且指定了 seriesChartType 时默认为『最后一个系列』")},
                    {"secondaryAxisSeriesIndex", New McpToolProperty("integer", "把指定系列放到**次坐标轴**（从 1 开始；留空=取最后一个系列；0 或负数=清空该系列的次轴设置，放回主轴）。典型：柱(主)+折(次) 看销量+同比")},
                    {"title", New McpToolProperty("string", "图表标题；传入空串可删除标题")},
                    {"position", New McpToolProperty("string", "位置与大小 JSON：{left,top,width,height}，缺某项不修改")},
                    {"seriesBy", New McpToolProperty("string", "数据系列方向：columns（按列分组）/ rows（按行分组）")},
                    {"switchRowColumn", New McpToolProperty("boolean", "True 反转行/列，等价于 Excel『切换行/列』按钮")},
                    {"chartStyle", New McpToolProperty("integer", "内置样式 1~208")},
                    {"hasLegend", New McpToolProperty("boolean", "是否显示图例")},
                    {"legendPosition", New McpToolProperty("string", "图例位置：xlLegendPositionRight / xlLegendPositionBottom / xlLegendPositionTop / xlLegendPositionLeft / xlLegendPositionCorner")},
                    {"xAxisTitle", New McpToolProperty("string", "横轴（分类轴）标题；空串=取消标题")},
                    {"yAxisTitle", New McpToolProperty("string", "纵轴（数值轴）标题；空串=取消标题")},
                    {"showDataLabels", New McpToolProperty("boolean", "是否显示数据标签")},
                    {"dataLabelPosition", New McpToolProperty("string", "数据标签位置：xlLabelPositionOutsideEnd / xlLabelPositionCenter / xlLabelPositionInsideEnd / xlLabelPositionInsideBase / xlLabelPositionBestFit 等")},
                    {"addTrendline", New McpToolProperty("boolean", "为第 1 个数值系列添加线性趋势线（R² + 预测公式常用于回归/预测）")}
                },
                .Required = New List(Of String) From {}
            }))

            ' [L3] 38d. excel_export_chart_image - 导出图表为 PNG/JPG（可在聊天窗内直接展示或贴入 PPT/Word）
            tools.Add(New McpTool("excel_export_chart_image", "将图表导出为 PNG/JPG 图片，返回保存路径；可直接把生成的图片在聊天中内联展示或用于制作报告", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")},
                    {"chartName", New McpToolProperty("string", "图表名称（可选，默认导出最后一个图表）；推荐先调用 excel_list_charts 获取")},
                    {"outputPath", New McpToolProperty("string", "完整输出文件路径（含扩展名）。为空时保存到系统临时目录，文件名=图表名.png")},
                    {"imageFormat", New McpToolProperty("string", "图片格式：png（默认）/ jpg / gif；outputPath 含扩展名时优先取扩展名")}
                },
                .Required = New List(Of String) From {}
            }))

            ' [L3] 38g. excel_set_pivot_field_layout - 配置已有透视表的字段布局（行/列/值），并可换数据源
            tools.Add(New McpTool("excel_set_pivot_field_layout", "配置已有数据透视表的字段布局：设置行字段(rowFields)、列字段(columnFields)、值字段(valueFields)与聚合函数(valueFunctions)；可用 sourceRange 同时更换数据源；resetFields=false 可在现有布局上追加字段（默认 true 先清空旧布局）。字段名必须取自该透视表实际字段——若传错，报错会列出全部可用字段。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "透视表所在工作表名（不传则用当前活动表；若该表没有透视表，报错会告诉你透视表真实所在工作表）")},
                    {"pivotTableName", New McpToolProperty("string", "透视表名称，如 PivotTable1；不传则操作该工作表上的最后一个透视表")},
                    {"sourceRange", New McpToolProperty("string", "【可选】更换数据源范围，如 Sheet1!A1:I101 或 A1:I101；留空表示不改动数据源")},
                    {"rowFields", New McpToolProperty("string", "行字段 JSON 数组，如 [""部门"",""月份""]；决定左侧逐层分行的维度")},
                    {"columnFields", New McpToolProperty("string", "列字段 JSON 数组，如 [""季度""]；决定顶部分列的维度，可留空")},
                    {"valueFields", New McpToolProperty("string", "值字段 JSON 数组，如 [""销售额""]；被汇总的度量列")},
                    {"valueFunctions", New McpToolProperty("string", "聚合函数 JSON 数组，与 valueFields 一一对应：Sum/Count/Average/Max/Min（默认 Sum）")},
                    {"resetFields", New McpToolProperty("boolean", "是否先清空现有字段布局（默认 true）；传 false 表示在现有布局上追加字段")}
                },
                .Required = New List(Of String) From {}
            }))

            ' [L3] 39. excel_refresh_pivot_table - 刷新数据透视表
            tools.Add(New McpTool("excel_refresh_pivot_table", "刷新数据透视表，同步数据源更新", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")},
                    {"pivotTableName", New McpToolProperty("string", "透视表名称（可选，不指定则刷新最后一个透视表）")}
                }
            }))

            ' [L3] 40. excel_apply_conditional_format - 应用条件格式
            tools.Add(New McpTool("excel_apply_conditional_format", "对指定范围应用条件格式（数据条、色阶、图标集、突出显示规则）。注意：range 与 formatType 可不传——系统会自动回退：优先对用户当前鼠标选中区域应用；未选中时则对当前工作表已使用数据区域（UsedRange）应用；formatType 空则默认使用三色阶（colorScale，最通用直观）。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认为当前活动工作表）")},
                    {"range", New McpToolProperty("string", "要应用条件格式的范围，如 A1:D100（可选；不传则自动使用用户选区→UsedRange 顺序回退）")},
                    {"formatType", New McpToolProperty("string", "条件格式类型：dataBar（数据条）、colorScale（色阶，默认）、iconSet（图标集）、highlight（突出显示规则）（可选；不传则默认 colorScale）")},
                    {"options", New McpToolProperty("string", "格式选项 JSON，如 {color: '#FF0000', operator: 'greaterThan', value: 100}（可选，有合理默认）")}
                },
                .Required = New List(Of String)()
            }))

            ' [L3] 41. excel_create_slicer - 创建 Excel 原生切片器（共享透视缓存 + 透视表 + SlicerCache/Slicers，图表 dataRange 引用透视表区域即可联动）
            tools.Add(New McpTool("excel_create_slicer", "创建 Excel 原生切片器（共享透视缓存 + 透视表 + SlicerCache/Slicers）。所有透视表共用同一 PivotCache，任一切片器筛选即联动全部透视表；图表通过 dataRange 引用透视表区域（本工具返回的 dataRange）即可随切片器联动", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"sourceRange", New McpToolProperty("string", "源数据范围（含表头），如 Sheet1!A1:G101；同一范围的所有透视表共享一个透视缓存")},
                    {"sheet", New McpToolProperty("string", "切片器所在工作表（看板工作表），省略则用当前活动表")},
                    {"slicerDimensions", New McpToolProperty("string", "切片器维度列表 JSON：[{""header"":""销售区域"",""top"":110,""left"":20,""width"":130,""height"":150}]；top/left 为磅坐标")},
                    {"pivotTables", New McpToolProperty("string", "透视表定义列表 JSON：[{""dimension"":""销售区域"",""value"":""订单金额(元)"",""func"":""sum"",""name"":""Pivot1"",""targetCell"":""A1""}]；每个图表一个透视表")}
                },
                .Required = New List(Of String) From {"sourceRange"}
            }))

            ' [L4] excel_add_shape - 绘制形状（圆角卡片/色块/分隔条，用 range 锚点，不合并单元格）
            tools.Add(New McpTool("excel_add_shape", "在看板上绘制形状：圆角卡片/矩形色块/圆形/细分隔条。用 range 锚点定位（取该单元格区域的左/上/宽/高换算成磅），支持 fill 填充色、line 边框色、可选多行文字（每行独立字号/颜色/加粗/对齐）。用于搭建专业卡片，不合并单元格", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"range", New McpToolProperty("string", "锚点单元格区域，如 A1:T2 或 D4:I9；形状贴合该区域绘制")},
                    {"sheet", New McpToolProperty("string", "目标工作表名，省略则用当前活动表（或 range 的 sheet! 前缀）")},
                    {"shapeType", New McpToolProperty("string", "形状类型：roundedRect=圆角卡片(默认)/rect=矩形/oval=圆形")},
                    {"fill", New McpToolProperty("string", "填充色 hex(#RRGGBB) 或 none 不填充")},
                    {"line", New McpToolProperty("string", "边框色 hex 或 none 无边框")},
                    {"lineWeight", New McpToolProperty("number", "边框粗细磅值，默认 1")},
                    {"cornerRadius", New McpToolProperty("number", "圆角半径 0~0.5，roundedRect 专用（默认 0.12）")},
                    {"noShadow", New McpToolProperty("boolean", "关闭形状阴影（默认 true）")},
                    {"lines", New McpToolProperty("string", "形状文字 JSON：[{text,size,color,bold,alignH}...]，每行一个对象，alignH=center/left/right")},
                    {"sendToBack", New McpToolProperty("boolean", "是否置于底层（默认 false）")}
                },
                .Required = New List(Of String) From {"range"}
            }))

            ' ── 通用框架层工具（非看板专用，多步骤复合操作通用）──
            ' excel_get_layout_snapshot — 每轮自动注入，让 LLM 知道当前工作表已占用区域/图表/KPI
            tools.Add(New McpTool("excel_get_layout_snapshot", "获取当前工作表的布局快照：已占用区域、图表列表(位置/大小/数据源)、已写 KPI、建议下一个锚点位置。框架自动在多步骤场景每轮调用，让 LLM 知道自己已经放了什么、还能放哪里", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"sheet", New McpToolProperty("string", "工作表名称（可选，默认当前活动工作表）")},
                    {"includeHeader", New McpToolProperty("boolean", "是否返回表头行样本（默认 true）")},
                    {"includeKpiCells", New McpToolProperty("boolean", "是否返回 KPI 单元格样本（默认 true，仅前 20 个非空单元格）")}
                },
                .Required = New List(Of String) From {}
            }))
        End Sub

    End Class

End Namespace