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

        ''' <summary>Layer1 核心基础 + Layer3 高级功能 工具定义</summary>
        Private Shared Sub BuildCoreTools(tools As List(Of McpTool))
            ' ========== <Layer>1 - 核心基础</Layer> ==========

            ' [L1] 1. excel_read_range
            tools.Add(New McpTool("excel_read_range", "读取 Excel 工作表中指定范围或当前选中区域的数据。\n📌 截断行为：超过 maxRows 行时自动截断，返回体带 truncated: true 标记，告知大模型数据不全需缩小 range 重读。\n📌 range 兜底：range 为空时自动按 → 当前选中区域(Selection) → 已使用区域(UsedRange) 逐级回退，返回体带上实际生效的 range 地址。\n📌 返回格式：array（默认二维数组）或 object（含表头的对象数组）。\n\n【默认行为】如果不指定range参数，则自动读取当前Excel中选中的单元格区域。这是读取/查看/获取数据的首选工具。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "工作表名称（默认为当前活动工作表）")},
                    {"range", New McpToolProperty("string", "单元格范围，如 A1:B10（可选，不填则读取当前选中区域）")},
                    {"autoProbe", New McpToolProperty("boolean", "【可选】range 为单格锚点（如 A1）时，是否自动向四周扩展为连续数据区（CurrentRegion），仅无连续区时回退 UsedRange。默认 true：不依赖工作表激活态即可定位真实数据区——拆分/新建的工作表 UsedRange 维度缓存可能不刷新，会返回光标单格")},
                    {"hasHeaders", New McpToolProperty("boolean", "第一行是否为表头")},
                    {"maxRows", New McpToolProperty("integer", "最大读取行数，默认 10000，防止大数据卡死")},
                    {"returnFormat", New McpToolProperty("string", "返回格式：array（默认）或 object（含表头的对象数组），仅允许二选一")},
                    {"skipEmptyRows", New McpToolProperty("boolean", "是否跳过全空行（默认 false，true 时自动过滤全部单元格均为空的行）")},
                    {"maxColumns", New McpToolProperty("integer", "最大读取列数，默认 100，防止列数过多导致单元格爆炸。与 maxRows 共同限制总单元格数 ≤ maxRows × maxColumns")},
                    {"readFormulas", New McpToolProperty("boolean", "是否读取单元格公式（默认 false，true 时返回原始公式文本如 =SUM(A1:A10)）")}
                }
            }))

            ' [L1] 2. excel_write_range
            tools.Add(New McpTool("excel_write_range", "将数据写入 Excel 工作表的指定范围", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"sheet", New McpToolProperty("string", "目标工作表名称")},
                    {"range", New McpToolProperty("string", "起始单元格，如 A1")},
                    {"values", New McpToolProperty("string", "要写入的数据，JSON 二维数组格式，如 [[列1,列2],[值1,值2]]")},
                    {"clearFirst", New McpToolProperty("boolean", "写入前是否清空目标范围")},
                    {"writeMode", New McpToolProperty("string", "写入模式：overwrite（覆盖，默认）或 append（追加）")}
                },
                .Required = New List(Of String) From {"values"}
            }))

            ' [L1] 3. excel_get_sheets
            tools.Add(New McpTool("excel_get_sheets", "获取工作簿中所有工作表的列表", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "可选：Excel 文件完整路径。不指定则使用当前活动工作簿")}
                }
            }))

            ' [L1] 4. excel_create_sheet
            tools.Add(New McpTool("excel_create_sheet", "创建新的空白工作表。仅用于创建空表，不涉及任何数据合并/复制/移动操作。用户说'创建新表'/'新建工作表'/'新建sheet'时使用本工具", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"filePath", New McpToolProperty("string", "【可选】Excel文件绝对路径。提供路径：直接读取磁盘上的文件，不依赖Excel是否打开；留空：操作Excel软件当前已经激活打开的工作簿。⚠️ 传入filePath时，将忽略Excel界面选中状态。")},
                    {"name", New McpToolProperty("string", "新工作表的名称")}
                },
                .Required = New List(Of String) From {"name"}
            }))

            ' ========== <Layer>3 - 高级功能</Layer> ==========

            ' [L3] 5. excel_execute_vba
            tools.Add(New McpTool("excel_execute_vba", "在 Excel 中执行 VBA 代码（经过安全检查）", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"code", New McpToolProperty("string", "要执行的 VBA 代码")}
                },
                .Required = New List(Of String) From {"code"}
            }))

            ' [L3] 5a. excel_read_vba_macros
            tools.Add(New McpTool("excel_read_vba_macros", "读取当前工作簿中所有 VBA 宏代码，按模块拆分输出。需要在 Excel 信任中心启用'信任对 VBA 项目对象模型的访问'。", New McpToolInputSchema() With {
                .Properties = New Dictionary(Of String, McpToolProperty) From {
                    {"moduleName", New McpToolProperty("string", "可选：只读取指定模块名称的代码，不指定则读取所有模块")},
                    {"includeCode", New McpToolProperty("boolean", "是否包含完整代码内容，默认 true。false 时只返回模块列表和摘要")},
                    {"splitByModule", New McpToolProperty("boolean", "是否按模块拆分输出，默认 true。true 时每个模块独立输出代码块")}
                }
            }))
        End Sub

    End Class

End Namespace