# ExcelAI VSTO Add-In

基于 Excel 的 VSTO（Visual Studio Tools for Office）插件，内置 WebView2 前端智能体层，对接本地桥接服务器与 MCP 工具，提供自然语言驱动的数据处理（合并 / 拆分 / 关联 / 清洗）、Excel 美化、图表生成与 Python 数据分析（确定性模板）、PPT 演示文稿生成（一键生成 / 分步构建）等能力。

## 目录结构

```
ExcelAIVSTOAddIn - 副本 (3)\
├── ExcelAIVSTOAddIn.vbproj      # VSTO 主工程（VB.NET）
├── ThisAddIn.vb                 # 插件入口 / 生命周期
├── ExcelBridge.vb               # COM 互操作桥
├── WebView2Manager.vb           # WebView2 前端宿主
├── Ribbon.vb / Ribbon.xml       # 功能区
├── Services/                    # VB 服务（Mcp / Rag / Ocr / 安全守卫 / 桥接）
├── Config/
│   ├── excel_tools.json         # Excel MCP 工具定义与元数据（单一事实源）
│   └── ppt_tools.json           # PPT 工具定义与元数据（主题 / 分步构建）
├── Bridges/                     # 资源型 Bridge（McpBridge / OcrBridge / RagBridge / MobileControlBridge）
├── UI/js/chat-v2/               # 前端智能体层（ECharts / 模板 / 路由 / 服务）
│   ├── pipelines/               # python-pipeline / skill-pipeline / text-pipeline …
│   ├── services/                # tool-registry / intent-detector / llm-io / …
│   └── ui/                      # chat-renderer 等渲染层
├── Tools/                       # 构建 / 同步 / 校验 / 测试脚本
├── Tests/                       # node:test 自动化测试（零依赖）+ VbTests
└── Docs/                        # 架构 / 审核 / 设计文档
```

## 快速开始

### 前置要求

| 依赖 | 版本 | 用途 |
|---|---|---|
| .NET SDK | 10.0.x | VB 测试工程（Tests/VbTests） |
| Node.js | >= 22 | 前端测试 / 工具脚本 |
| Visual Studio | + VSTO 工作负载 | 编译 VSTO 插件 |
| Python | 3.10（可选） | Python 数据分析模板（后端引擎） |

### 构建

在 Visual Studio 中打开 `ExcelAIVSTOAddIn.sln` 生成，或：

```powershell
# Debug 构建（自动触发 frontend 资源同步 + 版本戳）
msbuild ExcelAIVSTOAddIn.sln /p:Configuration=Debug
```

Debug 构建会通过 `SyncUIDebugStamp` AfterBuild target 调用 `Tools/sync-ui-to-debug.js`，把 `UI/` 下的运行时资源同步到 `bin/Debug/UI/` 并统一刷新 `?v=` 版本戳。

### 启动本地桥接服务器（仅独立网页版需要）

VSTO 插件内的前端通过 COM 宿主桥（`App.excelBridge`）直接通信，**无需手动启动桥**。
若以独立网页版方式使用（浏览器打开 `UI/index.html`），则需先启动本地桥（Python 版 `py-excel-bridge.py`，端口 9527）：

```powershell
# 方式一：一键启动（自动探测 Python，可用 PORT 环境变量改端口）
.\start.bat

# 方式二：手动
cd Tools
python py-excel-bridge.py
```

> 注：旧的 `Tools/local-bridge-server.js`（Node/winax）在 Node 20+ 无法编译，已被 `py-excel-bridge.py`（Python + pywin32）取代，二者 HTTP 契约一致。

## 测试

### 全部自动化测试（推荐）

```powershell
Tools\run-tests.ps1
```

该脚本与 CI `test` job 完全一致：
1. `node Tests/run-tests.js` —— node:test 前端 / 服务端套件（聚合运行器，进程内零 spawn）
2. `dotnet run --project Tests\VbTests\ExcelMcpCoreUtils.Tests.vbproj -c Release` —— VB CoreUtils 单元测试

任一失败即以非 0 退出码结束（`RESULT: ALL TESTS PASSED` / `RESULT: FAIL`）。

### 仅 JS 套件

```powershell
cd Tests
npm test            # 等效 node run-tests.js（聚合运行器）
```

### 分组子命令

| 命令 | 覆盖 |
|---|---|
| `npm run test:agent` | agent-loop / skill-agent-e2e |
| `npm run test:router` | pipeline-router |
| `npm run test:skills` | skill-route |
| `npm run test:python` | python-data-inject / python-pipeline-e2e |
| `npm run test:spec` | skill-spec-engine / skill-spec-e2e / skill-formatter / skill-chart-clean |
| `npm run test:e2e-script` | skill-chart-clean-e2e（脚本式，单独运行） |

测试套件明细见 `Tests/README.md`。

## 同步与校验命令

| 命令 | 作用 |
|---|---|
| `node Tools/sync-ui-to-debug.js` | 双端同步 `UI/` → `bin/Debug/UI/` 并刷版本戳（`--stamp` / `--dry-run` / `--prune`） |
| `node Tools/lint.js` | 前端零依赖静态检查（递归语法 + 规则子集，`UI/js` 默认）；CI `lint` job 硬闸门 |
| `Tools/check-tools-consistency.ps1` | MCP 工具一致性校验（VB 注册名 = JSON 名称、前端无硬编码映射、JSON 元数据完整性与 VB 兜底全字段一致） |
| `node Tests/run-tests.js` | 聚合运行所有 JS 测试套件 |

> 修改工具元数据时，必须三处同步：`ExcelMcpTools.vb` 注册名 + `Config/excel_tools.json` + 前端 `tool-registry.js` 显示名。CI 与本地 `check-tools-consistency.ps1` 都会拦截不一致。

## CI

`.github/workflows/build.yml` 定义了三个 job：

| Job | 触发 | 职责 |
|---|---|---|
| **build** | push / PR / 手动 | MSBuild 构建 + 语法检查（递归）+ JS 模块数校验 + MCP 工具一致性 |
| **test** | push / PR | JS 套件（`node Tests/run-tests.js`）+ VB Core 测试（dotnet run） |
| **lint** | push / PR | 前端静态检查（`node Tools/lint.js`），语法错误作为硬闸门 |

## 部署 / 发布

- **签名**：工程启用清单签名（`SignManifests`），默认使用开发证书 `ExcelAIVSTOAddIn_TemporaryKey.pfx`。对外分发前应替换为正式代码签名证书，并在 vbproj 中更新 `ManifestKeyFile` / `ManifestCertificateThumbprint`。
- **发布目录**：`PublishUrl` 指向 `bin\Release\`。使用 Visual Studio 的"发布"向导（ClickOnce）生成可分发安装包。
- **本地桥接**：装在 Excel / Word / PowerPoint 里的插件，对话和工具走 COM 宿主桥，最终用户不用单独启动桥。只有用浏览器打开 `UI/index.html` 的独立网页版，才需要先运行 `Tools/py-excel-bridge.py`（见上文「启动本地桥接服务器」）。不要把已废弃的 `local-bridge-server.js` 写进安装说明。
- **Python 引擎**：如启用 Python 数据分析模板，需部署目标机具备可被引擎调用的 Python 运行时。日常改表和票据识别不依赖 Python。
- **付费**：Stripe 价格和国内支付未配置完成前，账户页不显示升级按钮，授权服务也会拒绝模拟支付。试用注册仍然可用。

## PPT 分步构建工具（新增 14 个）

> 版本：2026-09 · 与 `Config/ppt_tools.json` 同步（自 `Docs/` 指南全文并入）

PPT 工具集由 **5 个核心工具**（`ppt_get_templates` 主题查询 / `ppt_analyze_data` 数据设计分析 / `ppt_generate` 一键生成 / `ppt_import_theme` 导入主题 / `ppt_browse_template` 选择模板）与 **14 个分步构建工具** 组成；本节覆盖后者，两类工具共享同一渲染引擎。

### 为什么需要分步构建

原有 `ppt_generate` 是"一键生成"：AI 一次性提交整份 PPT 的页面规格，适合需求明确、一次成型的场景。

分步构建工具则把"生成 PPT"拆成 **创建项目 → 逐页添加 → 随时调整 → 最终渲染** 四个阶段，AI 可以：

- 按用户需求**逐页拼装**，边构建边确认，中途改页、删页、调顺序
- 复杂 PPT（发布会、经营分析、培训课件）拆成多次调用，每次只负责一页
- 页面规格与 `ppt_generate` 的 slides 数组**完全同构**，渲染复用同一引擎，效果一致

**两类模式可以混用**：分步构建完成后，规格仍在项目里，也可以继续调整再重新渲染。

### 工作流程总览

```
┌─────────────┐      ┌────────────────────┐      ┌─────────────────┐
│ ① 创建项目   │ ───► │ ② 逐页添加（可多次） │ ───► │ ③ 渲染生成 .pptx │
│ project_start│      │ add_*_slide        │      │ project_build   │
└─────────────┘      └────────┬───────────┘      └─────────────────┘
                              │  随时：list / update / remove
                              └─────────────────────────────────────►
```

**典型链路（数据型 PPT）**：

1. `ppt_project_start`：创建项目，传 `theme` + 数据来源（`rangeAddress` 或 `headers`+`rows`）
2. `ppt_add_cover_slide`：封面
3. `ppt_add_kpi_slide`：核心指标页
4. `ppt_add_chart_slide`：趋势图（可多张）
5. `ppt_add_table_slide`：明细表
6. `ppt_project_build`：渲染生成

**典型链路（纯结构 PPT）**：

1. `ppt_project_start`：创建项目，传 `reportTitle`（纯结构可不传数据）
2. `ppt_add_cover_slide` → `ppt_add_structure_slide`（目录）→ `ppt_add_text_slide`（正文）→ `ppt_add_shapes_slide`（流程图）→ `ppt_add_smartart_slide`（架构图）→ `ppt_add_cover_slide`（致谢）
3. `ppt_project_build`：渲染生成

### 工具速查表

| 工具 | 类型 | 作用 |
| --- | --- | --- |
| `ppt_project_start` | 生命周期 | 创建 PPT 构建项目，返回 `projectId` |
| `ppt_add_cover_slide` | 页型添加 | 封面 / 标题页 / 致谢页 |
| `ppt_add_text_slide` | 页型添加 | 段落文本 / 要点列表 |
| `ppt_add_table_slide` | 页型添加 | 数据表格页 |
| `ppt_add_chart_slide` | 页型添加 | 图表页（柱/折线/饼/组合/气泡等 10 种） |
| `ppt_add_kpi_slide` | 页型添加 | 核心指标页（KPI 卡） |
| `ppt_add_structure_slide` | 页型添加 | 目录 / 章节 / 时间线 / 对比页 |
| `ppt_add_shapes_slide` | 页型添加 | 自由形状页（流程图/架构图） |
| `ppt_add_picture_slide` | 页型添加 | 图片页（本地路径或 URL） |
| `ppt_add_smartart_slide` | 页型添加 | SmartArt 替代页（流程/循环/层次/矩阵） |
| `ppt_project_list` | 生命周期 | 查看项目当前页面结构 |
| `ppt_project_update` | 生命周期 | 修改指定页 |
| `ppt_project_remove` | 生命周期 | 删除指定页 |
| `ppt_project_build` | 生命周期 | 渲染生成 .pptx 并自动打开 |

### 项目生命周期工具

#### ppt_project_start — 创建项目

一切分步构建的起点，**必先调用**。返回 `projectId`，后续所有工具都要携带它。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `reportTitle` | string | 否 | 报告/演示文稿标题，作为封面标题与文件名（默认"数据分析报告"；纯结构 PPT 建议必填） |
| `theme` | string | 否 | 主题 id（内置 10 个，用 `ppt_get_templates` 查询），不传用默认主题 |
| `orientation` | string | 否 | 页面方向：`landscape` 横版（默认）/ `portrait` 竖版 |
| `requirement` | string | 否 | 用户需求原文（供后续页面编排参考） |
| `rangeAddress` | string | 否 | Excel 选区地址（如 `A1:H27`，含表头）。数据表格/图表/KPI 页的数据来源 |
| `headers` | string | 否 | 列名数组的 JSON 字符串，与 `rows` 成对传入则优先用传入数据 |
| `rows` | string | 否 | 数据行数组的 JSON 字符串，每行为 `{列名: 值}` 的对象 |

**数据来源说明**：只有数据页（表格/图表/KPI）需要数据。`rangeAddress` 与 `headers+rows` 二选一即可；纯结构 PPT（封面/目录/文本/形状/图片/SmartArt）无需传数据。

```json
{
  "reportTitle": "2026 年 Q3 经营分析",
  "theme": "business_blue",
  "rangeAddress": "A1:H27"
}
```

#### ppt_project_list — 查看项目

构建过程中确认进度，或定位要修改/删除的页（`index` 从 1 起）。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `projectId` | string | 是 | 项目 id |

返回：页序（index）、每页类型与标题。

#### ppt_project_update — 修改指定页

传入的字段**覆盖**该页对应字段，`type` 也可改；未传的字段保持不变。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `projectId` | string | 是 | 项目 id |
| `index` | integer | 是 | 要修改的页序号（从 1 起，用 `ppt_project_list` 查询） |
| `type` | string | 否 | 改为其他页型 |
| `title` / `subtitle` | string | 否 | 新标题 / 新副标题 |
| `chartType` | string | 否 | 图表页：改图表类型 |
| `items` | string | 否 | 改条目/指标/节点数组 |
| `variant` | string | 否 | 改版式变体 |
| `transition` / `animation` / `animate` | - | 否 | 改动画设置 |

```json
{
  "projectId": "3f2a…",
  "index": 3,
  "title": "月度销售趋势（修正版）",
  "chartType": "combo"
}
```

#### ppt_project_remove — 删除指定页

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `projectId` | string | 是 | 项目 id |
| `index` | integer | 是 | 要删除的页序号（从 1 起） |

#### ppt_project_build — 渲染生成

分步构建的**最后一步**：把项目所有页面规格渲染为 .pptx 并自动打开预览。项目无页面时报错。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `projectId` | string | 是 | 项目 id |

生成后如需修改：`update/remove` 调整页面 → 再次 `project_build` 重新渲染即可。

### 页型添加工具

所有添加工具**共有的通用参数**：

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `projectId` | string | 是 | 项目 id（`ppt_project_start` 返回） |
| `transition` | string | 否 | 页切换动画：`fade`（默认）/`push`/`wipe`/`dissolve`/`cut`/`random`/`none` |
| `animation` | string | 否 | 元素进入动画：`fade`（默认）/`appear`/`rise`/`float`/`grow`/`none` |
| `animate` | boolean | 否 | 是否启用动画（`false` 关闭本页所有动画），默认 `true` |

#### ppt_add_cover_slide — 封面/结尾页

**页型**：`cover` 封面（默认）｜`title` 标题页｜`end`/`closing` 致谢/结束页

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `type` | string | 否 | `cover`/`title`/`end`/`closing`，默认 `cover` |
| `title` | string | 否 | 标题文字 |
| `subtitle` | string | 否 | 副标题/日期/部门等 |
| `variant` | string | 否 | cover 用 `centered`/`modern`/`split`；closing 用 `thanks`/`centered` |

```json
{
  "projectId": "3f2a…",
  "title": "2026 年 Q3 经营分析报告",
  "subtitle": "战略规划部 · 2026-09-06",
  "variant": "centered"
}
```

#### ppt_add_text_slide — 文本/要点页

**页型**：`text` 段落文本（默认）｜`bullets` 要点列表

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `type` | string | 否 | `text`/`bullets`，默认 `text` |
| `title` | string | 否 | 页标题 |
| `bullets` | string | 否 | 要点数组 JSON，如 `["第一点","第二点"]` |
| `content` | string | 否 | 整段文本内容（text 页用） |
| `variant` | string | 否 | `bullets` 要点 / `cards` 卡片 / `numbered` 编号 |

```json
{
  "projectId": "3f2a…",
  "type": "bullets",
  "title": "本季度核心结论",
  "bullets": "[\"营收同比增长 18%，创近三年新高\",\"毛利率提升至 42%，环比 +2.1pp\",\"华东区贡献 38% 增量，为主要引擎\"]",
  "variant": "numbered"
}
```

#### ppt_add_table_slide — 数据表格页

**页型**：`table`。数据来自项目级（创建时传入），或在本工具补充。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `title` | string | 否 | 表格页标题 |
| `variant` | string | 否 | 版式变体（`standard` 等） |
| `rangeAddress` | string | 否 | 【补充数据】Excel 选区地址；项目未带数据时在此补充 |
| `headers` | string | 否 | 【补充数据】列名数组 JSON |
| `rows` | string | 否 | 【补充数据】数据行数组 JSON，与 headers 成对 |

#### ppt_add_chart_slide — 图表页

**页型**：`chart`。支持 10 种图表类型，数据来自项目级或本工具补充。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `title` | string | 否 | 图表页标题 |
| `chartType` | string | 否 | `column` 柱（默认）/`bar` 条/`line` 折线/`pie` 饼/`doughnut` 环/`area` 面积/`radar` 雷达/`scatter` 散点/`combo` 组合图/`bubble` 气泡图 |
| `xField` | string | 否 | X 轴列名（默认第一列） |
| `series` | string | 否 | 数值系列列名数组 JSON，如 `["销售额","增长率"]` |
| `secondaryAxis` | string | 否 | 次坐标轴列名数组 JSON（量纲差异大时用） |
| `lineSeries` | string | 否 | 组合图中转为折线的系列（系列序号从 1 起，或列名数组） |
| `insight` | string | 否 | 数据洞察文字（图表下方结论） |
| `variant` | string | 否 | `standard`/`full`/`insight` |
| `rangeAddress`/`headers`/`rows` | string | 否 | 【补充数据】同表格页 |

```json
{
  "projectId": "3f2a…",
  "title": "营收与毛利率趋势",
  "chartType": "combo",
  "xField": "月份",
  "series": "[\"营收\",\"毛利率\"]",
  "secondaryAxis": "[\"毛利率\"]",
  "lineSeries": "[2]",
  "insight": "营收持续增长，毛利率在 8 月达峰后小幅回落"
}
```

#### ppt_add_kpi_slide — 核心指标页

**页型**：`kpi`。数据来自项目级或本工具补充。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `title` | string | 否 | 指标页标题 |
| `items` | string | 否 | 指标数组 JSON，每项 `{label,value,highlight}` |
| `variant` | string | 否 | `grid` 网格 / `hero` 大数字 / `row` 横向 |
| `rangeAddress`/`headers`/`rows` | string | 否 | 【补充数据】同表格页 |

```json
{
  "projectId": "3f2a…",
  "title": "Q3 核心指标",
  "items": "[{\"label\":\"总营收\",\"value\":\"28.6 亿\",\"highlight\":true},{\"label\":\"毛利率\",\"value\":\"42%\"},{\"label\":\"新增客户\",\"value\":\"1,286\"},{\"label\":\"NPS\",\"value\":\"62\"}]",
  "variant": "grid"
}
```

#### ppt_add_structure_slide — 结构页

**页型**：`agenda` 目录（默认）｜`section` 章节页｜`timeline` 时间线｜`comparison` 左右对比

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `type` | string | 否 | `agenda`/`section`/`timeline`/`comparison`，默认 `agenda` |
| `title` | string | 否 | 页标题 |
| `subtitle` | string | 否 | 副标题（section/comparison 用） |
| `number` | string | 否 | 章节序号（section 用，如 `01`） |
| `items` | string | 否 | 条目数组 JSON：agenda 用 `[{label,desc}]`，timeline 用 `[{time,text}]` |
| `leftTitle` / `rightTitle` | string | 否 | 对比页左/右侧标题 |
| `leftItems` / `rightItems` | string | 否 | 对比页左/右侧条目数组 JSON |
| `variant` | string | 否 | agenda 用 `numbered`/`cards`/`timeline`；section 用 `number`/`bar`/`centered` |

```json
{
  "projectId": "3f2a…",
  "type": "timeline",
  "title": "项目里程碑",
  "items": "[{\"time\":\"Q1\",\"text\":\"需求调研与立项\"},{\"time\":\"Q2\",\"text\":\"产品开发与内测\"},{\"time\":\"Q3\",\"text\":\"规模上线\"}]"
}
```

#### ppt_add_shapes_slide — 自由形状页

**页型**：`shapes`。适合流程图、架构图、示意图等用形状拼装的页面。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `title` | string | 否 | 页标题 |
| `items` | string | 否 | 形状数组 JSON，每项 `{type,x,y,w,h,fill,text,textColor,shadow,rotate,weight}` |

**形状 type 一览**：

| type | 形状 | type | 形状 |
| --- | --- | --- | --- |
| `rect` | 矩形 | `flowTerminator` | 流程起止框 |
| `roundrect` | 圆角矩形 | `flowDocument` | 流程文档框 |
| `oval` | 椭圆 | `callout` | 标注框 |
| `circle` | 圆 | `calloutOval` | 椭圆标注 |
| `diamond` | 菱形 | `triangle` | 三角 |
| `arrow` | 右箭头 | `arrowLeft`/`arrowUp`/`arrowDown` | 其他方向箭头 |
| `flow` | 流程框 | `flowDecision` | 判定框 |
| `line` | 线条 | | |

`fill` 支持主题令牌（`primary`/`secondary`/`accent`/`light`）或 `#RRGGBB` 颜色；坐标（x/y/w/h）缺失时自动网格布局。

```json
{
  "projectId": "3f2a…",
  "title": "数据接入流程",
  "items": "[{\"type\":\"flowTerminator\",\"text\":\"开始\",\"fill\":\"accent\"},{\"type\":\"flow\",\"text\":\"读取 Excel 数据\"},{\"type\":\"flowDecision\",\"text\":\"数据完整?\"},{\"type\":\"flow\",\"text\":\"生成 PPT\"}]"
}
```

#### ppt_add_picture_slide — 图片页

**页型**：`picture`。支持本地图片路径或 http(s) URL（自动下载后插入）。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `title` | string | 否 | 页标题 |
| `images` | string | 否 | 图片数组 JSON，每项 `{src,caption,x,y,w,h,rounded,border,shadow,fit}` |
| `src` | string | 否 | 单图模式：本地路径或 http(s) URL |
| `caption` | string | 否 | 单图模式：图片说明文字 |
| `x`/`y` | integer | 否 | 单图横/纵坐标 |
| `w`/`h` | integer | 否 | 单图宽/高 |
| `rounded` | boolean | 否 | 是否圆角 |
| `border` | string | 否 | 边框：`true`/`false` 或 `#RRGGBB` 颜色 |
| `shadow` | boolean | 否 | 是否阴影 |
| `fit` | string | 否 | 填充方式：`fill` 拉伸 / `fit` 等比 |

```json
{
  "projectId": "3f2a…",
  "title": "产品实拍",
  "images": "[{\"src\":\"C:\\\\images\\\\product.jpg\",\"caption\":\"旗舰版实拍\",\"rounded\":true,\"shadow\":true},{\"src\":\"https://example.com/team.jpg\",\"caption\":\"团队合影\",\"fit\":\"fit\"}]"
}
```

#### ppt_add_smartart_slide — SmartArt 替代页

**页型**：`smartart`。用原生形状绘制（不依赖 SmartArt 组件，兼容性更好）。

| 参数 | 类型 | 必填 | 说明 |
| --- | --- | --- | --- |
| `title` | string | 否 | 页标题 |
| `layout` | string | 否 | `flow` 流程（默认）/`cycle` 循环/`hierarchy` 层次组织架构/`matrix` 矩阵 |
| `items` | string | 否 | 节点数组 JSON：字符串或 `{title,desc}` |

```json
{
  "projectId": "3f2a…",
  "title": "组织架构",
  "layout": "hierarchy",
  "items": "[\"总经理\",\"研发中心\",\"市场中心\",\"运营中心\",\"研发一部\",\"研发二部\"]"
}
```

```json
{
  "projectId": "3f2a…",
  "title": "四象限矩阵",
  "layout": "matrix",
  "items": "[{\"title\":\"核心业务\",\"desc\":\"高价值高投入\"},{\"title\":\"增长业务\",\"desc\":\"高价值低投入\"},{\"title\":\"维持业务\",\"desc\":\"低价值高投入\"},{\"title\":\"退出业务\",\"desc\":\"低价值低投入\"}]"
}
```

### 典型场景示例

#### 场景 A：季度经营分析汇报（数据型）

```jsonc
// 步骤 1：创建项目，绑定数据与主题
ppt_project_start {
  "reportTitle": "2026 年 Q3 经营分析报告",
  "theme": "business_blue",
  "requirement": "季度经营分析，需含指标、趋势、区域对比与下季度展望"
}
// → 返回 projectId: "a1b2…"

// 步骤 2：封面
ppt_add_cover_slide { "projectId": "a1b2…", "title": "2026 年 Q3 经营分析", "subtitle": "战略规划部", "variant": "centered" }

// 步骤 3：核心指标
ppt_add_kpi_slide { "projectId": "a1b2…", "title": "Q3 核心指标", "items": "[…]", "variant": "grid" }

// 步骤 4：营收与毛利率组合图
ppt_add_chart_slide { "projectId": "a1b2…", "title": "营收与毛利率趋势", "chartType": "combo",
  "series": "[\"营收\",\"毛利率\"]", "secondaryAxis": "[\"毛利率\"]", "lineSeries": "[2]", "insight": "…" }

// 步骤 5：区域对比
ppt_add_chart_slide { "projectId": "a1b2…", "title": "各区域营收对比", "chartType": "column", "series": "[\"华东\",\"华南\",\"华北\",\"西南\"]" }

// 步骤 6：季度大事记
ppt_add_structure_slide { "projectId": "a1b2…", "type": "timeline", "title": "季度里程碑", "items": "[…]" }

// 步骤 7：下季度展望
ppt_add_text_slide { "projectId": "a1b2…", "type": "bullets", "title": "Q4 展望", "bullets": "[…]" }

// 步骤 8：渲染生成
ppt_project_build { "projectId": "a1b2…" }
```

#### 场景 B：产品发布会（纯结构型）

```jsonc
ppt_project_start { "reportTitle": "新产品发布会" }
ppt_add_cover_slide { "projectId": "p1…", "title": "Aurora 3.0 发布会", "subtitle": "2026 金秋 · 北京", "variant": "split" }
ppt_add_structure_slide { "projectId": "p1…", "type": "agenda", "items": "[{\"label\":\"01 行业洞察\"},{\"label\":\"02 产品发布\"},{\"label\":\"03 生态合作\"}]" }
ppt_add_text_slide { "projectId": "p1…", "title": "行业趋势", "type": "bullets", "bullets": "[…]" }
ppt_add_shapes_slide { "projectId": "p1…", "title": "产品架构", "items": "[…]" }
ppt_add_smartart_slide { "projectId": "p1…", "title": "组织与生态", "layout": "hierarchy", "items": "[…]" }
ppt_add_picture_slide { "projectId": "p1…", "title": "产品实拍", "images": "[…]" }
ppt_add_cover_slide { "projectId": "p1…", "type": "closing", "title": "谢谢", "variant": "thanks" }
ppt_project_build { "projectId": "p1…" }
```

#### 场景 C：构建中调整

```jsonc
// 查看当前页面结构（拿到 index）
ppt_project_list { "projectId": "a1b2…" }

// 把第 3 页的图表类型换成气泡图，并改标题
ppt_project_update { "projectId": "a1b2…", "index": 3, "title": "区域销售气泡图", "chartType": "bubble" }

// 删除第 5 页（时间线）
ppt_project_remove { "projectId": "a1b2…", "index": 5 }

// 重新渲染
ppt_project_build { "projectId": "a1b2…" }
```

### 注意事项

- **`projectId` 是必填参数**：所有 `ppt_add_*_slide` 与 `project_list/update/remove/build` 都需要；来自 `ppt_project_start` 的返回。
- **项目有效期为 2 小时**：超过 2 小时未访问的项目会被自动清理（防内存泄漏），需重新 `ppt_project_start`。
- **数据页必须要有数据**：`table`/`chart`/`kpi` 页在添加时会校验项目是否有数据来源；没有时可在添加调用中补充 `rangeAddress` 或 `headers`+`rows`，否则报错提示。
- **动画参数是页级通用字段**：每个添加工具都支持 `transition`/`animation`/`animate`；不需要动画的页面传 `"animate": false` 可整体关闭。
- **与 `ppt_generate` 的关系**：两者共享同一渲染引擎与页面规格，效果一致；`ppt_generate` 适合一次性成型，分步工具适合复杂/迭代式构建。
- **复杂 JSON 参数用 JSON 字符串传**：`items`/`series`/`bullets`/`images` 等数组参数以 JSON 字符串形式传入（如 `"[1,2,3]"`），与 `ppt_generate` 的 `slides` 参数风格一致。

## 关键约定

- **零依赖测试**：前端测试使用 Node 内置 `node:test`，不引入 node_modules / jest；`Tools/lint.js` 同样零依赖。
- **单一事实源**：MCP 工具元数据以 `Config/excel_tools.json`（Excel）与 `Config/ppt_tools.json`（PPT）为权威来源，VB 与前端均从它们构建/读取。
- **安全守卫**：自定义 Python 脚本需在 🔍深度 / 👑专家 思考模式下方可执行，其余模式一律拦截（见 `Docs/` 下安全指南）。

详细架构与设计见 [`Docs/`](Docs/)。