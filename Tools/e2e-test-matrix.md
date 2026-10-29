# 🔬 ExcelAIVSTOAddIn 端到端测试矩阵

> 阶段：D2 — 制定并运行端到端测试矩阵  
> 目标：覆盖 A/B/C 三阶段重构后的关键路径，确保静态约束、单元能力、集成链路、并发稳定性均可验证。  
> 最后更新：2026-07-23

---

## 1. 测试分层

| 层级 | 可自动化 | 依赖环境 | 说明 |
|------|----------|----------|------|
| L1 静态约束 | ✅ 是 | Node.js | 不启动 Excel/浏览器，校验代码约束、构建同步、版本戳 |
| L2 模块单元 | ✅ 是 | Node.js | VBA 安全验证器、路由规则、工具过滤规则等纯函数 |
| L3 浏览器集成 | ⚠️ 半自动 | Excel + 浏览器 | 需人工在 Excel 中触发，运行 L3 场景测试 |
| L4 端到端场景 | ❌ 手动 | Excel + 浏览器 + AI 后端 | 覆盖真实用户交互：纯文本、OCR、多文件、聚合图表、代码生成、编辑重发 |
| L5 并发/压力 | ❌ 手动 | Excel + 浏览器 | 快速连续发送、编辑重发、切换 Sheet 后发送 |

---

## 2. L4 端到端场景矩阵

### 2.1 纯文本对话

| 编号 | 用例 | 输入 | 预期结果 | 关键断言 |
|------|------|------|----------|----------|
| TXT-01 | 普通闲聊 | "你好" / "今天天气怎么样" | 走 STANDARD_AI，不渲染四阶段卡片，直接文本回复 | `skipExecutionFlow=true`，无 `[FLOW]` 日志 |
| TXT-02 | Excel 相关但无操作意图 | "Excel 是什么" | 走 STANDARD_AI，文本回复，不注入 MCP 工具 | 工具列表为空 |
| TXT-03 | 操作意图（无选中区） | "读取当前表格前5行" | 触发上下文感知 → 意图分析 → 决策规划 → 执行，使用 ActiveCell/UsedRange | 四阶段卡片正常显示，工具调用成功 |

### 2.2 单图 OCR

| 编号 | 用例 | 输入 | 预期结果 | 关键断言 |
|------|------|------|----------|----------|
| OCR-01 | 上传车票/票据图片 | 图片 + "提取这张车票信息" | 仅注入 `excel-local__excel_ocr_recognize`，前端 OCR 完成后直接基于文本分析，不调用 browse/read file | 工具列表仅含 OCR，无 `excel_browse_files` / `excel_read_file` |
| OCR-02 | 上传图片后要求结构化数据 | 图片 + "把内容整理成表格" | 输出业务字段表格（票号、出发地、目的地等），不含文件元数据 | 表格字段为业务字段，无 `fileType`/`ocrEngine`/`charCount` |
| OCR-03 | 图片场景要求写 VBA | 图片 + "根据图片内容写一段 VBA" | 走 CODE_GENERATION，不注入任何 MCP 工具 | `allowedTools=[]`，输出 `[EXECUTE:GenerateVBA]` 或标准代码块 |

### 2.3 多文件

| 编号 | 用例 | 输入 | 预期结果 | 关键断言 |
|------|------|------|----------|----------|
| MUL-01 | 批量文件上传 | 选择 3 个 Excel/文本文件 + "总结这些文件" | `processMultipleFiles` 复用同一个 bubble，不重复创建 | Console 无重复创建气泡日志 |
| MUL-02 | 多文件 + 图片混传 | 2 个文件 + 1 张图片 | 正确识别文件与图片上下文，不冲突 | 附件区与图片区同时存在，发送后无异常 |
| MUL-03 | 清空附件 | 点击清空按钮 | `clearAttachment()` 完全清空所有备份源 | `App.selectedFile`、`window.__backupSelectedFile`、`_storedFiles`、`file-input.value` 均为空 |

### 2.4 聚合图表

| 编号 | 用例 | 输入 | 预期结果 | 关键断言 |
|------|------|------|----------|----------|
| CHART-01 | 按部门聚合销售额 | 选中 A1:G22 + "按部门展示销售业绩" | 四阶段流程，执行 `excel_read_range` → `excel_aggregate` → `excel_create_chart` | 图表正常生成，上下文感知显示 5 行预览 |
| CHART-02 | 未指定聚合维度 | 选中数据 + "统计一下" | SmartRouter 默认按首列/表头推断，给出置信度 | 意图分析阶段显示规则匹配与置信度 |
| CHART-03 | 大数据量 | 选中 1000+ 行 | 上下文感知只取元信息和前 5 行预览，不拖慢 UI | 表格预览不超过 5 行，滚动流畅 |

### 2.5 VBA / Python 代码生成

| 编号 | 用例 | 输入 | 预期结果 | 关键断言 |
|------|------|------|----------|----------|
| CODE-01 | 写 VBA | "写一个宏，把选中区域去重" | 命中 `CODE_GENERATION`，调用 `GenerateVBA`，输出标准 markdown 代码块 | 按钮为“在编辑器中打开/执行宏/复制”，点击“执行宏”调用 `App.executeCode(this)` |
| CODE-02 | 写 Python | "预测下季度销售额" | 命中 `PYTHON_CODE_GENERATION`，调用 `GeneratePython`，代码仅使用标准库 | 无 `import pandas`、`import numpy`、`import sklearn` |
| CODE-03 | 写 VBA 但关键词不含“执行” | "生成一段 VBA 代码" | 工具注入过滤掉 `excel-local__excel_execute_vba` | 工具列表不含 execute_vba |
| CODE-04 | 直接要求执行 VBA | "执行这段代码：Sub ..." | 允许注入 `excel-local__excel_execute_vba` | 工具列表含 execute_vba（若后端支持） |

### 2.6 编辑重发与并发

| 编号 | 用例 | 输入 | 预期结果 | 关键断言 |
|------|------|------|----------|----------|
| EDIT-01 | 编辑历史消息重发 | 编辑上一条用户消息并重新发送 | 基于新消息重新路由，旧气泡被替换或更新，不残留旧结果 | 无重复气泡，最终答案更新 |
| EDIT-02 | 快速连续发送两条 | 连续发送消息 A、B | 第二条等待第一条完成或排队，不覆盖第一条气泡 | Console 显示两个独立 flowKey |
| EDIT-03 | 切换 Sheet 后立即发送 | 从 Sheet1 切到 Sheet2 再发送 | 上下文感知读取 Sheet2 的选中区域 | 上下文感知显示 Sheet2 名称 |

---

## 3. 自动化测试入口

```powershell
# 运行 L1 + L2 自动化测试
.\Tools\run-e2e-tests.ps1

# 仅运行静态校验
node .\Tools\e2e-static-validator.js

# 仅运行 VBA 验证器测试
node .\Tools\test-validator.js
```

---

## 4. 浏览器集成测试入口

在 Excel 插件 WebView2 中打开 DevTools，手动测试 L3 场景。

---

## 5. 结果记录模板

| 日期 | 测试员 | 版本戳 | 通过/失败 | 失败项 | 备注 |
|------|--------|--------|-----------|--------|------|
| 2026-07-23 | Auto | v2026072249 | ✅ L1 全部通过 | 无 | e2e-static-validator.js |
| 2026-07-23 | Auto | v2026072249 | ✅ L2 VBA 验证通过 | 无 | vba-validator.js |
| | | | | | |

---

## 6. D1 构建流程（已完成）

自动同步脚本已实现：

| 文件 | 说明 |
|------|------|
| [Tools/sync-ui-to-debug.js](file:///c:/Users/Administrator/Desktop/addin/ExcelAIVSTOAddIn/Tools/sync-ui-to-debug.js) | Node.js 同步脚本（跨平台，可被 CI 调用） |
| [Tools/sync-ui-to-debug.ps1](file:///c:/Users/Administrator/Desktop/addin/ExcelAIVSTOAddIn/Tools/sync-ui-to-debug.ps1) | PowerShell 入口 |

功能：
- 同步 `UI/js/` → `bin/Debug/UI/js/`
- 同步 `UI/css/` → `bin/Debug/UI/css/`
- 自动生成统一版本戳 `vYYYYMMDDHHMM`
- 同步更新 `UI/index.html` 与 `bin/Debug/UI/index.html` 中所有 `?v=` 戳
- 自动备份原文件，验证文件完整性（防损坏为只剩 `</html>`）

运行方式：

```powershell
# 实际同步
.\Tools\sync-ui-to-debug.ps1

# 仅预览
.\Tools\sync-ui-to-debug.ps1 -DryRun

# 强制指定版本戳
.\Tools\sync-ui-to-debug.ps1 -Stamp v2026072355

# 同步并清理 bin/Debug 端废弃文件（谨慎使用）
.\Tools\sync-ui-to-debug.ps1 -Prune
```
> 最近一次运行结果（2026-07-23）：版本戳统一为 `v202607231843`，109 个资源文件已同步，L1/L2 静态校验全部通过。
>
> 同步范围：
> - 目录：`js/`、`css/`、`images/`、`lib/`、`config/`、`prompts/`
> - 根目录：所有 `*.html`、`*.md`、`*.txt`、`*.js`、`*.json` 等入口与文档文件
> - 7 个 HTML 入口的版本戳统一更新
> - 可选 `--prune` 清理 bin/Debug 端废弃文件
