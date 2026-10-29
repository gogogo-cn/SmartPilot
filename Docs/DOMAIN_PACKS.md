# 专项识别模块（DomainPack）规范

> 状态：**识别改写入口已接入**。机械制图仍是内置模块。新的识别改写 = 一份 `.md` 清单
> 加一段 Python，丢进用户目录或在 OCR 面板里粘贴安装，**不需要重新编译插件**。
> 字段抽取继续用解析脚本，不要装成专项模块。

## 一、为什么要有它

以前"机械图纸"这个开关是硬编码的：UI 复选框、`localStorage['ocr-mechanical-drawing']`、
`OcrServiceManager.IsMechanicalMode` 的字符串比对，三处各写一遍。再加一个建筑图，
就得改三处 + 重新编译 + 重新发布。

现在改成注册表：

```
用户目录 %APPDATA%\ExcelAIVSTOAddIn\DomainPacks\*.md   ← 新增模块只放这里
        │
        ▼
DomainPackRegistry（VB，扫目录 + 内置登记 + 启用状态持久化）
        │
        ▼
OcrServiceManager.ApplyDomainPack()  ← 取代原来的 IsMechanicalMode 硬编码
```

## 二、模块 = 一个 .md 清单

```markdown
---
id: title_block
name: 标题栏抽取
description: 从工程图 OCR 文本里抽取图号、图名、比例、材料、日期
icon: 工程
category: 工程图
version: 1.0.0
author: SmartPilot
pipeline: text
parseScript: dimension_annotation
---

```options
[
  {"key": "strict", "label": "严格模式（字段缺失即报缺）", "type": "bool", "default": false},
  {"key": "unit", "label": "尺寸单位", "type": "select", "values": ["mm", "in"], "default": "mm"}
]
```
```

### frontmatter 字段

| 字段 | 必填 | 说明 |
|---|---|---|
| `id` | ✅ | 唯一标识，小写字母数字与 `-` `_`；同时是文件名。**不能与内置模块 `mechanical_drawing` 同名** |
| `name` | ✅ | 界面显示名 |
| `pipeline` | ✅ | 执行层：`python`（识别改写，已接入）/ `builtin`（内置，仅插件自身用）/ `text`（不在识别入口执行，字段抽取用解析脚本） |
| `module` | python/builtin 必填 | `builtin` 填识别器代号（如 `GdtSymbolRecognizer`）；`python` 填脚本文件名 |
| `parseScript` | 二选一 | `text` 流水线要复用的 ParseScript id，留空则用同 id 的解析脚本 |
| `requires` | 否 | python 流水线需要的 pip 包，逗号分隔 |
| `icon` / `category` / `version` / `author` / `description` | 否 | 展示信息 |

### options 代码块

JSON 数组，每个元素：

| 键 | 说明 |
|---|---|
| `key` | 取值键，模块内唯一 |
| `label` | 界面显示名 |
| `type` | `bool`（勾选）\| `select`（下拉）\| `number`（数字） |
| `default` | 默认值字符串，bool 用 `true`/`false` |
| `values` | `type=select` 时的候选项 |

勾选状态与选项取值存在 `%APPDATA%\ExcelAIVSTOAddIn\domain_packs_state.json`，**不写进清单文件**（便于升级清单不丢用户设置）。

## 三、三种执行层

| pipeline | 状态 | 能做什么 | 限制 |
|---|---|---|---|
| `python` | ✅ 已接入 | 识别完成后改写文字框。输入是原图路径和每个框的文字、坐标、置信度 | 不换模型权重。超时 15 秒或脚本失败时保留原识别结果。坐标只读，目前只回写 `text` |
| `builtin` | ✅ 仅内置 | 机械制图 GDT（2105 行 VB 实现，原地保留） | 第三方不能用（改了就要重编译插件） |
| `text` | 不在这条入口执行 | 字段抽取已经由解析脚本负责 | 装成专项模块不会改识别结果 |

### python 脚本约定

面板安装时，清单里的 python 代码块会写成 `module` 指定的文件，和 `.md` 放在同一目录。`module` 只能是文件名，不能带目录。

标准输入一个 JSON：

```json
{
  "imagePath": "C:\\...\\图.png",
  "packId": "diameter_mark",
  "options": {},
  "elements": [
    {"index": 0, "text": "O", "confidence": 0.9, "box": {"left": 1, "top": 2, "right": 20, "bottom": 18}}
  ]
}
```

标准输出一个 JSON（放在最后一行）。只列出要改的框：

```json
{"success": true, "elements": [{"index": 0, "text": "Ø"}]}
```

勾选即生效，不需要把 recognitionMode 设成模块 id。脚本可以用 Pillow / OpenCV 自己打开 `imagePath`。没装这些库时，只根据 `text` 和 `box` 改写也可以。可直接安装的模板见 `Config/DomainPacks/diameter_mark.md`，默认不改字。

## 四、安装方式

1. **文件安装**：把 `.md` 和同名 `.py` 一起复制到 `%APPDATA%\ExcelAIVSTOAddIn\DomainPacks\`，下次打开 OCR 面板自动出现。
2. **面板安装**：OCR 面板 → 「管理专项模块」→ 粘贴清单（含 python 代码块）→ 安装。清单和脚本会写入同一目录。

卸载：管理面板里点「卸载」（内置模块没有这个按钮）。停用只是不再参与识别，文件保留。

## 五、一个模块的生命周期

```
写清单（.md）→ 安装（复制/粘贴）→ 出现在面板勾选框
     → 勾选 = SetDomainPackEnabled(id, true)
     → 识别时 OcrServiceManager.ApplyDomainPack 按 pipeline 分派
     → 停用 / 卸载（文件删除，状态清理）
```

## 六、排错

| 现象 | 原因 |
|---|---|
| 装完不出现 | 清单不是以 `---` 开头；或缺 `id`/`name`/`pipeline`。管理面板会显示具体错误 |
| 提示 id 与内置冲突 | 你的 `id` 用了 `mechanical_drawing`，换一个 |
| python 模块勾了没反应 | 脚本不在模块目录、`module` 不是纯文件名、机器上没有 Python，或脚本没有往标准输出打 JSON。失败时识别结果保持原样 |
| 选项改了不生效 | 选项取值在 `domain_packs_state.json`，删除该文件可重置 |

## 七、代码位置

| 文件 | 作用 |
|---|---|
| `Services/DomainPack/DomainPackManifest.vb` | 清单模型 + 选项模型 + Pipeline 枚举 |
| `Services/DomainPack/DomainPackLoader.vb` | frontmatter / options 解析（永不抛异常，坏清单不拖垮 OCR） |
| `Services/DomainPack/DomainPackRegistry.vb` | 内置登记、目录扫描、启用持久化、安装/卸载 |
| `Services/DomainPack/DomainPackPythonRunner.vb` | 识别改写入口：调用已勾选的 python 脚本并回写文字 |
| `Services/OcrServiceManager.vb` | `ApplyDomainPack` / `ResolveDomainPack` 取代硬编码 |
| `Bridges/OcrBridge.vb` | `GetDomainPacks` / `InstallDomainPackFromMarkdown` / `SetDomainPackEnabled` / `UninstallDomainPack` / `SetDomainPackOption` |
| `UI/ocr-panel.html` | 勾选框动态渲染 + 管理弹窗 |
| `Config/DomainPacks/diameter_mark.md` | 可直接安装的识别改写模板（默认不改字） |
| `Config/DomainPacks/title_block.md` | 文本清单示例；识别改写不要走这条 |
| `Core/SmartPilot.Core.vbproj` | 三个新 .vb 的编译登记（漏登记 = 文件不存在） |

> 守栏：`Tests/domain-packs.test.js` 会检查以上接线（vbproj 登记、桥方法存在、
> 面板无硬编码复选框、示例清单字段齐全），防止改回去。
