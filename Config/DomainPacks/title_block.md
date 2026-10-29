---
id: title_block
name: 标题栏抽取
description: 从工程图 OCR 文本里抽取图号、图名、比例、材料、日期等标题栏字段（复用内置 dimension_annotation 解析器）
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

## 说明

这是一个**可直接安装**的示例模块：pipeline 为 `text`，表示复用插件自带的
ParseScript 解析引擎（`Config/DefaultParseScripts/dimension_annotation.md`），
不需要任何 Python 环境，安装后在 OCR 面板勾选即可生效。

把它复制到 `%APPDATA%\ExcelAIVSTOAddIn\DomainPacks\` 即可被识别，
或在 OCR 面板「管理专项模块」里粘贴本文件内容安装。
