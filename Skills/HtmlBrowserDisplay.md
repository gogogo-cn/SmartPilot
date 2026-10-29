---
hosts: [excel]
name: HtmlBrowserDisplay
displayName: HTML浏览器展示
description: 将选中的Excel数据以格式化表格形式在系统默认浏览器中展示。不做数据修改，不做数据分析，仅读取并展示当前选区。
keywords:
  - HTML展示
  - HTML
  - 浏览器展示
  - 浏览器
  - 网页展示
  - 网页
  - 在浏览器中打开
  - 用浏览器
  - 用HTML列出
  - 用html列出
  - 用HTML展示
  - 用html展示
  - HTML列出
  - html列出
  - 展示为表格
  - 表格形式
  - 网页表格
  - 浏览器表格
  - 列出数据
  - 看一下数据
  - 看下数据
  - 看看数据
  - 表格查看
  - 网页查看
  - 浏览器查看
  - 在网页中
  - 在浏览器中
examples:
  - 用HTML列出选中表格数据
  - 用浏览器展示这些数据
  - 在浏览器中查看
  - 展示为网页表格
  - 看一下选中的数据（表格形式）
isGeneric: true
category: 效率工具
needsConfig: false
processingMode: Hybrid
supportsImageInput: false
supportsFileInput: false
outputFormat: htmlBrowser
isCustom: false
pipeline: html
tags:
  - 数据可视化
  - HTML展示
  - 系统技能
quickPrompts:
  - label: 美化样式
    prompt: 用美化样式在浏览器中展示当前数据
  - label: 深色主题
    prompt: 用深色主题在浏览器中展示当前数据
  - label: 导出HTML
    prompt: 把当前数据导出为HTML文件
dependencies:
  - ExcelBridge
version: 2
---

# HTML浏览器展示

## 说明

此 Skill 通过 `pipeline: html` 声明路由到前端 html-display-pipeline，执行确定性流程：
1. 读取当前选中区域数据（excel_read_range）
2. 生成格式化 HTML 表格（含数字对齐、表头样式等）
3. 调用 excel_display_html_in_browser 在浏览器中打开

不经过 LLM 推理，不生成 VBA，不修改数据。
