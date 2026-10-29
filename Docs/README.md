# Docs 文档索引

> 最后更新：2026-09
> 本文档是 `Docs/` 目录的总索引，按用途分类并标注每份文档的状态。

---

## 🟢 活跃文档（当前版本，持续维护）

| 文档 | 说明 | 最近更新 |
|------|------|----------|
| [项目审核评价与扩展优化建议-2026-08-27.html](./项目审核评价与扩展优化建议-2026-08-27.html) | **主审计报告** — 架构评价、问题清单、优化路线图 | 2026-08-27 |
| [Priority1-Extension-Guide.md](./Priority1-Extension-Guide.md) | **第一优先级扩展** — RAG 知识库 + Pipeline 多步管道使用指南 | 2026-06 |
| [Priority2-Extension-Guide.md](./Priority2-Extension-Guide.md) | **第二优先级扩展** — 事件驱动/定时触发 + 增强输出格式 + 权限审批 | 2026-06 |
| [Priority3-Extension-Guide.md](./Priority3-Extension-Guide.md) | **第三优先级扩展** — 高级分析与外部集成 | 2026-06 |
| [Priority4-Extension-Guide.md](./Priority4-Extension-Guide.md) | **第四优先级扩展** — 企业级能力（多租户/审计/部署） | 2026-06 |
| [SmartPilot-改进实测验收清单.md](./SmartPilot-改进实测验收清单.md) | 改进项验收清单（测试通过/未通过状态） | 2026-08 |
| [项目审核评价与策划-2026-08.md](./项目审核评价与策划-2026-08.md) | 项目策划与迭代规划（文字版），含 2026-08-28 进度核对（次坐标轴/去重多列/Python 护栏） | 2026-08-28 |
| [设计-PPT图表次坐标轴-2026-08.md](./设计-PPT图表次坐标轴-2026-08.md) | **PPT 图表次坐标轴** 实现设计（chart slide-spec `secondaryAxis` 字段 + 后端 `AxisGroup=xlSecondary`） | 2026-08-28 |
| [规范-数据区域定位与多步骤展示-2026-09.md](./规范-数据区域定位与多步骤展示-2026-09.md) | **开发规范** — 统一数据区域定位（三级优先）、autoProbe 双分支回退、去重默认高亮、多步骤详情展示规范 | 2026-09 |

---

## 🎬 视频教程物料（教学 / 对外）

> SmartPilot（ExcelAI）操作视频教程全套文案，从入门到高阶共 24 集，含分集脚本、逐字口播、封面版式与口令速查卡。

| 文档 | 说明 |
|------|------|
| [video-tutorial-scripts.md](./video-tutorial-scripts.md) | **主稿** — 24 集分集脚本（目标/时长/分镜/必演口令/收尾） |
| [video-tutorial-scripts-basic-intro.md](./video-tutorial-scripts-basic-intro.md) | 第 1–13 集逐字口播稿 |
| [video-tutorial-scripts-highlevel.md](./video-tutorial-scripts-highlevel.md) | 第 14–22、24 集逐字口播稿 |
| [video-tutorial-script-ep23.md](./video-tutorial-script-ep23.md) | 第 23 集《综合实战：经营分析一条龙》完整逐字稿 |
| [video-tutorial-assets.md](./video-tutorial-assets.md) | 封面标题钩子 + 分平台策略 + 第 3 集口播范例 |
| [video-cover-templates.md](./video-cover-templates.md) | 竖屏封面版式规范（可套用） |
| [video-command-cheatsheet.md](./video-command-cheatsheet.md) | 完整口令速查卡（12 大类 50+ 条） |
| [video-command-cheatsheet-a4.md](./video-command-cheatsheet-a4.md) | A4 一页速查卡（可打印贴桌） |
| [SmartPilot-使用教程.md](./SmartPilot-使用教程.md) | **对外发布版** — 面向终端用户的完整使用教程（入门到高阶，不含视频脚本） |
| [SmartPilot-使用教程.html](./SmartPilot-使用教程.html) | 上述教程的发布级 HTML 帮助页（封面/侧边目录/页脚，可直接上线） |
| [SmartPilot-User-Guide.md](./SmartPilot-User-Guide.md) | English user guide for overseas users |
| [SmartPilot-快速上手指南.md](./SmartPilot-快速上手指南.md) | 对外终端用户快速上手指南（精简版） |

---

## 🟡 参考文档（有历史价值，但数字可能已过时）

> 以下文档记录了特定时点的审计结论或设计思考。核心观点仍然成立，但**具体行数、文件清单等细节可能已随代码演进失效**，引用时请核对当前代码。

| 文档 | 说明 | 快照日期 | 过时点 |
|------|------|----------|--------|
| [pipelines-审核评价.md](./pipelines-审核评价.md) | pipelines 目录架构专项审核 | ~2026-06 | 行数数据已过时（text-pipeline 已从 2030 行降至 ~830 行），P0 部分建议已落地 |
| [v3.0_ThinkingFramework_Implementation.md](./v3.0_ThinkingFramework_Implementation.md) | v3.0 思维框架实现记录 | 2026-08 | 实现阶段文档，功能已合入主线 |
| [设计-GeneratePython确定性预设-2026-08.md](./设计-GeneratePython确定性预设-2026-08.md) | Python 生成确定性预设设计稿 | 2026-08 | 设计阶段文档，已实现 |
| [设计-多轮消息继承机制-2026-08.md](./设计-多轮消息继承机制-2026-08.md) | 多轮消息继承机制设计稿 | 2026-08 | 设计阶段文档，已实现 |
| [对话流程架构图.md](./对话流程架构图.md) | 对话流程架构说明 | — | 架构概览，细节可能滞后 |
| [图片文件处理流程图.md](./图片文件处理流程图.md) | 图片处理流程说明 | — | 功能概览 |

---

## 🔵 专项解决方案

| 文档 | 说明 |
|------|------|
| [VBA-Error-424-Solution.md](./VBA-Error-424-Solution.md) | VBA 错误 424 专项解决方案 |
| [VBA-Safety-System-Guide.md](./VBA-Safety-System-Guide.md) | VBA 安全沙箱系统完整指南 |
| [VBA-Best-Practices-for-AI.md](./VBA-Best-Practices-for-AI.md) | AI 生成 VBA 代码的最佳实践 |
| [SOLUTION-SUMMARY.md](./SOLUTION-SUMMARY.md) | 错误 424 解决方案总清单 |
| [Web-Integration-Guide.md](./Web-Integration-Guide.md) | Web 集成指南 |
| [安全-沙箱镜像文件归档-2026-08-27.md](./安全-沙箱镜像文件归档-2026-08-27.md) | 安全沙箱镜像归档记录 |

---

## 📌 阅读建议

- **新成员上手**：从「项目审核评价与扩展优化建议」开始，建立全局认知
- **功能使用**：直接查阅对应 Priority 扩展指南
- **问题排查**：专项解决方案文档（VBA 错误、安全沙箱等）
- **历史设计决策**：参考「设计-*.md」系列，理解当时的取舍

---

_本索引由人工维护，新增文档时请同步更新此文件。_
