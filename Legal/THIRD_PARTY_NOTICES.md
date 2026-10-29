# 第三方组件声明

SmartPilot 使用下列第三方组件。完整许可证文本随对应的 NuGet 包或安装介质提供。版本以各工程 `PackageReference` 及运行环境为准。

| 组件 | 用途 | 许可证 |
|---|---|---|
| Microsoft WebView2 | 任务窗格界面 | Microsoft 软件许可 |
| Microsoft Office / VSTO Runtime | Excel、Word、PowerPoint 宿主 | Microsoft 软件许可 |
| Newtonsoft.Json | JSON 序列化 | MIT |
| NPOI | Office 文档读写 | Apache-2.0 |
| PDFsharp | PDF 生成 | MIT |
| YamlDotNet | 技能说明解析 | MIT |
| openpyxl | 独立服务读写 xlsx | MIT |
| Python（可选） | 数据分析 / OCR 脚本 | PSF License |

界面字体使用操作系统自带字体，不另附字体许可。

若你启用 Ollama、llama.cpp 或某一云端模型，该运行时或服务不随本安装包分发，适用其各自的条款。

OCR、嵌入式 Python 等可选能力若随包提供模型或二进制，其许可证以该目录内的 NOTICE / LICENSE 文件为准。
