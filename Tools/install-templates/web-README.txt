SmartPilot 桌面版（浏览器 + 本地桥）
====================================

这是 SmartPilot 的独立运行版本，无需安装 Office 插件：
在浏览器中使用，AI 通过本地桥直接驱动本机已安装的 Excel / PowerPoint。

【运行要求】
- Windows 10 / 11
- 已安装 Microsoft Excel（本地桥通过 COM 控制 Excel）
- Python 3.10 或更高版本（python.org 下载，安装时勾选 "Add Python to PATH"）

【首次运行】
双击 start.bat：
  1. 自动检测并安装 Python 依赖（pywin32、openpyxl，需要联网）
  2. 启动本地桥（端口 9527）并自动打开浏览器
  3. 页面地址：http://127.0.0.1:9527/

【使用说明】
- 桥服务运行期间请保持黑色窗口开启；关闭窗口即停止服务
- AI 模型的 API Key 等在页面「AI 设置」中配置，仅保存在本机
- 若需接入外部 MCP 工具：复制 Config\mcp.example.json 为
  Config\mcp.user.json 后修改

【与插件版的区别】
- 插件版（SmartPilotSetup.exe）：嵌入 Excel / Word / PowerPoint 任务窗格，
  无需 Python，推荐日常使用
- 桌面版（本包）：浏览器中使用，不依赖 Office 宿主，二者可共存
