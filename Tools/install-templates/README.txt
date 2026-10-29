SmartPilot for Office - 安装说明
====================================

本包为 Excel / Word / PowerPoint 三件套 AI 加载项的离线安装包（无需管理员权限）。

【安装方式一：免安装（zip 解压版）】
1. 把整个文件夹解压到任意位置（建议路径不含中文/空格，例如 D:\SmartPilot）。
2. 双击 install.bat（或右键 register.ps1 → 用 PowerShell 运行）。
   - 脚本会把证书导入"当前用户\受信任的发布者"，并在注册表 HKCU 下登记三个加载项。
3. 完全退出所有 Office 程序，重新打开 Excel / Word / PowerPoint，加载项自动出现。

【安装方式二：Inno Setup 安装包（setup.exe）】
双击 setup.exe，按向导安装到 %LocalAppData%\Programs\ExcelAIVSTOAddIn，安装结束自动注册。
卸载：控制面板 → 程序和功能 → SmartPilot for Office。

【依赖（目标机需具备，缺失时会提示）】
- .NET Framework 4.7.2 或更高
- Microsoft Visual Studio 2010 Tools for Office Runtime (VSTO Runtime)
  https://www.microsoft.com/en-us/download/details.aspx?id=56961
- Microsoft Edge WebView2 Runtime (Evergreen)
  https://go.microsoft.com/fwlink/p/?LinkId=2124703
- （可选）OCR 功能首次运行会自动下载嵌入式 Python 到 %AppData%\ExcelAIVSTOAddIn\PythonEmbed，
  无需手动安装 Python。

【关于 SmartScreen / 未知发布者警告】
本包使用开发用的自签名临时证书。安装器 exe 与首次加载时 Windows SmartScreen 可能提示
"未知发布者 / Windows 已保护你的电脑"。这是预期内的，点击"仍要运行 / 详细信息 → 仍要运行"即可。
若要消除警告，需使用受信任 CA 签发的 EV/OV 代码签名证书重新签名清单与安装器（见 Docs/DEPLOY.md）。

【卸载】
运行 uninstall.bat（或 setup.exe 的卸载）。

【目录结构】
  Excel\       Excel VSTO 加载项运行时
  Word\        Word VSTO 加载项运行时
  PowerPoint\  PowerPoint COM 加载项运行时
  certs\       签名证书（安装时导入受信任发布者）
  Legal\       许可协议 / 隐私说明 / 第三方声明
  register.ps1 注册脚本（装）
  install.bat  一键安装
  uninstall.bat 一键卸载
