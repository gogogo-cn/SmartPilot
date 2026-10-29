# SmartPilot 本地代理服务（Local Bridge）

允许 **Web 版 SmartPilot** 在浏览器里连接到本机桌面 Excel，**无需安装 VSTO 插件**。
前端 `UI/js/excel-connector.js` 会自动探测并在 VSTO 缺失时切换到本服务（默认 `http://localhost:9527/api/excel`）。

---

## ⚠️ 最重要的坑：Excel 驱动方式

本地桥要“真正操作 Excel”必须走 **COM 自动化**，由 `py-excel-bridge.py`（Python `win32com`）驱动，
与 Python 版本无关，稳定可用。

> 历史上的 Node 桥（`local-bridge-server.js` + winax）因 winax 3.6.9 在 Node 20+ 无法编译、
> 只能跑模拟模式，已于 2026-10 仓库清理时移除。请使用 `py-excel-bridge.py`。

Python 桥的真实 COM 只在 **Windows + 已装 Excel** 时生效，非 Windows 环境只会进入“无 Excel”的
降级模式。因此“跨平台”仅指 Web UI 本身可在任意系统打开，真正的 Excel 自动化依赖本机 Windows + Excel。

---

## 🚀 快速开始（推荐：Python 桥，真实操作 Excel）

### 前置要求
- **Python 3.8+**（本机已具备，Python 3.14 验证可用）
- **pywin32**：`pip install pywin32`（本机已装，已验证 `import win32com.client` 可用）
- **Microsoft Excel** 已安装并最好保持打开

### 启动
```bash
cd Tools
python py-excel-bridge.py
# 自定义端口： PORT=9531 python py-excel-bridge.py
```

启动成功显示：
```
============================================================
🚀 SmartPilot 本地代理服务 (pywin32 驱动真实 Excel COM)
📍 服务地址: http://localhost:9527
📄 Web界面:  http://localhost:9527/index.html
🔗 API端点:  http://localhost:9527/api/excel/*
============================================================
```

### 打开 Web 界面
浏览器访问 `http://localhost:9527/index.html`。页面加载时 `ExcelConnector.autoConnect()`
会轮询 `GET /api/excel/status` → 命中后 `connectLocalBridge()`，把 `App.excelBridge` 换成 HTTP 代理，
**现有聊天/插入逻辑无需任何改动即可走本地桥**。

### 验证真实连接
```bash
# 状态（pywin32Available=true, mockMode=false 即真实模式）
curl http://localhost:9527/api/excel/status

# 建连（返回 Excel 16.0 之类真实版本）
curl -X POST http://localhost:9527/api/excel/connect -H "Content-Type: application/json" -d "{\"action\":\"test\"}"

# 读真实选区 / 写数据
curl -X POST http://localhost:9527/api/excel/invoke -H "Content-Type: application/json" -d "{\"method\":\"GetOpenWorkbooks\",\"arguments\":[]}"
curl -X POST http://localhost:9527/api/excel/invoke -H "Content-Type: application/json" -d "{\"method\":\"WriteToRange\",\"arguments\":[\"A1\",\"{\\\"values\\\":[[\\\"网页版写入\\\",1,2]]}\"]}"
```

---

## 🔌 已实现的真实 Excel 方法（本地桥契约）

HTTP 调用：`POST /api/excel/invoke`  body `{ "method": "<名称>", "arguments": [...] }`

| 方法 | 参数 | 说明 |
|------|------|------|
| `GetSelectedRange` / `GetSelectedRangeAsync` | — | 读取当前选区（真实数据） |
| `WriteToRange` / `WriteToRangeAsync` | `rangeAddress, dataJson` | 写入二维数据到指定区域 |
| `GetWorkbookInfo` | — | 当前工作簿信息 + Sheet 列表 |
| `GetActiveSheetName` | — | 活动工作表名 |
| `GetOpenWorkbooks` | — | 所有已打开工作簿及 Sheet |
| `ActivateWorkbook` | `name` | 切换到指定工作簿 |
| `ExportToCsv` / `excel_export_csv` | `sheetName, filePath, delimiter, includeHeader` | 导出 CSV |
| `ExecuteVBA` / `ExecuteVBAAsync` | `code` | 写入宏模块并可自动运行首个 Sub |
| `DetectEnvironment` | — | 返回驱动/Excel 连接信息 |
| `GetAllSkills` / `GetMcpStatus` | — | 返回桥模式下的合理降级值 |

> 未实现的方法（如 VSTO 专有的 OCR / RAG / Qdrant / `InsertParsedData` 等）会返回
> `{"success":false,"error":"方法在本地桥(pywin32)模式下未实现: <method>"}`，
> 前端按错误优雅降级。要让这些重方法在 Web 模式可用，
> 需在 `py-excel-bridge.py` 中补充实现，或把前端硬编码的
> `window.chrome.webview.hostObjects.excelBridge` 调用改走 `App.excelBridge`。

---

## ⚙️ 配置

| 环境变量 | 默认值 | 说明 |
|----------|--------|------|
| `PORT` | `9527` | 服务端口（前端默认也探测 9527） |

---

## 🔐 安全

- ✅ 仅监听 `0.0.0.0:9527`，配合前端同源/localhost 访问，不外泄到公网
- ✅ 已开启 CORS（`Access-Control-Allow-Origin: *`）
- ⚠️ 本地使用足够；如需局域网/公网，请另行加认证与 HTTPS

---

## 🐛 故障排除

### 1. 提示“Excel 未运行”
- 先打开 Excel（本桥会尝试 `GetActiveObject`，没有运行实例时也可 `Dispatch` 新建）
- 确认 Excel 没有被“只读/保护”卡住

### 2. 端口被占用（`Address already in use`）
```bash
# 查看并释放 9527
(Get-NetTCPConnection -LocalPort 9527).OwningProcess | ForEach-Object { Stop-Process -Id $_ -Force }
# 或换端口
PORT=9531 python py-excel-bridge.py
```

### 3. `import win32com` 失败
```bash
pip install pywin32
```

### 4. 拿到的是样例数据而非真实 Excel
- 检查 `GET /api/excel/status` 的 `mockMode` 是否为 `false`、`pywin32Available` 是否为 `true`
- 若 `mockMode=true`：说明 pywin32 不可用（未安装或 Python 环境异常），服务降级为模拟模式
- 若 `pywin32Available=false`：本机没装 `pywin32`，见上一条

### 5. CORS 错误
- 务必通过 `http://localhost:9527/...` 访问，**不要**用 `file://` 直接打开页面

---

## 📊 模式对比

| 特性 | VSTO 插件 | Local Bridge (pywin32) | Office.js |
|------|----------|------------------------|-----------|
| 安装复杂度 | ⭐⭐⭐⭐ | ⭐⭐（Python+pywin32） | ⭐（需作为加载项侧载） |
| 真实 Excel 操作 | ✅ 完整 | ✅ 核心方法（选区/写入/工作簿/VBA/CSV） | ✅ 但**不能跑 VBA** |
| 离线可用 | ✅ | ✅ | ❌ 需 Excel 内加载 |
| 平台 | Windows | Windows（真实 COM） | 全平台（Excel Online/桌面） |
| 部署难度 | ⭐⭐⭐⭐ | ⭐⭐ | ⭐ |

---

**享受使用 SmartPilot! 🎉**
