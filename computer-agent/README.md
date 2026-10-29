# Computer Agent — MCP Server

让 LLM 通过 **PaddleOCR + pyautogui + mss** 操作你的电脑。
把每个原子操作做成 MCP Tool，LLM（Trae / Claude Desktop / 任意支持 MCP 的客户端）就能直接驱动电脑 agent。

## 架构

```
┌──────────────────┐   MCP stdio   ┌──────────────────────┐
│ LLM (Trae 等)    │ ◄────────────►│ mcp_server.py        │
└──────────────────┘               │ 19 个原子 Tool      │
                                   └────────┬─────────────┘
                                            │
                    ┌───────────────────────┼───────────────────────┐
                    ▼                       ▼                       ▼
            ┌──────────────┐         ┌──────────────┐         ┌──────────────┐
            │  感知层       │         │  操作层       │         │  定位层       │
            │ PaddleOCR   │         │ pyautogui    │         │ OCR搜索      │
            │ mss 截图    │         │ 键鼠控制      │         │ 图像匹配      │
            └──────────────┘         └──────────────┘         └──────────────┘
```

## 19 个原子 Tool

### 感知层 (4) — 让 Agent "看见"
| Tool | 说明 |
|------|------|
| `screen_capture` | 截图返回 base64 PNG |
| `screen_ocr` | 截图 + OCR，返回带坐标的文本块 |
| `screen_get_info` | 屏幕分辨率、鼠标位置、活动窗口 |
| `screen_describe` | 视觉模型描述（需额外配置） |

### 操作层 (8) — 让 Agent "控制"
| Tool | 说明 |
|------|------|
| `mouse_move` | 移动鼠标 |
| `mouse_click` | 点击（可选坐标） |
| `mouse_drag` | 拖拽 |
| `mouse_scroll` | 滚动滚轮 |
| `key_press` | 单键 |
| `key_hotkey` | 组合键 |
| `type_text` | 输入文本（支持中文） |
| `wait` | 等待若干毫秒 |

### 定位层 (3) — 让 Agent "精准找到"
| Tool | 说明 |
|------|------|
| `locate_text` | OCR 搜索文字，返回中心点坐标 |
| `locate_image` | 模板匹配找图标/图片 |
| `locate_color` | 颜色查找找状态指示器 |

### 辅助层 (4)
| Tool | 说明 |
|------|------|
| `execute_command` | 执行 shell 命令 |
| `open_application` | 启动应用/打开 URL |
| `clipboard_read` | 读剪贴板 |
| `clipboard_write` | 写剪贴板 |

### 系统控制 (1)
| Tool | 说明 |
|------|------|
| `set_safe_mode` | True=只打印不执行，调试用 |

## 快速开始

### 1. 安装依赖（在 sandbox 外或用 venv）

```powershell
cd computer-agent
python -m venv .venv
.\.venv\Scripts\activate
pip install mcp mss pyautogui Pillow pyperclip opencv-python numpy pyyaml

# OCR 需要 PaddleOCR（冷启动约 3-5s，模型约 100MB）
pip install paddlepaddle paddleocr
```

### 2. 配置 Trae MCP

在 Trae 设置 → MCP Servers 中添加：

```json
{
  "mcpServers": {
    "computer-agent": {
      "command": "python",
      "args": ["C:/Users/Administrator/Desktop/addin/ExcelAIVSTOAddIn/computer-agent/mcp_server.py"]
    }
  }
}
```

如果用 venv：

```json
{
  "mcpServers": {
    "computer-agent": {
      "command": "C:/Users/Administrator/Desktop/addin/ExcelAIVSTOAddIn/computer-agent/.venv/Scripts/python.exe",
      "args": ["C:/Users/Administrator/Desktop/addin/ExcelAIVSTOAddIn/computer-agent/mcp_server.py"]
    }
  }
}
```

### 3. 重启 Trae，开始使用

直接跟 LLM 说：
- "打开记事本"
- "在记事本里输入 Hello World"
- "截图看看屏幕上有什么"
- "找到屏幕上写着'文件'的按钮并点击"
- "把鼠标移动到 500, 300"

## 典型工作流

Agent 执行"点击文件菜单"的内部流程：
```
1. screen_get_info          → 确认屏幕状态
2. screen_ocr()             → 识别"文件"文字的位置
3. locate_text("文件")      → 精确拿到中心点坐标 [120, 45]
4. mouse_click(120, 45)     → 执行点击
5. wait(300)                → 等菜单弹出来
6. screen_ocr()             → 验证菜单是否打开
```

## 安全

| 机制 | 说明 |
|------|------|
| pyautogui.FAILSAFE | 鼠标甩到 (0,0) 自动中断所有操作 |
| 坐标校验 | 所有坐标自动 clamp 到屏幕范围内 |
| safe_mode | `set_safe_mode(true)` 后所有操作只打印不执行 |
| 剪贴板粘贴 | 中文输入不依赖 pyautogui 的 typewrite |

## 文件结构

```
computer-agent/
├── mcp_server.py         # MCP Server 入口（注册 19 个 tool）
├── pyproject.toml        # 依赖声明
├── config/default.yaml   # 配置文件
├── tools/                # 4 层 tool 实现
│   ├── perception.py     # 感知层
│   ├── control.py        # 操作层
│   ├── locate.py         # 定位层
│   └── utility.py        # 辅助层
└── core/                 # 底层引擎
    ├── utils.py          # 错误分类 / 坐标校验 / 配置加载
    ├── screen_capture.py # mss 截图
    ├── ocr_engine.py     # PaddleOCR 懒加载封装
    └── mouse_keyboard.py # pyautogui + 安全边界
```

## 设计原则

- **每个组件做它擅长的事**：OCR 引擎只做识别，mss 只做截图，pyautogui 只做键鼠
- **硬编码零容忍**：所有参数可配，安全边界开/关，OCR 引擎可替换
- **MCP 解耦 LLM**：tool schema 自动从 Python 函数签名生成，换 LLM 零成本
