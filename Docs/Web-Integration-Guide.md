# Web版 Excel 连接功能 - 集成指南

> ⚠️ **本文档已过期（2026-10-13 仓库清理时标注）**：文中"本地代理服务"指的是已被删除的
> Node 桥 `Tools/local-bridge-server.js`（winax 在 Node 20+ 只能跑模拟模式，且默认端口为旧的 5173）。
> **现行本地桥为 `Tools/py-excel-bridge.py`（Python + pywin32，端口 9527）**，启动方式与完整契约见
> [`Tools/README-Bridge.md`](../Tools/README-Bridge.md)，一键启动用项目根目录的 `start.bat`。
> 本文仅作为历史集成方案留存。

## 📋 概述

本指南说明如何将 **动态Excel连接功能** 集成到现有的 SmartPilot 插件中，实现 Web 版也能连接并操作桌面 Excel。

## 🎯 目标效果

用户可以在以下 **3种模式** 之间自由切换：

```
┌─────────────────────────────────────────────┐
│  模式1: VSTO原生插件（当前）                 │  ← 完整功能，无需额外配置
├─────────────────────────────────────────────┤
│  模式2: 本地代理服务（新增）✨               │  ← Web版+完整Excel功能
│  启动: python Tools/py-excel-bridge.py      │  ← 现已改用 Python 桥（原 Node 桥已删除）
├─────────────────────────────────────────────┤
│  模式3: Office.js插件（新增）✨              │  ← Excel Online支持
│  安装: Office Store 或 侧载加载             │
└─────────────────────────────────────────────┘
```

## 🔧 集成步骤

### 步骤 1：引入新文件

在 `UI/index.html` 的 `<head>` 中添加：

```html
<!-- 在其他JS文件之后引入 -->
<script src="js/excel-connector.js"></script>
<script src="js/excel-connector-ui.js"></script>
```

**建议位置**：在 `js-app-init.js` 之后

### 步骤 2：添加 UI 容器

在 `index.html` 的合适位置（建议在 header 之后或设置面板中）添加：

```html
<!-- Excel连接管理器 -->
<div id="excel-connector-ui"></div>
```

**推荐位置选项**：

#### 选项 A：Header 区域（推荐）
```html
<!-- 在 index.html 的 <div id="header"> 内部 -->
<div id="header">
    <!-- 现有内容 -->
    
    <!-- 新增：连接管理器 -->
    <div id="excel-connector-ui" style="margin-left: auto;"></div>
</div>
```

#### 选项 B：独立面板
```html
<!-- 在主应用容器之前 -->
<div id="app-container">
    <!-- 新增：顶部连接栏 -->
    <div class="connection-bar">
        <div id="excel-connector-ui"></div>
    </div>
    
    <!-- 现有内容 -->
</div>
```

#### 选项 C：设置面板内
```html
<!-- 在 settings-panel.html 中 -->
<div class="settings-section">
    <h3>🔗 Excel 连接</h3>
    <div id="excel-connector-ui"></div>
</div>
```

### 步骤 3：初始化连接管理器

在 `js-app-init.js` 的 `DOMContentLoaded` 事件处理函数中添加：

```javascript
document.addEventListener('DOMContentLoaded', async function() {
    // ... 现有的初始化代码 ...
    
    // 【新增】初始化 Excel 连接管理器
    if (typeof ExcelConnectorUI !== 'undefined') {
        try {
            await ExcelConnectorUI.init();
            console.log('[Init] ✅ Excel连接管理器已初始化');
        } catch(e) {
            console.warn('[Init] ⚠️ Excel连接管理器初始化失败:', e);
        }
    }
    
    // ... 继续后续初始化 ...
});
```

### 步骤 4：修改现有代码以支持多模式

#### 4.1 修改 Skill 管理 (`app-skill-mgmt.js`)

找到 `saveSkill()` 函数，修改为：

```javascript
async function saveSkill() {
    try {
        let skillData = collectSkillData();
        
        // 【修改】支持多种存储后端
        if (App.excelBridge) {
            // 方式A：通过Bridge保存（VSTO/Local Bridge/Office.js统一接口）
            const result = await App.excelBridge.SaveSkill(JSON.stringify(skillData));
            handleBackendResult(result);
        } else {
            // 方式B：本地存储降级（独立模式）
            await saveSkillToLocal(skillData);
            App.showToast('✅ Skill 已保存（本地模式）', 'success');
        }
        
        // 刷新列表
        refreshSkillsList();
        
    } catch (error) {
        App.showToast('❌ 保存失败: ' + error.message, 'error');
    }
}
```

#### 4.2 修改 MCP 管理 (`app-mcp.js`)

找到 `saveMcpClient()` 函数：

```javascript
async function saveMcpClient() {
    try {
        const config = collectMcpConfig();
        
        // 【修改】支持多种后端
        if (App.excelBridge) {
            const result = await App.excelBridge.AddMcpClientAsync(JSON.stringify(config));
            handleMcpResult(result);
        } else {
            // 本地存储
            saveMcpToLocal(config);
            App.showToast('✅ MCP配置已保存（本地模式）', 'success');
        }
        
    } catch (error) {
        App.showToast('❌ 保存失败: ' + error.message, 'error');
    }
}
```

#### 4.3 修改 VBA 执行 (`skill-executor.js` 或相关文件)

```javascript
async function executeVBA(code) {
    if (!App.excelBridge) {
        // 【新增】优雅降级：显示代码供手动执行
        showCodePreview(code);
        App.showToast('💡 请在Excel中手动执行此代码', 'info');
        return { success: false, code: code };
    }
    
    // 原有逻辑...
}
```

### 步骤 5：添加全局事件监听（可选）

在 `app-bootstrap.js` 或 `app.js` 中添加：

```javascript
// Excel连接状态变更回调
window.App.onConnectionChange = function(connected, mode) {
    console.log(`[Global] Excel连接状态变更: ${connected ? '已连接' : '已断开'} (${mode})`);
    
    if (connected) {
        // 连接成功后的操作
        enableExcelFeatures();
        loadExcelSpecificData();
    } else {
        // 断开连接后的操作
        disableExcelFeatures();
        showStandaloneModeNotice();
    }
};

window.App.onExcelConnected = function(mode) {
    console.log(`[Global] Excel已连接，模式: ${mode}`);
    // 可以在这里触发特定的初始化逻辑
};

window.App.onExcelDisconnected = function() {
    console.log('[Global] Excel已断开');
    // 清理资源等
};
```

## 📦 发布包准备

### Web版独立发布包结构

```
SmartPilot-Web-v3.0/
├── index.html                    # 主界面（web-version-example.html）
├── web-version.html              # 或者使用这个专用版本
├── js/
│   ├── app-bootstrap.js          # 核心框架
│   ├── excel-connector.js        # ✨ 新增：连接管理器
│   ├── excel-connector-ui.js     # ✨ 新增：连接UI组件
│   └── ... 其他现有文件
├── css/
│   └── styles.css
├── Tools/
│   ├── local-bridge-server.js    # ✨ 新增：本地代理服务
│   ├── package.json              # ✨ 新增：依赖配置
│   └── README-Bridge.md          # 使用文档
├── start-web.bat                 # Windows启动脚本
├── start-web.sh                  # Mac/Linux启动脚本
└── README.md                     # Web版使用说明
```

### 创建启动脚本

#### Windows: `start-web.bat`
```batch
@echo off
echo ============================================
echo   SmartPilot Web版 启动器
echo ============================================.

cd /d "%~dp0Tools"

:: 检查Node.js
where node >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [错误] 未检测到Node.js，请先安装:
    echo https://nodejs.org/
    pause
    exit /b 1
)

:: 检查依赖
if not exist "node_modules" (
    echo [提示] 首次运行，正在安装依赖...
    call npm install
)

echo.
echo [启动] 正在启动本地服务...
echo [地址] http://localhost:5173
echo.

call npm start

pause
```

#### Mac/Linux: `start-web.sh`
```bash
#!/bin/bash

echo "============================================"
echo "  SmartPilot Web版 启动器"
echo "============================================"
echo ""

cd "$(dirname "$0")/Tools"

# 检查Node.js
if ! command -v node &> /dev/null; then
    echo "[错误] 未检测到Node.js，请先安装:"
    echo "https://nodejs.org/"
    exit 1
fi

# 检查依赖
if [ ! -d "node_modules" ]; then
    echo "[提示] 首次运行，正在安装依赖..."
    npm install
fi

echo ""
echo "[启动] 正在启动本地服务..."
echo "[地址] http://localhost:5173"
echo ""

npm start
```

## 🧪 测试清单

### 功能测试

- [ ] **自动连接测试**
  - [ ] VSTO模式下自动识别并连接
  - [ ] Local Bridge模式下自动检测服务
  - [ ] Office.js模式下自动初始化
  
- [ ] **手动连接测试**
  - [ ] 点击各模式按钮能正确切换
  - [ ] 输入自定义URL能连接
  - [ ] 断开重连正常工作
  
- [ ] **功能完整性测试**
  - [ ] AI对话在所有模式下可用
  - [ ] OCR识别在所有模式下可用
  - [ ] Skill CRUD在各模式下正常
  - [ ] VBA代码生成和预览正常
  - [ ] 数据读写在连接状态下正常
  
- [ ] **边界情况测试**
  - [ ] Excel未启动时尝试连接
  - [ ] 服务端口被占用
  - [ ] 网络断开时重连
  - [ ] 多次快速切换模式
  
- [ ] **UI/UX测试**
  - [ ] 状态指示准确显示
  - [ ] 加载状态友好提示
  - [ ] 错误信息清晰易懂
  - [ ] 移动端响应式布局

## 🐛 常见问题

### Q1: 集成后原有功能受影响？

**A**: 不会。所有改动都是向后兼容的：
- 如果检测到 VSTO 环境，行为与原来完全一致
- 新增的连接管理器是可选增强
- 所有现有 API 接口保持不变

### Q2: 用户不安装Node.js能用吗？

**A**: 可以。有以下几种选择：
1. **纯独立模式**：不连接Excel，使用基础功能（AI、OCR等）
2. **VSTO模式**：在Excel中使用完整功能
3. **Office.js模式**：Excel Online中使用
4. **提供打包好的Electron版**（可选）：内置Node.js运行时

### Q3: 安全性如何保证？

**A**: 
- 默认仅监听 localhost
- 可选添加认证机制
- CORS 严格限制
- 建议生产环境使用 HTTPS + 反向代理

### Q4: 性能有影响吗？

**A**: 影响极小：
- 连接管理器代码仅 ~15KB（压缩后）
- 仅在需要时加载
- 不影响核心AI/OCR性能
- Local Bridge延迟 <50ms

## 📈 后续优化方向

### Phase 1 (已完成)
- ✅ 统一连接管理器
- ✅ Local Bridge 服务
- ✅ UI 组件
- ✅ 文档和示例

### Phase 2 (计划中)
- [ ] Electron 桌面应用封装
- [ ] 自动更新机制
- [ ] 用户账户系统
- [ ] 配置云同步

### Phase 3 (远期)
- [ ] 多Excel实例支持
- [ ] 远程协作功能
- [ ] 插件市场Web版
- [ ] AI模型云端部署

## 📞 获取帮助

- 📖 **完整文档**: `Tools/README-Bridge.md`
- 🐛 **问题反馈**: GitHub Issues
- 💬 **社区讨论**: Discord/GitHub Discussions
- 📧 **技术支持**: support@smartpilot.ai

---

**祝集成顺利！🎉 如有问题欢迎随时反馈。**