/**
 * ============================================================
 * app-mcp.js — MCP Protocol Manager
 * ============================================================
 * 功能：MCP 协议管理模块，负责 MCP 工具注册、调用、结果解析与本地/远程桥接管理。
 *
 * 核心类/函数：
 *   - registerMcpTools() — 注册 MCP 工具
 *   - callMcpTool() — 调用 MCP 工具
 *   - parseMcpResult() — 解析 MCP 结果
 *   - getMcpClient() — 获取 MCP 客户端
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 *   - bridge-config.js
 *
 * ⚠️ 单一事实源（SINGLE SOURCE OF TRUTH）
 *   本文件（UI/js/app-mcp.js）是 app-mcp.js 的唯一权威版本。
 *   各宿主的运行时副本 <Host>/bin/Debug/UI/js/app-mcp.js（以及 Excel 的
 *   bin/Debug/UI/js/app-mcp.js）均由 Tools/sync-ui-to-debug.js 在构建/同步时
 *   从本文件复制生成，**切勿直接编辑 bin 下的副本**——改动不会回流到源，且会在
 *   下次同步被覆盖，造成"改了没生效/线上用旧代码"的诡异问题。
 *   改动请只改本文件，然后跑 sync-ui-to-debug.js；可用
 *   Tools/verify-app-mcp-copies.js 校验所有 bin 副本是否与本文件一致。
 * ============================================================
 */

(function() {
    'use strict';

    // --- 安全辅助：转义动态内容，防止 innerHTML 拼接导致 XSS ---
    function escapeHtml(text) {
        if (text == null) return '';
        return String(text)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    }

    // --- 来自 app.js:L24129-L24147 ---
    function isVstoWebView2() {
        return !!(window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge && !window.chrome.webview._mockMode);
    }

    // ★ 统一提示函数：优先写入状态栏 fsb-text，降级用 toast
    function mcpToast(message, autoClear) {
        if (App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function') {
            App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">' + message + '</span>', !!autoClear);
        } else if (typeof App.showToast === 'function') {
            App.showToast(message);
        }
    }

    async function refreshMcpStatus(showToastMsg = false) {
        try {
            if (!isVstoWebView2()) {
                const bridgeBase = App.Endpoints ? App.Endpoints.excelBridge.baseUrl : (window.ExcelBridgeConfig ? window.ExcelBridgeConfig.baseUrl : 'http://localhost:9527');
                try {
                    const response = await fetch(`${bridgeBase}/api/mcp/status`);
                    const data = await response.json();
                    if (data && data.success && data.status) {
                        App.mcpStatus = data.status;
                        renderMcpUI();
                        if (showToastMsg) mcpToast('已刷新', true);
                        return;
                    }
                } catch (e) {
                    ChatLog.warn('MCP', '[MCP] Web API获取状态失败，尝试COM桥接:', e.message);
                }
            } else {
                ChatLog.info('MCP', '[MCP] VSTO WebView2 环境，直接使用 COM 桥接获取状态');
            }

            if (App.excelBridge) {
                try {
                    if (showToastMsg) mcpToast('刷新中');
                    const result = await App.excelBridge.GetMcpStatus();
                    const data = parseMcpResult(result);
                    if (data && data.success && data.status) {
                        App.mcpStatus = data.status;
                        renderMcpUI();
                        if (showToastMsg) mcpToast('已刷新', true);
                    } else if (data && !data.success) {
                        // Word/PPT 基类返回 not supported：不算故障，静默用默认空状态
                        const soft = /not supported|bridge-call-failed|0x80020006|未知名称|e is not a function/i.test(String(data.error || data.detail || ''));
                        if (showToastMsg && !soft) mcpToast('状态获取失败: ' + (data.error || ''), true);
                    } else {
                        if (showToastMsg) mcpToast('状态返回异常', true);
                    }
                    return;
                } catch (e) {
                    const msg = (e && e.message) ? e.message : String(e);
                    if (!/e is not a function|0x80020006|未知名称/i.test(msg)) {
                        ChatLog.warn('MCP', '[MCP] COM桥接获取状态失败:', msg);
                    }
                }
            }
            
            if (showToastMsg) {
                if (!isVstoWebView2()) {
                    // 网页版（非 VSTO）依赖本地桥 py-excel-bridge.py，给出明确指引
                    mcpToast('本地桥接服务未启动：请运行项目根目录 start.bat', true);
                } else {
                    mcpToast('本地桥接服务未启动', true);
                }
            }
        } catch (e) {
            if (showToastMsg) mcpToast('获取 MCP 状态失败: ' + (e.message || e), true);
        }

        // 【修复】无论成功失败，都要确保 UI 至少渲染一次默认空状态，避免页面卡白
        if (!App.mcpStatus) {
            App.mcpStatus = {
                serverRunning: false,
                serverPort: 9527,  // 统一为 MCP Server 实际端口（原 5173 是旧 local-bridge-server 遗留）
                serverToolCount: 0,
                totalAvailableTools: 0,
                clients: []
            };
            renderMcpUI();
        }
    }

    // --- 来自 app.js:L24150-L24211 ---
    function renderMcpUI() {
        // Server 状态
        const badge = document.getElementById('mcp-server-status-badge');
        const startBtn = document.getElementById('btn-start-mcp-server');
        const stopBtn = document.getElementById('btn-stop-mcp-server');
        const serverPort = document.getElementById('mcp-server-port');
        const serverTools = document.getElementById('mcp-server-tools');

        const ti = (k) => { try { return (window.App && App.i18n && App.i18n.t) ? App.i18n.t(k) : k; } catch (_) { return k; } };
        if (App.mcpStatus.serverRunning) {
            if (badge) { badge.textContent = ti('mcp.running'); badge.style.background = '#e8f5e9'; badge.style.color = '#4caf50'; }
            if (startBtn) startBtn.style.display = 'none';
            if (stopBtn) stopBtn.style.display = 'block';
        } else {
            if (badge) { badge.textContent = ti('mcp.notRunning'); badge.style.background = '#eee'; badge.style.color = '#999'; }
            if (startBtn) startBtn.style.display = 'block';
            if (stopBtn) stopBtn.style.display = 'none';
        }
        if (serverPort) serverPort.textContent = App.mcpStatus.serverPort;
        if (serverTools) serverTools.textContent = App.mcpStatus.serverToolCount;

        // Clients 列表（只显示外部MCP客户端，不含本地Excel）
        const clientsList = document.getElementById('mcp-clients-list');
        if (clientsList) {
            const externalClients = (App.mcpStatus.clients || []);
            if (externalClients.length > 0) {
                clientsList.innerHTML = externalClients.map(c => {
                    const statusColor = c.connected ? '#4caf50' : '#f44336';
                    const statusText = c.connected ? '已连接' : '未连接';
                    const errorMsg = c.errorMessage ? `<div style="color:#f44336; font-size:11px;">错误: ${escapeHtml(c.errorMessage)}</div>` : '';
                    return `<div style="display: flex; justify-content: space-between; align-items: center; padding: 10px; margin-bottom: 8px; background: white; border-radius: 6px; border: 1px solid #e0e0e0;">
                            <div style="flex: 1;">
                                <div style="font-weight: 500; font-size: 13px;">${escapeHtml(c.name)}</div>
                                <div style="font-size: 11px; color: #666;">${escapeHtml(c.description || '')} | 工具: ${escapeHtml(c.toolCount)} | ${escapeHtml(c.transport)}</div>
                                ${errorMsg}
                            </div>
                            <div style="display: flex; gap: 6px; align-items: center;">
                                <span style="font-size: 11px; color: ${statusColor};">${statusText}</span>
                                <button class="btn-secondary" onclick="editMcpClient('${escapeHtml(c.name)}')" style="padding: 4px 8px; font-size: 11px;">编辑</button>
                                <button class="btn-secondary" onclick="toggleMcpClient('${escapeHtml(c.name)}', ${c.connected})" style="padding: 4px 8px; font-size: 11px; ${c.connected ? 'background: #ff9800; color: white;' : 'background: #4caf50; color: white;'}">${c.connected ? '断开' : '连接'}</button>
                                <button class="btn-secondary" onclick="removeMcpClient('${escapeHtml(c.name)}')" style="padding: 4px 8px; font-size: 11px; background: #f44336; color: white;">移除</button>
                            </div>
                        </div>`;
                }).join('');
            } else {
                clientsList.innerHTML = '<div style="text-align: center; color: #999; padding: 20px; font-size: 13px;">暂无已连接的 MCP Client</div>';
            }
        }

        // 工具列表
        const toolsCount = document.getElementById('mcp-tools-count');
        const toolsList = document.getElementById('mcp-tools-list');
        if (toolsCount) toolsCount.textContent = App.mcpStatus.totalAvailableTools;
        if (toolsList) {
            // 【修复】三态显示：
            //   1) _mcpTools 已加载且非空 → 按 source 分组渲染
            //   2) totalAvailableTools > 0 但 _mcpTools 尚未加载 → 显示“工具加载中”
            //   3) 无任何工具 → 显示“暂无可用工具”
            const hasLoadedTools = window._mcpTools && window._mcpTools.length > 0;
            if (hasLoadedTools) {
                // 按 source 分组显示工具
                const toolsBySource = {};
                // 去重：使用工具名称作为唯一标识
                const seen = new Set();
                window._mcpTools.forEach(t => {
                    const fullName = t.name || '';
                    if (seen.has(fullName)) return;
                    seen.add(fullName);
                    const sepIdx = fullName.indexOf('__');
                    // 优先使用后端下发的 source 字段，兜底解析工具名前缀
                    const source = t.source || (sepIdx >= 0 ? fullName.substring(0, sepIdx) : 'other');
                    const toolName = sepIdx >= 0 ? fullName.substring(sepIdx + 2) : fullName;
                    if (!toolsBySource[source]) toolsBySource[source] = [];
                    toolsBySource[source].push({ name: fullName, shortName: toolName, description: t.description || '' });
                });

                let html = '<div style="padding: 5px;">';
                for (const [source, tools] of Object.entries(toolsBySource)) {
                    // 本地工具源（excel-local / ppt-local / word-local...）统一归为本地分组，新工具集自动适配
                    const isLocal = source.endsWith('-local');
                    const _t = (k, v) => (window.App && App.i18n && App.i18n.t) ? App.i18n.t(k, v) : k;
                    const sourceLabelMap = { 'excel-local': ' ' + _t('tools.localExcel'), 'ppt-local': ' ' + _t('tools.localPpt'), 'word-local': ' ' + _t('tools.localWord') };
                    const sourceLabel = isLocal ? (sourceLabelMap[source] || (' ' + _t('tools.localPrefix', { name: source.replace('-local', '') }))) : '🔗 ' + source;
                    const hasBridge = !!App.excelBridge;
                    const mockBadge = isLocal
                        ? (hasBridge
                            ? '<span style="background:#4caf50; color:white; font-size:9px; padding:1px 5px; border-radius:3px; margin-left:6px;">' + _t('tools.real') + '</span>'
                            : '<span style="background:#ff9800; color:white; font-size:9px; padding:1px 5px; border-radius:3px; margin-left:6px;">' + _t('tools.mock') + '</span>')
                        : '';
                    html += `<div style="margin-bottom: 8px;">
                        <div style="font-weight: 600; font-size: 12px; color: #333; cursor: pointer; padding: 4px 0;" 
                             onclick="this.nextElementSibling.style.display = this.nextElementSibling.style.display === 'none' ? 'block' : 'none'">
                            ${sourceLabel} (${tools.length}) ${mockBadge} <span style="color:#999; font-size:10px;">▼ ${_t('tools.expand')}</span>
                        </div>
                        <div style="display: none; padding-left: 8px;">`;
                    tools.forEach(t => {
                        const escapedName = escapeHtml(t.name);
                        const escapedDesc = escapeHtml((t.description || '').substring(0, 60));
                        html += `<div style="display: flex; justify-content: space-between; align-items: center; padding: 3px 0; border-bottom: 1px solid #f0f0f0;">
                            <div style="flex: 1; font-size: 11px; color: #555;" title="${escapedDesc}">
                                <b>${escapeHtml(t.shortName)}</b>
                            </div>
                            <button onclick="window.quickCallMcpTool('${escapedName}')" 
                                    style="padding: 2px 8px; font-size: 10px; background: #2196F3; color: white; border: none; border-radius: 3px; cursor: pointer;">${_t('tools.call')}</button>
                        </div>`;
                    });
                    html += `</div></div>`;
                }
                html += '</div>';
                toolsList.innerHTML = html;
            } else if (App.mcpStatus.totalAvailableTools > 0) {
                // 状态显示有工具但前端尚未加载完，提示加载中并触发一次加载
                toolsList.innerHTML = '<div style="text-align: center; color: #999; padding: 20px;">⏳ ' + ((window.App && App.i18n && App.i18n.t) ? App.i18n.t('tools.loading') : '工具加载中，请稍候...') + '</div>';
                if (typeof loadMcpToolsForAI === 'function' && !App._mcpToolsLoaded) {
                    loadMcpToolsForAI();
                }
            } else {
                toolsList.innerHTML = '<div style="text-align: center; color: #999; padding: 20px;">暂无可用工具</div>';
            }
        }
    }

    // --- 来自 app.js:L24214-L24217 ---
    function parseMcpResult(result) {
        if (!result || typeof result !== 'string') return null;
        try { return JSON.parse(result); } catch (e) { return null; }
    }

    // --- 来自 app.js:L24219-L24234 ---
    async function startMcpServer() {
        try {
            if (!App.excelBridge) { mcpToast('Excel 桥接未连接', true); return; }
            mcpToast('正在启动 MCP Server...');
            const result = await App.excelBridge.StartMcpServerAsync();
            const data = parseMcpResult(result);
            if (data) {
                mcpToast(data.message || data.error || 'Server 已启动', true);
            } else {
                mcpToast('Server 已启动: ' + (typeof result === 'string' ? result : 'OK'), true);
            }
            await refreshMcpStatus();
        } catch (e) {
            mcpToast('启动 MCP Server 失败: ' + (e.message || e), true);
        }
    }

    // --- 来自 app.js:L24237-L24251 ---
    async function stopMcpServer() {
        try {
            if (!App.excelBridge) { mcpToast('Excel 桥接未连接', true); return; }
            const result = await App.excelBridge.StopMcpServerAsync();
            const data = parseMcpResult(result);
            if (data) {
                mcpToast(data.message || data.error || 'Server 已停止', true);
            } else {
                mcpToast('Server 已停止', true);
            }
            await refreshMcpStatus();
        } catch (e) {
            mcpToast('停止 MCP Server 失败: ' + (e.message || e), true);
        }
    }

    // --- 来自 app.js:L24254-L24267 ---
    function showAddMcpClientDialog() {
        document.getElementById('mcp-client-dialog').style.display = 'flex';
        document.getElementById('mcp-client-dialog-title').textContent = '添加 MCP Server';
        document.getElementById('mcp-client-edit-name').value = '';
        document.getElementById('mcp-client-name').value = '';
        document.getElementById('mcp-client-description').value = '';
        document.getElementById('mcp-client-transport').value = 'stdio';
        document.getElementById('mcp-client-command').value = '';
        document.getElementById('mcp-client-args').value = '';
        document.getElementById('mcp-client-env').value = '';
        document.getElementById('mcp-client-url').value = '';
        document.getElementById('mcp-client-headers').value = '';
        document.getElementById('mcp-client-env-http').value = '';
        document.getElementById('mcp-client-enabled').checked = true;
        onMcpTransportChange();
    }

    // 编辑 MCP Client
    function editMcpClient(name) {
        // 从当前状态中找到该 client 的完整配置
        var clients = App.mcpStatus && App.mcpStatus.clients ? App.mcpStatus.clients : [];
        var client = clients.find(function(c) { return c.name === name; });
        if (!client) { mcpToast('未找到配置: ' + name, true); return; }

        document.getElementById('mcp-client-dialog').style.display = 'flex';
        document.getElementById('mcp-client-dialog-title').textContent = '编辑 MCP Server';
        document.getElementById('mcp-client-edit-name').value = name; // 记录原始名称
        document.getElementById('mcp-client-name').value = client.name || '';
        document.getElementById('mcp-client-description').value = client.description || '';
        document.getElementById('mcp-client-transport').value = client.transport || 'stdio';
        document.getElementById('mcp-client-command').value = client.command || '';
        document.getElementById('mcp-client-args').value = (client.args || []).join('\n');
        document.getElementById('mcp-client-url').value = client.url || '';

        // Headers
        var headersText = '';
        if (client.headers) {
            headersText = Object.keys(client.headers).map(function(k) { return k + '=' + client.headers[k]; }).join('\n');
        }
        document.getElementById('mcp-client-headers').value = headersText;

        // 环境变量
        var envText = '';
        if (client.env) {
            envText = Object.keys(client.env).map(function(k) { return k + '=' + client.env[k]; }).join('\n');
        }
        document.getElementById('mcp-client-env').value = envText;
        document.getElementById('mcp-client-env-http').value = envText;

        document.getElementById('mcp-client-enabled').checked = client.enabled !== false;
        onMcpTransportChange();
    }

    // --- 来自 app.js:L24270-L24272 ---
    function closeMcpClientDialog() {
        document.getElementById('mcp-client-dialog').style.display = 'none';
    }

    // --- 来自 app.js:L24275-L24279 ---
    function onMcpTransportChange() {
        const transport = document.getElementById('mcp-client-transport').value;
        document.getElementById('mcp-client-stdio-settings').style.display = (transport === 'stdio') ? 'block' : 'none';
        document.getElementById('mcp-client-http-settings').style.display = (transport === 'http') ? 'block' : 'none';
    }

    // --- 来自 app.js:L24282-L24331 ---
    async function saveMcpClient() {
        try {
            if (!App.excelBridge) { mcpToast('Excel 桥接未连接'); return; }

            const name = document.getElementById('mcp-client-name').value.trim();
            const editName = document.getElementById('mcp-client-edit-name').value.trim(); // 原始名称（编辑模式）
            const transport = document.getElementById('mcp-client-transport').value;

            if (!name) { mcpToast('请输入 Server 名称'); return; }

            const config = {
                name: name,
                description: document.getElementById('mcp-client-description').value.trim(),
                transport: transport,
                enabled: document.getElementById('mcp-client-enabled').checked
            };

            if (transport === 'stdio') {
                const command = document.getElementById('mcp-client-command').value.trim();
                if (!command) { mcpToast('请输入命令'); return; }
                config.command = command;
                config.args = document.getElementById('mcp-client-args').value.split('\n').map(s => s.trim()).filter(s => s);
            } else {
                const url = document.getElementById('mcp-client-url').value.trim();
                if (!url) { mcpToast('请输入 URL'); return; }
                config.url = url;
                const headersText = document.getElementById('mcp-client-headers').value.trim();
                if (headersText) {
                    config.headers = {};
                    headersText.split('\n').forEach(line => {
                        const idx = line.indexOf('=');
                        if (idx > 0) config.headers[line.substring(0, idx).trim()] = line.substring(idx + 1).trim();
                    });
                }
            }

            // 环境变量（两种模式通用，KEY=VALUE 格式，每行一条）
            const envElId = transport === 'stdio' ? 'mcp-client-env' : 'mcp-client-env-http';
            const envText = document.getElementById(envElId).value.trim();
            if (envText) {
                config.env = {};
                envText.split('\n').forEach(function(line) {
                    line = line.trim();
                    if (!line || line.startsWith('#')) return;
                    const idx = line.indexOf('=');
                    if (idx > 0) config.env[line.substring(0, idx).trim()] = line.substring(idx + 1).trim();
                });
            }

            mcpToast('正在保存 ' + name + '...');

            // 编辑模式：先移除旧配置（名称变了或相同都需要）
            if (editName && editName !== name) {
                await App.excelBridge.RemoveMcpClientAsync(editName);
            } else if (editName) {
                await App.excelBridge.RemoveMcpClientAsync(editName);
            }

            const result = await App.excelBridge.AddMcpClientAsync(JSON.stringify(config));
            const data = parseMcpResult(result);
            mcpToast((data && (data.message || data.error)) || (typeof result === 'string' ? result : '已保存，请在列表中点击「连接」'), true);
            closeMcpClientDialog();
            // 重置工具缓存，下次 AI 请求时重新加载
            App._mcpToolsLoaded = false;
            window._mcpTools = [];
            await refreshMcpStatus();
            await loadMcpToolsForAI();

        } catch (e) {
            mcpToast('保存 MCP Client 失败: ' + (e.message || e));
        }
    }

    // --- 来自 app.js:L24334-L24349 ---
    async function testMcpClient(name) {
        try {
            // 【优先Web模式】通过本地桥接API测试
            const bridgeBase = App.Endpoints ? App.Endpoints.excelBridge.baseUrl : (window.ExcelBridgeConfig ? window.ExcelBridgeConfig.baseUrl : 'http://localhost:9527');
            try {
                mcpToast('正在测试连接: ' + name + '...', true);
                const response = await fetch(`${bridgeBase}/api/mcp/test`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ name })
                });
                const data = await response.json();
                if (data && data.connected) {
                    mcpToast(name + ' 连接正常', true);
                } else {
                    mcpToast('连接失败: ' + ((data && data.error) || '未知'), true);
                }
                await refreshMcpStatus();
                return;
            } catch (e) {
                ChatLog.warn('MCP', '[MCP] Web API测试失败，尝试COM桥接:', e.message);
            }
            
            // 【后备VSTO模式】通过COM桥接
            if (App.excelBridge) {
                try {
                    mcpToast('正在测试连接: ' + name + '...', true);
                    const result = await App.excelBridge.TestMcpClientConnectionAsync(name);
                    const data = parseMcpResult(result);
                    if (data && data.connected) {
                        mcpToast(name + ' 连接正常', true);
                    } else {
                        mcpToast('连接失败: ' + ((data && data.error) || '未知'), true);
                    }
                    await refreshMcpStatus();
                    return;
                } catch (e) {
                    ChatLog.warn('MCP', '[MCP] COM桥接测试失败:', e.message);
                }
            }
            
            mcpToast('本地桥接服务未启动', true);
        } catch (e) {
            mcpToast('测试失败: ' + (e.message || e), true);
        }
    }

    async function testMcpClientFromDialog() {
        try {
            if (!App.excelBridge) { mcpToast('Excel 桥接未连接'); return; }
    
            const name = document.getElementById('mcp-client-name').value.trim();
            const transport = document.getElementById('mcp-client-transport').value;
            if (!name) { mcpToast('请先输入 Server 名称'); return; }
    
            const config = {
                name: name,
                description: document.getElementById('mcp-client-description').value.trim(),
                transport: transport,
                enabled: document.getElementById('mcp-client-enabled').checked
            };
    
            if (transport === 'stdio') {
                config.command = document.getElementById('mcp-client-command').value.trim();
                if (!config.command) { mcpToast('请输入命令'); return; }
                config.args = document.getElementById('mcp-client-args').value.split('\n').map(s => s.trim()).filter(s => s);
            } else {
                config.url = document.getElementById('mcp-client-url').value.trim();
                if (!config.url) { mcpToast('请输入 URL'); return; }
            }
    
            // 环境变量
            const envElId = transport === 'stdio' ? 'mcp-client-env' : 'mcp-client-env-http';
            const envText = document.getElementById(envElId).value.trim();
            if (envText) {
                config.env = {};
                envText.split('\n').forEach(function(line) {
                    line = line.trim();
                    if (!line || line.startsWith('#')) return;
                    const idx = line.indexOf('=');
                    if (idx > 0) config.env[line.substring(0, idx).trim()] = line.substring(idx + 1).trim();
                });
            }
    
            mcpToast('🧪 正在测试 ' + name + '...');
    
            // 直接发送完整配置进行测试（不保存到 mcp.user.json）
            const result = await App.excelBridge.TestMcpClientConnectionAsync(JSON.stringify(config));
            const data = parseMcpResult(result);
            if (data && data.connected) {
                mcpToast('✅ ' + name + ' 连接正常，工具数: ' + (data.toolCount || 0), true);
            } else {
                mcpToast('❌ 连接失败: ' + ((data && data.error) || '未知'), true);
            }
    
        } catch (e) {
            mcpToast('测试失败: ' + (e.message || e));
        }
    }

    // --- 来自 app.js:L24391-L24402 ---
    async function removeMcpClient(name) {
        if (!confirm('确定要移除 ' + name + ' 吗？')) return;
        try {
            if (!App.excelBridge) { mcpToast('Excel 桥接未连接'); return; }
            const result = await App.excelBridge.RemoveMcpClientAsync(name);
            const data = parseMcpResult(result);
            mcpToast((data && (data.message || data.error)) || (typeof result === 'string' ? result : '已移除'), true);
            await refreshMcpStatus();
        } catch (e) {
            mcpToast('移除失败: ' + (e.message || e));
        }
    }

    // 连接/断开切换
    async function toggleMcpClient(name, isConnected) {
        try {
            if (!App.excelBridge) { mcpToast('Excel 桥接未连接'); return; }

            if (isConnected) {
                // 断开
                mcpToast('正在断开 ' + name + '...');
                const result = await App.excelBridge.DisconnectMcpClientAsync(name);
                const data = parseMcpResult(result);
                mcpToast((data && (data.message || data.error)) || '已断开', true);
            } else {
                // 连接：从当前状态中取出完整配置重新发送
                var clients = App.mcpStatus && App.mcpStatus.clients ? App.mcpStatus.clients : [];
                var client = clients.find(function(c) { return c.name === name; });
                if (!client) { mcpToast('未找到配置: ' + name, true); return; }

                var config = { name: name, description: client.description || '', transport: client.transport || 'stdio', enabled: true };
                if (client.transport === 'stdio') {
                    config.command = client.command || '';
                    config.args = client.args || [];
                } else {
                    config.url = client.url || '';
                    if (client.headers) config.headers = client.headers;
                }
                if (client.env) config.env = client.env;

                mcpToast('正在连接 ' + name + '...');
                const result = await App.excelBridge.AddMcpClientAsync(JSON.stringify(config));
                const data = parseMcpResult(result);
                mcpToast((data && (data.message || data.error)) || '已连接', true);
            }

            await refreshMcpStatus();
            await loadMcpToolsForAI();
        } catch (e) {
            mcpToast('操作失败: ' + (e.message || e));
        }
    }

    // --- 来自 app.js:L24405-L24421 ---
    async function autoReconnectMcpClients() {
        try {
            // 【优先Web模式】通过本地桥接API连接
            const bridgeBase = App.Endpoints ? App.Endpoints.excelBridge.baseUrl : (window.ExcelBridgeConfig ? window.ExcelBridgeConfig.baseUrl : 'http://localhost:9527');
            try {
                mcpToast('正在连接 MCP Clients...');
                const response = await fetch(`${bridgeBase}/api/mcp/connect-all`, { method: 'POST' });
                const data = await response.json();
                if (data && data.success) {
                    mcpToast('正在后台连接，请稍候...');
                    // 等待3秒后刷新状态（给后台连接留时间）
                    setTimeout(async () => {
                        await refreshMcpStatus(true);
                    }, 3000);
                } else {
                    mcpToast('连接失败: ' + (data && data.error || '未知'), true);
                }
                return;
            } catch (e) {
                ChatLog.warn('MCP', '[MCP] Web API连接失败，尝试COM桥接:', e.message);
            }
            
            // 【后备VSTO模式】通过COM桥接
            if (App.excelBridge) {
                try {
                    const result = await App.excelBridge.AutoReconnectMcpClientsAsync();
                    const data = parseMcpResult(result);
                    if (!data || !data.success) {
                        mcpToast('重连失败: ' + ((data && data.error) || '未知'), true);
                    }
                    // 重置工具缓存以重新加载
                    App._mcpToolsLoaded = false;
                    window._mcpTools = [];
                    await refreshMcpStatus();
                    await loadMcpToolsForAI();
                    return;
                } catch (e) {
                    ChatLog.warn('MCP', '[MCP] COM桥接重连失败:', e.message);
                }
            }
            
            mcpToast('本地桥接服务未启动', true);
        } catch (e) {
            mcpToast('重连失败: ' + (e.message || e));
        }
    }

    // ==================== MCP 工具自动加载与 AI 集成 ====================

    // 缓存 MCP 工具列表，用于自动注入 AI 请求
    window._mcpTools = [];
    let _mcpToolsLoadingPromise = null;

    // --- 来自 app.js:L24430-L24452 ---
    async function loadMcpToolsForAI() {
        if (App._mcpToolsLoaded) return;
        if (_mcpToolsLoadingPromise) return _mcpToolsLoadingPromise;
        _mcpToolsLoadingPromise = (async () => {
            try {
                window._mcpTools = [];
                if (!isVstoWebView2()) {
                    const bridgeBase = App.Endpoints ? App.Endpoints.excelBridge.baseUrl : (window.ExcelBridgeConfig ? window.ExcelBridgeConfig.baseUrl : 'http://localhost:9527');
                    try {
                        const response = await fetch(`${bridgeBase}/api/mcp/tools`);
                        const data = await response.json();
                        if (data && data.success && data.tools && data.tools.length > 0) {
                            window._mcpTools = data.tools.map(buildMcpToolMeta);
                            ChatLog.info('MCP', '✅ [Web] 已加载 ' + window._mcpTools.length + ' 个 MCP 工具供 AI 使用');
                            App._mcpToolsLoaded = true;
                            renderMcpUI();
                            return;
                        }
                    } catch (e) {
                        ChatLog.warn('MCP', '[MCP] Web API加载工具失败，尝试COM桥接:', e.message);
                    }
                } else {
                    ChatLog.info('MCP', '[MCP] VSTO WebView2 环境，直接使用 COM 桥接加载工具');
                }

                if (App.excelBridge) {
                    try {
                        const result = await App.excelBridge.GetAvailableMcpTools();
                        const data = parseMcpResult(result);
                        if (data && data.success && data.tools && data.tools.length > 0) {
                            window._mcpTools = data.tools.map(buildMcpToolMeta);
                            ChatLog.info('MCP', '✅ 已加载 ' + window._mcpTools.length + ' 个 MCP 工具供 AI 使用');
                            App._mcpToolsLoaded = true;
                            renderMcpUI();
                            return;
                        }
                        // 桥包装失败返回的 JSON：留给外层重试，勿当空工具告警
                        if (data && data.error === 'bridge-call-failed') return;
                    } catch (e) {
                        const msg = (e && e.message) ? e.message : String(e);
                        if (!/e is not a function|0x80020006|未知名称/i.test(msg)) {
                            ChatLog.warn('MCP', '[MCP] COM桥接加载工具失败:', msg);
                        }
                        return; // 留给 loadMcpToolsAfterBridgeReady 延迟重试
                    }
                }
                
                ChatLog.warn('MCP', '️ MCP 工具为空：Web API 和 COM 桥接均未返回工具');
            } catch (e) {
                ChatLog.warn('MCP', '加载 MCP 工具失败:', e);
            } finally {
                _mcpToolsLoadingPromise = null;
                // 【修复】工具加载失败或为空时，仍要刷新 UI（让“暂无可用工具”显示出来）
                if (App.mcpStatus) {
                    renderMcpUI();
                }
            }
        })();
        return _mcpToolsLoadingPromise;
    }

    // --- 来自 app.js:L24455-L24470 ---
    function buildToolParameters(toolInfo) {
        const properties = {};
        const required = [];
        if (toolInfo.parameters && toolInfo.parameters.length > 0) {
            toolInfo.parameters.forEach(p => {
                properties[p.name] = {
                    type: p.type || 'string',
                    description: p.description || ''
                };
                if (p.required) required.push(p.name);
            });
        }
        const schema = { type: 'object', properties: properties };
        if (required.length > 0) schema.required = required;
        return schema;
    }

    // 构建工具元数据：保留 excel_tools.json 下发的完整元数据（displayName/intentKeywords/operationType 等）
    function buildMcpToolMeta(t) {
        const meta = {
            name: t.source + '__' + t.name,
            source: t.source || 'other',
            description: '[' + (t.source || 'mcp') + '] ' + (t.description || t.name),
            parameters: buildToolParameters(t)
        };
        if (t.displayName) meta.displayName = t.displayName;
        if (Array.isArray(t.intentKeywords)) meta.intentKeywords = t.intentKeywords;
        if (t.operationType) meta.operationType = t.operationType;
        if (typeof t.isWrite === 'boolean') meta.isWrite = t.isWrite;
        if (typeof t.isDangerous === 'boolean') meta.isDangerous = t.isDangerous;
        if (t.dangerSeverity) meta.dangerSeverity = t.dangerSeverity;
        if (t.dangerMessage) meta.dangerMessage = t.dangerMessage;
        if (typeof t.layer === 'number') meta.layer = t.layer;
        return meta;
    }

    // 核心实现
    // 遍历 tool.inputSchema.required，对"路径/文件"类必填参数在为空时从环境填默认。
    // 这是"永久可靠"的基础：避免每个导出/保存/截图类工具都单独 patch。
    const _PATH_FIELD_PATTERN = /path|路径|file|文件名|filename|output|输出|保存|directory|目录|image|图片|attachment|附件|dest|target.*file|save.*as|write.*to/i;
    function _isPathLikeField(prop) {
        if (!prop) return false;
        const desc = String(prop.description || '');
        const name = String(prop.name || '');
        return _PATH_FIELD_PATTERN.test(desc) || _PATH_FIELD_PATTERN.test(name) ||
            prop.format === 'uri' || prop.format === 'path';
    }
    function _isEmpty(v) { return v === undefined || v === null || (typeof v === 'string' && v.trim() === ''); }
    function _inferExtFromTool(toolName, prop) {
        const lower = String(toolName || '').toLowerCase();
        const d = String(prop?.description || '').toLowerCase();
        if (lower.includes('csv') || d.includes('csv')) return '.csv';
        if (lower.includes('pdf') || d.includes('pdf')) return '.pdf';
        if (lower.includes('image') || lower.includes('screenshot') || lower.includes('chart') || d.includes('图片') || d.includes('截图')) return '.png';
        if (lower.includes('xlsx') || lower.includes('workbook') || lower.includes('excel')) return '.xlsx';
        if (lower.includes('json')) return '.json';
        if (lower.includes('txt') || d.includes('文本')) return '.txt';
        return '.bin';
    }
    function _safeFilename(name) {
        return String(name || 'output').replace(/[\\/:*?"<>|]/g, '_').trim() || 'output';
    }
    function _timestamp() {
        const d = new Date();
        const pad = (n) => String(n).padStart(2, '0');
        return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
    }
    // 【环境路径解析】从 ctx.turn 拿用户原始消息，识别"桌面/下载/临时"意图，调用现有桥方法拿真实目录
    // 用户指令里的"桌面/下载"意图优先级高于工作簿目录——否则用户说"保存到桌面"永远被工作簿目录拦截
    async function _resolveDirForField(toolName, args, ctx) {
        const bridge = App?.excelBridge;
        const userText = String(ctx?.turn?.userMessage || ctx?.turn?.content || '');
        // 1) 用户指令意图优先（桌面/下载/临时）
        const wantsDesktop = /(桌面|desk\s*top|desktop)/i.test(userText);
        const wantsDownloads = /(下载|downloads?)/i.test(userText);
        const wantsTemp = /(临时|temp|tmp)/i.test(userText);
        if (wantsDesktop && typeof bridge?.GetDesktopPath === 'function') {
            try { return { dir: await bridge.GetDesktopPath(), src: 'desktop' }; } catch (_e) { /* fall through */ }
        }
        if (wantsDownloads && typeof bridge?.GetDownloadsPath === 'function') {
            try { return { dir: await bridge.GetDownloadsPath(), src: 'downloads' }; } catch (_e) { /* fall through */ }
        }
        if (wantsTemp && typeof bridge?.GetTempPath === 'function') {
            try { return { dir: await bridge.GetTempPath(), src: 'temp' }; } catch (_e) { /* fall through */ }
        }
        // 2) 降级：GetOpenWorkbooks 拿当前工作簿完整路径 → 拆目录
        try {
            if (bridge?.GetOpenWorkbooks) {
                const res = await bridge.GetOpenWorkbooks();
                const data = typeof res === 'string' ? safeJSON(res) : res;
                if (data && Array.isArray(data.workbooks)) {
                    const active = data.workbooks.find(w => w.isActive) || data.workbooks[0];
                    if (active && active.path) {
                        const sep = active.path.includes('\\') ? '\\' : '/';
                        return { dir: active.path.substring(0, active.path.lastIndexOf(sep)), src: 'workbook' };
                    }
                }
            }
        } catch (_e) { /* ignore */ }
        // 3) 最终降级
        if (typeof bridge?.GetTempPath === 'function') {
            try { return { dir: await bridge.GetTempPath(), src: 'temp' }; } catch (_e) { /* fall through */ }
        }
        return { dir: '.', src: 'fallback' };
    }
    function safeJSON(s) { try { return JSON.parse(s); } catch { return null; } }
    // 是否为"模型编造的占位符路径"（前端与后端同语义；toolRegistryUtil 未就绪时用等价内联兜底）
    function _isPlaceholderPathJS(v) {
        try {
            const u = App && App.ChatV2 && App.ChatV2.toolRegistryUtil;
            if (u && typeof u.isPlaceholderPath === 'function') return u.isPlaceholderPath(v);
        } catch (_e) { /* ignore */ }
        return /(your[\s_-]?user|yourname|your[\s_-]?name|<[^<>/]{1,24}>|\{[^{}/]{1,24}\}|%username%|用户名|用户目录|your[\s_-]?path|path[\\/]+to[\\/]|\.\.\.|…)/i.test(String(v || ''));
    }
    // 取路径里的"用户目录段"：C:\Users\X\... / /Users/X/... / /home/X/...
    function _pathUserSegment(p) {
        const s = String(p || '');
        const m = s.match(/^[A-Za-z]:[\\/]Users[\\/]([^\\/]+)/i) || s.match(/^[\\/](?:Users|home)[\\/]([^\\/]+)/i);
        return m ? m[1] : '';
    }
    async function _completeRequiredParamsIfNeeded(args, toolName, ctx) {
        const tool = ctx?.tool;
        if (!tool || !tool.inputSchema) return;
        const schema = tool.inputSchema;
        const props = (schema.properties && typeof schema.properties === 'object') ? schema.properties : {};
        const required = Array.isArray(schema.required) ? schema.required : [];
        if (required.length === 0) return;
        // 惰性取"真实用户目录段"，用于识破模型编造的 C:\Users\<假用户>\...（最多取一次）
        const _segCache = { done: false, seg: '' };
        const _realUserSeg = async () => {
            if (_segCache.done) return _segCache.seg;
            _segCache.done = true;
            try {
                if (typeof App?.excelBridge?.GetDesktopPath === 'function') {
                    const d = String(await App.excelBridge.GetDesktopPath() || '');
                    _segCache.seg = _pathUserSegment(d) || ((d.match(/[\\/]Users[\\/]([^\\/]+)/i) || [])[1] || '');
                }
            } catch (_e) { /* ignore */ }
            return _segCache.seg;
        };
        for (const key of required) {
            const prop = props[key] || {};
            if (!_isPathLikeField(prop)) continue;  // 非路径类必填（如 cell/value/range/script），不补
            if (!_isEmpty(args[key])) {
                // 已有值：仅当"可信"时才不动。模型编造的占位符路径、或指向伪造用户目录的路径
                // 必须当作"空"重解析——否则会带着假路径直接落盘（该值非空会跳过所有默认推断）。
                let bogus = _isPlaceholderPathJS(args[key]);
                if (!bogus) {
                    const seg = _pathUserSegment(args[key]);
                    if (seg) {
                        const real = await _realUserSeg();
                        if (real && seg.toLowerCase() !== real.toLowerCase()) bogus = true;
                    }
                }
                if (!bogus) continue;
                ChatLog.warn?.('MCP', `[通用补全] ${toolName}.${key} 为占位符/伪造用户目录 → 丢弃重解析: ${args[key]}`);
                args[key] = '';
            }
            // 解析目录 + 文件名
            const { dir, src } = await _resolveDirForField(toolName, args, ctx);
            const sheet = args.sheet || args.sheetName || ctx?.turn?.context?.activeSheet?.sheet || 'output';
            const safeSheet = _safeFilename(sheet);
            const ext = _inferExtFromTool(toolName, prop);
            const sep = (dir && dir.includes('\\')) ? '\\' : '/';
            const defaultPath = `${dir}${dir && dir !== '.' ? sep : ''}${safeSheet}_${_timestamp()}${ext}`;
            args[key] = defaultPath;
            ChatLog.info('MCP', `[通用补全] ${toolName}.${key}=空 → ${defaultPath} (src=${src})`);
        }
    }

    // 【可选导出路径前端补全】ppt/word 导出工具（ppt_export_pdf/word_export_pdf/ppt_export_image）
    // 的 filePath 是可选参数，留空时后端会自动落盘桌面，但工具卡片"参数"区只显示
    // `filePath:`（空），用户看不出文件会存到哪（实测 2026-10-06 00:47）。
    // 这里前端跑与 tool-registry 同一个 inferHostExportFilePath，把真实路径填进 args
    // → 卡片"参数"区直接显示完整落盘路径；失败（目录拿不到）留空交后端兜底，不反问不拦截。
    const _OPTIONAL_EXPORT_PATH_TOOLS = {
        ppt_export_pdf: { ext: 'pdf', base: '演示文稿' },
        word_export_pdf: { ext: 'pdf', base: '文档' },
        ppt_export_image: { ext: 'png', base: '幻灯片' }
    };
    async function _fillOptionalExportPathIfNeeded(args, toolName, ctx) {
        const conf = _OPTIONAL_EXPORT_PATH_TOOLS[String(toolName || '').toLowerCase()];
        if (!conf) return;
        const cur = args.filePath !== undefined ? args.filePath : args.file_path;
        if (!_isEmpty(cur) && !_isPlaceholderPathJS(cur)) return;
        try {
            const infer = App?.ChatV2?.toolRegistryUtil?.inferHostExportFilePath;
            if (typeof infer !== 'function') return;
            const p = await infer(ctx?.turn, conf);
            if (p) {
                args.filePath = p;
                ChatLog.info('MCP', `[可选路径补全] ${toolName}.filePath=空 → ${p}`);
            }
        } catch (_e) { /* 留空交后端兜底 */ }
    }

    // --- 来自 app.js:L24473-L24532 ---
    async function executeMcpToolCall(toolCall, ctx) {
        try {
            if (!App.excelBridge) {
                const toolName = toolCall.function.name || '';
                // 直接返回对象（单一事实源即对象，移除跨 JS 层的不必要 JSON 字符串包装）。
                // 调用方 tool-registry.execute / skill-spec-execute 均已对 string|object 做容错，
                // 且 normalizeMcpResult 会对 .result 字段自动解包一层，最终归一结构与旧字符串形态完全一致。
                return {
                    mock: true,
                    message: `【模拟执行】${toolName} 已模拟完成。\n\n⚠️ 以上为模拟结果，Web版无法真正操作Excel。\n📦 请下载【桌面版】或【VSTO插件版】获得真实Excel功能。`
                };
            }

            const fullName = toolCall.function.name;
            // 兼容 Ollama（arguments 可能是对象而非 JSON 字符串）
            let rawArgs = toolCall.function.arguments;
            let args;
            if (typeof rawArgs === 'string') {
                args = JSON.parse(rawArgs || '{}');
            } else if (typeof rawArgs === 'object' && rawArgs !== null) {
                args = rawArgs;
            } else {
                args = {};
            }

            // 解析工具名称：source__toolName
            const sepIdx = fullName.indexOf('__');
            let source, toolName;
            if (sepIdx < 0) {
                source = 'excel-local';
                toolName = fullName;
            } else {
                source = fullName.substring(0, sepIdx);
                toolName = fullName.substring(sepIdx + 2);
            }

            // 诊断日志：进入 executeMcpToolCall 时打印关键信息
            // 当后端报"源工作表名称不能为空"等"字段缺失"类错误时，可通过日志确认 JS 端 args 的真实内容
            try {
                const _ctx = ctx || {};
                const _turnCtx = _ctx.turn?.context || {};
                ChatLog.info('MCP-DEBUG', `[${toolName}] args=${JSON.stringify(args)} | activeSheet=${JSON.stringify(_turnCtx.activeSheet || null)} | hasSelection=${!!_turnCtx.selection} | selSheet=${_turnCtx.selection?.sheetName || ''}`);
            } catch (_e) { /* 日志失败不影响主流程 */ }

            // 从 ctx.turn.context 提取活动表名/工作簿名，备用
            // 用于"工作表名参数缺失或空"时主动兜底（不止占位符场景）
            const _ctx = ctx || {};
            const _turnCtx = _ctx.turn?.context || {};
            const _ctxActiveSheet = _turnCtx.activeSheet?.sheet || _turnCtx.selection?.sheetName || '';
            const _ctxActiveWorkbook = _turnCtx.activeSheet?.workbook || _turnCtx.selection?.workbookName || '';

            // workbook 参数格式归一化：后端期望数组 ["xxx.xlsx"]，AI 可能传对象 {"name":"xxx.xlsx"} 或字符串
            if (args.workbook) {
                if (typeof args.workbook === 'string') {
                    args.workbook = [args.workbook];
                } else if (typeof args.workbook === 'object' && !Array.isArray(args.workbook)) {
                    const wbName = args.workbook.name || args.workbook.path || args.workbook.value || '';
                    args.workbook = wbName ? [wbName] : [];
                }
            }

            // 【图片文件自动转 OCR】AI 误用 excel_read_file 读图片时，自动转为 excel_ocr_recognize
            if (toolName === 'excel_read_file' && args.filePath) {
                const imgExts = ['.png', '.jpg', '.jpeg', '.webp', '.bmp', '.gif', '.tiff', '.svg'];
                const ext = args.filePath.substring(args.filePath.lastIndexOf('.')).toLowerCase();
                if (imgExts.includes(ext)) {
                    ChatLog.info('MCP', '[自动转OCR] excel_read_file → excel_ocr_recognize，文件:', args.filePath);
                    toolName = 'excel_ocr_recognize';
                    args = { imagePath: args.filePath };
                }
            }

            // 本地源判定：所有 *-local 后缀（excel-local/ppt-local/computer-local/web-local/time-local）
            // 统一走 CallLocalTool，仅 excel-local 需做工作表名占位符补全等专属兜底。
            // 外部 MCP Client（ms365/filesystem 等）才走 CallMcpToolAsync。
            const _isLocalSource = /-local$/i.test(source);
            if (source === 'excel-local') {
                // AI 可能传入中文占位工作表名（如“当前工作表”、“当前工作表名称”——
                // 后者正是 excel_rename_sheet.oldName 的参数描述，小模型会照搬为值），导致 DISP_E_BADINDEX。
                // 统一校验/替换为实际活动工作表名。覆盖所有“被操作工作表”字段：
                //   sheet/dataSheet（通用）+ source(copy)/name(delete)/oldName(rename)，按工具名限定避免误伤。
                //
                // args 里 source/name/oldName 等工作表名参数不仅在"是占位符"时需替换，
                // 在"完全缺失"或"为空字符串"时也应主动从 ctx.activeSheet 兜底——后端 GetArg 找不到 key 就直接报
                // "源工作表名称不能为空"，必须从根上避免这种情况。优先级：ctx.activeSheet > GetSelection().sheetName。
                const _sheetPlaceholders = ['当前活动工作表', '当前工作表', '活动工作表', '当前sheet', '当前工作簿', '当前工作表名称'];
                const _isSheetPlaceholder = (v) => {
                    if (!v) return false;
                    const s = String(v);
                    // 仅识别明确的语义占位符（"当前工作表"等）与 "active"。
                    // 注意：不能用"纯中文"判断——真实中文工作表名（如"汇总"）会被误判为占位符，
                    // 导致 excel_format_range 等工具的 sheet 被错误替换成活动表名。
                    return _sheetPlaceholders.some(p => s.includes(p)) ||
                        s.toLowerCase() === 'active';
                };
                const _isEmpty = (v) => v === undefined || v === null || (typeof v === 'string' && v.trim() === '');
                // 工具专属“被操作工作表”字段：占位 OR 缺失/空 → 替换为真实活动表名
                const _sourceFieldByTool = {
                    excel_copy_sheet: 'source',
                    excel_delete_sheet: 'name',
                    excel_rename_sheet: 'oldName'
                };
                // 工具专属“新名字段”：占位 → 清空交后端处理（绝不替换为活动表名，避免 source==target）
                const _targetFieldByTool = {
                    excel_copy_sheet: 'target',
                    excel_rename_sheet: 'newName'
                };
                const _srcField = _sourceFieldByTool[toolName];
                const _tgtField = _targetFieldByTool[toolName];

                // 决定真实活动表名：ctx.activeSheet 优先（chat-entry 已捕获），否则 GetSelection 兜底
                let _realSheet = String(_ctxActiveSheet || '').trim();
                let _realSheetSource = 'ctx';
                if (!_realSheet) {
                    try {
                        const selection = await App.excelBridge.GetSelection();
                        const selData = JSON.parse(selection);
                        if (selData && selData.sheetName) {
                            _realSheet = String(selData.sheetName).trim();
                            _realSheetSource = 'GetSelection';
                        }
                    } catch (e) {
                        ChatLog.warn('MCP', '[MCP] GetSelection 兜底失败:', e);
                    }
                }

                // 1) sheet / dataSheet 占位替换（保持原行为）
                if (_realSheet && (_isSheetPlaceholder(args.sheet) || _isSheetPlaceholder(args.dataSheet))) {
                    if (_isSheetPlaceholder(args.sheet)) {
                        ChatLog.info('MCP', `[MCP] 替换占位工作表名 sheet=${args.sheet} → ${_realSheet}`);
                        args.sheet = _realSheet;
                    }
                    if (_isSheetPlaceholder(args.dataSheet)) {
                        ChatLog.info('MCP', `[MCP] 替换占位工作表名 dataSheet=${args.dataSheet} → ${_realSheet}`);
                        args.dataSheet = _realSheet;
                    }
                }

                // 2) 工具专属"被操作工作表"字段：占位 OR 缺失/空 → 用真实表名
                if (_srcField) {
                    if (_isSheetPlaceholder(args[_srcField])) {
                        if (_realSheet) {
                            ChatLog.info('MCP', `[MCP] 替换占位工作表名 ${_srcField}=${args[_srcField]} → ${_realSheet} (源=${_realSheetSource})`);
                            args[_srcField] = _realSheet;
                        }
                    } else if (_isEmpty(args[_srcField]) && _realSheet) {
                        // 字段缺失或空 → 用真实表名
                        ChatLog.info('MCP', `[MCP] 兜底补齐空工作表名 ${_srcField}=空 → ${_realSheet} (源=${_realSheetSource})`);
                        args[_srcField] = _realSheet;
                    }
                }

                // 3) "新名"字段：占位 → 清空（绝不能改成活动表名，避免 source==target）
                if (_tgtField && _isSheetPlaceholder(args[_tgtField])) {
                    ChatLog.info('MCP', `[MCP] 清空占位目标工作表名 ${_tgtField}=${args[_tgtField]}（交后端自动生成）`);
                    args[_tgtField] = '';
                }

                // 【关键修复】自动注入当前选中区域（多个工具共用）
                const _needsRange = ['excel_read_range', 'excel_create_pivot_table', 'excel_create_chart', 'excel_write_range', 'excel_aggregate'];
                const _rangeFields = { excel_read_range: 'range', excel_create_pivot_table: 'range', excel_create_chart: 'dataRange', excel_write_range: 'range', excel_aggregate: 'range' };
                // excel_copy_range 特殊处理：多步骤管道中步骤间活动工作表会切换（如先建新表再复制），
                // 实时 GetSelection() 会拿到新表的 A1（活动单元格）而非用户原始选区，注入反而有害。
                // 因此优先用请求时捕获的 turn.context.selection；为空则不注入，交 VB 端 UsedRange 兜底。
                if (toolName === 'excel_copy_range' && !args.sourceRange) {
                    const _ctxSel = _turnCtx.selection;
                    const _ctxAddr = (_ctxSel && (_ctxSel.address || _ctxSel.range)) ? String(_ctxSel.address || _ctxSel.range).trim() : '';
                    if (_ctxAddr) {
                        let _addr = _ctxAddr;
                        if (_addr.includes('!')) _addr = _addr.split('!').pop().trim();
                        _addr = _addr.replace(/\$/g, '');
                        if (/^[A-Za-z]+\d+/.test(_addr)) {
                            args.sourceRange = _addr;
                            if (_ctxSel.sheetName && !args.sourceSheet) {
                                args.sourceSheet = String(_ctxSel.sheetName).trim();
                            }
                            ChatLog.info('MCP', `[MCP] excel_copy_range 用请求时选区注入 sourceRange: ${_addr}, sheet: ${_ctxSel.sheetName || ''}`);
                        }
                    }
                }
                if (_needsRange.includes(toolName) && !args[_rangeFields[toolName]]) {
                    // 图表已带显式 dataValues（内嵌聚合数据源）时不再注入选区 dataRange——
                    // 否则 dataRange 会被覆盖成当前选中格（如看板标题 A1），图表数据源失效。
                    if (toolName === 'excel_create_chart' && args.dataValues) {
                        ChatLog.info?.('MCP', '[MCP excel_create_chart] 已带 dataValues，跳过选区 dataRange 注入');
                    } else {
                    try {
                        const selection = await App.excelBridge.GetSelection();
                        const selData = JSON.parse(selection);
                        ChatLog.info('MCP', `[MCP ${toolName}] GetSelection 返回:`, JSON.stringify(selData).substring(0, 200));
                        if (selData && selData.address && selData.address !== '') {
                            let addr = String(selData.address).trim();
                            // 去除可能的工作表限定符（如 "Sheet1!A1:B2" → "A1:B2"）
                            if (addr.includes('!')) {
                                addr = addr.split('!').pop().trim();
                            }
                            // 去除绝对引用符号 $（如 "$A$1:$H$6" → "A1:H6"）
                            addr = addr.replace(/\$/g, '');
                            // 仅当地址看起来像有效的单元格/范围地址时才注入
                            if (/^[A-Za-z]+\d+/.test(addr)) {
                                args[_rangeFields[toolName]] = addr;
                                // 同时注入选区所在工作表名，避免 Excel 活动工作表切换后导致 DISP_E_BADINDEX
                                if (selData.sheetName && !args.sheet && !args.dataSheet) {
                                    args.sheet = String(selData.sheetName).trim();
                                }
                                // excel_copy_range 额外注入 sourceSheet（选区所在工作表），
                                // 否则 sourceRange 有值但 sourceSheet 为空时 VB 端仍会报"源工作表不能为空"
                                if (toolName === 'excel_copy_range' && selData.sheetName && !args.sourceSheet) {
                                    args.sourceSheet = String(selData.sheetName).trim();
                                }
                                ChatLog.info('MCP', `[MCP] 注入选中区域到 ${toolName}:`, addr, 'sheet:', selData.sheetName);
                            } else {
                                ChatLog.warn('MCP', `[MCP] 地址格式异常，跳过注入:`, addr);
                            }
                        }
                    } catch (e) {
                        ChatLog.warn('MCP', '[MCP] 获取选中区域失败:', e);
                    }
                    }
                }

                // 在 sheet/range 占位符替换之后跑，确保拿到真实 sheet 拼文件名
                try {
                    await _completeRequiredParamsIfNeeded(args, toolName, ctx);
                } catch (_e) {
                    ChatLog.warn('MCP', '[MCP] 通用参数补全层失败（不影响主流程）:', _e);
                }

                return await executeLocalExcelTool(toolName, args);
            } else if (_isLocalSource) {
                // 本地非 Excel 工具（ppt-local/computer-local/web-local/time-local）→ 直接走 CallLocalTool
                // 后端 McpManager.CallLocalTool 已按工具名前缀（ppt_/computer_/web_/time_）正确分发
                ChatLog.info('MCP', `[executeMcpToolCall] 本地非Excel源: ${source} → CallLocalTool(${toolName})`);
                // 【参数补全】此前只有 excel-local 分支跑了补全 → ppt/word 导出类 filePath 留空时
                // 卡片"参数"区显示空值（实测 2026-10-06 00:47）。这里补齐：必填路径参数兜底 +
                // 可选导出路径前端填充（失败留空交后端兜底，不影响主流程）。
                try {
                    await _completeRequiredParamsIfNeeded(args, toolName, ctx);
                    await _fillOptionalExportPathIfNeeded(args, toolName, ctx);
                } catch (_e) {
                    ChatLog.warn('MCP', '[MCP] 本地工具参数补全失败（不影响主流程）:', _e);
                }
                return await executeLocalExcelTool(toolName, args);
            } else {
                // 外部 MCP Client 工具 — 使用异步方法避免 Excel COM STA 线程阻塞导致崩溃
                try {
                    // 【修复】使用 CallMcpToolAsync 而非 CallMcpTool
                    // CallMcpTool 同步版本内部调用 .GetAwaiter().GetResult() 会阻塞 COM STA 线程，导致 Excel 崩溃
                    let result = await App.excelBridge.CallMcpToolAsync(source, toolName, JSON.stringify(args));
                    ChatLog.info('MCP', '[MCP] 外部工具实际返回:', result);
                    
                    if (!result) {
                        return { error: '工具返回为空' };
                    }
                    let data;
                    try {
                        data = JSON.parse(result.toString());
                    } catch (parseErr) {
                        ChatLog.warn('MCP', '[MCP] JSON 解析失败，直接返回原始结果:', result);
                        return result.toString();
                    }
                    if (data.success) {
                        const contents = data.result?.content || [];
                        // 【修复-趋势图表】保留结构化数据：content 存在非 text 类型项（image/resource/结构化 JSON）
                        // 或文本含 [OUTPUT_JSON] 标记时，返回原始结构化数据（含图表配置），
                        // 避免纯文本提取把图表配置丢弃导致"趋势分析无图"。
                        const hasStructuredItem = contents.some(c => c && typeof c === 'object' && c.type !== 'text');
                        const textContents = contents.filter(c => c && c.type === 'text').map(c => c.text).join('\n');
                        if (hasStructuredItem || (typeof textContents === 'string' && textContents.indexOf('[OUTPUT_JSON]') >= 0)) {
                            return { ok: true, result: data.result };
                        }
                        // 返回对象：纯文本用 textContents；无文本时直接放对象 data.result。
                        // 注意：旧代码此处是 JSON.stringify(data.result) 裸字符串，经 tool-registry 的
                        // normalizeMcpResult 解包一层后 result 即对象；这里直接放对象，最终归一结构等价。
                        // （skill-spec-execute 实际从不调用外部工具，故其在该分支的"字符串 vs 对象"差异在真实调用图中不存在）
                        const payload = textContents || data.result;
                        return { ok: true, result: payload };
                    }
                    return { error: data.error || '工具调用失败' };
                    } catch (callErr) {
                    ChatLog.warn('MCP', '️ [MCP] 调用失败:', callErr.message);
                    return { error: callErr.message || 'MCP 调用失败' };
                    }
            }
        } catch (e) {
            return { error: e.message || '工具执行异常' };
        }
    }

    // --- 来自 app.js:L24535-L24564 ---
    async function executeLocalExcelTool(toolName, args) {
        try {
            // 优先使用 COM 桥接（VSTO 环境）
            const _fbCall = (window.App && App.ChatV2 && App.ChatV2.callLocalToolWithFallback) || null;
            if (((App.excelBridge && App.excelBridge.CallLocalTool) || typeof _fbCall === 'function') && !window.chrome?.webview?._mockMode) {
                const _argsJson = JSON.stringify(args);
                // CallLocalTool 实际下发前打印 args JSON，便于诊断后端 args 序列化/反序列化是否丢字段
                ChatLog.info('MCP-DEBUG', `[CallLocalTool] toolName=${toolName} argsJson=${_argsJson}`);
                // ---- P1-1：ppt_generate 走异步任务队列（SubmitGeneratePPT + 轮询 GetGeneratePPTStatus）----
                // 同步 CallLocalTool 会在 Excel 主线程执行 PowerPoint COM 生成，大 PPT 期间 Excel UI 假死；
                // 异步路径立即返回 taskId、后台生成、前端 1s 轮询，生成期间界面可交互。
                // 说明：返回的 output 与同步路径 data.result 形状一致（Generate 原始返回 JSON 字符串），
                // 上层既有解析逻辑无需改动；Web/HTTP 环境无异步接口，走下方同步分支。
                if (toolName === 'ppt_generate' && App.excelBridge && typeof App.excelBridge.SubmitGeneratePPT === 'function') {
                    try {
                        const subRaw = await App.excelBridge.SubmitGeneratePPT(_argsJson);
                        const sub = parseMcpResult(subRaw);
                        if (sub && sub.success && sub.result && sub.result.taskId) {
                            const taskId = sub.result.taskId;
                            ChatLog.info('MCP', `[executeLocalExcelTool] ppt_generate 已提交异步任务 ${taskId}，开始轮询`);
                            const deadline = Date.now() + 5 * 60 * 1000; // 上限 5 分钟
                            while (Date.now() < deadline) {
                                await new Promise(r => setTimeout(r, 1000));
                                const stRaw = await App.excelBridge.GetGeneratePPTStatus(taskId);
                                const st = parseMcpResult(stRaw);
                                const stData = st && st.success ? st.result : null;
                                if (!stData) continue;
                                if (stData.status === 'done') {
                                    ChatLog.info('MCP', `[executeLocalExcelTool] ppt_generate 任务 ${taskId} 完成`);
                                    return { result: stData.output, taskCompleted: true };
                                }
                                if (stData.status === 'error' || stData.status === 'cancelled') {
                                    return { error: stData.error || 'PPT 生成失败', success: false, method: toolName };
                                }
                                if (stData.status === 'running' || stData.status === 'queued') {
                                    ChatLog.info('MCP', `[executeLocalExcelTool] ppt_generate ${taskId} ${stData.status} ${stData.progress}% ${stData.stage || ''}`);
                                }
                            }
                            return { error: 'PPT 生成超时（5 分钟未完成）', success: false, method: toolName };
                        }
                        ChatLog.warn('MCP', '[executeLocalExcelTool] ppt_generate 异步提交失败，回退同步路径: ' + JSON.stringify(sub));
                    } catch (asyncErr) {
                        ChatLog.warn('MCP', '[executeLocalExcelTool] ppt_generate 异步路径异常，回退同步路径: ' + asyncErr.message);
                    }
                }
                let result;
                if (typeof _fbCall === 'function') {
                    // 多来源回退：规避包装桥对个别方法不可调用（e is not a function）
                    const _r = await _fbCall(toolName, _argsJson);
                    result = _r && _r.raw;
                } else {
                    result = await App.excelBridge.CallLocalTool(toolName, _argsJson);
                }
                ChatLog.info('MCP', '[executeLocalExcelTool] CallLocalTool 原始返回:', typeof result, result?.substring?.(0, 500));
                const data = parseMcpResult(result);
                if (data && data.success) {
                    const response = {
                        result: data.result,
                        taskCompleted: data.taskCompleted !== false
                    };
                    return response;
                }
                // 失败时统一返回 {error: ...} 格式，避免上层误判为成功
                const errMsg = (data && (data.message || data.error || data.Message || data.Error))
                    || result || '工具执行失败';
                return { error: String(errMsg), success: false, method: toolName };
            }
            
            // Web 浏览器环境：通过 HTTP API 调用
            const bridgeUrl = App.Endpoints ? App.Endpoints.excelBridge.baseUrl : (window.ExcelBridgeConfig ? window.ExcelBridgeConfig.baseUrl : 'http://localhost:9527/api/excel');
            const httpResponse = await fetch(bridgeUrl + '/invoke', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ method: toolName, arguments: [args || {}] })
            });
            const data = await httpResponse.json();
            ChatLog.info('MCP', '[executeLocalExcelTool] HTTP API 返回:', data);
            
            if (data.success) {
                const result = {
                    result: data.data !== undefined ? data.data : data,
                    taskCompleted: data.taskCompleted !== false
                };
                return result;
            }
            return { error: data.error || '工具执行失败', mock: true, method: toolName };
            
        } catch (e) {
            ChatLog.error('MCP', '[executeLocalExcelTool] 执行失败:', e);
            return { error: e.message || '工具执行失败', mock: true, method: toolName };
        }
    }

    // 快速调用 MCP 工具（从面板按钮触发）
    async function quickCallMcpTool(toolFullName) {
        try {
            // 查找工具信息
            const tool = (window._mcpTools || []).find(t => t.name === toolFullName);
            if (!tool) {
                mcpToast('未找到工具: ' + toolFullName);
                return;
            }

            const toolName = toolFullName.includes('__') ? toolFullName.substring(toolFullName.indexOf('__') + 2) : toolFullName;
            const description = tool.description || '无描述';

            // 解析 OpenAI schema 格式的参数
            const paramSchema = tool.parameters || {};
            const properties = paramSchema.properties || {};
            const requiredList = paramSchema.required || [];
            const paramNames = Object.keys(properties);

            let argsHtml = '';
            if (paramNames.length > 0) {
                argsHtml = paramNames.map(pName => {
                    const p = properties[pName];
                    const isRequired = requiredList.includes(pName);
                    const required = isRequired ? ' <span style="color:red">*</span>' : '';
                    return `<div style="margin-bottom: 8px;">
                        <label style="display: block; font-size: 12px; margin-bottom: 3px;">${escapeHtml(pName)}${required} <span style="color:#999">(${escapeHtml(p.type || 'string')})</span></label>
                        <input type="text" id="mcp-param-${escapeHtml(pName)}" placeholder="${escapeHtml((p.description || '').substring(0, 50))}" 
                               style="width: 100%; padding: 5px; font-size: 12px; border: 1px solid #ddd; border-radius: 4px; box-sizing: border-box;">
                    </div>`;
                }).join('');
            } else {
                argsHtml = '<div style="color: #999; font-size: 12px;">此工具无需参数</div>';
            }

            // 本地桥接工具（excel-local/ppt-local 等）在 Web 版无桥时均为模拟数据
            const isLocalTool = /(?:excel|ppt|word|pdf)-local__/.test(toolFullName) || toolFullName.startsWith('excel-local');
            // 【关键】仅 Web 版（非 VSTO）才显示模拟数据警告
            const mockWarning = (isLocalTool && !App.excelBridge) ? `
                <div style="background:#fff3e0; border:1px solid #ff9800; border-radius:4px; padding:10px; margin-bottom:12px; font-size:12px;">
                    <div style="font-weight:600; color:#e65100; margin-bottom:6px;">️ Web版为模拟数据</div>
                    <div style="color:#666; margin-bottom:8px;">此工具在Web版中返回模拟结果，完整本地功能请下载：</div>
                    <div style="display:flex; gap:8px;">
                        <a href="#" style="flex:1; text-align:center; padding:6px; background:#1976D2; color:white; border-radius:4px; text-decoration:none; font-size:12px;">📦 桌面版（推荐）</a>
                        <a href="#" style="flex:1; text-align:center; padding:6px; background:#388E3C; color:white; border-radius:4px; text-decoration:none; font-size:12px;">🔌 VSTO插件版</a>
                    </div>
                </div>` : '';

            const dialogHtml = `
                <div style="position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.5); z-index: 10000; display: flex; align-items: center; justify-content: center;">
                    <div style="background: white; border-radius: 8px; padding: 20px; max-width: 400px; width: 90%; max-height: 80vh; overflow-y: auto;">
                        <div style="font-weight: 600; font-size: 14px; margin-bottom: 5px;">调用工具: ${escapeHtml(toolName)}</div>
                        <div style="font-size: 11px; color: #666; margin-bottom: 15px;">${escapeHtml(description)}</div>
                        ${mockWarning}
                        ${argsHtml}
                        <div id="mcp-call-result" style="display: none; margin-top: 10px; padding: 10px; background: #f5f5f5; border-radius: 4px; font-size: 11px; max-height: 200px; overflow-y: auto; white-space: pre-wrap; word-break: break-all;"></div>
                        <div style="display: flex; gap: 8px; margin-top: 15px;">
                            <button id="mcp-call-execute" style="flex: 1; padding: 8px; background: #2196F3; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 13px;">执行</button>
                            <button id="mcp-call-cancel" style="flex: 1; padding: 8px; background: #999; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 13px;">取消</button>
                        </div>
                    </div>
                </div>`;

            // 添加到页面
            const container = document.createElement('div');
            container.innerHTML = dialogHtml;
            document.body.appendChild(container);

            // 绑定事件
            container.querySelector('#mcp-call-cancel').onclick = () => container.remove();

            container.querySelector('#mcp-call-execute').onclick = async () => {
                const args = {};
                paramNames.forEach(pName => {
                    const input = container.querySelector(`#mcp-param-${pName}`);
                    if (input && input.value.trim()) {
                        args[pName] = input.value.trim();
                    }
                });

                const execBtn = container.querySelector('#mcp-call-execute');
                execBtn.textContent = '执行中';
                execBtn.disabled = true;

                try {
                    const result = await App.executeMcpToolCall({
                        function: { name: toolFullName, arguments: JSON.stringify(args) }
                    });

                    const resultDiv = container.querySelector('#mcp-call-result');
                    resultDiv.style.display = 'block';
                    // result 现在可能是对象（executeMcpToolCall 已改为对象返回），需序列化展示，避免显示 [object Object]
                    resultDiv.textContent = (typeof result === 'string') ? result : JSON.stringify(result, null, 2);
                    execBtn.textContent = '执行完成';
                    execBtn.disabled = false;
                } catch (e) {
                    const resultDiv = container.querySelector('#mcp-call-result');
                    resultDiv.style.display = 'block';
                    resultDiv.textContent = '执行失败: ' + e.message;
                    execBtn.textContent = '重试';
                    execBtn.disabled = false;
                }
            };
        } catch (e) {
            mcpToast('调用失败: ' + e.message);
        }
    }

    // 等待 Excel Bridge 就绪后再加载 MCP 工具列表
    // 自动连接由 js-app-init.js 统一根据 autoStartClients 配置控制，避免重复连接
    async function loadMcpToolsAfterBridgeReady() {
        const maxRetries = 30;
        let retries = 0;
        while (!App.excelBridge && retries < maxRetries) {
            await new Promise(r => setTimeout(r, 200));
            retries++;
        }
        if (!App.excelBridge) {
            ChatLog.warn('MCP', '[MCP] Excel Bridge 未就绪，跳过工具加载');
            return;
        }

        // 等 initBridge 设好 defaultSyncMethod + BridgeAccessor 再包一层，避免首轮 e is not a function
        try {
            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects
                && window.chrome.webview.hostObjects.options) {
                window.chrome.webview.hostObjects.options.defaultSyncMethod = 'none';
            }
            if (window.BridgeAccessor && typeof window.BridgeAccessor.adoptHostBridge === 'function'
                && window.chrome && window.chrome.webview && window.chrome.webview.hostObjects
                && window.chrome.webview.hostObjects.excelBridge) {
                window.BridgeAccessor.adoptHostBridge(window.chrome.webview.hostObjects.excelBridge);
            }
        } catch (_e) { /* ignore */ }

        try {
            // 【修复】先刷新状态再加载工具，确保 UI 能同步 server/clients/tools 三态
            await refreshMcpStatus();
            await loadMcpToolsForAI();
            // 首轮若因桥未就绪失败，延迟再试一次（Word 上实测：500ms 太早会空工具，稍后能装上 30 个）
            if (!App._mcpToolsLoaded) {
                await new Promise(r => setTimeout(r, 800));
                await loadMcpToolsForAI();
            }
        } catch (e) {
            ChatLog.warn('MCP', '[MCP] 工具列表加载失败:', e);
        }
    }
    
    async function loadMcpToolsForBrowser() {
        if (isVstoWebView2()) {
            ChatLog.info('MCP', '[MCP] 跳过浏览器工具加载，VSTO环境');
            return;
        }
        if (App._mcpToolsLoaded) {
            ChatLog.info('MCP', '[MCP] 工具已加载，跳过浏览器工具加载');
            return;
        }
        const bridgeBase = App.Endpoints ? App.Endpoints.excelBridge.baseUrl : (window.ExcelBridgeConfig ? window.ExcelBridgeConfig.baseUrl : 'http://localhost:9527');
        try {
            const response = await fetch(`${bridgeBase}/api/mcp/tools`);
            const data = await response.json();
            if (data && data.success && data.tools && data.tools.length > 0) {
                window._mcpTools = data.tools.map(buildMcpToolMeta);
                ChatLog.info('MCP', '✅ [Browser] 已加载 ' + window._mcpTools.length + ' 个 MCP 工具');
                App._mcpToolsLoaded = true;
                renderMcpUI();
            }
        } catch (e) {
            ChatLog.warn('MCP', '[MCP] 浏览器工具加载失败:', e.message);
        }
    }

    // ==================== 模块导出 ====================
    
    // Web 浏览器环境：立即加载工具（不依赖 excelBridge）
    if (!isVstoWebView2()) {
        setTimeout(loadMcpToolsForBrowser, 100);
    }
    
    // VSTO 环境：等待 Bridge 就绪后加载工具列表，不主动发起 MCP 连接
    setTimeout(loadMcpToolsAfterBridgeReady, 500);

    // 将函数挂载到 App 对象
    App.McpModule = {
        refreshMcpStatus,
        renderMcpUI,
        parseMcpResult,
        startMcpServer,
        stopMcpServer,
        showAddMcpClientDialog,
        editMcpClient,
        closeMcpClientDialog,
        onMcpTransportChange,
        saveMcpClient,
        testMcpClient,
        testMcpClientFromDialog,
        removeMcpClient,
        toggleMcpClient,
        autoReconnectMcpClients,
        loadMcpToolsForAI,
        buildToolParameters,
        executeMcpToolCall,
        executeLocalExcelTool
    };

    // 同时暴露到 window 以便 HTML onclick 调用
    window.refreshMcpStatus = refreshMcpStatus;
    window.renderMcpUI = renderMcpUI;
    window.startMcpServer = startMcpServer;
    window.stopMcpServer = stopMcpServer;
    window.showAddMcpClientDialog = showAddMcpClientDialog;
    window.editMcpClient = editMcpClient;
    window.closeMcpClientDialog = closeMcpClientDialog;
    window.onMcpTransportChange = onMcpTransportChange;
    window.saveMcpClient = saveMcpClient;
    window.testMcpClient = testMcpClient;
    window.testMcpClientFromDialog = testMcpClientFromDialog;
    window.removeMcpClient = removeMcpClient;
    window.toggleMcpClient = toggleMcpClient;
    window.autoReconnectMcpClients = autoReconnectMcpClients;
    window.quickCallMcpTool = quickCallMcpTool;
    window.loadMcpToolsForAI = loadMcpToolsForAI;
    window.executeMcpToolCall = executeMcpToolCall;
    App.loadMcpToolsForAI = loadMcpToolsForAI;
    App.executeMcpToolCall = executeMcpToolCall;

})();