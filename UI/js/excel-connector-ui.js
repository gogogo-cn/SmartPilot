/**
 * ============================================================
 * excel-connector-ui.js — Excel Connection UI
 * ============================================================
 * 功能：Excel 连接界面模块，提供连接模式选择、高级设置与连接诊断界面。
 *
 * 核心类/函数：
 *   - showConnectionDialog() — 显示连接对话框
 *   - renderConnectionOptions() — 渲染连接选项
 *   - testConnection() — 测试连接
 *
 * 依赖：
 *   - app.js
 *   - excel-connector.js
 * ============================================================
 */

(function() {
    'use strict';

    const ExcelConnectorUI = {
        connector: null,
        container: null,
        
        /**
         * 本地 Toast 兜底（避免 App.showToast 未定义时报错）
         */
        _showToast(message, type = 'info') {
            if (typeof window.App !== 'undefined' && typeof App.showToast === 'function') {
                if (App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function') {
                    App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">' + message + '</span>', true);
                } else {
                    App.showToast(message, type);
                }
            } else if (typeof window.showToast === 'function') {
                window.showToast(message, type);
            } else {
                ChatLog.info('CONNECTOR', `[ConnectorUI][${type}] ${message}`);
            }
        },
        
        /**
         * 初始化连接管理UI
         */
        init() {
            this.container = document.getElementById('excel-connector-ui');
            if (!this.container) {
                ChatLog.warn('CONNECTOR', '[ConnectorUI] 未找到 #excel-connector-ui 容器');
                return;
            }
            
            // 初始化连接器
            this.connector = new ExcelConnector();
            
            // 渲染UI
            this.render();
            
            // 设置默认服务地址
            const bridgeUrlInput = document.getElementById('bridge-url');
            if (bridgeUrlInput && !bridgeUrlInput.value) {
                bridgeUrlInput.value = (window.ExcelBridgeConfig && window.ExcelBridgeConfig.defaultBridgeUrl)
                    || 'http://localhost:9527/api/excel';
                bridgeUrlInput.placeholder = (window.ExcelBridgeConfig && window.ExcelBridgeConfig.defaultBridgeUrl)
                    || 'http://localhost:9527/api/excel';
            }
            
            // 尝试自动连接
            this.autoConnect();
            
            ChatLog.info('CONNECTOR', '[ConnectorUI] ✅ 连接管理UI已初始化');
        },
        
        /**
         * 渲染UI
         */
        render() {
            this.container.innerHTML = `
                <div class="excel-connector-panel">
                    <div class="connector-header">
                        <h3>🔗 Excel 连接</h3>
                        <div class="connection-status" id="connection-status">
                            <span class="status-dot" id="status-dot"></span>
                            <span class="status-text" id="status-text">检测中</span>
                        </div>
                    </div>
                    
                    <div class="connector-body">
                        <!-- 手动连接选项 -->
                        <div class="connect-options" id="connect-options">
                            <h4>选择连接方式：</h4>
                            
                            <button class="btn-connect-mode" onclick="ExcelConnectorUI.connectMode('vsto')" 
                                    title="VSTO原生插件（需要Excel中运行插件）">
                                🟢 VSTO 插件
                                <small>原生集成</small>
                            </button>
                            
                            <button class="btn-connect-mode" onclick="ExcelConnectorUI.connectMode('local-bridge')"
                                    title="本地代理服务（需启动node服务）">
                                🟡 本地代理
                                <small>推荐Web版</small>
                            </button>
                            
                            <button class="btn-connect-mode" onclick="ExcelConnectorUI.connectMode('office-js')"
                                    title="Office.js浏览器插件">
                                🔵 Office.js
                                <small>Excel Online</small>
                            </button>
                        </div>
                        
                        <!-- 本地代理配置 -->
                        <div class="bridge-config" id="bridge-config" style="display: none;">
                            <label for="bridge-url">服务地址：</label>
                            <input type="text" id="bridge-url" 
                                   value=""
                                   placeholder="自动读取 bridge-config.js">
                            <button onclick="ExcelConnectorUI.testBridgeConnection()" 
                                    class="btn-test">测试连接</button>
                            <div class="config-help">
                                <a href="#" onclick="ExcelConnectorUI.showBridgeSetupGuide()">
                                    如何启动本地代理？
                                </a>
                            </div>
                        </div>
                        
                        <!-- 连接操作按钮 -->
                        <div class="connector-actions">
                            <button id="btn-connect" onclick="ExcelConnectorUI.toggleConnection()" 
                                    class="btn-primary">
                                🔌 连接 Excel
                            </button>
                            <button id="btn-disconnect" onclick="ExcelConnectorUI.disconnect()" 
                                    class="btn-secondary" style="display: none;">
                                ❌ 断开连接
                            </button>
                            <button onclick="ExcelConnectorUI.refreshStatus()" 
                                    class="btn-icon" title="刷新状态">
                                🔄
                            </button>
                        </div>
                    </div>
                    
                    <!-- 连接信息详情 -->
                    <div class="connection-details" id="connection-details" style="display: none;">
                        <h4>连接信息</h4>
                        <pre id="connection-info-json"></pre>
                    </div>
                </div>
                
                <style>
                    .excel-connector-panel {
                        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
                        border-radius: 12px;
                        padding: 20px;
                        margin-bottom: 20px;
                        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                    }
                    
                    .connector-header {
                        display: flex;
                        justify-content: space-between;
                        align-items: center;
                        margin-bottom: 15px;
                    }
                    
                    .connector-header h3 {
                        margin: 0;
                        font-size: 18px;
                        color: #333;
                    }
                    
                    .connection-status {
                        display: flex;
                        align-items: center;
                        gap: 8px;
                        padding: 6px 12px;
                        background: white;
                        border-radius: 20px;
                        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
                    }
                    
                    .status-dot {
                        width: 10px;
                        height: 10px;
                        border-radius: 50%;
                        background: #ffc107;
                        animation: pulse 2s infinite;
                    }
                    
                    .status-dot.connected {
                        background: #52c41a;
                        animation: none;
                    }
                    
                    .status-dot.disconnected {
                        background: #ff4d4f;
                        animation: none;
                    }
                    
                    @keyframes pulse {
                        0%, 100% { opacity: 1; }
                        50% { opacity: 0.5; }
                    }
                    
                    .status-text {
                        font-size: 13px;
                        color: #666;
                        font-weight: 500;
                    }
                    
                    .connect-options {
                        margin-bottom: 15px;
                    }
                    
                    .connect-options h4 {
                        font-size: 14px;
                        margin: 0 0 10px 0;
                        color: #555;
                    }
                    
                    .btn-connect-mode {
                        display: inline-flex;
                        flex-direction: column;
                        padding: 12px 16px;
                        margin: 5px;
                        border: 2px solid #e0e0e0;
                        border-radius: 8px;
                        background: white;
                        cursor: pointer;
                        transition: all 0.3s ease;
                        min-width: 120px;
                    }
                    
                    .btn-connect-mode:hover {
                        border-color: #1890ff;
                        transform: translateY(-2px);
                        box-shadow: 0 4px 8px rgba(24,144,255,0.2);
                    }
                    
                    .btn-connect-mode small {
                        font-size: 11px;
                        color: #999;
                        margin-top: 4px;
                    }
                    
                    .bridge-config {
                        background: white;
                        padding: 15px;
                        border-radius: 8px;
                        margin-bottom: 15px;
                        border-left: 4px solid #1890ff;
                    }
                    
                    .bridge-config label {
                        display: block;
                        font-weight: 600;
                        margin-bottom: 8px;
                        color: #333;
                    }
                    
                    .bridge-config input {
                        width: calc(100% - 100px);
                        padding: 8px 12px;
                        border: 1px solid #d9d9d9;
                        border-radius: 4px;
                        margin-right: 8px;
                        font-size: 13px;
                    }
                    
                    .btn-test {
                        padding: 8px 16px;
                        background: #52c41a;
                        color: white;
                        border: none;
                        border-radius: 4px;
                        cursor: pointer;
                        font-size: 13px;
                    }
                    
                    .btn-test:hover {
                        background: #389e0d;
                    }
                    
                    .config-help {
                        margin-top: 8px;
                        font-size: 12px;
                    }
                    
                    .config-help a {
                        color: #1890ff;
                        text-decoration: none;
                    }
                    
                    .config-help a:hover {
                        text-decoration: underline;
                    }
                    
                    .connector-actions {
                        display: flex;
                        gap: 10px;
                        align-items: center;
                    }
                    
                    .btn-primary {
                        flex: 1;
                        padding: 12px 24px;
                        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                        color: white;
                        border: none;
                        border-radius: 8px;
                        cursor: pointer;
                        font-size: 15px;
                        font-weight: 600;
                        transition: all 0.3s ease;
                    }
                    
                    .btn-primary:hover {
                        transform: translateY(-2px);
                        box-shadow: 0 6px 16px rgba(102,126,234,0.4);
                    }
                    
                    .btn-primary:disabled {
                        background: #ccc;
                        cursor: not-allowed;
                        transform: none;
                    }
                    
                    .btn-secondary {
                        padding: 12px 24px;
                        background: #ff4d4f;
                        color: white;
                        border: none;
                        border-radius: 8px;
                        cursor: pointer;
                        font-size: 15px;
                        font-weight: 600;
                    }
                    
                    .btn-secondary:hover {
                        background: #cf1322;
                    }
                    
                    .btn-icon {
                        padding: 12px;
                        background: white;
                        border: 1px solid #d9d9d9;
                        border-radius: 8px;
                        cursor: pointer;
                        font-size: 18px;
                    }
                    
                    .btn-icon:hover {
                        background: #f5f5f5;
                    }
                    
                    .connection-details {
                        margin-top: 15px;
                        background: white;
                        padding: 15px;
                        border-radius: 8px;
                        max-height: 200px;
                        overflow-y: auto;
                    }
                    
                    .connection-details h4 {
                        margin: 0 0 10px 0;
                        font-size: 14px;
                        color: #333;
                    }
                    
                    .connection-details pre {
                        margin: 0;
                        padding: 10px;
                        background: #f6f8fa;
                        border-radius: 4px;
                        font-size: 12px;
                        overflow-x: auto;
                    }
                </style>
            `;
        },
        
        /**
         * 自动连接
         */
        async autoConnect() {
            this.updateStatus('detecting', '正在检测...');
            
            try {
                const result = await this.connector.autoConnect();
                
                if (result.success) {
                    this.onConnected(result);
                } else {
                    this.onDisconnected(result.message);
                }
            } catch (error) {
                ChatLog.error('CONNECTOR', '[ConnectorUI] 自动连接失败:', error);
                this.onDisconnected(error.message);
            }
        },
        
        /**
         * 选择连接模式
         */
        async connectMode(mode) {
            this.updateStatus('connecting', '连接中');
            
            // 显示对应的配置项
            document.querySelectorAll('.btn-connect-mode').forEach(btn => {
                btn.style.borderColor = '#e0e0e0';
            });
            event.target.closest('.btn-connect-mode').style.borderColor = '#1890ff';
            
            if (mode === 'local-bridge') {
                document.getElementById('bridge-config').style.display = 'block';
            } else {
                document.getElementById('bridge-config').style.display = 'none';
            }
            
            try {
                let config = {};
                if (mode === 'local-bridge') {
                    config.url = document.getElementById('bridge-url').value.trim();
                }
                
                const result = await this.connector.connect(mode, config);
                
                if (result.success) {
                    this.onConnected(result);
                    this._showToast(`✅ ${result.message}`, 'success');
                } else {
                    this.onDisconnected(result.message);
                    this._showToast(`❌ ${result.message}`, 'error');
                }
            } catch (error) {
                this.onDisconnected(error.message);
                this._showToast(`❌ 连接失败: ${error.message}`, 'error');
            }
        },
        
        /**
         * 测试本地代理连接
         */
        async testBridgeConnection() {
            const url = document.getElementById('bridge-url').value.trim();
            const btn = event.target;
            
            btn.textContent = '测试中';
            btn.disabled = true;
            
            try {
                const response = await fetch(`${url}/status`, {
                    method: 'GET',
                    signal: AbortSignal.timeout(3000)
                });
                
                if (response.ok) {
                    const data = await response.json();
                    this._showToast(`✅ 代理服务可用 (版本: ${data.version})`, 'success');
                    btn.textContent = '✓ 可用';
                    setTimeout(() => { btn.textContent = '测试连接'; }, 2000);
                } else {
                    throw new Error(`HTTP ${response.status}`);
                }
            } catch (error) {
                this._showToast(`❌ 无法连接到代理服务: ${error.message}`, 'error');
                btn.textContent = '测试连接';
            }
            
            btn.disabled = false;
        },
        
        /**
         * 切换连接/断开
         */
        async toggleConnection() {
            if (this.connector.isConnected) {
                this.disconnect();
            } else {
                // 如果已经选择了模式，直接连接
                const activeMode = document.querySelector('.btn-connect-mode[style*="border-color: rgb(24, 144, 255)"]');
                if (activeMode) {
                    await this.connectMode(activeMode.onclick.toString().match(/'(\w+)'/)[1]);
                } else {
                    // 默认尝试自动连接
                    await this.autoConnect();
                }
            }
        },
        
        /**
         * 断开连接
         */
        disconnect() {
            this.connector.disconnect();
            this.onDisconnected('已断开连接');
            this._showToast('已断开Excel连接', 'info');
        },
        
        /**
         * 刷新状态
         */
        async refreshStatus() {
            await this.autoConnect();
        },
        
        /**
         * 连接成功回调
         */
        onConnected(result) {
            this.updateStatus('connected', result.message || '已连接');
            
            // 更新按钮状态
            document.getElementById('btn-connect').style.display = 'none';
            document.getElementById('btn-disconnect').style.display = 'inline-block';
            
            // 显示连接详情
            const details = document.getElementById('connection-details');
            const infoJson = document.getElementById('connection-info-json');
            details.style.display = 'block';
            infoJson.textContent = JSON.stringify(this.connector.connectionInfo, null, 2);
            
            // 触发全局事件
            if (typeof window.App !== 'undefined' && typeof App.onExcelConnected === 'function') {
                App.onExcelConnected(this.connector.mode);
            }
        },
        
        /**
         * 断开连接回调
         */
        onDisconnected(message = '未连接') {
            this.updateStatus('disconnected', message);
            
            // 更新按钮状态
            document.getElementById('btn-connect').style.display = 'inline-block';
            document.getElementById('btn-disconnect').style.display = 'none';
            
            // 隐藏连接详情
            document.getElementById('connection-details').style.display = 'none';
            
            // 触发全局事件
            if (typeof window.App !== 'undefined' && typeof App.onExcelDisconnected === 'function') {
                App.onExcelDisconnected();
            }
        },
        
        /**
         * 更新状态显示
         */
        updateStatus(status, text) {
            const dot = document.getElementById('status-dot');
            const textEl = document.getElementById('status-text');
            
            dot.className = 'status-dot ' + status;
            textEl.textContent = text;
        },
        
        /**
         * 显示本地代理设置指南
         */
        showBridgeSetupGuide() {
            const bridgeBaseUrl = (window.ExcelBridgeConfig && window.ExcelBridgeConfig.baseUrl) || 'http://localhost:9527';
            const guideHtml = `
                <div style="max-width: 600px; line-height: 1.8;">
                    <h3>🚀 本地代理服务启动指南</h3>
                    
                    <h4>步骤 1：启动本地桥</h4>
                    <pre style="background: #f6f8fa; padding: 10px; border-radius: 4px;">双击项目根目录的 start.bat
（或：cd Tools &amp;&amp; python py-excel-bridge.py）</pre>
                    <p>首次运行需 Python 3.8+ 与 pywin32（缺失时脚本会提示 <code>pip install pywin32</code>）。</p>
                    
                    <h4>步骤 2：确认服务已启动</h4>
                    <pre style="background: #f6f8fa; padding: 10px; border-radius: 4px;">📍 服务地址: ${bridgeBaseUrl}</pre>
                    
                    <h4>步骤 3：连接使用</h4>
                    <ol>
                        <li>确保 Excel 已打开</li>
                        <li>点击上方 "🟡 本地代理" 按钮</li>
                        <li>点击 "🔌 连接 Excel"</li>
                        <li>开始享受完整功能！</li>
                    </ol>
                    
                    <h4>💡 常见问题</h4>
                    <ul>
                        <li><strong>端口被占用？</strong> 先执行 <code>set PORT=9531</code> 再运行 start.bat</li>
                        <li><strong>CORS 错误？</strong> 确保通过 localhost 访问，不是 file:// 协议</li>
                        <li><strong>无法连接 Excel？</strong> 以管理员身份运行命令行</li>
                    </ul>
                    
                    <p><a href="../Tools/README-Bridge.md" target="_blank">📖 查看完整文档 →</a></p>
                </div>
            `;
            
            // 显示模态框或对话框
            if (typeof App !== 'undefined' && typeof App.showDialog === 'function') {
                App.showDialog('设置指南', guideHtml);
            } else {
                alert(guideHtml.replace(/<[^>]+>/g, '\n'));
            }
        }
    };

    // 导出到全局
    window.ExcelConnectorUI = ExcelConnectorUI;

})();