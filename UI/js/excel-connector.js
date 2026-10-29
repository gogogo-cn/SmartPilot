﻿﻿/**
 * ============================================================
 * excel-connector.js — Excel Connector
 * ============================================================
 * 功能：Excel 连接器，按优先级自动检测 VSTO / Local Bridge / Office.js / Mock 连接模式。
 *
 * 核心类/函数：
 *   - detectConnection() — 检测连接
 *   - getExcelBridge() — 获取 Excel 桥接
 *   - executeExcelMethod() — 执行 Excel 方法
 *   - shutdown() — 关闭连接
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 *   - bridge-config.js
 * ============================================================
 */

(function() {
    'use strict';

    class ExcelConnector {
        constructor() {
            this.mode = null; // 'vsto' | 'local-bridge' | 'office-js' | 'standalone'
            this.bridge = null;
            this.isConnected = false;
            this.connectionInfo = {};
            
            // 绑定到全局App对象
            window.App.ExcelConnector = this;
        }

        /**
         * 默认连接配置（优先读取 bridge-config.js 中的全局配置，避免硬编码）
         */
        static get DEFAULT_CONFIG() {
            const cfg = window.ExcelBridgeConfig;
            if (cfg) {
                return {
                    bridgeUrls: cfg.bridgeUrls,
                    defaultBridgeUrl: cfg.defaultBridgeUrl,
                    statusTimeout: cfg.statusTimeout,
                    connectTimeout: cfg.connectTimeout,
                    invokeTimeout: cfg.invokeTimeout
                };
            }
            return {
                bridgeUrls: [
                    App.Endpoints ? App.Endpoints.excelBridge.baseUrl : 'http://localhost:9527/api/excel',
                    'http://127.0.0.1:9527/api/excel'
                ],
                defaultBridgeUrl: App.Endpoints ? App.Endpoints.excelBridge.baseUrl : 'http://localhost:9527/api/excel',
                statusTimeout: 2000,
                connectTimeout: 5000,
                invokeTimeout: 30000
            };
        }

        /**
         * 自动检测并连接最佳可用模式
         * @returns {Promise<{success: boolean, mode: string, message: string}>}
         */
        async autoConnect() {
            ChatLog.info('CONNECTOR', '[ExcelConnector] 开始自动检测连接模式...');
            
            // 优先级1: VSTO WebView2 (原生插件)
            if (this.detectVSTOMode()) {
                const result = await this.connectVSTO();
                if (result.success) return result;
            }
            
            // 优先级2: Office.js (浏览器插件)
            if (this.detectOfficeJSMode()) {
                const result = await this.connectOfficeJS();
                if (result.success) return result;
            }
            
            // 优先级3: Local Bridge (本地代理服务)
            const localResult = await this.tryLocalBridge();
            if (localResult.available) {
                const result = await this.connectLocalBridge();
                if (result.success) return result;
            }
            
            // 最终降级为独立模式
            this.mode = 'standalone';
            this.isConnected = false;
            ChatLog.info('CONNECTOR', '[ExcelConnector] 未检测到Excel连接，使用独立模式');
            
            return {
                success: false,
                mode: 'standalone',
                message: '未连接Excel（独立运行模式）'
            };
        }

        /**
         * 手动指定连接模式
         * @param {string} mode - 'vsto' | 'local-bridge' | 'office-js'
         * @param {Object} config - 连接配置
         */
        async connect(mode, config = {}) {
            switch(mode) {
                case 'vsto':
                    return await this.connectVSTO();
                case 'local-bridge':
                    return await this.connectLocalBridge(config);
                case 'office-js':
                    return await this.connectOfficeJS();
                default:
                    throw new Error(`未知的连接模式: ${mode}`);
            }
        }

        // ==================== 模式1: VSTO WebView2 ====================
        
        detectVSTOMode() {
            return !!(window.chrome && window.chrome.webview && window.chrome.webview.hostObjects);
        }

        async connectVSTO() {
            try {
                ChatLog.info('CONNECTOR', '[ExcelConnector] 尝试VSTO WebView2连接...');
                
                // 配置COM桥接选项
                // ⚠️ 切勿把 defaultSyncMethod 设为 'sync'：
                // WebView2 在 sync 模式下调用宿主对象方法会在代理内部抛 "e is not a function"
                // （sync 代理 apply 路径的已知缺陷），而本项目所有桥调用都使用 await（异步），
                // 因此必须保留默认的异步模式。强制 sync 会导致全盘桥调用失败（工作簿/MCP/OCR 全挂）。
                // 合法取值为 'none'（默认·始终异步）与 'sync'，没有 'async'。
                if (window.chrome.webview.hostObjects.options) {
                    window.chrome.webview.hostObjects.options.defaultSyncMethod = 'none';
                }
                
                // 获取bridge对象
                this.bridge = window.chrome.webview.hostObjects.excelBridge;
                App.excelBridge = this.bridge;
                App.state.excelBridge = this.bridge;
                
                // 测试连接
                const testResult = await this._callBridge('TestConnection');
                
                if (testResult && !testResult.error) {
                    this.mode = 'vsto';
                    this.isConnected = true;
                    this.connectionInfo = {
                        type: 'VSTO WebView2',
                        version: testResult.version || 'unknown'
                    };
                    
                    ChatLog.info('CONNECTOR', '[ExcelConnector] ✅ VSTO连接成功', this.connectionInfo);
                    this._updateUI(true);
                    
                    return {
                        success: true,
                        mode: 'vsto',
                        message: `已连接Excel (VSTO插件)`
                    };
                } else {
                    throw new Error(testResult?.error || '连接测试失败');
                }
            } catch (error) {
                ChatLog.warn('CONNECTOR', '[ExcelConnector] ❌ VSTO连接失败:', error.message);
                return {
                    success: false,
                    mode: null,
                    message: `VSTO连接失败: ${error.message}`
                };
            }
        }

        // ==================== 模式2: Local Bridge Service ====================
        
        /**
         * 检测本地代理服务是否可用
         * @returns {Promise<{available: boolean, url: string, version?: string}>}
         */
        async tryLocalBridge() {
            const config = ExcelConnector.DEFAULT_CONFIG;
            
            for (const url of config.bridgeUrls) {
                try {
                    const response = await fetch(`${url}/status`, {
                        method: 'GET',
                        mode: 'cors',
                        signal: AbortSignal.timeout(config.statusTimeout)
                    });
                    
                    if (response.ok) {
                        const data = await response.json();
                        ChatLog.info('CONNECTOR', '[ExcelConnector] 发现本地代理服务:', url);
                        return {
                            available: true,
                            url: url,
                            version: data.version,
                            winaxAvailable: data.winaxAvailable,
                            mockMode: data.mockMode
                        };
                    }
                } catch (e) {
                    // 继续尝试下一个URL
                }
            }
            
            return { available: false };
        }

        async connectLocalBridge(config = {}) {
            const baseUrl = config.url || ExcelConnector.DEFAULT_CONFIG.defaultBridgeUrl;
            
            try {
                ChatLog.info('CONNECTOR', '[ExcelConnector] 尝试本地代理服务连接...', baseUrl);
                
                // 测试连接
                const response = await fetch(`${baseUrl}/connect`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ action: 'test' }),
                    signal: AbortSignal.timeout(ExcelConnector.DEFAULT_CONFIG.connectTimeout)
                });
                
                let responseData;
                try {
                    responseData = await response.json();
                } catch (parseError) {
                    responseData = null;
                }
                
                if (!response.ok) {
                    const serverMessage = responseData?.message || responseData?.error || `HTTP ${response.status}`;
                    throw new Error(serverMessage);
                }
                
                const result = responseData;
                
                if (result && result.success) {
                    this.mode = 'local-bridge';
                    this.isConnected = true;
                    this.baseUrl = baseUrl;
                    
                    // 创建模拟的bridge对象，接口与VSTO一致
                    this.bridge = this._createLocalBridgeProxy(baseUrl);
                    App.excelBridge = this.bridge;
                    App.state.excelBridge = this.bridge;
                    this._installHostObjectShim();
                    
                    this.connectionInfo = {
                        type: 'Local Bridge Service',
                        url: baseUrl,
                        excelVersion: result.excelVersion,
                        version: result.serviceVersion,
                        winaxAvailable: result.winaxAvailable,
                        mockMode: result.mockMode,
                        mockReason: result.mockReason
                    };
                    
                    ChatLog.info('CONNECTOR', '[ExcelConnector] ✅ 本地代理连接成功', this.connectionInfo);
                    this._updateUI(true);
                    
                    const mockHint = result.mockMode ? '（模拟模式）' : '';
                    return {
                        success: true,
                        mode: 'local-bridge',
                        message: result.message || `已连接Excel (本地代理 ${result.excelVersion || ''}${mockHint})`
                    };
                } else {
                    throw new Error((result && result.error) || '连接失败');
                }
            } catch (error) {
                ChatLog.warn('CONNECTOR', '[ExcelConnector] ❌ 本地代理连接失败:', error.message);
                return {
                    success: false,
                    mode: null,
                    message: `本地代理连接失败: ${error.message}`
                };
            }
        }

        /**
         * 创建本地代理的Bridge代理对象
         * 模拟VSTO bridge接口，使现有代码无需修改
         */
        _createLocalBridgeProxy(baseUrl) {
            const self = this;
            const proxy = new Proxy({}, {
                get(target, prop) {
                    return async function(...args) {
                        return await self._callLocalBridge(prop, args);
                    };
                },
                
                // 支持同步调用（某些旧代码可能使用）
                apply(target, thisArg, args) {
                    ChatLog.warn('CONNECTOR', '[ExcelConnector] 同步调用不支持，请使用异步版本');
                    return Promise.reject(new Error('同步调用不支持'));
                }
            });
            
            return proxy;
        }

        async _callLocalBridge(methodName, args = []) {
            try {
                const response = await fetch(`${this.baseUrl}/invoke`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        method: methodName,
                        arguments: args
                    }),
                    signal: AbortSignal.timeout(ExcelConnector.DEFAULT_CONFIG.invokeTimeout)
                });
                
                let result;
                try {
                    result = await response.json();
                } catch (parseError) {
                    result = null;
                }
                
                if (!response.ok) {
                    const serverMessage = result?.message || result?.error || `HTTP ${response.status}`;
                    throw new Error(serverMessage);
                }
                
                if (result && result.error) {
                    throw new Error(result.error);
                }
                
                return result && result.data;
            } catch (error) {
                ChatLog.error('CONNECTOR', `[ExcelConnector] 调用失败 [${methodName}]:`, error);
                throw error;
            }
        }

        /**
         * 安装 hostObjects 垫片（仅非 WebView2 环境）。
         * 前端大量旧代码直接调用 window.chrome.webview.hostObjects.excelBridge.xxx()，
         * 该对象只在 VSTO(WebView2) 真环境存在。这里把硬编码路径重定向到当前 bridge
         * （Local Bridge 代理 或 Office.js 适配器），让现有代码无需逐处改写即可走桥。
         * 仅当不存在真实 WebView2 宿主对象时才写入，绝不覆盖 VSTO 提供的原生对象。
         */
        _installHostObjectShim() {
            try {
                if (!App || !App.excelBridge) return;
                const real = window.chrome && window.chrome.webview &&
                             window.chrome.webview.hostObjects &&
                             window.chrome.webview.hostObjects.excelBridge;
                if (real) return; // 已在 WebView2 / VSTO 真环境，保留原生对象

                window.chrome = window.chrome || {};
                window.chrome.webview = window.chrome.webview || {};
                window.chrome.webview.hostObjects = window.chrome.webview.hostObjects || {};
                window.chrome.webview.hostObjects.excelBridge = App.excelBridge;

                ChatLog.info('CONNECTOR', '[ExcelConnector] 已安装 hostObjects 垫片，硬编码调用将走本地桥');
            } catch (e) {
                ChatLog.warn('CONNECTOR', '[ExcelConnector] 安装 hostObjects 垫片失败:', e && e.message);
            }
        }

        // ==================== 模式3: Office.js Add-in ====================
        
        detectOfficeJSMode() {
            return typeof Office !== 'undefined' && Office.context && Office.context_excel;
        }

        async connectOfficeJS() {
            try {
                ChatLog.info('CONNECTOR', '[ExcelConnector] 尝试Office.js连接...');
                
                // 等待Office.js就绪
                await new Promise((resolve, reject) => {
                    if (Office.context && Office.context_excel) {
                        resolve();
                    } else {
                        Office.onReady((info) => {
                            if (info.host === Office.HostType.Excel) {
                                resolve();
                            } else {
                                reject(new Error('非Excel环境'));
                            }
                        });
                        
                        setTimeout(() => reject(new Error('Office.js初始化超时')), 10000);
                    }
                });
                
                this.mode = 'office-js';
                this.isConnected = true;
                
                // 创建Office.js适配器
                this.bridge = this._createOfficeJSAdapter();
                App.excelBridge = this.bridge;
                App.state.excelBridge = this.bridge;
                this._installHostObjectShim();
                
                this.connectionInfo = {
                    type: 'Office.js Add-in',
                    platform: 'Excel Online / Desktop'
                };
                
                ChatLog.info('CONNECTOR', '[ExcelConnector] ✅ Office.js连接成功', this.connectionInfo);
                this._updateUI(true);
                
                return {
                    success: true,
                    mode: 'office-js',
                    message: '已连接Excel (Office.js)'
                };
            } catch (error) {
                ChatLog.warn('CONNECTOR', '[ExcelConnector] ❌ Office.js连接失败:', error.message);
                return {
                    success: false,
                    mode: null,
                    message: `Office.js连接失败: ${error.message}`
                };
            }
        }

        /**
         * 创建Office.js适配器，将Office.js API映射到统一接口
         */
        _createOfficeJSAdapter() {
            const adapter = {};
            
            // 核心Excel操作
            adapter.GetSelectedRangeAsync = async () => {
                return await Excel.run(async (context) => {
                    const range = context.workbook.getSelectedRange();
                    range.load('address, text, values, columnCount, rowCount');
                    await context.sync();
                    
                    return JSON.stringify({
                        address: range.address,
                        values: range.values,
                        text: range.text,
                        columnCount: range.columnCount,
                        rowCount: range.rowCount
                    });
                });
            };
            
            adapter.WriteToRangeAsync = async (rangeAddress, dataJson) => {
                const data = JSON.parse(dataJson);
                return await Excel.run(async (context) => {
                    const range = context.workbook.getRange(rangeAddress);
                    range.values = data.values || data;
                    await context.sync();
                    return JSON.stringify({ success: true });
                });
            };
            
            adapter.ExecuteVBAAsync = async (code) => {
                // Office.js不支持直接执行VBA，返回提示信息
                ChatLog.warn('CONNECTOR', '[Office.js] 不支持VBA执行，将显示代码供用户复制');
                return JSON.stringify({
                    success: false,
                    code: code,
                    message: 'Office.js不支持VBA执行，请在桌面版Excel中使用'
                });
            };
            
            // Skill相关操作（简化实现）
            adapter.GetAllSkills = async () => {
                // 返回默认Skill列表
                return JSON.stringify(this._getDefaultSkillsForOfficeJS());
            };
            
            adapter.SaveSkill = async (skillJson) => {
                // 保存到localStorage
                const skills = JSON.parse(localStorage.getItem('officejs_skills') || '[]');
                const skill = JSON.parse(skillJson);
                skills.push(skill);
                localStorage.setItem('officejs_skills', JSON.stringify(skills));
                return JSON.stringify({ success: true });
            };
            
            // MCP相关（Office.js模式下禁用或简化）
            adapter.GetMcpStatus = async () => {
                return JSON.stringify({
                    status: 'limited',
                    message: 'Office.js模式下MCP功能受限'
                });
            };
            
            return adapter;
        }

        _getDefaultSkillsForOfficeJS() {
            return [
                { name: 'ExtractText', displayName: '提取文字(OCR)', isGeneric: true },
                { name: 'GenerateVBA', displayName: '写 VBA 代码', isGeneric: true },
                { name: 'SmartRecognize', displayName: '智能识别', isGeneric: true },
                { name: 'TableToExcel', displayName: '图片表格转Excel', isGeneric: true }
            ];
        }

        // ==================== 通用方法 ====================
        
        /**
         * 通过当前bridge调用方法（自动选择正确的调用方式）
         */
        async call(methodName, ...args) {
            if (!this.isConnected || !this.bridge) {
                throw new Error('Excel未连接');
            }
            
            try {
                switch(this.mode) {
                    case 'vsto':
                        // VSTO支持同步和异步
                        const asyncMethodName = methodName + 'Async';
                        if (typeof this.bridge[asyncMethodName] === 'function') {
                            return await this.bridge[asyncMethodName](...args);
                        } else if (typeof this.bridge[methodName] === 'function') {
                            return this.bridge[methodName](...args);
                        } else {
                            throw new Error(`方法不存在: ${methodName}`);
                        }
                        
                    case 'local-bridge':
                    case 'office-js':
                        // 这两种模式已经通过proxy/adapter处理
                        if (typeof this.bridge[methodName] === 'function') {
                            return await this.bridge[methodName](...args);
                        } else {
                            throw new Error(`方法不存在: ${methodName}`);
                        }
                        
                    default:
                        throw new Error(`不支持的连接模式: ${this.mode}`);
                }
            } catch (error) {
                ChatLog.error('CONNECTOR', `[ExcelConnector] 方法调用失败 [${methodName}]:`, error);
                throw error;
            }
        }

        /**
         * 断开连接
         */
        disconnect() {
            this.mode = 'standalone';
            this.isConnected = false;
            this.bridge = null;
            App.excelBridge = null;
            App.state.excelBridge = null;
            
            this._updateUI(false);
            ChatLog.info('CONNECTOR', '[ExcelConnector] 已断开连接');
        }

        /**
         * 获取连接状态信息
         */
        getStatus() {
            return {
                isConnected: this.isConnected,
                mode: this.mode,
                connectionInfo: this.connectionInfo,
                availableModes: this.getAvailableModes()
            };
        }

        /**
         * 获取所有可用的连接模式
         */
        getAvailableModes() {
            const modes = [];
            
            if (this.detectVSTOMode()) {
                modes.push({ id: 'vsto', name: 'VSTO原生插件', available: true });
            }
            
            modes.push({ 
                id: 'local-bridge', 
                name: '本地代理服务', 
                available: true, 
                needsSetup: true 
            });
            
            if (this.detectOfficeJSMode()) {
                modes.push({ id: 'office-js', name: 'Office.js插件', available: true });
            }
            
            return modes;
        }

        /**
         * 更新UI状态显示
         */
        _updateUI(connected) {
            const dot = document.getElementById('status-dot');
            const text = document.getElementById('status-text');
            
            if (dot && text) {
                if (connected) {
                    dot.classList.remove('disconnected');
                    
                    let modeLabel = '';
                    switch(this.mode) {
                        case 'vsto': modeLabel = '🟢 VSTO'; break;
                        case 'local-bridge': modeLabel = '🟡 本地代理'; break;
                        case 'office-js': modeLabel = '🔵 Office.js'; break;
                        default: modeLabel = '已连接';
                    }
                    
                    text.textContent = modeLabel;
                    text.title = `${this.connectionInfo.type}\n${JSON.stringify(this.connectionInfo, null, 2)}`;
                } else {
                    dot.classList.add('disconnected');
                    text.textContent = '未连接';
                    text.title = '点击连接Excel';
                }
                // 同步 tooltip 到连接状态容器（大悬停区域）
                const cs = document.getElementById('connection-status');
                if (cs) cs.title = text.textContent;
            }
            
            // 触发状态变更事件
            if (window.App && typeof App.onConnectionChange === 'function') {
                App.onConnectionChange(connected, this.mode);
            }
        }

        /**
         * 内部方法：通过VSTO bridge调用
         */
        async _callBridge(methodName, ...args) {
            const asyncMethodName = methodName + 'Async';
            
            try {
                if (typeof this.bridge[asyncMethodName] === 'function') {
                    return await this.bridge[asyncMethodName](...args);
                } else if (typeof this.bridge[methodName] === 'function') {
                    return this.bridge[methodName](...args);
                } else {
                    throw new Error(`Bridge方法不存在: ${methodName}`);
                }
            } catch (error) {
                ChatLog.error('CONNECTOR', `[ExcelConnector] Bridge调用失败 [${methodName}]:`, error);
                return { error: error.message };
            }
        }
    }

    // 导出
    window.ExcelConnector = ExcelConnector;
    window.ExcelBridgeConfig = window.ExcelBridgeConfig || ExcelConnector.DEFAULT_CONFIG;

})();