/**
 * ============================================================
 * app-mobile.js — Mobile Adaptation
 * ============================================================
 * 功能：移动端适配模块，检测移动设备、调整布局、优化触摸交互与屏幕尺寸适配。
 *
 * 核心类/函数：
 *   - isMobile() — 检测移动设备
 *   - adaptLayout() — 适配布局
 *   - handleTouchEvents() — 处理触摸事件
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 * ============================================================
 */

(function() {
    'use strict';

    const App = window.App = window.App || {};

    /** 是否正在运行 */
    let _isRunning = false;
    /** 服务 URL */
    let _serverUrl = '';
    /** 是否已自动启动过 */
    let _autoStarted = false;

    /**
     * DOM 元素快捷获取函数
     * @param {string} id - 元素 ID
     * @returns {HTMLElement|null} DOM 元素
     */
    function $(id) { return document.getElementById(id); }

    /**
     * 获取 Excel Bridge 对象
     * 尝试多种获取方式，按优先级排序
     * @returns {object|null} Bridge 对象或 null
     */
    function getBridge() {
        // 方式1：WebView2 hostObjects
        if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
            return window.chrome.webview.hostObjects.excelBridge;
        }
        // 方式2：全局 excelBridge
        if (window.excelBridge) return window.excelBridge;
        // 方式3：App.bridge
        if (window.App && window.App.bridge) return window.App.bridge;
        return null;
    }

    /**
     * 更新 UI 状态
     * @param {boolean} running - 服务是否运行
     * @param {string} [url] - 服务 URL
     * @param {string} [ip] - 局域网 IP
     * @param {number} [port] - 服务端口
     */
    function updateUI(running, url, ip, port) {
        _isRunning = running;
        _serverUrl = url || '';

        const badge = $('mobile-status-badge');
        const qrSection = $('mobile-qr-section');
        const startBtn = $('mobile-start-btn');
        const stopBtn = $('mobile-stop-btn');
        const infoText = $('mobile-info-text');

        if (running) {
            // 运行中状态
            badge.textContent = '● 服务运行中';
            badge.style.background = '#e8f5e9';
            badge.style.color = '#2e7d32';

            qrSection.style.display = 'block';
            startBtn.style.display = 'none';
            stopBtn.style.display = 'inline-block';

            if (url) {
                $('mobile-url-text').textContent = url;
                generateQR(url);
            }

            infoText.innerHTML = `
                <div>局域网 IP: <b>${ip || '-'}</b></div>
                <div>端口: <b>${port || '-'}</b></div>
                <div style="margin-top: 8px; color: #aaa;">提示: 手机和电脑需连接同一 WiFi</div>
            `;
        } else {
            // 未运行状态
            badge.textContent = '○ 服务未启动';
            badge.style.background = '#f5f5f5';
            badge.style.color = '#999';

            qrSection.style.display = 'none';
            startBtn.style.display = 'inline-block';
            stopBtn.style.display = 'none';
            infoText.innerHTML = '<div>点击"启动服务"开始</div>';
        }
    }

    /**
     * 生成 QR 码
     * 使用 qrcodejs（davidshimjs）构造器模式：库会在容器内创建 canvas/table
     * @param {string} text - 要编码的文本
     */
    function generateQR(text) {
        const container = $('mobile-qr-canvas');
        if (!container || typeof QRCode === 'undefined') return;

        try {
            container.innerHTML = '';
            new QRCode(container, {
                text: text,
                width: 280,
                height: 280,
                colorDark: '#000000',
                colorLight: '#ffffff',
                correctLevel: QRCode.CorrectLevel.H
            });
        } catch (e) {
            console.error('[App.mobile] QR 码生成失败:', e);
            container.style.display = 'none';
            $('mobile-url-text').innerHTML = `<span style="color:#f44336">QR 码生成失败: ${e.message}</span><br>${text}`;
        }
    }

    /**
     * 调用 Bridge 方法
     * @param {string} method - 方法名
     * @param {...*} args - 参数列表
     * @returns {object|null} 调用结果或 null
     */
    function callBridge(method, ...args) {
        const bridge = getBridge();
        if (!bridge) {
            console.warn('[App.mobile] Bridge 未连接');
            return null;
        }
        try {
            const result = bridge[method](...args);
            return result ? JSON.parse(result) : null;
        } catch (e) {
            console.error(`[App.mobile] Bridge.${method} 调用失败:`, e);
            return null;
        }
    }

    /**
     * 公开 API - 移动端遥控模块
     */
    App.mobile = {
        /**
         * 初始化模块
         * 检查服务状态并更新 UI
         */
        init() {
            this.checkStatus();
        },

        /**
         * 检查服务状态
         * 通过 Bridge 获取服务运行状态
         */
        checkStatus() {
            const result = callBridge('GetMobileControlStatus');
            if (result && result.success) {
                updateUI(result.isRunning, result.url, result.ip, result.port);
            } else {
                updateUI(false);
            }
        },

        /**
         * 启动服务
         * 调用 Bridge 启动移动端遥控 HTTP 服务
         */
        start() {
            const badge = $('mobile-status-badge');
            badge.textContent = '⏳ 启动中';
            badge.style.background = '#fff3e0';
            badge.style.color = '#e65100';

            // 使用 setTimeout 避免阻塞 UI
            setTimeout(() => {
                const result = callBridge('StartMobileControl', -1);
                if (result && result.success) {
                    updateUI(true, result.url, result.ip, result.port);
                } else {
                    const errMsg = (result && result.error) || '启动失败';
                    badge.textContent = '✕ 启动失败';
                    badge.style.background = '#fde7e7';
                    badge.style.color = '#c62828';
                    console.error('[App.mobile] 启动失败:', errMsg);
                    alert('启动移动端服务失败:\n' + errMsg + '\n\n请尝试以管理员身份运行 Excel。');
                }
            }, 100);
        },

        /**
         * 停止服务
         * 弹出确认框后调用 Bridge 停止服务
         */
        stop() {
            if (!confirm('确定要停止移动端遥控服务吗？')) return;

            const result = callBridge('StopMobileControl');
            if (result && result.success) {
                updateUI(false);
            } else {
                alert('停止服务失败');
            }
        },

        /**
         * 刷新状态
         * 重新检查服务状态
         */
        refresh() {
            this.checkStatus();
        },

        /**
         * 切换到手机标签时调用
         * 首次激活时自动检查状态
         */
        onTabActivate() {
            if (!_autoStarted) {
                _autoStarted = true;
                this.checkStatus();
            }
        }
    };

})();