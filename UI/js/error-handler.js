/**
 * ============================================================
 * error-handler.js — Error Handler
 * ============================================================
 * 功能：全局错误处理模块，统一捕获、分类、展示错误信息并提供用户友好的恢复建议。
 *
 * 核心类/函数：
 *   - handleError() — 处理错误
 *   - showErrorToast() — 显示错误提示
 *   - classifyError() — 错误分类
 *   - reportError() — 上报错误
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 * ============================================================
 */

(function() {
    'use strict';

    // 错误码映射表（中文兜底；运行时优先走 App.i18n err.*）
    const ERROR_MESSAGES = {
        'network_error': '网络连接失败，请检查您的网络设置',
        'timeout': '请求超时，请稍后重试',
        'fetch_failed': '无法连接到服务器',
        'cors_error': '跨域请求被拒绝，请联系管理员',
        'excel_not_ready': 'Office 宿主未就绪，请确保应用正常运行',
        'excel_access_denied': '无法访问宿主应用，请检查权限',
        'excel_operation_failed': '操作失败',
        'cell_write_failed': '单元格写入失败',
        'ocr_failed': '文字识别失败，请尝试重新上传',
        'ocr_timeout': '文字识别超时',
        'ocr_engine_not_available': 'OCR 引擎不可用，请检查配置',
        'image_invalid': '图片格式无效，请上传 PNG/JPG 格式',
        'image_too_large': '图片过大，请压缩后重试',
        'llm_failed': 'AI 响应失败，请稍后重试',
        'llm_timeout': 'AI 响应超时',
        'llm_api_error': 'API 调用失败',
        'token_limit_exceeded': '上下文过长，请开启新对话',
        'config_load_failed': '配置加载失败',
        'config_save_failed': '配置保存失败',
        'config_invalid': '配置格式错误',
        'unknown_error': '发生未知错误',
        'permission_denied': '权限不足',
        'file_not_found': '文件不存在',
        'operation_cancelled': '操作已取消'
    };

    function localizedError(code, hostLabel) {
        const host = hostLabel || 'Office';
        try {
            if (window.App && App.i18n && typeof App.i18n.t === 'function') {
                const key = 'err.' + code;
                const s = App.i18n.t(key, { host: host });
                if (s && s !== key) return s;
            }
        } catch (_e) { /* ignore */ }
        return ERROR_MESSAGES[code] || ERROR_MESSAGES.unknown_error;
    }

    /**
     * 获取用户友好的错误消息
     * @param {Error|string} error - 错误对象或消息
     * @param {string} [context] - 错误上下文(用于日志)
     * @returns {string} 用户友好的错误消息
     */
    function getUserFriendlyMessage(error, context) {
        if (!error) return localizedError('unknown_error');
        
        const errorMessage = typeof error === 'string' ? error : (error.message || '');
        const errorCode = typeof error === 'string' ? error : (error.code || '');
        
        // 记录详细错误到控制台(开发调试用)
        if (context) {
            ChatLog.error('ERROR', `[${context}] 详细错误:`, error);
        }
        
        const hostLabel = (window.SmartPilotHost && window.SmartPilotHost.label) || 'Office';

        // 优先使用错误码映射（宿主相关文案动态替换）
        if (errorCode && ERROR_MESSAGES[errorCode]) {
            return localizedError(errorCode, hostLabel);
        }
        
        // 根据错误消息关键词匹配
        const lowerMessage = errorMessage.toLowerCase();
        
        // 网络相关
        if (lowerMessage.includes('failed to fetch') || 
            lowerMessage.includes('network') || 
            lowerMessage.includes('cors')) {
            return localizedError('network_error', hostLabel);
        }
        
        if (lowerMessage.includes('timeout')) {
            return localizedError('timeout', hostLabel);
        }
        
        // 宿主 COM / 互操作。必须按词边界匹配，避免 "completed" 里的 "com" 被当成 Excel 故障。
        if (/\bexcel\b|excelbridge|interop|comexception|\bcom\b/i.test(errorMessage)) {
            return localizedError('excel_operation_failed', hostLabel);
        }
        
        if (lowerMessage.includes('permission') || 
            lowerMessage.includes('access denied')) {
            return localizedError('permission_denied', hostLabel);
        }
        
        // OCR相关
        if (lowerMessage.includes('ocr') || 
            lowerMessage.includes('recognize')) {
            return localizedError('ocr_failed', hostLabel);
        }
        
        if (lowerMessage.includes('image') || 
            lowerMessage.includes('picture')) {
            return localizedError('image_invalid', hostLabel);
        }
        
        // LLM相关
        if (lowerMessage.includes('llm') || 
            lowerMessage.includes('api') || 
            lowerMessage.includes('token')) {
            return localizedError('llm_failed', hostLabel);
        }
        
        // 文件相关
        if (lowerMessage.includes('not found') || 
            lowerMessage.includes('不存在')) {
            return localizedError('file_not_found', hostLabel);
        }
        
        // 默认返回原始消息(但去除敏感信息)
        return sanitizeErrorMessage(errorMessage);
    }

    /**
     * 清理错误消息,移除敏感信息
     * @param {string} message - 原始错误消息
     * @returns {string} 清理后的消息
     */
    function sanitizeErrorMessage(message) {
        if (!message) return ERROR_MESSAGES.unknown_error;
        
        // 移除文件路径
        message = message.replace(/C:\\[^\s]+/gi, '[文件路径]');
        message = message.replace(/\/[^\s]+/g, '[文件路径]');
        
        // 移除API密钥
        message = message.replace(/sk-[a-zA-Z0-9]{20,}/g, '[API密钥]');
        message = message.replace(/Bearer [^\s]+/g, '[认证令牌]');
        
        // 移除堆栈信息
        message = message.split('\n')[0]; // 只保留第一行
        
        // 限制长度
        if (message.length > 200) {
            message = message.substring(0, 200) + '...';
        }
        
        return message;
    }

    /**
     * 输入区上方错误横幅（主可见面，避免只出现在底部状态栏被忽略）
     */
    function showErrorBanner(message, type) {
        const banner = document.getElementById('app-error-banner');
        if (!banner) return;
        const text = String(message || '').trim();
        if (!text) {
            banner.style.display = 'none';
            banner.textContent = '';
            return;
        }
        const icon = type === 'warning' ? '⚠️' : (type === 'info' ? 'ℹ️' : '❌');
        banner.className = 'app-error-banner' + (type === 'warning' ? ' is-warning' : (type === 'info' ? ' is-info' : ''));
        banner.innerHTML = '';
        const span = document.createElement('span');
        span.className = 'app-error-banner-text';
        span.textContent = icon + ' ' + text;
        const close = document.createElement('button');
        close.type = 'button';
        close.className = 'app-error-banner-close';
        close.title = '关闭';
        close.textContent = '×';
        close.onclick = function () { showErrorBanner(''); };
        banner.appendChild(span);
        banner.appendChild(close);
        banner.style.display = 'flex';
        clearTimeout(showErrorBanner._timer);
        showErrorBanner._timer = setTimeout(function () {
            if (banner.querySelector('.app-error-banner-text') &&
                banner.querySelector('.app-error-banner-text').textContent === span.textContent) {
                showErrorBanner('');
            }
        }, type === 'error' ? 8000 : 5000);
    }

    /**
     * 显示错误提示(统一入口)
     * 主可见面：Toast + 输入区横幅；次要：底部状态栏。三者同步，避免漏看。
     * @param {Error|string} error - 错误对象或消息
     * @param {string} [context] - 错误上下文
     * @param {string} [type='error'] - 提示类型(error/warning/info)
     */
    function showError(error, context, type = 'error') {
        const userMessage = getUserFriendlyMessage(error, context);
        const toastType = (type === 'warning' || type === 'info') ? type : 'error';

        // 1) Toast（主）
        try {
            if (window.App && typeof App.showToast === 'function') {
                App.showToast(userMessage, toastType);
            } else if (typeof window.showToast === 'function') {
                window.showToast(userMessage, toastType);
            }
        } catch (_) { /* ignore */ }

        // 2) 输入区横幅（主，error/warning）
        if (toastType === 'error' || toastType === 'warning') {
            try { showErrorBanner(userMessage, toastType); } catch (_) { /* ignore */ }
        }

        // 3) 底部状态栏（次）
        try {
            const icon = toastType === 'error' ? '❌' : (toastType === 'warning' ? '⚠️' : 'ℹ️');
            if (window.App && App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function') {
                App.fixedStatusBar.setExtraStatus(
                    ' <span style="color:#999;font-size:11px;margin-left:8px;">' + icon + ' ' + userMessage + '</span>',
                    true
                );
            }
        } catch (_) { /* ignore */ }
    }

    /**
     * 包装异步函数,自动处理错误
     * @param {Function} asyncFn - 异步函数
     * @param {string} context - 错误上下文
     * @returns {Function} 包装后的函数
     */
    function withErrorHandling(asyncFn, context) {
        return async function(...args) {
            try {
                return await asyncFn.apply(this, args);
            } catch (error) {
                showError(error, context);
                throw error; // 重新抛出,调用者可继续处理
            }
        };
    }

    // 导出到全局
    window.ErrorHandler = {
        getUserFriendlyMessage: getUserFriendlyMessage,
        showError: showError,
        showErrorBanner: showErrorBanner,
        withErrorHandling: withErrorHandling,
        ERROR_MESSAGES: ERROR_MESSAGES
    };
    window.showAppError = showError;
    window.showErrorBanner = showErrorBanner;

    // =========================================================
    // 全局未捕获异常边界 (P2: 前端错误边界)
    // 捕获未被组件级 try/catch 处理的错误，防止白屏/静默失败
    // =========================================================

    /** @type {number} 全局错误限流：30秒内最多3次 */
    let _globalErrorCount = 0;
    let _globalErrorTimer = 0;

    function errorTextOf(value) {
        if (value == null) return '';
        if (typeof value === 'string') return value;
        if (value.message) return String(value.message);
        try {
            if (value.parameters && value.parameters.error) return String(value.parameters.error);
        } catch (_) { /* ignore */ }
        return '';
    }

    /** 浏览器布局观察器的良性提示，不是程序异常，也不该弹「操作失败」。 */
    function isBenignBrowserError(text) {
        return typeof text === 'string' && text.indexOf('ResizeObserver loop') !== -1;
    }

    window.addEventListener('error', function(event) {
        // 跳过资源加载失败（由各脚本文件自身处理）
        // 资源加载错误的 event.target 是元素节点，且 event.error 为空
        // 此类错误不应作为"未捕获异常"上报，否则空 src="" 的 <img> 等会刷屏
        if (event.target && event.target.nodeType === 1) {
            var tag = (event.target.tagName || '').toUpperCase();
            var RESOURCE_TAGS = { SCRIPT: 1, LINK: 1, IMG: 1, IFRAME: 1, AUDIO: 1, VIDEO: 1, SOURCE: 1, OBJECT: 1 };
            if (RESOURCE_TAGS[tag]) return;
        }

        var rawText = errorTextOf(event && event.message) || errorTextOf(event && event.error);
        if (isBenignBrowserError(rawText)) return;

        _globalErrorCount++;
        if (_globalErrorTimer) clearTimeout(_globalErrorTimer);
        _globalErrorTimer = setTimeout(function() { _globalErrorCount = 0; }, 30000);

        if (_globalErrorCount > 3) return; // 限流：30秒内超过3次静默丢弃

        var userMsg = getUserFriendlyMessage(event.error || event.message, 'GLOBAL_ERROR');
        ChatLog.error('ERROR', '[GlobalErrorBoundary] 未捕获异常:', event.error || event.message);
        showError(userMsg, 'GLOBAL_ERROR', 'error');
    }, true);

    window.addEventListener('unhandledrejection', function(event) {
        var reason = event.reason;
        var detail = '';
        try {
            if (!reason) detail = '';
            else if (typeof reason === 'string') detail = reason;
            else if (reason.message) detail = String(reason.message);
            else if (reason.parameters && reason.parameters.error) detail = String(reason.parameters.error);
            else detail = JSON.stringify(reason);
        } catch (_) { detail = String(reason || ''); }

        if (isBenignBrowserError(detail) || isBenignBrowserError(errorTextOf(reason))) {
            event.preventDefault();
            return;
        }

        // WebView2 宿主桥：缺方法(0x80020006) / 投影失败(e is not a function) — 非用户可见故障
        if (/0x80020006|未知名称|e is not a function|bridge-call-failed/i.test(detail)) {
            event.preventDefault();
            return;
        }

        _globalErrorCount++;
        if (_globalErrorTimer) clearTimeout(_globalErrorTimer);
        _globalErrorTimer = setTimeout(function() { _globalErrorCount = 0; }, 30000);

        if (_globalErrorCount > 3) {
            ChatLog.warn('ERROR', '[GlobalErrorBoundary] 限流丢弃 Promise 拒绝:', event.reason);
            event.preventDefault();
            return;
        }

        var userMsg = getUserFriendlyMessage(
            (reason && reason.message) || detail || reason || 'Promise 异步操作失败',
            'PROMISE_REJECTION'
        );
        ChatLog.error('ERROR', '[GlobalErrorBoundary] 未处理Promise拒绝:', reason);
        showError(userMsg, 'PROMISE_REJECTION', 'error');

        // 防止浏览器控制台报 UnhandledPromiseRejection
        event.preventDefault();
    });

    ChatLog.info('ERROR', '[ErrorHandler] 前端错误处理工具 + 全局错误边界已加载');

})();
