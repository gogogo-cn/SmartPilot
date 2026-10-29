/**
 * ============================================================
 * scope-wrapper.js — Scope Wrapper
 * ============================================================
 * 功能：作用域包装器，为回调函数绑定执行上下文，防止变量泄漏与作用域污染。
 *
 * 核心类/函数：
 *   - wrap() — 包装函数
 *   - bindScope() — 绑定作用域
 *   - createSafeCallback() — 创建安全回调
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 * ============================================================
 */

(function() {
    'use strict';
    
    ChatLog.info('SCOPE', '初始化函数作用域包装器...');
    
    /**
     * 关键函数列表
     * 需要确保这些函数能从全局作用域访问
     */
    const criticalFunctions = [
        'sendMessage',
        'toggleEditor', 
        'initTabs',
        'loadHistory',
        'renderHistory',
        'switchTab',
        'testAIConfig',
        'showAddSkillDialog',
        'refreshModelsList',
        'onChatModelChange',
        'executeCode',
        'saveCodeSnippet',
        'showCodeLibrary',
        'clearEditor',
        'undoLastAction',
        'redoLastAction',
        'clearAllCache',
        'onProviderChange',
        'saveSettings',
        'saveOcrSettings',
        'testCloudOcrConfig',
        'detectLocalModels'
    ];
    
    /**
     * DOM 加载完成后重写事件处理器
     * 将 onclick 属性转换为 addEventListener 调用
     */
    window.addEventListener('DOMContentLoaded', function() {
        ChatLog.info('SCOPE', '📦 包装器: DOM加载完成,开始重写事件处理器');
        
        // 遍历所有带 onclick 属性的元素
        document.querySelectorAll('[onclick]').forEach(function(element) {
            const onclickAttr = element.getAttribute('onclick');
            if (!onclickAttr) return;
            
            // 提取函数调用表达式
            // 匹配格式: functionName('arg1', 'arg2') 或 functionName(var) 或 functionName()
            const match = onclickAttr.match(/^(\w+)\((.*)\)\s*;?\s*$/);
            if (!match) return;
            
            const functionName = match[1];
            const argsStr = match[2].trim();
            
            // 移除 onclick 属性，防止浏览器默认执行
            element.removeAttribute('onclick');
            
            // 添加点击事件监听器
            element.addEventListener('click', function(e) {
                e.preventDefault();
                e.stopPropagation();
                
                ChatLog.info('SCOPE', '🖱️ 触发:', functionName, argsStr ? '(参数: ' + argsStr + ')' : '');
                
                try {
                    if (typeof window[functionName] === 'function') {
                        // 函数存在于全局作用域，执行调用
                        if (argsStr) {
                            // 解析参数
                            var args = [];
                            var argParts = argsStr.split(',').map(function(s) { return s.trim(); });
                            argParts.forEach(function(arg) {
                                // 字符串字面量
                                if ((arg.startsWith("'") && arg.endsWith("'")) ||
                                    (arg.startsWith('"') && arg.endsWith('"'))) {
                                    args.push(arg.slice(1, -1));
                                } else if (arg === 'true') {
                                    // 布尔值 true
                                    args.push(true);
                                } else if (arg === 'false') {
                                    // 布尔值 false
                                    args.push(false);
                                } else if (!isNaN(arg) && arg !== '') {
                                    // 数字
                                    args.push(Number(arg));
                                } else {
                                    // 变量引用，从 window 获取
                                    args.push(window[arg]);
                                }
                            });
                            window[functionName].apply(window, args);
                        } else {
                            // 无参数调用
                            window[functionName]();
                        }
                    } else {
                        // 函数未定义，显示错误提示
                        ChatLog.error('SCOPE', `❌ 函数 ${functionName} 未定义`);
                        if (App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function') {
                            App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">功能暂时不可用: ' + functionName + '</span>', true);
                        } else {
                            showToast(`功能暂时不可用: ${functionName}`, 'error');
                        }
                    }
                } catch (error) {
                    // 执行异常处理
                    ChatLog.error('SCOPE', `执行 ${functionName} 失败:`, error);
                    if (App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function') {
                        App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">执行失败: ' + error.message + '</span>', true);
                    } else {
                        showToast(`执行失败: ${error.message}`, 'error');
                    }
                }
            }, true);
        });
        
        ChatLog.info('SCOPE', '✅ 函数作用域包装器初始化完成');
    });
    
    /**
     * 显示提示消息
     * 降级方案：当 fixedStatusBar 不可用时使用 toast 显示
     * @param {string} message - 提示消息
     * @param {string} [type='info'] - 提示类型：info, error, success
     */
    function showToast(message, type) {
        const toast = document.getElementById('toast');
        if (toast) {
            toast.textContent = message;
            toast.className = 'toast ' + (type || 'info');
            toast.style.display = 'block';
            setTimeout(() => {
                toast.style.display = 'none';
            }, 3000);
        } else {
            ChatLog.info('SCOPE', 'Toast:', message);
        }
    }
    
})();