/**
 * ============================================================
 * app-history.js — History & Storage
 * ============================================================
 * 功能：历史记录与存储模块，管理聊天记录、配置持久化、欢迎面板与轮播展示。
 *
 * 核心类/函数：
 *   - loadHistory() — 加载历史记录
 *   - saveHistory() — 保存历史记录
 *   - renderHistoryList() — 渲染历史列表
 *   - showWelcomePanel() — 显示欢迎面板
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 *   - app-utils.js
 * ============================================================
 */

(function() {
    'use strict';

    function wt(key, vars) {
        try {
            if (window.App && App.i18n && typeof App.i18n.t === 'function') return App.i18n.t(key, vars);
        } catch (_e) { /* ignore */ }
        return key;
    }

    // HTML 转义（防止 XSS）
    function escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.appendChild(document.createTextNode(text));
        return div.innerHTML;
    }

    // --- 来自 app.js:L4299-L4315 ---
    function updateContextStats() {
        // 安全检查：确保 ContextManager 已定义
        if (typeof ContextManager === 'undefined' || !ContextManager) {
            ChatLog.info('CONTEXT', '[Context] ContextManager 尚未初始化，跳过统计更新');
            return;
        }

        const stats = ContextManager.getContextStats();
        const statsDiv = document.getElementById('context-stats');
        if (statsDiv) {
            statsDiv.innerHTML = `
                当前对话: ${stats.messageCount} 条消息 |
                预估 Token: ${stats.totalTokens.toLocaleString()} / ${stats.limit.toLocaleString()} |
                使用率: ${stats.usage}%
            `;
        }
    }

    // ==================== 历史记录 ====================

    // --- 来自 app.js:L4320-L4329 ---
    // 检查 localStorage 是否可用
    function isLocalStorageAvailable() {
        try {
            const test = '__storage_test__';
            localStorage.setItem(test, test);
            localStorage.removeItem(test);
            return true;
        } catch (e) {
            return false;
        }
    }

    // --- 来自 app.js:L4332-L4365 ---
    // 初始化存储模式
    // 优先级: localStorage > indexedDB > memory
    function initStorageMode() {
        if (isLocalStorageAvailable()) {
            App.state.storageMode = 'localStorage';
            ChatLog.info('STORAGE', '使用 localStorage 存储模式');
            hideStorageIndicator();
            return;
        }

        // localStorage 不可用，尝试 IndexedDB
        if (typeof StorageDB !== 'undefined' && StorageDB.isAvailable()) {
            App.state.storageMode = 'indexedDB';
            ChatLog.info('STORAGE', '使用 IndexedDB 持久化存储模式（localStorage 不可用）');
            showStorageIndicator('IndexedDB');
            return;
        }

        // 最终降级到内存存储
        App.state.storageMode = 'memory';
        ChatLog.info('STORAGE', '使用内存存储模式（IDE内置浏览器环境）');
        showStorageIndicator('内存');
    }

    // --- 来自 app.js:L4345-L4360 ---
    // 显示存储模式指示器
    function showStorageIndicator(mode) {
        const indicator = document.getElementById('storage-mode-indicator');
        if (indicator) {
            indicator.style.display = 'block';
            // 显示当前存储模式
            const label = indicator.querySelector('.storage-mode-label');
            if (label && mode) {
                label.textContent = '存储: ' + mode;
                label.title = '当前使用 ' + mode + ' 存储模式' +
                    (mode === '内存' ? '（刷新页面后数据会丢失）' :
                     mode === 'IndexedDB' ? '（数据持久化保存）' : '');
            }
        }
    }

    // --- 来自 app.js:L4353-L4358 ---
    // 隐藏存储模式指示器
    function hideStorageIndicator() {
        const indicator = document.getElementById('storage-mode-indicator');
        if (indicator) {
            indicator.style.display = 'none';
        }
    }

    // --- 来自 app.js:L4360-L4400 ---
    async function loadHistory() {
        initStorageMode();

        if (App.state.storageMode === 'memory') {
            // 使用内存存储
            App.state.chatHistory = App.state.memoryStorage.history || [];
            ChatLog.info('HISTORY', '从历史记录已从内存加载, 数量:', App.state.chatHistory.length);
            return;
        }

        if (App.state.storageMode === 'indexedDB') {
            // 使用 IndexedDB 持久化存储
            try {
                const saved = await StorageDB.get(App.state.HISTORY_KEY);
                if (saved && Array.isArray(saved)) {
                    App.state.chatHistory = saved;
                    ChatLog.info('HISTORY', '历史记录已从 IndexedDB 加载, 数量:', App.state.chatHistory.length);
                } else {
                    App.state.chatHistory = [];
                    ChatLog.info('HISTORY', 'IndexedDB 中无历史记录，使用空数组');
                }
            } catch (e) {
                ChatLog.error('HISTORY', '从 IndexedDB 加载历史记录失败:', e);
                App.state.chatHistory = [];
            }
            return;
        }

        // 使用 localStorage
        const saved = localStorage.getItem(App.state.HISTORY_KEY);
        if (saved) {
            try {
                App.state.chatHistory = JSON.parse(saved);
                ChatLog.info('HISTORY', '历史记录已加载, 数量:', App.state.chatHistory.length);
            } catch (e) {
                ChatLog.error('HISTORY', '解析历史记录失败:', e);
                App.state.chatHistory = [];
            }
        }
    }

    // --- 来自 app.js:L4383-L4480 ---
    async function saveHistory() {
        if (App.state.storageMode === 'memory') {
            // 使用内存存储
            App.state.memoryStorage.history = [...App.state.chatHistory];
            return true;
        }

        if (App.state.storageMode === 'indexedDB') {
            // 使用 IndexedDB 持久化存储
            try {
                const result = await StorageDB.set(App.state.HISTORY_KEY, App.state.chatHistory);
                if (!result) {
                    ChatLog.warn('HISTORY', '[StorageDB] 保存历史记录返回 false');
                }
                return result;
            } catch (e) {
                ChatLog.error('HISTORY', '[StorageDB] 保存历史记录失败:', e);
                return false;
            }
        }

        // 使用 localStorage
        try {
            localStorage.setItem(App.state.HISTORY_KEY, JSON.stringify(App.state.chatHistory));
            return true;
        } catch (e) {
            ChatLog.error('HISTORY', '保存历史记录失败:', e);

            // 如果是配额超限错误，尝试迁移到 IndexedDB（而不是清空数据！）
            if (e.name === 'QuotaExceededError' || e.code === 22) {
                ChatLog.warn('HISTORY', '⚠️ localStorage 存储空间已满，尝试迁移到 IndexedDB...');

                // 尝试迁移到 IndexedDB
                if (typeof StorageDB !== 'undefined' && StorageDB.isAvailable()) {
                    try {
                        await StorageDB.set(App.state.HISTORY_KEY, App.state.chatHistory);
                        // 切换存储模式
                        App.state.storageMode = 'indexedDB';
                        ChatLog.info('HISTORY', '✅ 已成功迁移到 IndexedDB 持久化存储');
                        App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">存储空间已满，已自动切换到持久化存储模式</span>', true) : App.showToast('存储空间已满，已自动切换到持久化存储模式', 'success');
                        showStorageIndicator('IndexedDB');
                        return true;
                    } catch (migrateErr) {
                        ChatLog.error('HISTORY', '迁移到 IndexedDB 也失败:', migrateErr);
                    }
                }

                // IndexedDB 也不可用，才执行原来的清理策略
                ChatLog.warn('HISTORY', '⚠️ IndexedDB 不可用，执行原有清理策略...');

                // 策略1：减少历史记录数量（从30条减到10条）
                if (App.state.chatHistory.length > 10) {
                    const removed = App.state.chatHistory.splice(10);
                    ChatLog.info('HISTORY', `🗑️ 已清理 ${removed.length} 条旧记录`);
                    try {
                        localStorage.setItem(App.state.HISTORY_KEY, JSON.stringify(App.state.chatHistory));
                        App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">️ 存储空间不足，已自动清理旧记录</span>', true) : App.showToast('️ 存储空间不足，已自动清理旧记录', 'warning');
                        return true;
                    } catch (e2) {
                        ChatLog.error('HISTORY', '清理后仍然超限:', e2);
                    }
                }

                // 策略2：清空所有历史记录（最后手段）
                ChatLog.warn('HISTORY', '🗑️ 所有策略失败，只能清空所有历史记录');
                App.state.chatHistory = [];
                try {
                    localStorage.setItem(App.state.HISTORY_KEY, '[]');
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">️ 存储空间严重不足，历史记录已被清空</span>', true) : App.showToast('️ 存储空间严重不足，历史记录已被清空', 'warning');
                    renderHistory();
                    return true;
                } catch (e3) {
                    ChatLog.error('HISTORY', '清空历史记录失败:', e3);
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;"> 存储空间严重不足，请清除浏览器缓存</span>', true) : App.showToast('❌ 存储空间严重不足，请清除浏览器缓存', 'error');
                    return false;
                }
            }

            return false;
        }
    }

    // ==================== 清空缓存功能 ====================

    // --- 来自 app.js:L4434-L4523 ---
    function clearAllCache() {
        // 创建自定义选择对话框
        const dialogHtml = `
            <div id="cache-clear-dialog" style="
                position: fixed;
                top: 0;
                left: 0;
                width: 100%;
                height: 100%;
                background: rgba(0,0,0,0.5);
                display: flex;
                align-items: center;
                justify-content: center;
                z-index: 10000;
            ">
                <div style="
                    background: white;
                    border-radius: 8px;
                    padding: 24px;
                    max-width: 400px;
                    width: 90%;
                    box-shadow: 0 4px 20px rgba(0,0,0,0.3);
                ">
                    <h3 style="margin: 0 0 16px 0; font-size: 18px; color: #333;" data-i18n="clear.title">选择要清空的数据</h3>
                    <div style="margin-bottom: 20px;">
                        <label style="display: flex; align-items: center; padding: 10px; border-radius: 4px; cursor: pointer; transition: background 0.2s;" onmouseover="this.style.background='#f5f5f5'" onmouseout="this.style.background='transparent'">
                            <input type="checkbox" id="clear-history" checked style="margin-right: 10px; width: 18px; height: 18px;">
                            <span style="font-size: 14px; color: #555;">
                                <strong data-i18n="clear.history">对话历史记录</strong><br>
                                <small style="color: #999;" data-i18n="clear.historyDesc">所有聊天记录和对话历史</small>
                            </span>
                        </label>
                        <label style="display: flex; align-items: center; padding: 10px; border-radius: 4px; cursor: pointer; transition: background 0.2s;" onmouseover="this.style.background='#f5f5f5'" onmouseout="this.style.background='transparent'">
                            <input type="checkbox" id="clear-config" style="margin-right: 10px; width: 18px; height: 18px;">
                            <span style="font-size: 14px; color: #555;">
                                <strong data-i18n="clear.config">应用配置</strong><br>
                                <small style="color: #999;" data-i18n="clear.configDesc">API设置、模型选择等配置</small>
                            </span>
                        </label>
                        <label style="display: flex; align-items: center; padding: 10px; border-radius: 4px; cursor: pointer; transition: background 0.2s;" onmouseover="this.style.background='#f5f5f5'" onmouseout="this.style.background='transparent'">
                            <input type="checkbox" id="clear-custom-types" style="margin-right: 10px; width: 18px; height: 18px;">
                            <span style="font-size: 14px; color: #555;">
                                <strong data-i18n="clear.types">自定义识别类型</strong><br>
                                <small style="color: #999;" data-i18n="clear.typesDesc">用户添加的自定义识别类型</small>
                            </span>
                        </label>
                        <label style="display: flex; align-items: center; padding: 10px; border-radius: 4px; cursor: pointer; transition: background 0.2s;" onmouseover="this.style.background='#f5f5f5'" onmouseout="this.style.background='transparent'">
                            <input type="checkbox" id="clear-snippets" style="margin-right: 10px; width: 18px; height: 18px;">
                            <span style="font-size: 14px; color: #555;">
                                <strong data-i18n="clear.snippets">代码片段</strong><br>
                                <small style="color: #999;" data-i18n="clear.snippetsDesc">保存的代码片段和代码库</small>
                            </span>
                        </label>

                    </div>
                    <div style="display: flex; gap: 10px; justify-content: flex-end;">
                        <button onclick="App.closeCacheClearDialog()" style="
                            padding: 8px 16px;
                            border: 1px solid #ddd;
                            background: white;
                            border-radius: 4px;
                            cursor: pointer;
                            font-size: 14px;
                            color: #666;
                        " data-i18n="clear.cancel">取消</button>
                        <button onclick="App.executeCacheClear()" style="
                            padding: 8px 16px;
                            border: none;
                            background: #f44336;
                            color: white;
                            border-radius: 4px;
                            cursor: pointer;
                            font-size: 14px;
                        " data-i18n="clear.confirm">确认清空</button>
                    </div>
                </div>
            </div>
        `;

        // 添加对话框到页面
        const div = document.createElement('div');
        div.innerHTML = dialogHtml;
        document.body.appendChild(div);
        try { if (window.App && App.i18n) App.i18n.apply(div); } catch (e) {}
    }

    // --- 来自 app.js:L4526-L4533 ---
    // 关闭清空缓存对话框
    function closeCacheClearDialog() {
        const dialog = document.getElementById('cache-clear-dialog');
        if (dialog) {
            dialog.remove();
        }
    }
    // 暴露到全局作用域（onclick 需要）
    window.App.closeCacheClearDialog = closeCacheClearDialog;

    // 自定义确认对话框（替代原生 confirm()，解决 WebView2 中 confirm() 不工作的问题）
    function showCustomConfirm(message) {
        return new Promise((resolve) => {
            const overlay = document.createElement('div');
            overlay.id = 'custom-confirm-overlay';
            overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:100000;display:flex;align-items:center;justify-content:center;';
            
            const dialog = document.createElement('div');
            dialog.style.cssText = 'background:white;border-radius:8px;padding:24px;max-width:400px;width:90%;box-shadow:0 4px 20px rgba(0,0,0,0.3);';
            dialog.innerHTML = `
                <div style="font-size:16px;font-weight:500;margin-bottom:20px;color:#333;">${message}</div>
                <div style="display:flex;gap:10px;justify-content:flex-end;">
                    <button id="custom-confirm-cancel" style="padding:8px 16px;border:1px solid #ddd;background:white;border-radius:4px;cursor:pointer;font-size:14px;color:#666;">取消</button>
                    <button id="custom-confirm-ok" style="padding:8px 16px;border:none;background:#f44336;color:white;border-radius:4px;cursor:pointer;font-size:14px;">确定</button>
                </div>
            `;
            
            overlay.appendChild(dialog);
            document.body.appendChild(overlay);
            
            const cleanup = () => {
                if (overlay.parentNode) overlay.parentNode.removeChild(overlay);
            };
            
            document.getElementById('custom-confirm-cancel').onclick = () => {
                cleanup();
                resolve(false);
            };
            
            document.getElementById('custom-confirm-ok').onclick = () => {
                cleanup();
                resolve(true);
            };
        });
    }
    // 暴露到全局作用域
    window.App.showCustomConfirm = showCustomConfirm;

    // --- 来自 app.js:L4536-L4640 ---
    // 执行清空缓存操作
    async function executeCacheClear() {
        const clearHistory = document.getElementById('clear-history').checked;
        const clearConfig = document.getElementById('clear-config').checked;
        const clearCustomTypes = document.getElementById('clear-custom-types').checked;
        const clearSnippets = document.getElementById('clear-snippets').checked;

        // 检查是否至少选择了一项
        if (!clearHistory && !clearConfig && !clearCustomTypes && !clearSnippets) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">请至少选择一项要清空的数据</span>', true) : App.showToast('请至少选择一项要清空的数据');
            return;
        }

        // 二次确认（使用自定义对话框，避免 WebView2 中 confirm() 不工作的问题）
        if (!await showCustomConfirm('确定要清空选中的数据吗？此操作不可恢复！')) {
            return;
        }

        try {
            let clearedItems = [];

            // 1. 清空历史记录（同时清理 localStorage 和 IndexedDB）
            if (clearHistory) {
                App.state.chatHistory = [];
                // 【修复】memoryStorage在App对象根级别，不在state内
                if (App.memoryStorage) {
                    App.memoryStorage.history = [];
                }
                // 【修复】HISTORY_KEY在App对象根级别
                const historyKey = App.HISTORY_KEY || 'excel_ai_history';
                localStorage.removeItem(historyKey);
                // 同时清理 IndexedDB
                if (typeof StorageDB !== 'undefined' && StorageDB.isAvailable()) {
                    await StorageDB.remove(historyKey);
                }
                clearedItems.push('对话历史记录');

                // 刷新历史记录显示
                renderHistory();

                // 清空当前对话
                document.getElementById('chat-messages').innerHTML = '';
            }

            // 2. 清空配置
            if (clearConfig) {
                // 【修复】CONFIG_KEY在App对象根级别
                const configKey = App.CONFIG_KEY || 'excel_ai_config';
                localStorage.removeItem(configKey);
                // 【修复】memoryStorage在App对象根级别
                if (App.memoryStorage) {
                    App.memoryStorage.config = null;
                }
                // 【修复】App.DEFAULT_CONFIG不存在，直接清空配置对象
                if (App.config) {
                    App.config = {};
                }
                clearedItems.push('应用配置');
            }

            // 3. 清空自定义识别类型
            if (clearCustomTypes) {
                localStorage.removeItem('customRecognitionTypes');
                localStorage.removeItem('modifiedBuiltInTypes');
                localStorage.removeItem('excel_ai_custom_types');
                clearedItems.push('自定义识别类型');
            }

            // 4. 清空代码片段
            if (clearSnippets) {
                App.state.codeSnippets = [];
                // 【修复】memoryStorage在App对象根级别
                if (App.memoryStorage) {
                    App.memoryStorage.snippets = [];
                }
                // 【修复】SNIPPETS_KEY在App对象根级别
                const snippetsKey = App.SNIPPETS_KEY || 'excel_ai_code_snippets';
                localStorage.removeItem(snippetsKey);
                clearedItems.push('代码片段');
            }

            // 关闭对话框
            closeCacheClearDialog();

            // 添加系统消息
            if (clearedItems.length > 0) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">已清空以下数据：' + clearedItems.join('、') + '</span>', true) : App.showToast('已清空以下数据：' + clearedItems.join('、'));
            }

            ChatLog.info('CACHE', '✅ 已清空数据:', clearedItems);
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">成功清空 ' + clearedItems.length + ' 项数据</span>', true) : App.showToast('成功清空 ' + clearedItems.length + ' 项数据');

            // 询问是否刷新页面（使用自定义对话框）
            if (clearConfig && await showCustomConfirm('配置已重置，是否刷新页面以应用更改？')) {
                location.reload();
            }
        } catch (e) {
            ChatLog.error('CACHE', '清空缓存失败:', e);
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">清空缓存失败: ' + (e.message || '未知错误') + '</span>', true) : App.showToast('清空缓存失败: ' + (e.message || '未知错误'));
        }
    }
    // 暴露到全局作用域（onclick 需要）
    window.App.executeCacheClear = executeCacheClear;

    // --- 来自 app.js:L4625-L4647 ---
    // 将 base64 图片缩放为小缩略图（用于历史记录显示）
    function resizeBase64Image(base64Src, maxWidth, maxHeight) {
        try {
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            const img = new Image();
            img.src = base64Src;
            
            // 等待图片加载完成（使用 onload 回调，兼容异步加载）
            if (!img.complete || img.naturalWidth === 0) {
                // 异步方式：使用 Promise
                return new Promise((resolve) => {
                    img.onload = function() {
                        try {
                            let w = img.naturalWidth || img.width || 200;
                            let h = img.naturalHeight || img.height || 150;
                            if (w > maxWidth || h > maxHeight) {
                                const ratio = Math.min(maxWidth / w, maxHeight / h);
                                w = Math.round(w * ratio);
                                h = Math.round(h * ratio);
                            }
                            canvas.width = w;
                            canvas.height = h;
                            ctx.drawImage(img, 0, 0, w, h);
                            resolve(canvas.toDataURL('image/jpeg', 0.7));
                        } catch (e) {
                            ChatLog.warn('HISTORY', '[缩略图] 异步生成失败:', e.message);
                            resolve(null);
                        }
                    };
                    img.onerror = function() {
                        ChatLog.warn('HISTORY', '[缩略图] 图片加载失败');
                        resolve(null);
                    };
                });
            }
            
            // 同步方式：图片已加载完成
            let w = img.naturalWidth || img.width || 200;
            let h = img.naturalHeight || img.height || 150;
            if (w > maxWidth || h > maxHeight) {
                const ratio = Math.min(maxWidth / w, maxHeight / h);
                w = Math.round(w * ratio);
                h = Math.round(h * ratio);
            }
            canvas.width = w;
            canvas.height = h;
            ctx.drawImage(img, 0, 0, w, h);
            return canvas.toDataURL('image/jpeg', 0.7);
        } catch (e) {
            ChatLog.warn('HISTORY', '[缩略图] 生成失败:', e.message);
            return null;
        }
    }

    // --- 来自 app.js:L4649-L4720 ---
    async function addToHistory(title, messages) {
        // 在压缩前，从用户消息中提取图片并生成小缩略图
        let thumbnailSrc = null;
        for (const msg of messages) {
            if (msg.role === 'user') {
                // 优先从 attachments 中提取图片
                if (msg.attachments && msg.attachments.length > 0) {
                    for (const att of msg.attachments) {
                        if (att.type === 'image' && att.base64) {
                            const b64 = att.base64.startsWith('data:') ? att.base64 : 'data:' + (att.mime || 'image/png') + ';base64,' + att.base64;
                            const result = resizeBase64Image(b64, 200, 150);
                            thumbnailSrc = result instanceof Promise ? await result : result;
                            break;
                        }
                    }
                    if (thumbnailSrc) break;
                }
                // 兼容旧版：从 content 中的 <img> 标签提取
                if (!thumbnailSrc && msg.content) {
                    const imgMatch = msg.content.match(/<img\s+src="(data:image\/[^"]+)"/i);
                    if (imgMatch && imgMatch[1]) {
                        const result = resizeBase64Image(imgMatch[1], 200, 150);
                        thumbnailSrc = result instanceof Promise ? await result : result;
                        break;
                    }
                }
            }
        }

        // 压缩消息数据，避免localStorage超限
        const compressedMessages = messages.map(msg => {
            // 归一化角色：assistant -> ai
            const normalizedRole = msg.role === 'assistant' ? 'ai' : msg.role;
            
            // 如果是AI消息且包含超长内容，进行截断
            if (normalizedRole === 'ai' && msg.content) {
                if (msg.content.length > 50000) {
                    ChatLog.warn('HISTORY', `⚠️ 消息过大(${msg.content.length}字节),进行截断`);
                    return {
                        ...msg,
                        role: normalizedRole,
                        content: msg.content.substring(0, 50000) + '\n\n...(内容过长已截断)',
                        truncated: true
                    };
                }
            }

            // 如果是用户消息，处理图片数据
            if (normalizedRole === 'user') {
                let compressedContent = msg.content || '';
                
                // 移除 content 中的 base64 图片标签
                if (compressedContent) {
                    compressedContent = compressedContent.replace(
                        /<img\s+src="data:image\/[^;]+;base64,[^"]+"/gi,
                        '<img src="[image]"'
                    );
                }
                
                // 处理 attachments 中的图片 base64（压缩存储）
                let compressedAttachments = msg.attachments;
                if (msg.attachments && msg.attachments.length > 0) {
                    compressedAttachments = msg.attachments.map(att => {
                        if (att.type === 'image' && att.base64) {
                            // 缩略图已存在 thumbnail 字段，这里只保留引用标识
                            return {
                                type: 'image',
                                fileName: att.fileName || att.name || 'image',
                                mime: att.mime || 'image/png',
                                size: att.size || 0,
                                base64: null, // base64 数据移到 thumbnail，这里清空以节省空间
                                _hasImage: true // 标记有图片，加载时从 thumbnail 恢复
                            };
                        }
                        // 非图片附件，清理 base64
                        const { base64, data, ...rest } = att;
                        return rest;
                    });
                }

                // 用户消息也做50KB截断保护
                if (compressedContent.length > 50000) {
                    ChatLog.warn('HISTORY', `⚠️ 用户消息过大(${compressedContent.length}字节),进行截断`);
                    compressedContent = compressedContent.substring(0, 50000) + '\n\n...(内容过长已截断)';
                }

                const result = {
                    ...msg,
                    role: normalizedRole,
                    content: compressedContent
                };
                if (compressedContent !== msg.content) {
                    result.imageRemoved = true;
                }
                if (compressedAttachments && compressedAttachments.length > 0) {
                    result.attachments = compressedAttachments;
                }
                return result;
            }

            return msg;
        });

        const item = {
            id: Date.now(),
            title: title,
            date: new Date().toLocaleString('zh-CN'),
            messages: compressedMessages,
            thumbnail: thumbnailSrc || null
        };

        // 检查总大小
        const itemSize = JSON.stringify(item).length;
        if (itemSize > 100000) { // 100KB
            ChatLog.warn('HISTORY', `⚠️ 历史记录项过大(${itemSize}字节),跳过保存`);
            return;
        }

        App.state.chatHistory.unshift(item);
        if (App.state.chatHistory.length > 30) App.state.chatHistory.pop(); // 减少到30条
        saveHistory();
        renderHistory();
    }
    App.addToHistory = addToHistory;

    // 历史多选状态
    var _historySelectMode = false;
    var _selectedHistoryIds = Object.create(null);

    function syncHistorySelectUI() {
        var selectBar = document.getElementById('history-select-bar');
        var deleteBtn = document.getElementById('history-delete-selected-btn');
        var selectBtn = document.getElementById('history-select-btn');
        var countEl = document.getElementById('history-selected-count');
        var allCb = document.getElementById('history-select-all');
        var n = Object.keys(_selectedHistoryIds).length;
        if (selectBar) selectBar.style.display = _historySelectMode ? 'flex' : 'none';
        if (deleteBtn) deleteBtn.style.display = _historySelectMode ? '' : 'none';
        if (selectBtn) {
            selectBtn.classList.toggle('active', !!_historySelectMode);
            selectBtn.title = _historySelectMode ? '退出多选' : '多选删除';
        }
        if (countEl) countEl.textContent = wt('history.selectedCount', { n: n });
        if (allCb && !_historySelectMode) allCb.checked = false;
    }

    function toggleHistorySelectMode(force) {
        if (typeof force === 'boolean') _historySelectMode = force;
        else _historySelectMode = !_historySelectMode;
        if (!_historySelectMode) _selectedHistoryIds = Object.create(null);
        syncHistorySelectUI();
        renderHistory();
    }

    function toggleSelectAllHistory(checked) {
        var list = document.getElementById('history-list');
        if (!list) return;
        var items = list.querySelectorAll('.history-item[data-history-id]');
        _selectedHistoryIds = Object.create(null);
        if (checked) {
            items.forEach(function (el) {
                var id = parseInt(el.dataset.historyId, 10);
                if (!isNaN(id)) _selectedHistoryIds[id] = true;
            });
        }
        syncHistorySelectUI();
        renderHistory();
    }

    async function togglePinHistory(id) {
        var item = App.state.chatHistory.find(function (h) { return h.id === id; });
        if (!item) return;
        item.pinned = !item.pinned;
        item.pinnedAt = item.pinned ? Date.now() : 0;
        var ok = await saveHistory();
        if (ok) {
            renderHistory();
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function'
                ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">' + (item.pinned ? wt('history.pinnedToast') : wt('history.unpinnedToast')) + '</span>', true)
                : App.showToast(item.pinned ? wt('history.pinnedToast') : wt('history.unpinnedToast'));
        }
    }

    async function deleteSelectedHistory() {
        var ids = Object.keys(_selectedHistoryIds).map(function (k) { return parseInt(k, 10); }).filter(function (n) { return !isNaN(n); });
        if (!ids.length) {
            App.showToast ? App.showToast('请先勾选要删除的记录', 'info') : null;
            return;
        }
        if (!confirm('确定删除选中的 ' + ids.length + ' 条历史记录？')) return;
        var idSet = {};
        ids.forEach(function (id) { idSet[id] = true; });
        var backup = App.state.chatHistory.slice();
        App.state.chatHistory = App.state.chatHistory.filter(function (h) { return !idSet[h.id]; });
        var ok = await saveHistory();
        if (ok) {
            _selectedHistoryIds = Object.create(null);
            _historySelectMode = false;
            syncHistorySelectUI();
            renderHistory();
            App.showToast ? App.showToast('已删除 ' + ids.length + ' 条', 'success') : null;
        } else {
            App.state.chatHistory = backup;
            App.showToast ? App.showToast('删除失败，请检查存储权限', 'error') : null;
        }
    }

    // --- 来自 app.js:L4717-L4783 ---
    function renderHistory() {
        const list = document.getElementById('history-list');
        if (!list) return;
        const searchBox = document.getElementById('search-box');
        const searchTerm = (searchBox && searchBox.value ? searchBox.value : '').toLowerCase();

        const filtered = App.state.chatHistory.filter(h =>
            (h.title || '').toLowerCase().includes(searchTerm)
        );

        syncHistorySelectUI();

        if (filtered.length === 0) {
            list.innerHTML = '<div style="text-align: center; color: #999; padding: 20px;">' +
                (typeof wt === 'function' ? wt('history.empty') : '暂无历史记录') + '</div>';
            return;
        }

        // 按日期分组（置顶单独成组，排在最前）
        const grouped = groupHistoryByDate(filtered);

        list.innerHTML = Object.entries(grouped).map(([dateLabel, items]) => `
            <div class="history-group">
                <div class="history-group-title">${dateLabel}</div>
                ${items.map(h => {
                    // 从消息中提取第一张图片作为缩略图
                    const thumbnail = extractImageThumbnail(h);
                    const pinned = !!h.pinned;
                    const selected = !!_selectedHistoryIds[h.id];
                    const checkHtml = _historySelectMode
                        ? `<label class="history-check" onclick="event.stopPropagation()"><input type="checkbox" class="history-check-input" data-id="${h.id}" ${selected ? 'checked' : ''}></label>`
                        : '';
                    return `
                    <div class="history-item${pinned ? ' is-pinned' : ''}${selected ? ' is-selected' : ''}" data-history-id="${h.id}">
                        ${checkHtml}
                        <div class="history-item-actions">
                            <button class="history-action-btn pin-btn" data-action="pin" data-id="${h.id}" title="${pinned ? wt('history.unpin') : wt('history.pin')}">${pinned ? '📌' : '📍'}</button>
                            <button class="history-action-btn export-btn" data-action="export" data-id="${h.id}" title="${wt('history.export')}">⬇️</button>
                            <button class="history-action-btn edit-btn" data-action="edit" data-id="${h.id}" title="${wt('history.edit')}">✏️</button>
                            <button class="history-action-btn delete-btn" data-action="delete" data-id="${h.id}" title="${wt('history.delete')}">🗑️</button>
                        </div>
                        ${thumbnail ? `<div class="history-thumbnail"><img src="${thumbnail}" alt="缩略图"></div>` : ''}
                        <div class="history-content">
                            <div class="history-title">${pinned ? '<span class="history-pin-badge">' + wt('history.pinnedBadge') + '</span> ' : ''}${escapeHtml(h.title)}</div>
                            <div class="history-date">${h.date}</div>
                        </div>
                    </div>
                `}).join('')}
            </div>
        `).join('');

        // 使用事件委托绑定事件（更可靠，避免内存泄漏）
        list.onclick = function(e) {
            const item = e.target.closest('.history-item');
            if (!item) return;

            const id = parseInt(item.dataset.historyId);
            if (isNaN(id)) return;

            // 多选：点复选框或点整行切换勾选
            if (_historySelectMode) {
                if (e.target.classList && e.target.classList.contains('history-check-input')) {
                    if (e.target.checked) _selectedHistoryIds[id] = true;
                    else delete _selectedHistoryIds[id];
                    syncHistorySelectUI();
                    item.classList.toggle('is-selected', !!_selectedHistoryIds[id]);
                    return;
                }
                if (!e.target.closest('.history-action-btn')) {
                    if (_selectedHistoryIds[id]) delete _selectedHistoryIds[id];
                    else _selectedHistoryIds[id] = true;
                    const cb = item.querySelector('.history-check-input');
                    if (cb) cb.checked = !!_selectedHistoryIds[id];
                    item.classList.toggle('is-selected', !!_selectedHistoryIds[id]);
                    syncHistorySelectUI();
                    return;
                }
            }

            // 检查是否点击了操作按钮
            const actionBtn = e.target.closest('.history-action-btn');
            if (actionBtn) {
                e.preventDefault();
                e.stopPropagation();
                const action = actionBtn.dataset.action;
                ChatLog.info('HISTORY', '点击操作按钮:', action, 'id:', id);
                if (action === 'edit') {
                    editHistoryTitle(id);
                } else if (action === 'delete') {
                    deleteHistoryItem(id);
                } else if (action === 'pin') {
                    togglePinHistory(id);
                } else if (action === 'export') {
                    if (typeof window.promptExportConversation === 'function') {
                        window.promptExportConversation({ historyId: id });
                    } else if (typeof window.exportHistoryConversation === 'function') {
                        window.exportHistoryConversation(id, { format: 'md' });
                    } else if (window.App && typeof window.App.showToast === 'function') {
                        window.App.showToast('导出功能暂不可用', 'warning');
                    }
                }
                return;
            }

            // 点击整个项加载历史（排除操作按钮区域）
            if (!e.target.closest('.history-item-actions')) {
                loadHistoryItem(id);
            }
        };
    }

    // --- 来自 app.js:L4786-L4813 ---
    // 从历史记录项中提取缩略图（优先使用预存的小缩略图）
    function extractImageThumbnail(historyItem) {
        // 优先使用保存时生成的缩略图
        if (historyItem && historyItem.thumbnail) {
            return historyItem.thumbnail;
        }
        // fallback：从消息内容中提取（可能已被压缩为 [image] 占位符）
        const messages = historyItem ? historyItem.messages : null;
        if (!messages || !Array.isArray(messages)) return null;

        for (const msg of messages) {
            if (msg.role === 'user' && msg.content) {
                // 查找 img 标签
                const imgMatch = msg.content.match(/<img\s+src="([^"]+)"/i);
                if (imgMatch && imgMatch[1]) {
                    const src = imgMatch[1];
                    // 跳过占位符和无效URL
                    if (src === '[image]' || src.startsWith('[')) {
                        continue;
                    }
                    // 只返回有效的base64图片或完整URL
                    if (src.startsWith('data:image/') || src.startsWith('http')) {
                        return src;
                    }
                }
            }
        }
        return null;
    }

    // --- 来自 app.js:L4816-L4842 ---
    // 按日期分组历史记录（置顶优先独立成组）
    function groupHistoryByDate(items) {
        const groups = {};
        const today = new Date().toDateString();
        const yesterday = new Date(Date.now() - 86400000).toDateString();
        const pinned = [];
        const rest = [];

        items.forEach(item => {
            if (item && item.pinned) pinned.push(item);
            else rest.push(item);
        });

        // 置顶：按 pinnedAt 新→旧
        pinned.sort(function (a, b) {
            return (b.pinnedAt || b.id || 0) - (a.pinnedAt || a.id || 0);
        });
        if (pinned.length) groups[wt('history.groupPinned')] = pinned;

        var dateLocale = 'zh-CN';
        try {
            var localeMeta = window.App && App.i18n && typeof App.i18n.localeMeta === 'function' && App.i18n.localeMeta();
            if (localeMeta && localeMeta.dateLocale) dateLocale = localeMeta.dateLocale;
        } catch (_) { /* ignore */ }
        rest.forEach(item => {
            const itemDate = new Date(item.id).toDateString();
            let label;

            if (itemDate === today) {
                label = wt('history.groupToday');
            } else if (itemDate === yesterday) {
                label = wt('history.groupYesterday');
            } else {
                label = new Date(item.id).toLocaleDateString(dateLocale, {
                    month: 'long',
                    day: 'numeric',
                    weekday: 'short'
                });
            }

            if (!groups[label]) groups[label] = [];
            groups[label].push(item);
        });

        return groups;
    }

    // --- 来自 app.js:L4845-L4870 ---
    // 删除单条历史记录
    async function deleteHistoryItem(id) {
        ChatLog.info('HISTORY', 'deleteHistoryItem 被调用, id:', id);
        const originalLength = App.state.chatHistory.length;
        const deletedItem = App.state.chatHistory.find(h => h.id === id);
        ChatLog.info('HISTORY', '找到要删除的项:', deletedItem);
        App.state.chatHistory = App.state.chatHistory.filter(h => h.id !== id);
        ChatLog.info('HISTORY', `删除前: ${originalLength} 条, 删除后: ${App.state.chatHistory.length} 条`);

        const saveResult = await saveHistory();
        ChatLog.info('HISTORY', 'saveHistory 结果:', saveResult);
        if (saveResult) {
            delete _selectedHistoryIds[id];
            renderHistory();
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">已删除</span>', true) : App.showToast('已删除');
        } else {
            // 保存失败，恢复数据
            if (deletedItem) {
                App.state.chatHistory.push(deletedItem);
            }
            if (window.ErrorHandler && typeof ErrorHandler.showError === 'function') {
                ErrorHandler.showError('删除失败，请检查浏览器存储权限', '历史记录', 'error');
            } else {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">删除失败，请检查浏览器存储权限</span>', true) : App.showToast('删除失败，请检查浏览器存储权限');
            }
        }
    }

    // --- 来自 app.js:L4868-4895 ---
    // 清空所有历史记录
    async function clearAllHistory() {
        ChatLog.info('HISTORY', 'clearAllHistory 被调用');
        if (App.state.chatHistory.length === 0) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">没有历史记录</span>', true) : App.showToast('没有历史记录');
            return;
        }
        const backupHistory = [...App.state.chatHistory];
        App.state.chatHistory = [];
        ChatLog.info('HISTORY', '已清空 chatHistory，准备保存');

        const saveResult = await saveHistory();
        ChatLog.info('HISTORY', 'saveHistory 结果:', saveResult);
        if (saveResult) {
            renderHistory();
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">已清空所有历史记录</span>', true) : App.showToast('已清空所有历史记录');
        } else {
            // 保存失败，恢复数据
            App.state.chatHistory = backupHistory;
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">清空失败，请检查浏览器存储权限</span>', true) : App.showToast('清空失败，请检查浏览器存储权限');
        }
    }

    // --- 来自 app.js:L4891-L4910 ---
    // 编辑历史记录标题
    async function editHistoryTitle(id) {
        ChatLog.info('HISTORY', 'editHistoryTitle 被调用, id:', id);
        const item = App.state.chatHistory.find(h => h.id === id);
        if (!item) return;

        const newTitle = prompt('编辑标题:', item.title);
        if (newTitle && newTitle.trim()) {
            item.title = newTitle.trim();
            await saveHistory();
            renderHistory();
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">标题已更新</span>', true) : App.showToast('标题已更新');
        }
    }

    // --- 来自 app.js:L4906-L4950 ---
    // 新建对话
    function newChat() {
        // 优先委托给 ChatV2 入口的 startNewChat（已含保存历史/清空会话/重置上下文等逻辑）
        try {
            if (App && App.ChatV2 && App.ChatV2.chatEntry && typeof App.ChatV2.chatEntry.startNewChat === 'function') {
                ChatLog.info?.('HISTORY', '[newChat] 委托 ChatV2 startNewChat');
                App.ChatV2.chatEntry.startNewChat();
                // 新建对话后回到对话状态（收起左侧子面板 + 激活聊天 tab）
                if (typeof window.collapseSidePanels === 'function') window.collapseSidePanels();
                if (typeof window.activateChatView === 'function') window.activateChatView();
                return;
            }
        } catch (e) {
            ChatLog.warn?.('HISTORY', '[newChat] ChatV2 委托失败，回退旧逻辑:', e);
        }

        // 【修复】新建对话前，先保存当前对话到历史记录
        try {
            const currentMsgs = App.getChatMessages();
            if (currentMsgs.length > 0) {
                const container = document.getElementById('chat-messages');
                const hasRealMessages = !container.querySelector('.welcome-message') &&
                                       container.querySelectorAll('.message-wrapper:not(.thinking-only)').length > 0;
                if (hasRealMessages) {
                    const lastUserMsg = currentMsgs.find(m => m.role === 'user');
                    const title = (lastUserMsg && lastUserMsg.content)
                        ? lastUserMsg.content.replace(/<[^>]+>/g, '').substring(0, 30)
                        : '对话';
                    addToHistory(title || '对话', currentMsgs);
                    ChatLog.info('HISTORY', '[历史保存] 📝 新建对话前已保存当前聊天');
                }
            }
        } catch (e) {
            ChatLog.warn('HISTORY', '[历史保存] 新建对话前保存失败:', e.message);
        }

        // 清空当前对话
        document.getElementById('chat-messages').innerHTML = '';

        // 【完全回到对话状态】收起左侧代码编辑器 / 任务规划面板
        if (typeof window.collapseSidePanels === 'function') window.collapseSidePanels();

        // 清空代码编辑器
        document.getElementById('code-editor').value = '';
        if (typeof App.updateLineNumbers === 'function') {
            App.updateLineNumbers();
        }

        // 重置当前对话模型
        App.state.currentChatModel = '';
        document.getElementById('chat-model-select').value = '';

        // ✨ 重置所有附件和请求状态（解决新建对话后上传没反应的问题）
        App.clearAttachment();
        App.state.activeMode = null;  // 重置发票模式

        // 清除待处理技能上下文
        App._pendingImageSkill = null;
        App._pendingImageSkillMessage = '';

        // 清空所有待处理的请求（避免状态残留）
        if (typeof App.state.pendingImageRequest !== 'undefined') App.state.pendingImageRequest = null;
        if (typeof App.state.pendingCardRequest !== 'undefined') App.state.pendingCardRequest = null;
        if (typeof App.state.pendingInvoiceRequest !== 'undefined') App.state.pendingInvoiceRequest = null;

        // 重置后端选区上下文（防止跨会话选区污染）
        try {
            if (App.excelBridge && App.excelBridge.ResetRangeContext) {
                App.excelBridge.ResetRangeContext();
            }
        } catch(e) { /* 忽略COM调用异常 */ }

        // 显示欢迎消息
        addWelcomeMessage();

        App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">已新建对话</span>', true) : App.showToast('已新建对话');
    }

    // --- 来自 app.js:L4955-L5044 ---
    /**
     * 按宿主返回欢迎页文案与功能点（Excel / Word / PowerPoint）
     */
    function getWelcomeCopy() {
        const kind = (window.SmartPilotHost && window.SmartPilotHost.kind) || 'excel';
        const hostLabel = (window.SmartPilotHost && window.SmartPilotHost.label) || 'Excel';
        const hostKey = kind === 'word' ? 'word' : (kind === 'powerpoint' ? 'ppt' : 'excel');
        const icons = {
            excel: ['📷', '🔗', '📊', '⚡'],
            word: ['📝', '📑', '🔍', '✔️', '⚡', '💾'],
            ppt: ['🗂️', '🎨', '📊', '🎤', '⚡', '💾']
        };
        const features = (icons[hostKey] || icons.excel).map(function (icon, i) {
            return {
                icon: icon,
                title: wt('feat.' + hostKey + '.' + i + '.t'),
                desc: wt('feat.' + hostKey + '.' + i + '.d')
            };
        });
        return {
            hostLabel: hostLabel,
            subtitle: wt('welcome.' + hostKey + '.subtitle'),
            desc: wt('welcome.' + hostKey + '.desc'),
            features: features
        };
    }

    function escapeWelcomeHtml(s) {
        return String(s == null ? '' : s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    /**
     * 刷新欢迎页「去配置模型」引导条（已配置则隐藏）
     */
    function refreshWelcomeSetupCta() {
        const cta = document.getElementById('welcome-setup-cta');
        if (!cta) return;
        let ready = true;
        try {
            if (typeof App !== 'undefined' && typeof App.isLlmConfigured === 'function') {
                ready = !!App.isLlmConfigured();
            }
        } catch (e) {
            ready = true;
        }
        cta.style.display = ready ? 'none' : 'flex';
    }

    /**
     * 显示欢迎面板(进度条消失后 / 新建对话 / 空聊天区)
     */
    function showWelcomePanel() {
        const chatMessages = document.getElementById('chat-messages');
        if (!chatMessages) {
            ChatLog.warn('WELCOME', '[WelcomePanel] chat-messages 元素不存在');
            return;
        }

        // 已有真实消息则不插欢迎页
        const hasRealMsg = Array.prototype.some.call(chatMessages.children, function(el) {
            return el && el.nodeType === 1
                && !el.classList.contains('welcome-panel')
                && !el.classList.contains('pomodoro-welcome')
                && !el.classList.contains('welcome-message');
        });
        if (hasRealMsg) {
            ChatLog.info('WELCOME', '[WelcomePanel] 已有对话消息，跳过欢迎面板');
            return;
        }

        // 去掉历史番茄钟空状态
        try {
            chatMessages.classList.remove('has-pomodoro');
            const pomo = chatMessages.querySelector('.pomodoro-welcome');
            if (pomo) pomo.parentNode.removeChild(pomo);
        } catch (e) { /* ignore */ }

        // 已存在则只刷新引导条与快捷提示
        const existing = document.querySelector('.welcome-panel');
        if (existing) {
            refreshWelcomeSetupCta();
            if (window.App && typeof window.App.renderQuickPrompts === 'function') {
                try { window.App.renderQuickPrompts(); } catch (e) { /* ignore */ }
            }
            return;
        }

        ChatLog.info('WELCOME', '[WelcomePanel] 开始创建欢迎面板...');

        const copy = getWelcomeCopy();
        const featuresHtml = (copy.features || []).map(function(f) {
            return '<div class="wf-item">'
                + '<span class="wf-icon">' + escapeWelcomeHtml(f.icon) + '</span>'
                + '<div class="wf-info"><strong>' + escapeWelcomeHtml(f.title) + '</strong>'
                + '<span>' + escapeWelcomeHtml(f.desc) + '</span></div></div>';
        }).join('');

        const welcomePanel = document.createElement('div');
        welcomePanel.className = 'welcome-panel';
        welcomePanel.innerHTML = ''
            + '<div class="welcome-header">'
            +   '<h2>' + escapeWelcomeHtml(wt('app.title')) + '</h2>'
            +   '<p class="welcome-subtitle">' + escapeWelcomeHtml(copy.subtitle) + '</p>'
            +   '<p class="welcome-desc">' + escapeWelcomeHtml(copy.desc) + '</p>'
            + '</div>'
            + '<div class="welcome-setup-cta" id="welcome-setup-cta" style="display:none">'
            +   '<div class="welcome-setup-text">' + escapeWelcomeHtml(wt('welcome.setup')) + '</div>'
            +   '<div class="welcome-setup-actions">'
            +     '<button type="button" class="welcome-setup-btn" id="welcome-ocr-btn">' + escapeWelcomeHtml(wt('welcome.openOcr')) + '</button>'
            +     '<button type="button" class="welcome-setup-btn welcome-setup-btn-secondary" id="welcome-setup-btn">' + escapeWelcomeHtml(wt('welcome.openSettings')) + '</button>'
            +   '</div>'
            + '</div>'
            + '<div class="welcome-features-grid">' + featuresHtml + '</div>'
            + '<div class="quick-prompts" id="welcome-quick-prompts">'
            +   '<div class="quick-prompts-title">' + escapeWelcomeHtml(wt('welcome.tryThese')) + '</div>'
            +   '<div class="prompts-grid" id="quick-prompts-grid"></div>'
            +   '<div class="welcome-help-link"><button type="button" class="welcome-help-btn" id="welcome-help-btn">' + escapeWelcomeHtml(wt('welcome.helpLink')) + '</button></div>'
            + '</div>';

        chatMessages.insertBefore(welcomePanel, chatMessages.firstChild);

        const ocrBtn = welcomePanel.querySelector('#welcome-ocr-btn');
        if (ocrBtn) {
            ocrBtn.addEventListener('click', function() {
                if (typeof window.openOcrPanel === 'function') window.openOcrPanel();
            });
        }
        const setupBtn = welcomePanel.querySelector('#welcome-setup-btn');
        if (setupBtn) {
            setupBtn.addEventListener('click', function() {
                if (typeof window.openAiSettings === 'function') window.openAiSettings();
                else if (typeof App !== 'undefined' && typeof App.openAiSettings === 'function') App.openAiSettings();
            });
        }

        refreshWelcomeSetupCta();

        var helpBtn = welcomePanel.querySelector('#welcome-help-btn');
        if (helpBtn) {
            helpBtn.addEventListener('click', function () {
                if (typeof window.showProductHelp === 'function') window.showProductHelp({ tab: 'start' });
            });
        }

        if (window.App && typeof window.App.renderQuickPrompts === 'function') {
            setTimeout(function() {
                try {
                    window.App.renderQuickPrompts();
                } catch (e) {
                    ChatLog.warn('WELCOME', '[WelcomePanel] 快捷提示词渲染失败:', e);
                }
            }, 100);
        }

        ChatLog.info('WELCOME', '✅ 欢迎面板已显示（宿主: ' + (copy.hostLabel || 'Excel') + '）');
    }

    window.refreshWelcomeSetupCta = refreshWelcomeSetupCta;

    /**
     * 发送快捷提示词（仅填入输入框，不自动发送，方便用户修改）
     */
    window.App.sendQuickPrompt = function(promptText) {
        const messageInput = document.getElementById('message-input');
        if (messageInput) {
            messageInput.value = promptText;
            messageInput.focus();
            // 不再自动发送，用户可在基础上修改后再手动发送
        }
    };

    // ==================================================================
    // 快捷技能绑定（Quick Skill Binding）
    // ------------------------------------------------------------------
    // 背景（实测 2026-10-02 23:37）：快捷入口原先只把 quickShortcutPrompt 填进输入框，
    // **技能名被丢掉**（旧代码注释："skillName 保留给未来扩展"）。用户按发送后这条消息
    // 走的是通用意图路由 —— 于是"为火车票创建一个解析器"（GenerateOCRParseScript 快捷
    // 入口的提示词）被 skill_SkillCreator 抢走执行，用户："我都直接用的快捷键调用的
    // skill 也能匹配到别的技能？"
    // 现在把技能身份显式带上：
    //   triggerQuickSkill → App._quickSkillBinding
    //   → chat-entry.sendMessage 写入 context._forcedSkillName
    //   → orchestrator【快捷入口直连钉死】该技能，跳过 regex 路由与 LLM 覆盖。
    // ==================================================================
    window.App._quickSkillBinding = null;

    /**
     * 记录"本条消息必须用这个技能"（用户显式选择优先于任何自动路由）
     * @param {object} skill - 技能对象（name/displayName）
     * @param {string} promptText - 已填入输入框的提示词
     */
    window.App.bindQuickSkill = function(skill, promptText) {
        if (!skill) return;
        var name = skill.name || skill.Name || '';
        if (!name) return;
        App._quickSkillBinding = {
            name: name,
            displayName: skill.displayName || skill.DisplayName || name,
            prompt: promptText || '',
            ts: Date.now()
        };
        // 注意：**不要**在这里写 window.__TRIGGER_QUICK_SKILL_CTX。
        // app-file-batch.js 是它的唯一读取方，而那段代码调用的 App.executeImageProcessingSkill /
        // App.executeCustomSkill 在前端**根本没有定义**（历史遗留），命中后只会 `return` 掉，
        // 导致"选了文件却连预览/状态都不更新"。图片+技能场景由本处的 _quickSkillBinding 同样能钉死。
        App.renderQuickSkillChip();
        ChatLog.info('QUICK', `[bindQuickSkill] 已绑定技能 ${name}（本条消息不再走自动路由）`);
    };

    /** 解除绑定（用户点 × 取消，或消息已发出后自动清理） */
    window.App.clearQuickSkillBinding = function() {
        App._quickSkillBinding = null;
        window.__TRIGGER_QUICK_SKILL_CTX = null;
        App.renderQuickSkillChip();
    };

    /**
     * 渲染"本条消息将使用技能：XXX ×"提示条（挂在输入区顶部）
     * 目的：绑定必须对用户可见、可取消 —— 否则用户无从知道这次发送用的是哪个技能。
     */
    window.App.renderQuickSkillChip = function() {
        var inputArea = document.getElementById('input-area');
        if (!inputArea) return;
        var chip = document.getElementById('quick-skill-chip');
        var binding = App._quickSkillBinding;
        // 超过 30 分钟的绑定视为过期（用户早切去干别的了，不该继续锁死技能）
        if (binding && (Date.now() - binding.ts) > 30 * 60 * 1000) {
            App._quickSkillBinding = null;
            binding = null;
        }
        if (!binding) {
            if (chip) chip.remove();
            return;
        }
        if (!chip) {
            chip = document.createElement('div');
            chip.id = 'quick-skill-chip';
            chip.style.cssText = 'display:flex;align-items:center;gap:6px;margin:2px 6px 4px 6px;' +
                'padding:3px 8px;border-radius:10px;background:rgba(0,120,212,0.10);' +
                'border:1px solid rgba(0,120,212,0.35);color:#0a6cbf;font-size:11px;line-height:16px;';
            inputArea.insertBefore(chip, inputArea.firstChild);
        }
        chip.innerHTML = '';
        var label = document.createElement('span');
        label.textContent = '⚡ 本条消息将使用技能：' + (binding.displayName || binding.name);
        label.title = binding.name;
        var cancel = document.createElement('span');
        cancel.textContent = '×';
        cancel.title = '取消技能绑定，改回自动识别';
        cancel.style.cssText = 'cursor:pointer;font-weight:bold;font-size:13px;padding:0 2px;';
        cancel.onclick = function() { App.clearQuickSkillBinding(); };
        chip.appendChild(label);
        chip.appendChild(cancel);
    };

    /**
     * 触发快捷 Skill（保持"仅填入输入框、不自动发送"的交互，方便用户改词/传附件后再发送）
     * @param {string} skillName - Skill 名称（如 GenerateOCRParseScript）
     * @param {string} promptText - 提示词文本
     * @param {object} [skillObj] - 技能对象；缺省时按名字从 App.cachedSkills 反查
     */
    window.App.triggerQuickSkill = function(skillName, promptText, skillObj) {
        // 【关键修复】绑定技能身份：只填提示词会让这条消息走通用路由，被别的技能匹配走。
        var sk = skillObj;
        if (!sk && Array.isArray(App.cachedSkills)) {
            sk = App.cachedSkills.find(function(s) {
                return String(s.name || s.Name || '') === String(skillName || '');
            });
        }
        App.bindQuickSkill(sk || { name: skillName }, promptText);
        App.sendQuickPrompt(promptText || skillName || '');
    };

    /**
     * 渲染快捷提示词（统一配置模块 App.QuickPrompts，与侧边栏 ⚡ 弹层同源）
     * 快捷 Skill 由 renderQuickShortcutBar 单独渲染在模型选择器上方的图标栏
     */
    window.App.renderQuickPrompts = async function() {
        const grid = document.getElementById('quick-prompts-grid');
        if (!grid) return;

        const items = (window.App.QuickPrompts && window.App.QuickPrompts.getItems)
            ? window.App.QuickPrompts.getItems()
            : [];

        grid.innerHTML = '';
        if (!items.length) {
            const empty = document.createElement('div');
            empty.style.cssText = 'color:#999;font-size:12px;padding:6px;text-align:center;';
            empty.textContent = '暂无快捷提示词';
            grid.appendChild(empty);
            return;
        }
        items.forEach(function(p) {
            const item = document.createElement('div');
            item.className = 'prompt-item';
            item.title = p.prompt;
            item.onclick = function() {
                if (p.action === 'ocr' && typeof window.openOcrPanel === 'function') window.openOcrPanel();
                else App.sendQuickPrompt(p.prompt);
            };
            const icon = document.createElement('span');
            icon.className = 'prompt-icon';
            icon.textContent = p.icon || '⚡';
            const text = document.createElement('span');
            text.className = 'prompt-text';
            text.textContent = p.label || p.prompt;
            item.appendChild(icon);
            item.appendChild(text);
            grid.appendChild(item);
        });
    };

    /**
     * 渲染常驻快捷图标栏（模型选择器上方，发消息后不消失）
     * 从后端获取 IsQuickShortcut=true 的 Skill，渲染为紧凑图标按钮
     */
    window.App.renderQuickShortcutBar = async function() {
        const barInner = document.getElementById('quick-shortcut-bar-inner');
        if (!barInner) return;

        try {
            const bridge = (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge)
                || (App.state && App.state.excelBridge);
            if (bridge && typeof bridge.GetQuickShortcutSkills === 'function') {
                const result = await bridge.GetQuickShortcutSkills();
                const data = JSON.parse(result);

                if (data.success && data.skills && data.skills.length > 0) {
                    barInner.innerHTML = '';
                    data.skills.forEach(function(skill) {
                        const icon = skill.quickShortcutIcon || skill.displayName.charAt(0) || '⚡';
                        const promptText = skill.quickShortcutPrompt ||
                            (skill.examples && skill.examples.length > 0 ? skill.examples[0] : skill.displayName);
                        const item = document.createElement('span');
                        item.className = 'qs-item';
                        item.setAttribute('data-qs-name', skill.name);
                        item.title = promptText;
                        item.innerHTML = '<span class="qs-icon">' + icon + '</span><span class="qs-label">' + skill.displayName + '</span>';
                        item.onclick = function() { App.triggerQuickSkill(skill.name, promptText, skill); };
                        barInner.appendChild(item);
                    });
                    return;
                }
            }
        } catch (e) {
            ChatLog.warn('WELCOME', '⚠️ 快捷图标栏加载失败:', e.message);
        }

        // 空状态：没有快捷入口时显示提示
        barInner.innerHTML = '<span class="qs-empty-hint">在 Skill 管理中标记「⭐快捷入口」即可在此显示</span>';
    };

    // 绑定快捷图标栏鼠标滚轮横向滚动（只绑定一次）
    (function initQuickShortcutBarScroll() {
        var bar = document.getElementById('quick-shortcut-bar');
        if (!bar) return;
        // 防止重复绑定
        if (bar._qsWheelBound) return;
        bar._qsWheelBound = true;
        bar.addEventListener('wheel', function(e) {
            // 只有当内容溢出时才拦截滚轮
            if (bar.scrollWidth <= bar.clientWidth) return;
            e.preventDefault();
            bar.scrollLeft += e.deltaY || e.deltaX || e.wheelDelta * -1;
        }, { passive: false });
    })();

    // ==================== 稳定的循环轮播 ====================
    window.App = window.App || {};
    window.App._carouselState = { index: 0, timer: null, autoInterval: 3500 };

    window.App.initCarousel = function(panel) {
        var track = panel.querySelector('.carousel-track');
        var dotsContainer = panel.querySelector('.carousel-dots');
        if (!track || !dotsContainer) return;
        var cards = track.querySelectorAll('.feature-card');
        if (cards.length === 0) return;

        // 使用原生滚动方式
        track.style.overflowX = 'auto';
        track.style.scrollBehavior = 'smooth';

        // 生成圆点
        dotsContainer.innerHTML = '';
        for (var i = 0; i < cards.length; i++) {
            (function(idx) {
                var dot = document.createElement('span');
                dot.className = 'dot' + (idx === 0 ? ' active' : '');
                dot.onclick = function() { window.App.carouselGoTo(idx); };
                dotsContainer.appendChild(dot);
            })(i);
        }

        window.App._carouselState.index = 0;
        window.App._scrollToIndex(panel, 0);
        window.App._carouselAutoPlay(panel);
    };

    window.App._scrollToIndex = function(panel, idx) {
        var track = panel.querySelector('.carousel-track');
        var cards = track.querySelectorAll('.feature-card');
        var targetCard = cards[idx];
        if (targetCard && track) {
            var scrollLeft = targetCard.offsetLeft - track.offsetLeft - 16;
            track.scrollTo({ left: scrollLeft, behavior: 'smooth' });
        }
    };

    window.App._updateDots = function(panel) {
        var dots = panel.querySelectorAll('.carousel-dots .dot');
        var cards = panel.querySelectorAll('.carousel-track .feature-card');
        var currentIdx = window.App._carouselState.index % cards.length;
        dots.forEach(function(dot, i) {
            dot.className = 'dot' + (i === currentIdx ? ' active' : '');
        });
    };

    window.App.carouselNext = function() {
        var panel = document.querySelector('.welcome-panel');
        if (!panel) return;
        var cards = panel.querySelectorAll('.carousel-track .feature-card');
        window.App._carouselState.index = (window.App._carouselState.index + 1) % cards.length;
        window.App._scrollToIndex(panel, window.App._carouselState.index);
        window.App._updateDots(panel);
        window.App._carouselResetAutoPlay(panel);
    };

    window.App.carouselPrev = function() {
        var panel = document.querySelector('.welcome-panel');
        if (!panel) return;
        var cards = panel.querySelectorAll('.carousel-track .feature-card');
        window.App._carouselState.index = (window.App._carouselState.index - 1 + cards.length) % cards.length;
        window.App._scrollToIndex(panel, window.App._carouselState.index);
        window.App._updateDots(panel);
        window.App._carouselResetAutoPlay(panel);
    };

    window.App.carouselGoTo = function(idx) {
        var panel = document.querySelector('.welcome-panel');
        if (!panel) return;
        window.App._carouselState.index = idx;
        window.App._scrollToIndex(panel, idx);
        window.App._updateDots(panel);
        window.App._carouselResetAutoPlay(panel);
    };

    window.App._carouselAutoPlay = function(panel) {
        clearInterval(window.App._carouselState.timer);
        window.App._carouselState.timer = setInterval(function() {
            window.App.carouselNext();
        }, window.App._carouselState.autoInterval);
    };

    window.App._carouselResetAutoPlay = function(panel) {
        clearInterval(window.App._carouselState.timer);
        window.App._carouselAutoPlay(panel);
    };

    // --- 欢迎消息：产品欢迎面板（番茄钟不再作为空状态默认页，仍可通过任务规划使用） ---
    function addWelcomeMessage() {
        showWelcomePanel();
    }

    // 兜底：document 级别委托，无论 welcome DOM 来自静态 HTML 还是动态创建，按钮点击都能命中
    function bindPomodoroButtonDelegates() {
        try {
            if (window.App._pomodoroDelegatesBound) return;
            window.App._pomodoroDelegatesBound = true;
            document.addEventListener('click', function(e) {
                const t = e && e.target ? (e.target.closest ? e.target.closest('button') : e.target) : null;
                if (!t) return;
                if (t.classList && t.classList.contains('pomodoro-start')) { try { pomodoroStart(); } catch(_){ if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); } }
                else if (t.classList && t.classList.contains('pomodoro-pause')) { try { pomodoroPause(); } catch(_){ if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); } }
                else if (t.classList && t.classList.contains('pomodoro-reset')) { try { pomodoroReset(); } catch(_){ if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); } }
            }, true);
        } catch(_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '操作失败:', _ && _.message); }
    }

        // 番茄钟累计数据（配合 SharedPomodoro 同步）
    let pomodoroCounters = (function initPomodoroCounters() {
      let c = {};
      try { const s = localStorage.getItem('pomodoro_counters_v1'); if (s) c = JSON.parse(s) || {}; } catch(_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '加载番茄钟计数器失败:', _ && _.message); }
      const today = new Date().toISOString().slice(0, 10);
      if (c.todayKey && c.todayKey !== today) { c.todayPomodoros = 0; c.todayMinutes = 0; c.todayKey = today; }
      if (!c.todayKey) c.todayKey = today;
      c.completedPomodoros = c.completedPomodoros|0;
      c.completedMinutes = c.completedMinutes|0;
      c.todayPomodoros = c.todayPomodoros|0;
      c.currentCycle = c.currentCycle|0;
      try { localStorage.setItem('pomodoro_counters_v1', JSON.stringify(c)); } catch(_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '保存番茄钟计数器失败:', _ && _.message); }
      return c;
    })();

    // SharedPomodoro：内嵌 + 独立端 双向同步引擎（HTTP API 优先，localStorage 兜底）
    const SharedPomodoro = (function() {
      const WRITER_ID = 'excel-embed-' + Math.random().toString(36).slice(2, 10);
      const HTTP = 'http://localhost:9528';
      // 【守卫修正·B】VSTO(WebView2) 无 9528 本地服务，跳过 HTTP 同步，仅用 localStorage。
      const _VSTO = !!(window.chrome && window.chrome.webview && window.chrome.webview.hostObjects);
      const KEY = 'shared_pomodoro_v1';
      let knownVersion = -1;
      let pollTimer = null;
      let applyCb = null;
      let httpOk = false;
      let cfg = null;
      let cachedToken = '';

      // 获取 session token：优先 COM 桥（内嵌 WebView2），其次公开端点 /api/status
      function getToken() {
        return new Promise((resolve) => {
          if (cachedToken) return resolve(cachedToken);
          try {
            if (window.ExcelBridge && typeof window.ExcelBridge.GetMobileControlStatus === 'function') {
              const r = window.ExcelBridge.GetMobileControlStatus();
              if (r) { const j = JSON.parse(r); if (j && j.sessionToken) { cachedToken = j.sessionToken; return resolve(cachedToken); } }
            }
          } catch(_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '操作失败:', _ && _.message); }
          try {
            const c = new AbortController(); const to = setTimeout(() => c.abort(), 1500);
            fetch(HTTP + '/api/status', { cache: 'no-store', signal: c.signal })
              .then((res) => { clearTimeout(to); return res.json(); })
              .then((j) => { if (j && j.sessionToken) cachedToken = j.sessionToken; resolve(cachedToken); })
              .catch(() => { clearTimeout(to); resolve(''); });
          } catch(_) { resolve(''); }
        });
      }
      function apiUrl(path) { return HTTP + path + '?token=' + encodeURIComponent(cachedToken || ''); }

      function loadConfig() {
        if (cfg) return cfg;
        try { const s = localStorage.getItem('pomodoro_config_v1'); if (s) cfg = JSON.parse(s); } catch(_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); }
        if (!cfg) cfg = { workDur: 25, shortDur: 5, longDur: 15, longEvery: 4 };
        return cfg;
      }
      function saveConfig(c) {
        cfg = Object.assign({}, cfg, c || {});
        try { localStorage.setItem('pomodoro_config_v1', JSON.stringify(cfg)); } catch(_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); }
      }

      function toShared(state, counters) {
        const _cfg = loadConfig();
        const runtime = {
          isRunning: !!(state && state.isRunning),
          isPaused: !!(state && !state.isRunning && ((state.minutes|0) !== (_cfg.workDur|0) || (state.seconds|0) !== 0)),
          isBreak: !!(state && state.isBreak),
          minutes: (state && state.minutes)|0,
          seconds: (state && state.seconds)|0,
          remainingSeconds: ((state && state.minutes)|0) * 60 + ((state && state.seconds)|0),
          sessionCount: (state && state.sessionCount)|0,
          maxSessions: (state && state.maxSessions)|0,
          updatedAt: Date.now()
        };
        const counters2 = {
          completedPomodoros: (counters && counters.completedPomodoros)|0,
          completedMinutes:  (counters && counters.completedMinutes)|0,
          todayPomodoros:    (counters && counters.todayPomodoros)|0,
          currentCycle:      (counters && counters.currentCycle)|0,
          todayKey:          (counters && counters.todayKey) || new Date().toISOString().slice(0,10)
        };
        return { runtime: runtime, counters: counters2, config: _cfg };
      }

      function toEmbedded(shared) {
        const r = shared && shared.runtime;
        if (!r) return null;
        return {
          minutes: r.minutes|0, seconds: r.seconds|0,
          isBreak: !!r.isBreak,
          sessionCount: r.sessionCount|0,
          maxSessions: r.maxSessions|0
        };
      }

      function readLocal() {
        try {
          const s = localStorage.getItem(KEY);
          if (!s) return null;
          const j = JSON.parse(s);
          if (!j || typeof j !== 'object') return null;
          return j;
        } catch(_) { return null; }
      }

      function writeLocal(payload) {
        try {
          const cur = readLocal() || {};
          const next = Object.assign({}, cur, payload, { version: (cur.version|0) + 1, lastWriterId: payload.writerId || WRITER_ID, updatedAt: Date.now() });
          localStorage.setItem(KEY, JSON.stringify(next));
          knownVersion = next.version|0;
        } catch(_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '操作失败:', _ && _.message); }
      }

      async function httpGet() {
        if (_VSTO) return null;
        try {
          await getToken();
          const r = await fetch(apiUrl('/api/pomodoro/get'), { cache: 'no-store' });
          if (!r.ok) return null;
          const j = await r.json();
          return j;
        } catch(_) { return null; }
      }

      async function httpSet(payload) {
        if (_VSTO) return null;
        try {
          await getToken();
          const body = JSON.stringify(Object.assign({ writerId: WRITER_ID, fromVersion: knownVersion }, payload));
          const r = await fetch(apiUrl('/api/pomodoro/set'), {
            method: 'POST', headers: { 'Content-Type': 'application/json' }, body,
            cache: 'no-store'
          });
          if (!r.ok) return null;
          return await r.json();
        } catch(_) { return null; }
      }

      async function push(state, counters, evt) {
        const data = toShared(state, counters);
        writeLocal({ data: data, event: evt || '' });
        if (httpOk) {
          const resp = await httpSet({ data: data, event: evt || '' });
          if (resp && resp.version >= 0) {
            knownVersion = resp.version|0;
            if (resp.conflict && resp.data && typeof applyCb === 'function') {
              try { applyCb({ runtime: resp.data.runtime, counters: resp.data.counters, config: resp.data.config }); } catch(_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); }
            }
          }
        }
      }

      async function tickHttp() {
        const resp = await httpGet();
        if (!resp || resp.success === false) { httpOk = false; return; }
        httpOk = true;
        if ((resp.version|0) === knownVersion) return;
        knownVersion = resp.version|0;
        if (resp.lastWriterId === WRITER_ID) return;
        if (!resp.data) return;
        if (typeof applyCb === 'function') {
          try { applyCb({ runtime: resp.data.runtime, counters: resp.data.counters, config: resp.data.config }); } catch(_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); }
        }
        try {
          localStorage.setItem(KEY, JSON.stringify({
            version: knownVersion, lastWriterId: resp.lastWriterId,
            data: resp.data, updatedAt: Date.now()
          }));
        } catch(_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '操作失败:', _ && _.message); }
      }

      function tickLocal() {
        const cur = readLocal();
        if (!cur) return;
        const v = cur.version|0;
        if (v === knownVersion) return;
        knownVersion = v;
        if (cur.lastWriterId === WRITER_ID) return;
        if (!cur.data) return;
        if (typeof applyCb === 'function') {
          try { applyCb({ runtime: cur.data.runtime, counters: cur.data.counters, config: cur.data.config }); } catch(_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); }
        }
      }

      function start(applyFn) {
        applyCb = applyFn;
        (async function init() {
          loadConfig();
          let remote = null;
          if (!_VSTO) {
            try {
              await getToken();
              const ctrl = new AbortController();
              const to = setTimeout(() => ctrl.abort(), 1500);
              const r = await fetch(apiUrl('/api/pomodoro/get'), { cache: 'no-store', signal: ctrl.signal });
              clearTimeout(to);
              if (r.ok) { remote = await r.json(); httpOk = true; }
            } catch(_) { httpOk = false; }
          }
          if (remote && remote.success !== false && remote.data && Object.keys(remote.data).length > 0) {
            knownVersion = remote.version|0;
            if (typeof applyCb === 'function') {
              try { applyCb({ runtime: remote.data.runtime, counters: remote.data.counters, config: remote.data.config }); } catch(_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); }
            }
            try { localStorage.setItem(KEY, JSON.stringify({ version: knownVersion, lastWriterId: remote.lastWriterId, data: remote.data, updatedAt: Date.now() })); } catch(_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); }
          } else {
            const local = readLocal();
            if (local && local.data && Object.keys(local.data).length > 0) {
              knownVersion = local.version|0;
              if (typeof applyCb === 'function') {
                try { applyCb({ runtime: local.data.runtime, counters: local.data.counters, config: local.data.config }); } catch(_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); }
              }
              if (httpOk) { try { await httpSet({ data: local.data, event: 'sync_from_local' }); } catch(_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); } }
            }
          }
          pollTimer = setInterval(function() { if (httpOk) tickHttp(); else tickLocal(); }, 2000);
        })();
      }

      return { start: start, push: push, toShared: toShared, toEmbedded: toEmbedded, loadConfig: loadConfig, saveConfig: saveConfig, writerId: WRITER_ID };
    })();

    // 持久化 pomodoroCounters（写入内嵌端内存改动后立即存 localStorage）
    function persistPomodoroCounters() {
      try { localStorage.setItem('pomodoro_counters_v1', JSON.stringify(pomodoroCounters)); } catch(_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); }
    }
// 番茄钟状态
    let pomodoroState = {
        minutes: 25,
        seconds: 0,
        isRunning: false,
        isBreak: false,
        sessionCount: 0,
        maxSessions: 4,
        timerId: null
    };

    // 应用共享状态（来自独立端 / 本地）到内嵌端
    function pomodoroApplyShared(shared) {
        try {
            if (!shared) return;
            // counters 合并：取最大值（累计不回退）
            if (shared.counters) {
                pomodoroCounters.completedPomodoros = Math.max(pomodoroCounters.completedPomodoros|0, shared.counters.completedPomodoros|0);
                pomodoroCounters.completedMinutes  = Math.max(pomodoroCounters.completedMinutes|0,  shared.counters.completedMinutes|0);
                pomodoroCounters.todayKey          = shared.counters.todayKey || pomodoroCounters.todayKey || new Date().toISOString().slice(0,10);
                pomodoroCounters.todayPomodoros    = Math.max(pomodoroCounters.todayPomodoros|0,   shared.counters.todayPomodoros|0);
                pomodoroCounters.currentCycle      = Math.max(pomodoroCounters.currentCycle|0,     shared.counters.currentCycle|0);
            }
            // 若内嵌端不在运行，则把共享 runtime 覆盖过来；若正在运行，则仅当远端推进更快时覆盖
            if (shared.runtime) {
                const remRun = !!shared.runtime.isRunning && !shared.runtime.isPaused;
                const locRun = pomodoroState.isRunning;
                if (!locRun) {
                    const emb = SharedPomodoro.toEmbedded(shared);
                    if (emb) {
                        pomodoroState.minutes = emb.minutes;
                        pomodoroState.seconds = emb.seconds;
                        pomodoroState.isBreak = emb.isBreak;
                        pomodoroState.sessionCount = emb.sessionCount;
                        pomodoroState.maxSessions = emb.maxSessions;
                        // 如果共享端是 running 状态，这里本地启动 tick
                        if (remRun && !pomodoroState.isRunning) {
                            pomodoroState.isRunning = true;
                            if (pomodoroState.timerId) clearInterval(pomodoroState.timerId);
                            pomodoroState.timerId = setInterval(pomodoroTick, 1000);
                            // 按钮同步
                            try {
                                const sb = document.getElementById('pomodoro-start');
                                const pb = document.getElementById('pomodoro-pause');
                                if (sb) sb.disabled = true;
                                if (pb) pb.disabled = false;
                            } catch(_){ if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); }
                        }
                    }
                } else {
                    // 都在跑：选推进更快的 remaining
                    const rem = shared.runtime.remainingSeconds|0;
                    const loc = pomodoroState.minutes*60 + pomodoroState.seconds;
                    if (rem < loc) {
                        pomodoroState.minutes = Math.floor(rem/60);
                        pomodoroState.seconds = rem % 60;
                    }
                }
            }
            pomodoroUpdateDisplay();
            // post-remote-sync：远端改动后立刻存盘 + 同步按钮状态（避免按钮状态和显示滞后）
            try { persistPomodoroCounters(); } catch(_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); }
            try {
              const sb2 = document.getElementById('pomodoro-start');
              const pb2 = document.getElementById('pomodoro-pause');
              if (sb2) sb2.disabled = !!pomodoroState.isRunning;
              if (pb2) pb2.disabled = !pomodoroState.isRunning;
            } catch(_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '操作失败:', _ && _.message); }
        } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '操作失败:', e && e.message); }
    }
    // 启动共享同步（页面加载后尽快）
    try { SharedPomodoro.start(pomodoroApplyShared); } catch(e){ if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '操作失败:', e && e.message); }
    // static-welcome-init-sync：若 index.html 已先把番茄钟静态 HTML 渲染出来，JS 初始化完成后立刻同步显示与按钮 disabled
    (function staticWelcomeInitSync() {
      try {
        const sb = document.getElementById('pomodoro-start');
        const pb = document.getElementById('pomodoro-pause');
        if (sb) sb.disabled = !!pomodoroState.isRunning;
        if (pb) pb.disabled = !pomodoroState.isRunning;
        pomodoroUpdateDisplay();
      } catch(_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); }
    })();
    // 尽快绑定番茄钟按钮（document 委托），避免按钮提前出现却点不动
    // ========== 番茄钟：DOM 安全工具 + 按钮反馈兜底（杜绝"点不动"） ==========
    function $_pid(id) { try { return document.getElementById(id); } catch(_) { return null; } }
    function pomToast(text, kind) {
      try {
        if (window.App && typeof window.App.showToast === 'function') { window.App.showToast(text, kind || 'info'); return; }
      } catch(_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); }
      try {
        var d = document.createElement('div');
        d.textContent = text;
        var color = (kind === 'error') ? '#d93025' : (kind === 'warn') ? '#d48806' : (kind === 'ok') ? '#137333' : '#1a73e8';
        d.style.cssText = 'position:fixed;z-index:100000;left:50%;bottom:48px;transform:translateX(-50%);padding:8px 16px;border-radius:8px;background:' + color + ';color:#fff;font-size:13px;box-shadow:0 4px 16px rgba(0,0,0,0.2);pointer-events:none;opacity:0;transition:opacity 0.2s;';
        if (document.body) document.body.appendChild(d);
        requestAnimationFrame(function(){ d.style.opacity = '1'; });
        setTimeout(function(){ d.style.opacity='0'; setTimeout(function(){ try{ d.remove(); }catch(_) { if (typeof ChatLog !== 'undefined') ChatLog.debug('HISTORY', '[history] toast 移除失败:', _ && _.message); } }, 250); }, 1800);
      } catch(_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); }
    }
    function pomSetBtn(startDisabled, pauseDisabled, startText) {
      try {
        var s = $_pid('pomodoro-start'); if (s) { s.disabled = !!startDisabled; if (startText) s.textContent = startText; }
        var p = $_pid('pomodoro-pause'); if (p) p.disabled = !!pauseDisabled;
        var r = $_pid('pomodoro-reset'); if (r) r.disabled = false;
      } catch(_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); }
    }
    try { bindPomodoroButtonDelegates(); } catch(e){ if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '操作失败:', e && e.message); }

    // 番茄钟欢迎页
    function buildPomodoroTimer() {
        const wrap = document.createElement('div');
        wrap.className = 'pomodoro-welcome';

        const scan = document.createElement('div');
        scan.className = 'pomodoro-scanline';
        wrap.appendChild(scan);

        const icon = document.createElement('div');
        icon.className = 'pomodoro-icon';
        icon.textContent = '\ud83c\udf45';
        wrap.appendChild(icon);

        const timer = document.createElement('div');
        timer.className = 'pomodoro-timer';
        timer.id = 'pomodoro-timer';
        timer.textContent = '25:00';
        wrap.appendChild(timer);

        const status = document.createElement('div');
        status.className = 'pomodoro-status';
        status.id = 'pomodoro-status';
        status.textContent = '\u4e13\u6ce8\u4e2d';
        wrap.appendChild(status);

        const session = document.createElement('div');
        session.className = 'pomodoro-session';
        session.id = 'pomodoro-session';
        session.textContent = '\u7b2c 0 / 4 \u4e2a\u756a\u8304';
        wrap.appendChild(session);

        const controls = document.createElement('div');
        controls.className = 'pomodoro-controls';

        const startBtn = document.createElement('button');
        startBtn.className = 'pomodoro-btn pomodoro-start';
        startBtn.id = 'pomodoro-start';
        startBtn.textContent = '\u5f00\u59cb';
        startBtn.addEventListener('click', pomodoroStart);

        const pauseBtn = document.createElement('button');
        pauseBtn.className = 'pomodoro-btn pomodoro-pause';
        pauseBtn.id = 'pomodoro-pause';
        pauseBtn.textContent = '\u6682\u505c';
        pauseBtn.disabled = true;
        pauseBtn.addEventListener('click', pomodoroPause);

        const resetBtn = document.createElement('button');
        resetBtn.className = 'pomodoro-btn pomodoro-reset';
        resetBtn.id = 'pomodoro-reset';
        resetBtn.textContent = '\u91cd\u7f6e';
        resetBtn.addEventListener('click', pomodoroReset);

        controls.appendChild(startBtn);
        controls.appendChild(pauseBtn);
        controls.appendChild(resetBtn);
        wrap.appendChild(controls);

        // ---- 任务规划入口（打开侧边栏内嵌任务规划面板） ----
        const taskEntry = document.createElement('div');
        taskEntry.style.cssText = 'margin-top:24px; border-top:1px solid #e0e0e0; padding-top:16px; text-align:center;';

        const openBtn = document.createElement('button');
        openBtn.textContent = '\u279c \u6253\u5f00\u4efb\u52a1\u89c4\u5212';
        openBtn.style.cssText = 'padding:10px 24px; border:none; border-radius:8px; background:#d32f2f; color:#fff; cursor:pointer; font-size:14px; font-weight:600; transition:all 0.15s;';
        openBtn.addEventListener('mouseenter', function() { openBtn.style.background = '#b71c1c'; });
        openBtn.addEventListener('mouseleave', function() { openBtn.style.background = '#d32f2f'; });
        openBtn.addEventListener('click', function() {
            // 绑定到侧边栏内嵌任务规划面板（打开/切换）
            if (typeof window.toggleTaskPlanner === 'function') {
                window.toggleTaskPlanner();
            }
        });

        taskEntry.appendChild(openBtn);
        wrap.appendChild(taskEntry);

        // ---- 独立番茄钟入口（优先 HTTP/9528 双向同步，失败降级 file://） ----
        const standEntry = document.createElement('div');
        standEntry.style.cssText = 'margin-top:12px; text-align:center;';
        const standBtn = document.createElement('button');
        standBtn.textContent = '🍅 独立番茄钟（数据同步）';
        standBtn.style.cssText = 'padding:8px 18px; border:1px solid #d32f2f; border-radius:8px; background:#fff; color:#d32f2f; cursor:pointer; font-size:13px; font-weight:600; transition:all 0.15s;';
        standBtn.addEventListener('mouseenter', function() { standBtn.style.background = '#ffebee'; });
        standBtn.addEventListener('mouseleave', function() { standBtn.style.background = '#fff'; });
        standBtn.addEventListener('click', function() {
            if (typeof window.openStandalonePomodoro === 'function') {
                window.openStandalonePomodoro();
            } else if (typeof window.ExcelBridge !== 'undefined' && typeof window.ExcelBridge.GetUIPath === 'function') {
                try { window.open(window.ExcelBridge.GetUIPath('pomodoro.html'), '_blank'); } catch(_){ if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); }
            } else {
                try { window.open('http://localhost:9528/pomodoro', '_blank'); } catch(_){ if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); }
            }
        });
        standEntry.appendChild(standBtn);
        wrap.appendChild(standEntry);

        return wrap;
    }

    // 更新计时器显示
    function pomodoroUpdateDisplay() {
        const timerEl = document.getElementById('pomodoro-timer');
        const statusEl = document.getElementById('pomodoro-status');
        const sessionEl = document.getElementById('pomodoro-session');
        if (!timerEl) return;
        const m = String(pomodoroState.minutes).padStart(2, '0');
        const s = String(pomodoroState.seconds).padStart(2, '0');
        timerEl.textContent = m + ':' + s;
        if (statusEl) {
            statusEl.textContent = pomodoroState.isBreak ? '\u4f11\u606f\u4e2d' : '\u4e13\u6ce8\u4e2d';
        }
        if (sessionEl) {
            sessionEl.textContent = '\u7b2c ' + pomodoroState.sessionCount + ' / ' + pomodoroState.maxSessions + ' \u4e2a\u756a\u8304';
        }
    }

    // 开始计时
    function pomodoroStart() {
        try {
          if (pomodoroState.isRunning) { pomToast('番茄钟已在运行', 'warn'); return; }
          if (pomodoroState.minutes === 0 && pomodoroState.seconds === 0) {
              pomodoroReset();
              pomToast('已重置并重新开始专注 25 分钟', 'ok');
              // reset 后不继续：用户需要再次点开始（避免自动连续误解）
              return;
          }
          pomodoroState.isRunning = true;
          pomSetBtn(true, false);
          if (pomodoroState.timerId) { clearInterval(pomodoroState.timerId); pomodoroState.timerId = null; }
          pomodoroState.timerId = setInterval(pomodoroTick, 1000);
          pomodoroUpdateDisplay();
          try { SharedPomodoro.push(pomodoroState, pomodoroCounters, 'start'); } catch(_){ if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); }
          persistPomodoroCounters();
          pomToast('▶ 专注开始，加油！', 'ok');
        } catch(e) { pomToast('开始失败：' + (e && e.message ? e.message : e), 'error'); }
    }

    // 暂停计时
    function pomodoroPause() {
        try {
          if (!pomodoroState.isRunning) { pomToast('当前并未在运行', 'warn'); return; }
          pomodoroState.isRunning = false;
          if (pomodoroState.timerId) {
              clearInterval(pomodoroState.timerId);
              pomodoroState.timerId = null;
          }
          pomSetBtn(false, true, '继续');
          pomodoroUpdateDisplay();
          try { SharedPomodoro.push(pomodoroState, pomodoroCounters, 'pause'); } catch(_){ if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); }
          persistPomodoroCounters();
          pomToast('⏸ 已暂停', 'info');
        } catch(e) { pomToast('暂停失败：' + (e && e.message ? e.message : e), 'error'); }
    }

    // 重置计时
    function pomodoroReset() {
        try {
          pomodoroState.isRunning = false;
          if (pomodoroState.timerId) { clearInterval(pomodoroState.timerId); pomodoroState.timerId = null; }
          pomodoroState.minutes = 25;
          pomodoroState.seconds = 0;
          pomodoroState.isBreak = false;
          pomodoroUpdateDisplay();
          pomSetBtn(false, true, '开始');
          try { SharedPomodoro.push(pomodoroState, pomodoroCounters, 'reset'); } catch(_){ if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '番茄钟操作失败:', _ && _.message); }
          persistPomodoroCounters();
          pomToast('↺ 已重置为 25:00', 'ok');
        } catch(e) { pomToast('重置失败：' + (e && e.message ? e.message : e), 'error'); }
    }

    // 每秒 tick
    function pomodoroTick() {
        if (pomodoroState.seconds > 0) {
            pomodoroState.seconds--;
        } else if (pomodoroState.minutes > 0) {
            pomodoroState.minutes--;
            pomodoroState.seconds = 59;
        } else {
            pomodoroPause();
            if (!pomodoroState.isBreak) {
                pomodoroState.sessionCount++;
                pomodoroCounters.currentCycle = Math.max(pomodoroCounters.currentCycle|0, pomodoroState.sessionCount);
                pomodoroCounters.completedPomodoros++;
                pomodoroCounters.todayPomodoros++;
                const cfgDur = (SharedPomodoro && SharedPomodoro.loadConfig) ? (SharedPomodoro.loadConfig().workDur|0 || 25) : 25;
                pomodoroCounters.completedMinutes += cfgDur;
                pomodoroState.isBreak = true;
                if (pomodoroState.sessionCount >= pomodoroState.maxSessions) {
                    pomodoroState.minutes = 15;
                    pomodoroState.sessionCount = 0;
                    pomodoroCounters.currentCycle = 0;
                } else {
                    pomodoroState.minutes = 5;
                }
                pomodoroState.seconds = 0;
                pomodoroUpdateDisplay();
                pomodoroCounters.todayKey = pomodoroCounters.todayKey || new Date().toISOString().slice(0,10);
                persistPomodoroCounters();
                try { SharedPomodoro.push(pomodoroState, pomodoroCounters, 'complete_focus'); } catch(e){ if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '操作失败:', e && e.message); }
                pomodoroStart();
            } else {
                pomodoroState.isBreak = false;
                pomodoroState.minutes = 25;
                pomodoroState.seconds = 0;
                pomodoroUpdateDisplay();
                persistPomodoroCounters();
                try { SharedPomodoro.push(pomodoroState, pomodoroCounters, 'complete_break'); } catch(e){ if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '操作失败:', e && e.message); }
                pomodoroStart();
            }
        }
        // 每 10 秒同步一次（减少网络压力）
        const remaining = pomodoroState.minutes * 60 + pomodoroState.seconds;
        if (remaining % 10 === 0) { try { SharedPomodoro.push(pomodoroState, pomodoroCounters, 'tick'); } catch(e){ if (typeof ChatLog !== 'undefined') ChatLog.warn('HISTORY', '操作失败:', e && e.message); } persistPomodoroCounters(); }
        pomodoroUpdateDisplay();
    }

    window.pomodoroReset = pomodoroReset;

    // 像素字体（已弃用，由番茄钟替代）
    const PIXEL_LETTERS = {};

    // --- 来自 app.js:L5083-L5106 ---
    function loadHistoryItem(id) {
        const item = App.state.chatHistory.find(h => h.id === id);
        if (!item) return;

        // 检查 chat-v2 是否可用
        const sessionStore = App.ChatV2?.sessionStore;
        const renderer = App.ChatV2?.renderer;
        const isV2Mode = !!(sessionStore && renderer);

        if (isV2Mode) {
            // chat-v2 渲染路径：通过 session-store 创建 turn 触发渲染
            loadHistoryItemV2(item, sessionStore);
        } else {
            // 旧版渲染路径：使用 App.addMessage 兼容存根
            document.getElementById('chat-messages').innerHTML = '';
            item.messages.forEach(m => {
                let content;
                if (m.role === 'user') {
                    const hasImages = (m.attachments || []).some(a => a.type === 'image');
                    // 优先使用 originalContent（不含 OCR 文本）
                    const original = m.originalContent || '';
                    if (original && original.trim()) {
                        content = original.trim();
                    } else if (hasImages) {
                        // 用户只发图片没打字，不显示系统生成的文字
                        content = '';
                    } else {
                        // 纯文本消息
                        content = m.content || '';
                    }
                    // 如果有 attachments 中的图片，添加缩略图
                    if (hasImages && item.thumbnail) {
                        const imgHtml = `<img src="${item.thumbnail}" alt="历史图片" style="max-width:200px;max-height:160px;border-radius:6px;margin-top:4px;">`;
                        content = (content ? content + '\n' : '') + imgHtml;
                    }
                } else {
                    content = restoreImagePlaceholders(m.content, item.thumbnail);
                }
                const hasCodeBlock = /```[a-zA-Z0-9+_ -]*\s*\n/.test(content);
                const hasVBACode = (m.role === 'ai' || m.role === 'assistant') && App.containsVBACode(content);
                const parseCode = (m.role === 'ai' || m.role === 'assistant') && (hasCodeBlock || hasVBACode);
                App.addMessage(m.role === 'assistant' ? 'ai' : m.role, content, parseCode, m.timestamp || item.id);
            });
        }

        // 切换到聊天面板
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
        document.querySelector('[data-tab="chat"]').classList.add('active');
        document.getElementById('chat-panel').classList.add('active');

        var inputArea = document.getElementById('input-area');
        if (inputArea) inputArea.style.display = '';
        var modelSelector = document.getElementById('model-selector-area');
        if (modelSelector) modelSelector.style.display = '';
        var quickShortcutBar = document.getElementById('quick-shortcut-bar');
        if (quickShortcutBar) quickShortcutBar.style.display = '';
    }

    /**
     * 使用 chat-v2 渲染器加载历史记录。
     * 将历史消息按 user-assistant 配对分组，逐一创建 turn 并渲染。
     * 注意事项：
     *   - 加载期间抑制自动保存（防止重复写入旧历史系统）
     *   - 标记当前会话为"从历史加载"，避免 startNewChat 重复保存
     *   - 旧格式不存储 toolCall 数据，多步骤执行内容以文本形式保留在 content 中
     */
    function restoreImagePlaceholders(content, thumbnail) {
            if (!content || !content.includes('[image]')) return content;
            if (thumbnail) {
                return content.replace(/src="\[image\]"/g, 'src="' + thumbnail + '"');
            }
            return content.replace(/<img\s+src="\[image\]"[^>]*>/gi, 
                '<div style="display:inline-flex;align-items:center;gap:4px;padding:4px 10px;background:#f5f5f5;border:1px dashed #ccc;border-radius:6px;font-size:12px;color:#999;margin:4px 0;">🖼️ [图片已压缩保存，无法预览]</div>'
            );
        }

        function buildUserAttachments(msg, item) {
            const atts = (msg.attachments || []).map(att => {
                if (att.type === 'image' && att._hasImage) {
                    // 从 thumbnail 恢复图片数据
                    const base64Match = (item.thumbnail || '').match(/^data:(image\/\w+);base64,(.+)$/);
                    return {
                        type: 'image',
                        mime: att.mime || (base64Match ? base64Match[1] : 'image/png'),
                        base64: base64Match ? base64Match[2] : '',
                        fileName: att.fileName || '历史图片',
                        size: att.size || 0
                    };
                }
                // 非图片附件直接传递
                return att;
            });
            return atts;
        }

        function extractImageFromContent(content, thumbnail) {
            if (!content || !content.includes('[image]') || !thumbnail) {
                return { cleanContent: content, attachments: [] };
            }
            const cleanContent = content.replace(/<img\s+src="\[image\]"[^>]*>/gi, '').trim();
            const base64Match = thumbnail.match(/^data:(image\/\w+);base64,(.+)$/);
            if (base64Match) {
                return {
                    cleanContent: cleanContent,
                    attachments: [{
                        type: 'image',
                        mime: base64Match[1],
                        base64: base64Match[2],
                        fileName: '历史图片',
                        size: 0
                    }]
                };
            }
            return { cleanContent: content, attachments: [] };
        }

        function getUserDisplayMessage(msg, item) {
            // 检查是否有图片附件
            const hasImages = (msg.attachments || []).some(a => a.type === 'image');
            
            // 优先使用 originalContent（用户原始输入，不含 OCR 文本）
            const original = msg.originalContent || '';
            if (original && original.trim()) {
                return original.trim();
            }
            
            // 没有 originalContent：
            // - 有图片附件 → 用户可能只发了图片没输入文字 → 不显示系统拼接的文字
            // - 无图片附件 → 纯文本消息，直接显示 content
            if (hasImages) {
                // 用户只发图片没打字，不显示系统生成的 OCR 提示语
                return '';
            }
            
            // 纯文本消息（非 OCR 流程），直接显示
            return msg.content || '';
        }

        function loadHistoryItemV2(item, sessionStore) {
        const ChatLog = window.ChatLog || console;

        // 1. 抑制自动保存 + 标记为历史加载
        sessionStore.setSuppressSave(true);
        sessionStore.setLoadedFromHistory(true);

        // 2. 清空当前会话（DOM + store）
        sessionStore.clearHistory();
        const container = document.getElementById('chat-messages');
        if (container) container.innerHTML = '';

        // 3. 将消息按 user-assistant 配对分组
        const messages = item.messages || [];
        const pairs = []; // [{ userMessage, content, attachments }]
        let currentUser = null;
        for (const m of messages) {
            const content = m.content || '';
            if (m.role === 'user') {
                // 如果有上一个未配对的用户消息，作为独立用户消息保留
                if (currentUser) {
                    const atts = buildUserAttachments(currentUser, item);
                    const userMsg = getUserDisplayMessage(currentUser, item);
                    pairs.push({
                        userMessage: userMsg,
                        content: '',
                        attachments: atts
                    });
                }
                currentUser = m;
            } else if (m.role === 'assistant' || m.role === 'ai') {
                let aiContent = restoreImagePlaceholders(content, item.thumbnail);
                // 还原富内容：思考区 / 工具卡 / 步骤面板 的结构化数据
                const rich = {
                    reasoning: m.reasoning || '',
                    toolCalls: Array.isArray(m.toolCalls) ? m.toolCalls : [],
                    thoughtList: Array.isArray(m.thoughtList) ? m.thoughtList : [],
                    steps: Array.isArray(m.steps) ? m.steps : (Array.isArray(m._steps) ? m._steps : []),
                    planHtml: (typeof m.planHtml === 'string' ? m.planHtml : (m._planHtml || '')),
                    enableSteps: !!(m._enableSteps || m.enableSteps),
                    kind: m.kind,
                    usage: (m.usage && (Number(m.usage.in) || Number(m.usage.out))) ? m.usage : null
                };
                if (currentUser) {
                    const atts = buildUserAttachments(currentUser, item);
                    const userMsg = getUserDisplayMessage(currentUser, item);
                    pairs.push({
                        userMessage: userMsg,
                        content: aiContent,
                        attachments: atts,
                        rich
                    });
                    currentUser = null;
                } else {
                    // 孤立 AI 回复：没有对应的用户消息
                    pairs.push({
                        userMessage: '',
                        content: aiContent,
                        attachments: [],
                        rich
                    });
                }
            }
        }
        // 末尾未配对的用户消息
        if (currentUser) {
            const atts = buildUserAttachments(currentUser, item);
            const userMsg = getUserDisplayMessage(currentUser, item);
            pairs.push({
                userMessage: userMsg,
                content: '',
                attachments: atts
            });
        }

        // 4. 逐一创建 turn 并渲染（还原富内容：思考区/工具卡/步骤面板 走原生流式方法）
        for (const p of pairs) {
            try {
                const rich = p.rich || {};
                const turn = sessionStore.startTurn({
                    message: p.userMessage || ' ',
                    kind: rich.kind || 'text',
                    attachments: p.attachments || [],
                    context: { _enableSteps: !!rich.enableSteps }
                });
                // 还原思考区（reasoning 走原生 delta 通道渲染到思考区）
                if (rich.reasoning && typeof sessionStore.appendDelta === 'function') {
                    sessionStore.appendDelta(turn.id, 'reasoning', rich.reasoning);
                }
                // 还原 token 用量（汇总盒展示）
                if (rich.usage && typeof sessionStore.addTurnUsage === 'function') {
                    sessionStore.addTurnUsage(turn.id, rich.usage);
                }
                // 还原工具卡：先 addToolCall 建卡，再 updateToolCall 汇出结果（与原生顺序一致）
                if (rich.toolCalls && rich.toolCalls.length && typeof sessionStore.addToolCall === 'function') {
                    for (const tc of rich.toolCalls) {
                        if (!tc || !tc.name) continue;
                        const created = sessionStore.addToolCall(turn.id, {
                            id: tc.id, name: tc.name, args: tc.args || {}, status: tc.status || 'pending'
                        });
                        if (typeof sessionStore.updateToolCall === 'function') {
                            sessionStore.updateToolCall(turn.id, (created && created.id) || tc.id, {
                                status: tc.status === 'failed' ? 'failed' : 'completed',
                                result: tc.result,
                                error: tc.error
                            });
                        }
                    }
                }
                // 还原步骤面板数据（finalizeAIBubble 读取 turn._steps / _planHtml）
                turn._steps = rich.steps;
                turn._planHtml = rich.planHtml;
                // 还原 thoughtList（过程列表的思想/路由/工具行），让历史重载与实时终态完全一致
                if (Array.isArray(rich.thoughtList) && rich.thoughtList.length) {
                    turn.thoughtList = rich.thoughtList;
                }
                if (rich.enableSteps) turn._enableSteps = true;
                sessionStore.updateTurn(turn.id, { content: p.content });
                sessionStore.completeTurn(turn.id, {
                    content: p.content,
                    _steps: rich.steps,
                    _planHtml: rich.planHtml
                });
            } catch (e) {
                ChatLog.warn?.('HISTORY', '[loadHistoryItemV2] 渲染 turn 失败', e);
            }
        }

        // 5. 恢复自动保存
        sessionStore.setSuppressSave(false);

        ChatLog.info?.('HISTORY', `[loadHistoryItemV2] 已加载 ${pairs.length} 轮对话`);
    }

    // ==================== 模块导出 ====================
    window.AppHistory = {
        updateContextStats,
        isLocalStorageAvailable,
        initStorageMode,
        showStorageIndicator,
        hideStorageIndicator,
        loadHistory,
        saveHistory,
        addToHistory,
        renderHistory,
        extractImageThumbnail,
        groupHistoryByDate,
        deleteHistoryItem,
        clearAllHistory,
        editHistoryTitle,
        loadHistoryItem,
        newChat,
        showWelcomePanel,
        addWelcomeMessage,
        clearAllCache,
        closeCacheClearDialog,
        executeCacheClear,
        togglePinHistory,
        toggleHistorySelectMode,
        toggleSelectAllHistory,
        deleteSelectedHistory
    };

    // 兼容 index.html 中 onclick 直接调用
    window.clearAllCache = clearAllCache;
    window.newChat = newChat;
    window.showWelcomePanel = showWelcomePanel;
    window.addWelcomeMessage = addWelcomeMessage;
    window.togglePinHistory = togglePinHistory;
    window.toggleHistorySelectMode = toggleHistorySelectMode;
    window.toggleSelectAllHistory = toggleSelectAllHistory;
    window.deleteSelectedHistory = deleteSelectedHistory;

})();