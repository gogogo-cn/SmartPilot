/**
 * ============================================================
 * app-chat-compat.js — 与旧 API 的兼容性存根
 * ============================================================
 * 背景：chat-v2 重写删除了 5 个 god 文件（含 App.addMessage / App.renderMessageContent 等），
 *       但这些函数被其他保留模块（app-editor.js / app-recognition.js / app-skill-mgmt.js
 *       / app-image-table.js 等）调用。本文件提供最小兼容实现，让调用方不崩溃。
 *
 * 设计原则：
 *   - 只做「不崩溃」的最小实现（no-op 或简化版）
 *   - 不反向耦合 chat-v2 核心逻辑
 *   - 每个 stub 都在 console 打一次 warn，便于识别
 *
 * 放在 chat-v2 之后加载（在 index.html 中 chat-v2/index.js 之后）
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};

    const ChatLog = window.ChatLog || console;
    let warned = new Set();

    function warnOnce(name) {
        if (warned.has(name)) return;
        warned.add(name);
        ChatLog.warn?.('COMPAT', `[app-chat-compat] 调用了已删除的函数 App.${name}，使用空实现兜底`);
    }

    // ============ 核心 UI 函数 ============

    /**
     * App.addMessage(role, content, streaming, messageId, isHtml)
     * 旧版用于向 #chat-messages 追加用户/AI 气泡。
     * 新版已由 chat-renderer.js 完全接管（订阅 sessionStore）。
     * 这里做最小可用实现：创建基础气泡 + 支持简单 ID 更新
     */
    App.addMessage = function(role, content, streaming, messageId, isHtml) {
        warnOnce('addMessage');

        const container = document.getElementById('chat-messages');
        if (!container) return { id: 'compat_' + Date.now(), wrapper: null };

        const wrapper = document.createElement('div');
        wrapper.className = 'message-wrapper ' + (role === 'user' ? 'user' : 'ai');

        const main = document.createElement('div');
        main.className = 'message-main';

        const body = document.createElement('div');
        body.className = 'message-content';
        if (isHtml) {
            body.innerHTML = content || '';
        } else {
            body.textContent = content || '';
        }

        main.appendChild(body);
        wrapper.appendChild(main);

        const id = messageId || ('compat_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6));
        wrapper.id = id;
        container.appendChild(wrapper);
        container.scrollTop = container.scrollHeight;

        // 旧代码依赖返回的对象具有 { id, wrapper, contentEl/messageEl/update } 等
        return {
            id,
            wrapper,
            messageEl: wrapper,
            contentEl: body,
            update: function(newContent, newIsHtml) {
                if (newIsHtml) body.innerHTML = newContent;
                else body.textContent = newContent;
            },
            remove: function() { wrapper.remove(); }
        };
    };

    /**
     * App.renderMessageContent(text, skipMarkdown)
     * 旧版把 Markdown 文本渲染为 HTML。
     * 新版 chat-renderer 内部使用 marked.parse。此处提供外部调用入口。
     */
    App.renderMessageContent = function(text, skipMarkdown) {
        warnOnce('renderMessageContent');
        if (text == null) return '';
        if (skipMarkdown) {
            return String(text).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
        }
        try {
            if (window.marked && typeof marked.parse === 'function') {
                return marked.parse(String(text), { breaks: true, gfm: true });
            }
        } catch (e) { ChatLog.warn?.('UI', '[chat-compat] marked 渲染失败，回退纯文本:', e); }
        return '<p>' + String(text).replace(/\n/g, '<br>') + '</p>';
    };

    // ============ 编辑器相关 ============

    App.copyCode = function() {
        warnOnce('copyCode');
        try {
            const editor = document.getElementById('code-editor');
            if (!editor) return;
            const text = editor.value || '';
            navigator.clipboard?.writeText(text).then(
                () => App.showToast?.('已复制代码到剪贴板', 'info'),
                () => App.showToast?.('复制失败', 'error')
            );
        } catch (e) { ChatLog.warn?.('UI', '[chat-compat] 复制代码失败:', e); }
    };

    App.toEditor = function(code, language) {
        warnOnce('toEditor');
        const editor = document.getElementById('code-editor');
        if (!editor) return;
        editor.value = typeof code === 'string' ? code : (code?.code || '');
        const langSel = document.getElementById('editor-language');
        if (language && langSel) langSel.value = language;
        App.showToast?.('代码已载入编辑器', 'info');
    };

    App.previewHtml = function(html) {
        warnOnce('previewHtml');
        const w = window.open('', '_blank', 'width=800,height=600');
        if (w) {
            w.document.write(html || '');
            w.document.close();
        }
    };

    // ============ AI 调用兼容 ============

    /**
     * App.callAIUnified({ messages, stream, maxTokens, temperature, signal, model })
     * 旧版 AI 调用入口，当前仅被 app-model.js（连接测试）使用。
     * 新版由 App.ChatV2.llmClient 接管。
     * 这里提供最小兼容实现：非流式调用 complete()，返回 fetch-like Response 对象。
     */
    App.callAIUnified = async function({ messages = [], stream = false, maxTokens, temperature, signal, model } = {}) {
        warnOnce('callAIUnified');

        const llmClient = App.ChatV2?.llmClient;
        if (!llmClient) {
            const errText = JSON.stringify({ error: 'llmClient not loaded' });
            return createCompatResponse(errText, false, 503);
        }

        const cfg = { ...(App.config || {}) };
        if (temperature != null) cfg.temperature = temperature;
        if (maxTokens != null) cfg.maxTokens = maxTokens;
        if (model) cfg.model = model;

        try {
            const events = await llmClient.complete({ messages, config: cfg, signal });
            let content = '';
            let modelName = cfg.model;
            for (const evt of events) {
                if (evt.type === 'token') content += evt.delta || '';
                if (evt.type === 'model') modelName = evt.model;
            }

            const responseBody = {
                model: modelName,
                choices: [{
                    message: { role: 'assistant', content },
                    finish_reason: 'stop'
                }]
            };
            return createCompatResponse(JSON.stringify(responseBody), true, 200);
        } catch (err) {
            const errText = JSON.stringify({ error: err?.message || String(err) });
            return createCompatResponse(errText, false, 500);
        }
    };

    function createCompatResponse(bodyText, ok, status) {
        const encoder = new TextEncoder();
        const bytes = encoder.encode(bodyText);
        let consumed = false;
        return {
            ok,
            status,
            statusText: ok ? 'OK' : 'Error',
            json: async () => JSON.parse(bodyText),
            text: async () => bodyText,
            body: {
                getReader: () => ({
                    read: () => {
                        if (consumed) return Promise.resolve({ done: true, value: undefined });
                        consumed = true;
                        return Promise.resolve({ done: false, value: bytes });
                    },
                    cancel: () => Promise.resolve()
                })
            }
        };
    }

    // ============ Skill 相关 ============

    App.loadCustomSkills = function() {
        warnOnce('loadCustomSkills');
        // 新版通过 App.skillMgmt / App.ChatV2.toolRegistry 加载；这里触发一次刷新
        try {
            if (App.skillMgmt?.refresh) App.skillMgmt.refresh();
            if (App.ChatV2?.toolRegistry?.syncFromSkills) {
                App.ChatV2.toolRegistry.syncFromSkills().then(n => {
                    ChatLog.info?.('COMPAT', `[loadCustomSkills] synced ${n} skills to toolRegistry`);
                });
            }
        } catch (e) {
            ChatLog.warn?.('COMPAT', '[loadCustomSkills] failed', e);
        }
        return App.cachedSkills || [];
    };

    // ============ 输入框自适应 ============

    App.autoResizeTextarea = function(textarea) {
        warnOnce('autoResizeTextarea');
        const el = textarea || document.getElementById('message-input');
        if (!el) return;
        el.style.height = 'auto';
        el.style.height = Math.min(el.scrollHeight, 200) + 'px';
    };

    // ============ 历史记录兼容 ============

    /**
     * App.getChatMessages()
     * 旧版从 DOM 读取消息列表，新版从 session-store 读取已完成 turn 的消息。
     * 被 app-history.js(newChat)、app-file-batch.js、app-recognition.js 等调用。
     * @returns {Array<{role: string, content: string}>}
     */
    App.getChatMessages = function() {
        const sessionStore = App.ChatV2?.sessionStore;
        if (!sessionStore) {
            // 兜底：从 DOM 读取（兼容旧模式）
            const container = document.getElementById('chat-messages');
            if (!container) return [];
            const wrappers = container.querySelectorAll('.message-wrapper');
            const msgs = [];
            wrappers.forEach(w => {
                const role = w.classList.contains('user') ? 'user' : 'assistant';
                const contentEl = w.querySelector('.message-content');
                const content = contentEl ? contentEl.textContent || '' : '';
                if (content) msgs.push({ role, content });
            });
            return msgs;
        }
        // 从 session-store 读取已完成 turn 的消息
        const messages = sessionStore.getCompletedTurnsMessages();
        ChatLog.info?.('COMPAT', `[getChatMessages] 从 session-store 读取 ${messages.length} 条消息`);
        return messages;
    };

    // 也挂到 window 供老代码直接调用
    window.addMessage = App.addMessage;
    window.renderMessageContent = App.renderMessageContent;
    window.copyCode = App.copyCode;
    window.toEditor = App.toEditor;
    window.previewHtml = App.previewHtml;
    window.loadCustomSkills = App.loadCustomSkills;
    window.autoResizeTextarea = App.autoResizeTextarea;
    window.callAIUnified = App.callAIUnified;

    ChatLog.info?.('COMPAT', '✅ app-chat-compat: 8 个兼容存根已挂载');
})();
