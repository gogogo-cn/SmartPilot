/**
 * product-help.js — 应用内帮助 + 会话导出
 *
 * - showProductHelp({ tab })：起步三步 / 常见问题 / 快捷键
 * - exportCurrentConversation({ format })：导出当前会话 Markdown 或纯文本
 * - exportHistoryConversation(id, { format })：导出历史列表中的某条会话
 * - promptExportConversation({ historyId? })：弹出格式选择
 */
(function () {
    'use strict';

    var ChatLog = window.ChatLog || console;

    function hostInfo() {
        var h = window.SmartPilotHost || {};
        return {
            kind: h.kind || 'excel',
            label: h.label || 'Excel',
            isExcel: h.isExcel !== false && h.kind !== 'word' && h.kind !== 'powerpoint'
        };
    }

    function esc(s) {
        return String(s == null ? '' : s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function ht(key, vars) {
        try {
            if (window.App && App.i18n && typeof App.i18n.t === 'function') return App.i18n.t(key, vars);
        } catch (_) { /* ignore */ }
        return key;
    }

    function getFaq(host) {
        var common = [0, 1, 2, 3, 4, 5].map(function (i) {
            return { q: ht('faq.common.' + i + '.q'), a: ht('faq.common.' + i + '.a') };
        });
        if (host.kind === 'word') {
            return [
                { q: ht('faq.word.0.q'), a: ht('faq.word.0.a') },
                { q: ht('faq.word.1.q'), a: ht('faq.word.1.a') }
            ].concat(common);
        }
        if (host.kind === 'powerpoint') {
            return [
                { q: ht('faq.ppt.0.q'), a: ht('faq.ppt.0.a') },
                { q: ht('faq.ppt.1.q'), a: ht('faq.ppt.1.a') }
            ].concat(common);
        }
        return [
            { q: ht('faq.excel.0.q'), a: ht('faq.excel.0.a') },
            { q: ht('faq.excel.1.q'), a: ht('faq.excel.1.a') }
        ].concat(common);
    }

    function getStartSteps(host) {
        return [
            {
                n: '1',
                title: ht('help.start.1.t'),
                desc: ht('help.start.1.d'),
                actionLabel: ht('help.start.1.a'),
                action: 'ocr'
            },
            {
                n: '2',
                title: ht('help.start.2.t'),
                desc: ht('help.start.2.d', { host: host.label }),
                actionLabel: ht('help.start.2.a'),
                action: 'chat'
            },
            {
                n: '3',
                title: ht('help.start.3.t'),
                desc: ht('help.start.3.d'),
                actionLabel: ht('help.start.3.a'),
                action: 'export'
            }
        ];
    }

    function collectShortcutRows() {
        try {
            if (window.KeyboardShortcuts && window.KeyboardShortcuts.shortcuts) {
                return Array.from(window.KeyboardShortcuts.shortcuts.entries()).map(function (entry) {
                    var v = entry[1];
                    return { key: v.keyCombo || entry[0], description: v.description || '' };
                });
            }
        } catch (e) { /* ignore */ }
        return [
            { key: 'Ctrl+Enter', description: ht('help.sc.send') },
            { key: 'F1', description: ht('help.sc.help') },
            { key: 'Ctrl+/', description: ht('help.sc.list') },
            { key: 'Esc', description: ht('help.sc.esc') }
        ];
    }

    function doc() {
        return (typeof document !== 'undefined' && document)
            || (window && window.document)
            || null;
    }

    function closeHelp() {
        var d = doc();
        if (!d) return;
        var el = d.getElementById('product-help-overlay');
        if (el) {
            try { el.remove(); } catch (e) { /* ignore */ }
        }
        try {
            if (d._productHelpEsc) {
                d.removeEventListener('keydown', d._productHelpEsc, true);
                d._productHelpEsc = null;
            }
        } catch (e) { /* ignore */ }
    }

    function runHelpAction(action) {
        if (action === 'settings') {
            closeHelp();
            if (typeof window.openAiSettings === 'function') window.openAiSettings();
            return;
        }
        if (action === 'ocr') {
            closeHelp();
            if (typeof window.openOcrPanel === 'function') window.openOcrPanel();
            return;
        }
        if (action === 'chat') {
            closeHelp();
            if (typeof window.sideToolbarNav === 'function') window.sideToolbarNav('chat');
            var input = document.getElementById('message-input');
            if (input) input.focus();
            return;
        }
        if (action === 'export') {
            closeHelp();
            exportCurrentConversation({ format: 'md' });
            return;
        }
        if (action === 'diagnostic') {
            exportDiagnostic();
        }
    }

    /**
     * @param {{ tab?: 'start'|'faq'|'shortcuts'|'about' }} [opts]
     */
    function showProductHelp(opts) {
        opts = opts || {};
        var initial = opts.tab || 'start';
        if (initial !== 'start' && initial !== 'faq' && initial !== 'shortcuts' && initial !== 'about') initial = 'start';

        closeHelp();
        var d = doc();
        if (!d || !d.body) {
            ChatLog.warn?.('HELP', 'showProductHelp: document 不可用');
            return;
        }
        var host = hostInfo();
        var steps = getStartSteps(host);
        var faq = getFaq(host);
        var shortcuts = collectShortcutRows();

        var overlay = d.createElement('div');
        overlay.id = 'product-help-overlay';
        overlay.className = 'product-help-overlay';
        overlay.setAttribute('role', 'presentation');

        var stepsHtml = steps.map(function (s) {
            return '<div class="ph-step">' +
                '<div class="ph-step-n">' + esc(s.n) + '</div>' +
                '<div class="ph-step-body">' +
                '<div class="ph-step-title">' + esc(s.title) + '</div>' +
                '<div class="ph-step-desc">' + esc(s.desc) + '</div>' +
                '<button type="button" class="ph-action-btn" data-ph-action="' + esc(s.action) + '">' + esc(s.actionLabel) + '</button>' +
                '</div></div>';
        }).join('');

        var faqHtml = faq.map(function (item) {
            return '<details class="ph-faq-item"><summary>' + esc(item.q) + '</summary>' +
                '<div class="ph-faq-a">' + esc(item.a) + '</div></details>';
        }).join('');

        var scHtml = shortcuts.map(function (s) {
            return '<div class="ph-shortcut-row"><kbd>' + esc(s.key) + '</kbd><span>' + esc(s.description) + '</span></div>';
        }).join('');

        overlay.innerHTML =
            '<div class="product-help-backdrop" data-ph-close="1"></div>' +
            '<div class="product-help-panel" role="dialog" aria-modal="true" aria-labelledby="product-help-title">' +
            '  <div class="product-help-header">' +
            '    <div><div id="product-help-title" class="product-help-title">' + esc(ht('help.title')) + '</div>' +
            '    <div class="product-help-sub">' + esc(ht('help.host', { label: host.label })) + '</div></div>' +
            '    <button type="button" class="product-help-close" data-ph-close="1" title="' + esc(ht('help.close')) + '">×</button>' +
            '  </div>' +
            '  <div class="product-help-tabs">' +
            '    <button type="button" class="ph-tab' + (initial === 'start' ? ' active' : '') + '" data-ph-tab="start">' + esc(ht('help.tabStart')) + '</button>' +
            '    <button type="button" class="ph-tab' + (initial === 'faq' ? ' active' : '') + '" data-ph-tab="faq">' + esc(ht('help.tabFaq')) + '</button>' +
            '    <button type="button" class="ph-tab' + (initial === 'shortcuts' ? ' active' : '') + '" data-ph-tab="shortcuts">' + esc(ht('help.tabShortcuts')) + '</button>' +
            '    <button type="button" class="ph-tab' + (initial === 'about' ? ' active' : '') + '" data-ph-tab="about">' + esc(ht('help.tabAbout')) + '</button>' +
            '  </div>' +
            '  <div class="product-help-body">' +
            '    <div class="ph-pane' + (initial === 'start' ? ' active' : '') + '" data-ph-pane="start">' + stepsHtml + '</div>' +
            '    <div class="ph-pane' + (initial === 'faq' ? ' active' : '') + '" data-ph-pane="faq">' + faqHtml + '</div>' +
            '    <div class="ph-pane' + (initial === 'shortcuts' ? ' active' : '') + '" data-ph-pane="shortcuts">' +
            '      <div class="ph-shortcut-list">' + scHtml + '</div>' +
            '    </div>' +
            '    <div class="ph-pane' + (initial === 'about' ? ' active' : '') + '" data-ph-pane="about">' +
            '      <div class="ph-kicker">' + esc(ht('help.about.version', { version: appVersion() })) + '</div>' +
            '      <p class="ph-about-p">' + esc(ht('settings.data.cloud')) + '</p>' +
            '      <p class="ph-about-p">' + esc(ht('settings.data.local')) + '</p>' +
            '      <p class="ph-about-p">' + esc(ht('settings.data.tools')) + '</p>' +
            '      <p class="ph-about-p">' + esc(ht('settings.data.keys')) + '</p>' +
            '      <p class="ph-about-p">' + esc(ht('help.about.diagHint')) + '</p>' +
            '      <button type="button" class="ph-action-btn" data-ph-action="diagnostic">' + esc(ht('help.about.diag')) + '</button>' +
            '    </div>' +
            '  </div>' +
            '</div>';

        d.body.appendChild(overlay);

        function switchTab(name) {
            overlay.querySelectorAll('.ph-tab').forEach(function (t) {
                t.classList.toggle('active', t.getAttribute('data-ph-tab') === name);
            });
            overlay.querySelectorAll('.ph-pane').forEach(function (p) {
                p.classList.toggle('active', p.getAttribute('data-ph-pane') === name);
            });
        }

        overlay.addEventListener('click', function (e) {
            var t = e.target;
            if (!t) return;
            if (t.getAttribute('data-ph-close') === '1' || t.classList.contains('product-help-backdrop')) {
                closeHelp();
                return;
            }
            var tab = t.getAttribute('data-ph-tab') || (t.closest && t.closest('[data-ph-tab]') && t.closest('[data-ph-tab]').getAttribute('data-ph-tab'));
            if (tab) {
                switchTab(tab);
                return;
            }
            var actionBtn = t.closest ? t.closest('[data-ph-action]') : null;
            if (actionBtn) {
                runHelpAction(actionBtn.getAttribute('data-ph-action'));
            }
        });

        var onEsc = function (e) {
            if (e && e.key === 'Escape') {
                try { e.preventDefault(); } catch (_) { /* ignore */ }
                closeHelp();
            }
        };
        d._productHelpEsc = onEsc;
        d.addEventListener('keydown', onEsc, true);
    }

    function timestamp() {
        var d = new Date();
        var p = function (n) { return String(n).padStart(2, '0'); };
        return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '_' + p(d.getHours()) + p(d.getMinutes()) + p(d.getSeconds());
    }

    function downloadText(filename, content, mime) {
        var d = doc();
        var BlobCtor = (typeof Blob !== 'undefined' && Blob) || (window && window.Blob);
        var URLApi = (typeof URL !== 'undefined' && URL) || (window && window.URL);
        if (!d || !BlobCtor || !URLApi) {
            ChatLog.warn?.('HELP', 'downloadText: 环境不支持下载');
            return;
        }
        var blob = new BlobCtor([content], { type: mime || 'text/plain;charset=utf-8' });
        var url = URLApi.createObjectURL(blob);
        var a = d.createElement('a');
        a.href = url;
        a.download = filename;
        d.body.appendChild(a);
        a.click();
        setTimeout(function () {
            try { d.body.removeChild(a); } catch (_) { /* ignore */ }
            try { URLApi.revokeObjectURL(url); } catch (_) { /* ignore */ }
        }, 0);
    }

    function collectExportMessages() {
        var ss = window.App && window.App.ChatV2 && window.App.ChatV2.sessionStore;
        var messages = [];
        if (ss && typeof ss.getCompletedTurnsMessages === 'function') {
            messages = ss.getCompletedTurnsMessages() || [];
        }
        // 补充当前进行中的用户消息（若有）
        try {
            var turn = ss && typeof ss.getCurrentTurn === 'function' ? ss.getCurrentTurn() : null;
            if (turn && turn.userMessage) {
                var last = messages[messages.length - 1];
                if (!last || last.role !== 'user' || last.content !== turn.userMessage) {
                    messages = messages.concat([{ role: 'user', content: turn.userMessage }]);
                    if (turn.content) {
                        messages = messages.concat([{ role: 'assistant', content: turn.content }]);
                    }
                }
            }
        } catch (e) { /* ignore */ }
        return messages.filter(function (m) { return m && m.content && String(m.content).trim(); });
    }

    function buildMarkdown(messages, hostLabel) {
        var lines = [];
        lines.push('# SmartPilot 对话导出');
        lines.push('');
        lines.push('- 宿主：' + hostLabel);
        lines.push('- 导出时间：' + new Date().toLocaleString());
        lines.push('');
        lines.push('---');
        lines.push('');
        messages.forEach(function (m, i) {
            var role = (m.role === 'user') ? '用户' : '助手';
            lines.push('## ' + (i + 1) + '. ' + role);
            lines.push('');
            lines.push(String(m.content || '').trim());
            lines.push('');
        });
        return lines.join('\n');
    }

    function buildPlainText(messages, hostLabel) {
        var lines = [];
        lines.push('SmartPilot 对话导出');
        lines.push('宿主：' + hostLabel);
        lines.push('导出时间：' + new Date().toLocaleString());
        lines.push('');
        messages.forEach(function (m, i) {
            var role = (m.role === 'user') ? '用户' : '助手';
            lines.push('[' + (i + 1) + '] ' + role);
            lines.push(String(m.content || '').trim());
            lines.push('');
        });
        return lines.join('\n');
    }

    /**
     * @param {Array} messages
     * @param {{ format?: 'md'|'txt', title?: string, filePrefix?: string }} [opts]
     */
    function exportMessagesList(messages, opts) {
        opts = opts || {};
        var format = (opts.format === 'txt') ? 'txt' : 'md';
        var host = hostInfo();
        var list = (messages || []).filter(function (m) {
            return m && m.content && String(m.content).trim();
        });
        if (!list.length) {
            if (window.App && typeof window.App.showToast === 'function') {
                window.App.showToast('没有可导出的对话内容', 'info');
            }
            return false;
        }
        var body = format === 'txt'
            ? buildPlainText(list, host.label)
            : buildMarkdown(list, host.label);
        if (opts.title) {
            body = (format === 'txt'
                ? ('标题：' + opts.title + '\n')
                : ('> 标题：' + opts.title + '\n\n')) + body;
        }
        var prefix = opts.filePrefix || ('SmartPilot对话_' + host.label);
        var safePrefix = String(prefix).replace(/[\\/:*?"<>|]+/g, '_').slice(0, 60);
        var name = safePrefix + '_' + timestamp() + (format === 'txt' ? '.txt' : '.md');
        var mime = format === 'txt' ? 'text/plain;charset=utf-8' : 'text/markdown;charset=utf-8';
        downloadText(name, body, mime);
        if (window.App && typeof window.App.showToast === 'function') {
            window.App.showToast(ht('toast.exported', { n: list.length }), 'success');
        }
        ChatLog.info?.('HELP', '[product-help] 已导出会话: ' + name);
        return true;
    }

    /**
     * @param {{ format?: 'md'|'txt' }} [opts]
     */
    function exportCurrentConversation(opts) {
        return exportMessagesList(collectExportMessages(), opts || {});
    }

    /**
     * 导出历史记录中的某条会话
     * @param {number|string} historyId
     * @param {{ format?: 'md'|'txt' }} [opts]
     */
    function exportHistoryConversation(historyId, opts) {
        opts = opts || {};
        var id = typeof historyId === 'string' ? parseInt(historyId, 10) : historyId;
        var hist = (window.App && window.App.state && window.App.state.chatHistory) || [];
        var item = hist.find(function (h) { return h && h.id === id; });
        if (!item) {
            if (window.App && typeof window.App.showToast === 'function') {
                window.App.showToast('未找到该历史记录', 'warning');
            }
            return false;
        }
        var messages = Array.isArray(item.messages) ? item.messages : [];
        var title = item.title || ('会话_' + id);
        return exportMessagesList(messages, {
            format: opts.format || 'md',
            title: title,
            filePrefix: 'SmartPilot_' + title
        });
    }

    /** 弹出导出格式选择（轻量） */
    function promptExportConversation(opts) {
        opts = opts || {};
        var historyId = opts.historyId;
        var title = historyId != null ? ht('help.exportHistoryTitle') : ht('help.exportTitle');
        var d = doc();
        var doExport = function (fmt) {
            if (historyId != null) exportHistoryConversation(historyId, { format: fmt });
            else exportCurrentConversation({ format: fmt });
        };
        if (!d || !d.body) {
            doExport('md');
            return;
        }
        var overlay = d.createElement('div');
        overlay.className = 'product-help-overlay';
        overlay.innerHTML =
            '<div class="product-help-backdrop" data-ex-close="1"></div>' +
            '<div class="product-help-panel ph-export-panel" role="dialog" aria-modal="true">' +
            '  <div class="product-help-header"><div class="product-help-title">' + esc(title) + '</div>' +
            '  <button type="button" class="product-help-close" data-ex-close="1">×</button></div>' +
            '  <div class="product-help-body" style="padding:16px;">' +
            '    <p style="margin:0 0 12px;font-size:13px;color:#475569;">' + esc(ht('help.exportHint')) + '</p>' +
            '    <div style="display:flex;gap:8px;flex-wrap:wrap;">' +
            '      <button type="button" class="ph-action-btn" data-ex-fmt="md">' + esc(ht('help.exportMd')) + '</button>' +
            '      <button type="button" class="ph-action-btn ph-action-secondary" data-ex-fmt="txt">' + esc(ht('help.exportTxt')) + '</button>' +
            '    </div>' +
            '  </div>' +
            '</div>';
        d.body.appendChild(overlay);
        overlay.addEventListener('click', function (e) {
            var t = e.target;
            if (!t) return;
            if (t.getAttribute('data-ex-close') === '1' || t.classList.contains('product-help-backdrop')) {
                try { overlay.remove(); } catch (_) { /* ignore */ }
                return;
            }
            var fmt = t.getAttribute('data-ex-fmt');
            if (fmt) {
                try { overlay.remove(); } catch (_) { /* ignore */ }
                doExport(fmt);
            }
        });
    }

    var DISCLOSURE_KEY = 'smartpilot.dataDisclosure.v1';

    function appVersion() {
        try { return window.__APP_VERSION__ || 'unknown'; } catch (e) { return 'unknown'; }
    }

    function disclosureAccepted() {
        try { return localStorage.getItem(DISCLOSURE_KEY) === '1'; } catch (e) { return false; }
    }

    function rememberDisclosure() {
        try { localStorage.setItem(DISCLOSURE_KEY, '1'); } catch (e) { /* 隐私模式仍允许本次继续 */ }
    }

    function redactDiagnostic(text) {
        return String(text || '')
            .replace(/sk-[A-Za-z0-9_\-]{8,}/g, 'sk-***')
            .replace(/("apiKey"\s*:\s*")[^"]*"/gi, '$1***"');
    }

    function apiHostOnly() {
        var cfg = (window.App && window.App.config) || {};
        var raw = cfg.apiUrl || cfg.baseUrl || '';
        if (!raw) return '';
        try { return new URL(raw).host || ''; } catch (e) { return ''; }
    }

    function exportDiagnostic() {
        var cfg = (window.App && window.App.config) || {};
        var host = hostInfo();
        var lines = [
            'SmartPilot diagnostic',
            'version: ' + appVersion(),
            'host: ' + host.label + ' (' + host.kind + ')',
            'time: ' + (new Date()).toISOString(),
            'provider: ' + (cfg.apiType || cfg.provider || ''),
            'model: ' + (cfg.model || ''),
            'apiHost: ' + apiHostOnly(),
            'locale: ' + (cfg.uiLocale || ''),
            'disclosureAccepted: ' + (disclosureAccepted() ? 'yes' : 'no'),
            'userAgent: ' + ((typeof navigator !== 'undefined' && navigator.userAgent) || ''),
            '',
            'Keys and document text are omitted. Do not paste workbook contents into this file before sending it.'
        ];
        downloadText('smartpilot-diagnostic-' + timestamp() + '.txt', redactDiagnostic(lines.join('\n')), 'text/plain');
    }

    /**
     * 首次发送前确认数据会发往用户配置的模型地址。已确认则直接通过。
     * @returns {Promise<boolean>}
     */
    function ensureDataDisclosureAccepted() {
        if (disclosureAccepted()) return Promise.resolve(true);
        var d = doc();
        if (!d || !d.body) return Promise.resolve(true);
        return new Promise(function (resolve) {
            var overlay = d.createElement('div');
            overlay.id = 'data-disclosure-overlay';
            overlay.className = 'product-help-overlay';
            overlay.setAttribute('role', 'presentation');
            overlay.innerHTML =
                '<div class="product-help-backdrop"></div>' +
                '<div class="product-help-panel" role="dialog" aria-modal="true">' +
                '  <div class="product-help-header">' +
                '    <div><div class="product-help-title">' + esc(ht('help.disclosure.title')) + '</div></div>' +
                '  </div>' +
                '  <div class="product-help-body">' +
                '    <p class="ph-about-p">' + esc(ht('settings.data.cloud')) + '</p>' +
                '    <p class="ph-about-p">' + esc(ht('settings.data.local')) + '</p>' +
                '    <p class="ph-about-p">' + esc(ht('settings.data.tools')) + '</p>' +
                '    <div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:8px;">' +
                '      <button type="button" class="ph-action-btn" data-disclose="yes">' + esc(ht('help.disclosure.accept')) + '</button>' +
                '      <button type="button" class="ph-action-btn ph-action-secondary" data-disclose="no">' + esc(ht('help.disclosure.cancel')) + '</button>' +
                '    </div>' +
                '  </div>' +
                '</div>';
            function finish(ok) {
                try { overlay.remove(); } catch (e) { /* ignore */ }
                if (ok) rememberDisclosure();
                resolve(!!ok);
            }
            overlay.addEventListener('click', function (e) {
                var t = e.target;
                if (!t || !t.getAttribute) return;
                var choice = t.getAttribute('data-disclose');
                if (choice === 'yes') finish(true);
                else if (choice === 'no') finish(false);
            });
            d.body.appendChild(overlay);
        });
    }

    window.showProductHelp = showProductHelp;
    window.ensureDataDisclosureAccepted = ensureDataDisclosureAccepted;
    window.exportSmartPilotDiagnostic = exportDiagnostic;
    window.closeProductHelp = closeHelp;
    window.exportCurrentConversation = exportCurrentConversation;
    window.exportHistoryConversation = exportHistoryConversation;
    window.promptExportConversation = promptExportConversation;

    if (window.App) {
        window.App.showProductHelp = showProductHelp;
        window.App.exportCurrentConversation = exportCurrentConversation;
        window.App.exportHistoryConversation = exportHistoryConversation;
        window.App.promptExportConversation = promptExportConversation;
    }

    ChatLog.info?.('HELP', '✅ product-help 已加载');
})();
