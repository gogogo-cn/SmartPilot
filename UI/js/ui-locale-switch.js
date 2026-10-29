/**
 * ui-locale-switch.js — 界面语言统一切换入口
 * 语言列表来自 App.i18n.localeCatalog()，追加语言不用改本文件。
 *
 * 职责（把语言切换从「进设置面板→改下拉框→点保存」三步收敛成一键）：
 *   1) window.toggleUiLanguage() —— 侧工具栏 🌐 按钮 / Alt+L 快捷键入口，即时生效；
 *   2) 语言变化时同步所有指示器：侧工具栏徽标（中/EN）、AI 设置面板里的下拉框、
 *      任意位置的 [data-locale-indicator] 元素；
 *   3) 兜底重渲染那些在启动后才由 JS 生成的动态区域（快捷提示、连接状态条等）。
 *
 * 语言本身的分发由 App.i18n 负责（localStorage + 事件 + postMessage 跨窗口广播），
 * 本模块只做「UI 收尾」，不重复持久化。
 */
(function () {
    'use strict';

    /** 语言变化后可以尝试重渲染的动态面板（存在且是无参函数才调用） */
    var DYNAMIC_RERENDER = [
        'renderQuickPrompts',
        'updateConnectionStatus',
        'refreshWelcomeSetupCta',
        'renderSkillShortcuts',
        'refreshOcrLocaleTexts'
    ];

    function i18n() {
        return (window.App && window.App.i18n) || null;
    }

    function localeMeta(loc) {
        var api = i18n();
        if (api && typeof api.localeMeta === 'function') {
            try { return api.localeMeta(loc); } catch (_) { /* ignore */ }
        }
        return null;
    }

    /** 工具栏徽标。目录未就绪时保持原来的中 / EN。 */
    function localeBadge(loc) {
        var meta = localeMeta(loc);
        if (meta && meta.badge) return meta.badge;
        return loc === 'en' ? 'EN' : '中';
    }

    function localeNativeName(loc) {
        var meta = localeMeta(loc);
        if (meta && meta.nativeName) return meta.nativeName;
        return loc === 'en' ? 'English' : '中文';
    }

    /** 下拉框按目录补齐选项，新增语言不必再手写 <option>。 */
    function ensureLocaleOptions(select) {
        var api = i18n();
        if (!select || !api || typeof api.localeCatalog !== 'function') return;
        var list;
        try { list = api.localeCatalog(); } catch (_) { return; }
        if (!list) return;
        for (var i = 0; i < list.length; i++) {
            var id = list[i] && list[i].id;
            if (!id) continue;
            var exists = false;
            for (var j = 0; j < select.options.length; j++) {
                if (select.options[j].value === id) { exists = true; break; }
            }
            if (exists) continue;
            var opt = document.createElement('option');
            opt.value = id;
            opt.textContent = list[i].nativeName || id;
            select.appendChild(opt);
        }
    }

    function setAttrIfChanged(el, attr, value) {
        if (!el) return;
        if (el.getAttribute(attr) === value) return;
        el.setAttribute(attr, value);
    }

    /** 同步本窗口（必要时含父/子 iframe）中所有语言下拉框的选中值 */
    function syncLocaleSelects(loc) {
        var docs = [document];
        try {
            if (window.parent && window.parent !== window && window.parent.document) docs.push(window.parent.document);
        } catch (_) { /* 跨域 iframe 忽略 */ }
        docs.forEach(function (d) {
            var list;
            try { list = d.querySelectorAll('[data-locale-select], #ui-locale'); } catch (_) { return; }
            for (var i = 0; i < list.length; i++) {
                try {
                    ensureLocaleOptions(list[i]);
                    list[i].value = loc;
                } catch (_) { /* ignore */ }
            }
        });
    }

    /** 刷新所有语言指示器（徽标文本 + title 提示） */
    function refreshIndicators() {
        var api = i18n();
        if (!api) return;
        var loc = api.getLocale();
        var badge = localeBadge(loc);
        try {
            var list = document.querySelectorAll('[data-locale-indicator]');
            for (var i = 0; i < list.length; i++) {
                var el = list[i];
                if (el.textContent !== badge) el.textContent = badge;
                var titled = el.getAttribute('data-locale-indicator') !== 'silent';
                if (titled) setAttrIfChanged(el, 'title', api.t('nav.language') + ' (Alt+L)');
            }
        } catch (_) { /* ignore */ }
        syncLocaleSelects(loc);
    }

    function rerenderDynamicParts() {
        for (var i = 0; i < DYNAMIC_RERENDER.length; i++) {
            var fn = window[DYNAMIC_RERENDER[i]];
            if (typeof fn === 'function' && fn.length === 0) {
                try { fn(); } catch (_) { /* 单个面板失败不影响切换 */ }
            }
        }
    }

    /** 按语言目录轮换（并返回切换后的语言码） */
    function toggleUiLanguage(opts) {
        var api = i18n();
        if (!api) return null;
        var loc = typeof api.toggleLocale === 'function'
            ? api.toggleLocale(opts)
            : api.setLocale(
                typeof api.nextLocale === 'function' ? api.nextLocale(api.getLocale()) : (api.getLocale() === 'en' ? 'zh' : 'en'),
                opts
            );
        try {
            var msg = api.t('nav.languageSwitched', { lang: localeNativeName(loc) });
            if (window.App && typeof window.App.showToast === 'function') window.App.showToast(msg, 'info');
        } catch (_) { /* ignore */ }
        return loc;
    }

    function bindShortcut() {
        try {
            window.addEventListener('keydown', function (e) {
                if (!e.altKey || e.ctrlKey || e.metaKey) return;
                if (e.key !== 'l' && e.key !== 'L' && e.code !== 'KeyL') return;
                var tgt = e.target || {};
                var tag = (tgt.tagName || '').toUpperCase();
                if (tag === 'INPUT' || tag === 'TEXTAREA' || tgt.isContentEditable) return;
                e.preventDefault();
                toggleUiLanguage();
            });
        } catch (_) { /* ignore */ }
    }

    function boot() {
        var api = i18n();
        if (!api) return;
        if (typeof api.addLocaleChangeListener === 'function') {
            api.addLocaleChangeListener(function () {
                refreshIndicators();
                rerenderDynamicParts();
            });
        }
        try {
            window.addEventListener('smartpilot:locale-changed', function () {
                refreshIndicators();
                rerenderDynamicParts();
            });
        } catch (_) { /* ignore */ }
        bindShortcut();
        refreshIndicators();
    }

    window.toggleUiLanguage = toggleUiLanguage;
    window.refreshUiLanguageIndicators = refreshIndicators;
    window.SmartPilotLocaleSwitch = {
        toggle: toggleUiLanguage,
        refreshIndicators: refreshIndicators,
        badge: localeBadge
    };

    if (document.readyState === 'loading') {
        try { document.addEventListener('DOMContentLoaded', boot); } catch (_) { boot(); }
    } else {
        boot();
    }
})();
