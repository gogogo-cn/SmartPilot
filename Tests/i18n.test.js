/**
 * i18n.test.js — 中英文主路径文案 + 语言切换链路回归
 *
 * 覆盖：
 *   1) 默认语言 / 切换后文案变化 / 占位符替换（原有）；
 *   2) 设置·帮助·错误·历史等英文键齐全（原有）；
 *   3) 【新增】词条表双向齐平（zh↔en 无缺失、无重复、占位符一致）；
 *   4) 【新增】HTML 里声明的 data-i18n* 必须都能在词条表查到（防「挂了 key 忘了补词条」）；
 *   5) 【新增】toggleLocale 一键切换 + 监听器回调 + missingKeys 自检；
 *   6) 【新增】apply(root) 按作用域刷新 data-i18n / title / placeholder；
 *   7) 【新增】语言切换 UI 层（ui-locale-switch.js）可加载并暴露 toggleUiLanguage。
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const UI_ROOT = path.join(__dirname, '..', 'UI');

function loadI18n() {
    const code = fs.readFileSync(path.join(UI_ROOT, 'js', 'i18n.js'), 'utf8');
    const store = {};
    const win = {
        App: {},
        localStorage: {
            getItem(k) { return store[k] || null; },
            setItem(k, v) { store[k] = String(v); }
        },
        navigator: { language: 'zh-CN' },
        document: {
            documentElement: { lang: 'zh-CN' },
            readyState: 'complete',
            querySelectorAll() { return []; },
            getElementById() { return null; },
            addEventListener() {}
        },
        dispatchEvent() {},
        addEventListener() {},
        CustomEvent: function (type, init) { this.type = type; this.detail = init && init.detail; }
    };
    win.window = win;
    vm.runInNewContext(code, { console, window: win, document: win.document, localStorage: win.localStorage, navigator: win.navigator, CustomEvent: win.CustomEvent });
    return win.App.i18n;
}

function localeStringTable() {
    const code = fs.readFileSync(path.join(UI_ROOT, 'js', 'i18n.js'), 'utf8');
    const grab = (lang) => {
        const start = code.indexOf('    ' + lang + ': {');
        assert.ok(start > 0, '找不到语言表: ' + lang);
        let i = code.indexOf('{', start), depth = 0, quote = null;
        for (let p = i; p < code.length; p++) {
            const ch = code[p];
            if (quote) { if (ch === '\\') { p++; continue; } if (ch === quote) quote = null; continue; }
            if (ch === '"' || ch === "'") { quote = ch; continue; }
            if (ch === '{') depth++;
            else if (ch === '}') { depth--; if (depth === 0) return code.slice(i + 1, p); }
        }
        return '';
    };
    return { zh: grab('zh'), en: grab('en') };
}

function tableKeys(body) {
    return [...body.matchAll(/^\s{12}'([^']+)'\s*:/gm)].map(m => m[1]);
}

// ---------------------------------------------------------------- 原有用例
test('i18n: 默认中文，切换英文后进度文案变化', () => {
    const i18n = loadI18n();
    assert.strictEqual(i18n.getLocale(), 'zh');
    assert.ok(i18n.t('phase.thinking').includes('分析'));
    i18n.setLocale('en', { apply: false, reloadWelcome: false });
    assert.strictEqual(i18n.getLocale(), 'en');
    assert.ok(i18n.t('phase.thinking').toLowerCase().includes('understanding'));
    assert.ok(i18n.t('phase.stepOf', { n: 2, total: 5 }).includes('2'));
});

test('i18n: 缺失 key 回退中文或 key 本身', () => {
    const i18n = loadI18n();
    i18n.setLocale('en', { apply: false, reloadWelcome: false });
    assert.strictEqual(typeof i18n.t('input.send'), 'string');
    assert.ok(i18n.t('input.send').length > 0);
});

test('i18n: 设置/帮助/错误/历史英文键齐全', () => {
    const i18n = loadI18n();
    i18n.setLocale('en', { apply: false, reloadWelcome: false });
    const keys = [
        'settings.title', 'settings.save', 'settings.dangerAskLabel',
        'help.tabFaq', 'help.start.1.t', 'faq.common.0.q', 'faq.excel.0.a',
        'err.network_error', 'err.excel_not_ready', 'history.groupToday',
        'history.selectedCount', 'progress.running', 'nav.undo',
        'ext.title', 'kb.tabList', 'tp.newTask', 'editor.run',
        'tool.openFolder', 'tool.editRetry', 'mcp.start', 'kb.saveCfg'
    ];
    for (const k of keys) {
        const s = i18n.t(k, { host: 'Excel', n: 3, label: 'Excel', path: 'C:\\x' });
        assert.notStrictEqual(s, k, 'missing en key: ' + k);
        assert.ok(s.length > 0, k);
    }
    assert.ok(!/[\u4e00-\u9fff]/.test(i18n.t('settings.save')), 'settings.save should be English');
    assert.ok(!/[\u4e00-\u9fff]/.test(i18n.t('help.tabFaq')), 'help.tabFaq should be English');
    assert.ok(!/[\u4e00-\u9fff]/.test(i18n.t('ext.title').replace(/🛠️/g, '')), 'ext.title should be English');
    assert.ok(!/[\u4e00-\u9fff]/.test(i18n.t('tool.openFolder')), 'tool.openFolder should be English');
});

// ---------------------------------------------------------------- 新增用例
test('i18n: 中英词条表齐平（无缺失、无重复）', () => {
    const { zh, en } = localeStringTable();
    const zhKeys = tableKeys(zh), enKeys = tableKeys(en);
    assert.deepStrictEqual(enKeys.length, zhKeys.length);
    const missing = zhKeys.filter(k => !enKeys.includes(k));
    assert.deepStrictEqual(missing, [], '英文缺失词条: ' + missing.join(', '));
    const dup = zhKeys.filter((k, i) => zhKeys.indexOf(k) !== i);
    assert.deepStrictEqual(dup, [], '中文表重复 key: ' + dup.join(', '));
});

test('i18n: 同一 key 的占位符在中英文里保持一致', () => {
    const { zh, en } = localeStringTable();
    const pick = (body, key) => {
        const re = new RegExp("^\\s{12}'" + key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + "'\\s*:\\s*((?:'(?:[^'\\\\]|\\\\.)*')|(?:\"(?:[^\"\\\\]|\\\\.)*\"))", 'm');
        const m = re.exec(body);
        return m ? m[1] : null;
    };
    const vars = s => [...s.matchAll(/\{(\w+)\}/g)].map(m => m[1]).sort().join(',');
    const bad = [];
    for (const k of tableKeys(zh)) {
        const z = pick(zh, k), e = pick(en, k);
        if (!z || !e) continue;
        if (vars(z) !== vars(e)) bad.push(`${k}: zh[${vars(z)}] vs en[${vars(e)}]`);
    }
    assert.deepStrictEqual(bad, [], '占位符不一致:\n' + bad.join('\n'));
});

test('i18n: HTML 声明的 data-i18n* 都能查到词条（防漏翻译回归）', () => {
    const i18n = loadI18n();
    const htmlFiles = ['index.html', 'ai-settings-panel.html', 'mobile-control-panel.html',
        'ocr-panel.html', 'ocr-settings-panel.html', 'pomodoro.html', 'ppt-theme-editor.html'];
    const declared = new Set();
    for (const f of htmlFiles) {
        const p = path.join(UI_ROOT, f);
        if (!fs.existsSync(p)) continue;
        const html = fs.readFileSync(p, 'utf8').replace(/<script[\s\S]*?<\/script>/gi, '');
        const re = /data-i18n(?:-(?:title|placeholder|aria-label|value|html))?="([^"]+)"/g;
        let m;
        while ((m = re.exec(html))) declared.add(m[1]);
    }
    assert.ok(declared.size > 300, 'HTML 里声明的 i18n key 过少，可能扫描逻辑变了: ' + declared.size);
    const missing = [...declared].filter(k => !i18n.has(k));
    assert.deepStrictEqual(missing, [], 'HTML 声明了但词条表没有的 key: ' + missing.join(', '));
});

test('i18n: missingKeys 自检为空，toggleLocale 可双向切换并通知监听者', () => {
    const i18n = loadI18n();
    // 注意：i18n 跑在 vm 沙箱里，返回的数组是「另一个 realm」的实例，
    // 与宿主 [] 直接 deepStrictEqual 会因原型不同而失败 —— 统一转成宿主数组再断言。
    assert.deepStrictEqual(Array.from(i18n.missingKeys('en')), [], '英文仍有缺词');
    assert.deepStrictEqual(Array.from(i18n.missingKeys('zh')), [], '中文仍有缺词');

    const events = [];
    const off = i18n.addLocaleChangeListener(d => events.push(d.locale));
    assert.strictEqual(i18n.getLocale(), 'zh');
    let next = i18n.toggleLocale({ apply: false, reloadWelcome: false });
    assert.strictEqual(next, 'en');
    assert.strictEqual(i18n.getLocale(), 'en');
    next = i18n.toggleLocale({ apply: false, reloadWelcome: false });
    assert.strictEqual(next, 'zh');
    assert.deepStrictEqual(events, ['en', 'zh']);
    off();
    i18n.setLocale('zh', { apply: false, reloadWelcome: false });
    assert.deepStrictEqual(events, ['en', 'zh'], 'off() 后不应再收到回调');
});

test('i18n: 非法语言码收敛到 zh，getLocale 与 App.config 同步', () => {
    const i18n = loadI18n();
    assert.strictEqual(i18n.normalizeLocale('en-US'), 'en');
    assert.strictEqual(i18n.normalizeLocale('zh-CN'), 'zh');
    assert.strictEqual(i18n.normalizeLocale('fr'), 'zh');
    assert.strictEqual(i18n.setLocale('klingon', { apply: false, reloadWelcome: false }), 'zh');
});

test('i18n: 语言目录归一化别名，未知标签不当成已登记语言', () => {
    const i18n = loadI18n();
    assert.deepStrictEqual(Array.from(i18n.locales()), ['zh', 'en']);
    assert.strictEqual(i18n.resolveLocale('zh-Hans'), 'zh');
    assert.strictEqual(i18n.resolveLocale('en_GB'), 'en');
    assert.strictEqual(i18n.resolveLocale('EN'), 'en');
    assert.strictEqual(i18n.resolveLocale('zh-TW'), 'zh');
    assert.strictEqual(i18n.resolveLocale('fr'), null);
    assert.strictEqual(i18n.resolveLocale('ja-JP'), null);
    assert.strictEqual(i18n.localeMeta('en-US').badge, 'EN');
    assert.strictEqual(i18n.localeMeta('zh').nativeName, '中文');
    assert.strictEqual(i18n.localeMeta('en').htmlLang, 'en');
    assert.strictEqual(i18n.localeMeta('zh-CN').dateLocale, 'zh-CN');
    assert.strictEqual(i18n.nextLocale('zh'), 'en');
    assert.strictEqual(i18n.nextLocale('en'), 'zh');
    const catalog = Array.from(i18n.localeCatalog()).map(d => d.id);
    assert.deepStrictEqual(catalog, ['zh', 'en']);
});

test('i18n: apply(root) 按作用域刷新文本/title/placeholder', () => {
    const i18n = loadI18n();
    const mk = (tag, attrs) => ({
        tagName: tag,
        _a: Object.assign({}, attrs),
        textContent: '', innerHTML: '', value: '',
        getAttribute(n) { return this._a[n] != null ? this._a[n] : null; },
        setAttribute(n, v) { this._a[n] = v; },
        hasAttribute(n) { return this._a[n] !== undefined; }
    });
    const els = [
        mk('BUTTON', { 'data-i18n': 'input.send' }),
        mk('INPUT', { 'data-i18n-placeholder': 'history.search', type: 'text' }),
        mk('DIV', { 'data-i18n-title': 'nav.settings', 'aria-label': 'x' })
    ];
    const scope = {
        getElementById() { return null; },
        querySelectorAll(sel) {
            const key = /\[data-i18n(?:-([a-z-]+))?\]/.exec(sel);
            return els.filter(e => e.hasAttribute(key[1] ? 'data-i18n-' + key[1] : 'data-i18n'));
        }
    };
    i18n.setLocale('en', { apply: false, reloadWelcome: false });
    i18n.apply(scope);
    assert.ok(/send/i.test(els[0].textContent), '按钮文本应变成英文: ' + els[0].textContent);
    assert.ok(/search/i.test(els[1].getAttribute('placeholder')), els[1].getAttribute('placeholder'));
    assert.ok(/settings|model/i.test(els[2].getAttribute('title')), els[2].getAttribute('title'));
});

test('UI 层: ui-locale-switch.js 暴露一键切换入口', () => {
    const p = path.join(UI_ROOT, 'js', 'ui-locale-switch.js');
    assert.ok(fs.existsSync(p), '缺少 UI/js/ui-locale-switch.js');
    const code = fs.readFileSync(p, 'utf8');
    const win = {
        App: {},
        document: { readyState: 'complete', querySelectorAll: () => [], addEventListener() {} },
        addEventListener() {}
    };
    win.window = win;
    vm.runInNewContext(code, { console, window: win, document: win.document });
    assert.strictEqual(typeof win.toggleUiLanguage, 'function', 'toggleUiLanguage 未导出');
    assert.strictEqual(typeof win.SmartPilotLocaleSwitch.toggle, 'function');
    assert.strictEqual(win.SmartPilotLocaleSwitch.badge('zh'), '中');
    assert.strictEqual(win.SmartPilotLocaleSwitch.badge('en'), 'EN');
});

test('入口: index.html 引入 ui-locale-switch 并带语言切换按钮', () => {
    const html = fs.readFileSync(path.join(UI_ROOT, 'index.html'), 'utf8');
    assert.match(html, /js\/ui-locale-switch\.js\?v=v[0-9]{12}/, 'index.html 未引入（或缺少版本戳）ui-locale-switch.js');
    assert.match(html, /onclick="toggleUiLanguage\(\)"/, '侧工具栏缺少一键切换按钮');
    assert.match(html, /data-locale-indicator/, '缺少语言徽标元素');
});

test('入口: AI 设置面板的语言下拉改即时生效（带 onchange）', () => {
    const html = fs.readFileSync(path.join(UI_ROOT, 'ai-settings-panel.html'), 'utf8');
    assert.match(html, /id="ui-locale"[^>]*onchange="onUiLocaleChange/, '设置面板语言下拉未立即生效');
    assert.match(html, /js\/ui-locale-switch\.js\?v=v[0-9]{12}/, '设置面板未加载语言切换模块');
    assert.match(html, /function onUiLocaleChange|window\.onUiLocaleChange/, '缺少 onUiLocaleChange 实现');
});

test('入口: AI 设置面板不再引用已删除的 app-ai-* 链路，并加载 bridge-accessor', () => {
    const html = fs.readFileSync(path.join(UI_ROOT, 'ai-settings-panel.html'), 'utf8');
    assert.ok(!/app-ai-core\.js|app-ai-context\.js|app-ai-stream-ui\.js|app-ai-calls\.js/.test(html),
        '设置面板仍引用已删除的旧 AI 脚本（会 404）');
    assert.match(html, /js\/bridge-accessor\.js\?v=v[0-9]{12}/, '设置面板需加载 bridge-accessor 以安全调用宿主桥');
    assert.match(html, /js\/app-model\.js/, '设置面板需 app-model（testAIConfig/saveSettings）');
});
