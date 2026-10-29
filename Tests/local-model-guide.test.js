/**
 * local-model-guide.test.js — AI 设置面板「本地模型安装引导」契约回归
 *
 * 背景（2026-10-22）：引导卡片的「下载 Ollama / 下载 llama.cpp」按钮点击无反应。
 * 根因：设置面板是 file:// iframe（与主页面虚拟主机源跨域），拿不到 WebView2
 * 宿主桥，桥调用静默落到本地 9527 桥（未实现 OpenExternalUrl）→ 无任何反馈。
 * 修复：内嵌模式下 postMessage 转交父页面代开链接 / 代复制，失败时状态行
 * 给出可手动复制的链接。本测试锁死该契约，防回归。
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const UI_ROOT = path.join(__dirname, '..', 'UI');
const read = (p) => fs.readFileSync(path.join(UI_ROOT, p), 'utf8').replace(/\r\n/g, '\n');

// 与 i18n.test.js 相同的沙箱加载方式：i18n.js 以 IIFE 挂在 window.App 上
function loadI18n() {
    const code = fs.readFileSync(path.join(UI_ROOT, 'js', 'i18n.js'), 'utf8');
    const store = {};
    const win = {
        App: {},
        localStorage: {
            getItem(k) { return store[k] || null; },
            setItem(k, v) { store[k] = String(v); }
        },
        document: {
            readyState: 'complete',
            documentElement: { lang: '' },
            querySelectorAll() { return []; },
            getElementById() { return null; },
            addEventListener() {}
        },
        dispatchEvent() {},
        addEventListener() {},
        CustomEvent: function (type, init) { this.type = type; this.detail = init && init.detail; }
    };
    win.window = win;
    vm.runInNewContext(code, {
        console, window: win, document: win.document, localStorage: win.localStorage,
        navigator: win.navigator, CustomEvent: win.CustomEvent
    });
    return win.App.i18n;
}

test('引导卡片：仅 Ollama / llama.cpp 显示，含下载/复制/检测入口', () => {
    const html = read('ai-settings-panel.html');
    assert.match(html, /id="local-model-guide"[^>]*style="display:none;"/,
        '引导卡片应默认隐藏（由 onProviderChange 按提供商显隐）');
    assert.match(html, /id="local-guide-ollama"/, '缺少 Ollama 引导块');
    assert.match(html, /id="local-guide-llamacpp"/, '缺少 llama.cpp 引导块');
    assert.match(html, /id="local-guide-status"/, '缺少检测状态行');
    assert.match(html, /openLocalModelDownload\('https:\/\/ollama\.com\/download'\)/,
        'Ollama 引导缺少下载按钮');
    assert.match(html, /openLocalModelDownload\('https:\/\/github\.com\/ggml-org\/llama\.cpp\/releases'\)/,
        'llama.cpp 引导缺少下载按钮');
    assert.match(html, /copyGuideCommand\(this, 'ollama pull qwen2\.5:7b'\)/,
        'Ollama 拉模型命令缺少复制按钮');
    assert.match(html, /copyGuideCommand\(this, 'llama-server -m models\\\\qwen2\.5-7b-instruct-q4_k_m\.gguf --port 8080'\)/,
        'llama.cpp 启动命令缺少复制按钮');
    assert.match(html, /onclick="detectLocalModels\(\)"/, '引导卡片缺少检测本地模型按钮');
});

test('引导卡片：内嵌模式经 postMessage 转交父页面，失败必须可见', () => {
    const html = read('ai-settings-panel.html');
    // 内嵌抽屉（file:// iframe 跨域，无宿主桥）：必须走 postMessage 代开
    assert.match(html, /postMessage\(\{ type: 'smartpilot:open-external-url', url: url, token: token \}, '\*'\)/,
        '内嵌模式未通过 postMessage 请求父页面代开链接');
    // 失败回包 / 3 秒超时都要给出可手动复制的链接（绝不做无反应点击）
    assert.match(html, /'smartpilot:open-external-url-result'/, '缺少代开链接结果回包处理');
    assert.match(html, /_pendingGuideOpen\[token\] = url;/, '缺少待回包登记（超时兜底依赖它）');
    assert.match(html, /setTimeout\(function \(\) \{[\s\S]*?showGuideOpenFail\(url\);/, '缺少父页面超时未回包的失败兜底');
    assert.match(html, /if \(!data\.ok\) showGuideOpenFail\(data\.url \|\| pendingUrl, data\.error\);/,
        '父页面代开失败时未给出可见反馈');
    // 复制同样优先父页面代复制（iframe 剪贴板可能被权限策略拦）
    assert.match(html, /postMessage\(\{ type: 'smartpilot:copy-text', text: cmd, token: token \}, '\*'\)/,
        '复制未优先走父页面代复制');
    assert.match(html, /'smartpilot:copy-text-result'/, '缺少代复制结果回包处理');
});

test('父页面（side-toolbar）：处理代开链接与代复制消息并回包', () => {
    const js = read('js/side-toolbar.js');
    assert.match(js, /data\.type === 'smartpilot:open-external-url' && data\.url/,
        '父页面未处理代开链接消息');
    assert.match(js, /callHost\('OpenExternalUrl', openUrl\)/,
        '父页面未通过宿主桥打开链接');
    assert.match(js, /token: data\.token, url: openUrl, ok: !!ok/,
        '代开链接回包未携带 token（面板无法匹配待回包）');
    assert.match(js, /data\.type === 'smartpilot:copy-text'/,
        '父页面未处理代复制消息');
    assert.match(js, /type: 'smartpilot:copy-text-result',\s*\n\s*token: data\.token/,
        '代复制回包未携带 token');
});

test('app-model：检测到本地模型时隐藏引导卡片，未检测到才显示', () => {
    const js = read('js/app-model.js');
    // onProviderChange：云端直接隐藏；本地只切换引导块（整体显隐交给检测结果）
    assert.match(js, /if \(guideCard && !providerConfig\.isLocal\) \{\s*guideCard\.style\.display = 'none';/,
        'onProviderChange 未在云端提供商下隐藏引导卡片');
    assert.match(js, /ollamaGuide\) ollamaGuide\.style\.display = providerValue === 'llamacpp' \? 'none' : '';/,
        '未按提供商切换 Ollama / llama.cpp 引导块');
    // updateLocalGuideStatus：count>0 隐藏（已装好不再显示引导），count=0 显示
    assert.match(js, /function updateLocalGuideStatus\(count\) \{[\s\S]*?const card = document\.getElementById\('local-model-guide'\);\s*if \(card\) card\.style\.display = count > 0 \? 'none' : '';/,
        '引导卡片未按检测结果显隐（已检测到模型时应隐藏）');
    assert.match(js, /updateLocalGuideStatus\(localModels\.length\);/, '检测到模型时未刷新状态行');
    assert.match(js, /updateLocalGuideStatus\(0\);/, '未检测到模型时未刷新状态行');
});

test('i18n：引导卡片新增词条 zh/en 齐平且占位符一致', () => {
    const i18n = loadI18n();
    const keys = [
        'settings.localGuide.title',
        'settings.localGuide.statusNone',
        'settings.localGuide.statusOk',
        'settings.localGuide.ollamaHint',
        'settings.localGuide.ollamaStep1',
        'settings.localGuide.downloadOllama',
        'settings.localGuide.ollamaStep2',
        'settings.localGuide.ollamaPullCmd',
        'settings.localGuide.copy',
        'settings.localGuide.copied',
        'settings.localGuide.copyFail',
        'settings.localGuide.openFail',
        'settings.localGuide.ollamaStep3',
        'settings.localGuide.ollamaServeCmd',
        'settings.localGuide.llamacppHint',
        'settings.localGuide.llamacppStep1',
        'settings.localGuide.downloadLlamacpp',
        'settings.localGuide.llamacppStep2',
        'settings.localGuide.llamacppStep3',
        'settings.localGuide.llamacppServeCmd',
        'settings.localGuide.detectNow'
    ];
    i18n.setLocale('en', { apply: false, reloadWelcome: false });
    const missing = Array.from(i18n.missingKeys('en')).filter(k => k.indexOf('settings.localGuide.') === 0);
    assert.deepStrictEqual(missing, [], '英文引导词条缺失: ' + missing.join(', '));
    for (const k of keys) {
        assert.ok(i18n.has(k), `词条表缺少 ${k}`);
        assert.ok(i18n.t(k).length > 0, `${k} 英文为空`);
    }
    // statusOk 占位符 {count} 双向可用
    assert.ok(i18n.t('settings.localGuide.statusOk', { count: 3 }).indexOf('3') !== -1,
        'statusOk 的 {count} 未正确插值');
    i18n.setLocale('zh', { apply: false, reloadWelcome: false });
    assert.ok(i18n.t('settings.localGuide.statusOk', { count: 3 }).indexOf('3') !== -1,
        '中文 statusOk 的 {count} 未正确插值');
});

test('i18n：HTML 声明的 data-i18n 都能在词条表查到（含新增引导键）', () => {
    const i18n = loadI18n();
    const html = read('ai-settings-panel.html').replace(/<script[\s\S]*?<\/script>/gi, '');
    const declared = new Set();
    const re = /data-i18n(?:-(?:title|placeholder|aria-label|value|html))?="([^"]+)"/g;
    let m;
    while ((m = re.exec(html))) declared.add(m[1]);
    assert.ok(declared.size > 20, '扫描到的 data-i18n 键过少，正则可能失效: ' + declared.size);
    const missing = [...declared].filter(k => !i18n.has(k));
    assert.deepStrictEqual(missing, [], 'HTML 声明了但词条表没有的 key: ' + missing.join(', '));
});
