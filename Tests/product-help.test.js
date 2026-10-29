/**
 * product-help.test.js — 帮助文案与导出会话纯函数行为（通过加载模块验证导出）
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

function loadProductHelp(hostKind) {
    const code = fs.readFileSync(
        path.join(__dirname, '..', 'UI', 'js', 'product-help.js'),
        'utf8'
    );
    const downloads = [];
    const fakeDoc = {
        createElement(tag) {
            const el = {
                tagName: String(tag).toUpperCase(),
                style: {},
                classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
                setAttribute() {},
                appendChild() {},
                remove() {},
                click() {},
                addEventListener() {},
                querySelector() { return null; },
                querySelectorAll() { return []; },
                closest() { return null; }
            };
            Object.defineProperty(el, 'innerHTML', { set() {}, get() { return ''; } });
            return el;
        },
        body: { appendChild() {}, removeChild() {} },
        getElementById() { return null; },
        addEventListener() {},
        removeEventListener() {}
    };
    const BlobCtor = function Blob(parts, opts) {
        this.parts = parts;
        this.type = opts && opts.type;
    };
    const URLApi = {
        createObjectURL(blob) {
            downloads.push(blob);
            return 'blob:test';
        },
        revokeObjectURL() {}
    };
    const win = {
        SmartPilotHost: {
            kind: hostKind || 'excel',
            label: hostKind === 'word' ? 'Word' : (hostKind === 'powerpoint' ? 'PowerPoint' : 'Excel'),
            isExcel: !hostKind || hostKind === 'excel'
        },
        ChatLog: { info() {}, warn() {} },
        App: {
            ChatV2: {
                sessionStore: {
                    getCompletedTurnsMessages() {
                        return [
                            { role: 'user', content: '你好' },
                            { role: 'assistant', content: '您好，有什么可以帮您？' }
                        ];
                    },
                    getCurrentTurn() { return null; }
                }
            },
            showToast() {}
        },
        document: fakeDoc,
        URL: URLApi,
        Blob: BlobCtor,
        setTimeout(fn) { try { fn(); } catch (_) { /* ignore */ } return 0; }
    };
    win.window = win;
    const sandbox = {
        console,
        window: win,
        document: fakeDoc,
        Blob: BlobCtor,
        URL: URLApi,
        setTimeout: win.setTimeout
    };
    vm.runInNewContext(code, sandbox);
    return { api: win, downloads };
}

test('product-help: 导出 Markdown 含用户与助手内容', () => {
    const { api, downloads } = loadProductHelp('excel');
    assert.strictEqual(typeof api.exportCurrentConversation, 'function');
    const ok = api.exportCurrentConversation({ format: 'md' });
    assert.strictEqual(ok, true);
    assert.ok(downloads.length >= 1, '应触发下载');
});

test('product-help: showProductHelp 可调用不抛错', () => {
    const { api } = loadProductHelp('word');
    assert.strictEqual(typeof api.showProductHelp, 'function');
    assert.doesNotThrow(() => api.showProductHelp({ tab: 'faq' }));
});

test('product-help: 导出历史会话', () => {
    const { api, downloads } = loadProductHelp('excel');
    api.App.state = {
        chatHistory: [{
            id: 42,
            title: '测试会话',
            messages: [
                { role: 'user', content: '历史问题' },
                { role: 'assistant', content: '历史回答' }
            ]
        }]
    };
    assert.strictEqual(typeof api.exportHistoryConversation, 'function');
    const ok = api.exportHistoryConversation(42, { format: 'md' });
    assert.strictEqual(ok, true);
    assert.ok(downloads.length >= 1, '应触发历史会话下载');
});

test('product-help: 空会话导出返回 false', () => {
    const code = fs.readFileSync(path.join(__dirname, '..', 'UI', 'js', 'product-help.js'), 'utf8');
    const sandbox = {
        console,
        window: {
            SmartPilotHost: { kind: 'excel', label: 'Excel', isExcel: true },
            ChatLog: { info() {}, warn() {} },
            App: {
                ChatV2: { sessionStore: { getCompletedTurnsMessages() { return []; }, getCurrentTurn() { return null; } } },
                showToast() {}
            },
            document: {
                createElement() { return { style: {}, classList: { add() {}, remove() {} }, setAttribute() {}, appendChild() {}, remove() {}, addEventListener() {}, click() {} }; },
                body: { appendChild() {}, removeChild() {} },
                getElementById() { return null; },
                addEventListener() {},
                removeEventListener() {}
            },
            URL: { createObjectURL() { return 'x'; }, revokeObjectURL() {} },
            Blob: function () {},
            setTimeout() { return 0; }
        }
    };
    sandbox.window.window = sandbox.window;
    vm.runInNewContext(code, sandbox);
    assert.strictEqual(sandbox.window.exportCurrentConversation({ format: 'md' }), false);
});
