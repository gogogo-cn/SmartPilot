/**
 * excel-connection-widget 返回值归一化 / VSTO 契约
 */
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const ROOT = path.join(__dirname, '..');
const SRC = path.join(ROOT, 'UI', 'js', 'excel-connection-widget.js');

function loadHelpers() {
    const code = fs.readFileSync(SRC, 'utf8');
    const logs = [];
    const fakeDoc = {
        readyState: 'complete',
        getElementById: () => null,
        createElement: () => ({ style: {}, appendChild() {}, remove() {} }),
        addEventListener() {},
        body: { appendChild() {} },
        head: { appendChild() {} }
    };
    const win = {
        App: {},
        ChatLog: { info() {}, warn() {}, error() {} },
        console,
        document: fakeDoc,
        setTimeout,
        clearTimeout,
        AbortSignal: { timeout: () => ({}) },
        fetch: async () => ({ json: async () => ({}) })
    };
    win.window = win;
    // 没有 #connection-status → initialize 早退，但 createWidget 不会跑，helpers 未挂载。
    // 直接用正则抽出函数体并 eval 太脆；改为在沙箱里临时挂容器再执行。
    fakeDoc.getElementById = (id) => {
        if (id === 'connection-status') {
            return {
                style: {},
                title: '',
                set innerHTML(v) { this._html = v; },
                get innerHTML() { return this._html || ''; }
            };
        }
        // createWidget 里大量 getElementById：返回带 style 的桩，避免炸
        return {
            style: {},
            textContent: '',
            innerHTML: '',
            value: '',
            appendChild() {},
            querySelector() { return null; },
            addEventListener() {}
        };
    };
    vm.runInNewContext(code, win, { filename: 'excel-connection-widget.js' });
    // 等待同步 initialize 跑完
    assert.ok(win.App.ExcelConnection, 'ExcelConnection 未挂载');
    assert.equal(typeof win.App.ExcelConnection._parseBridgeResult, 'function');
    return win.App.ExcelConnection;
}

test('parseBridgeResult: VSTO 裸工作表名 / JSON / 失败结构', () => {
    const ex = loadHelpers();
    const p = ex._parseBridgeResult;
    const bare = p('销售数据');
    assert.equal(bare.success, true);
    assert.equal(bare.data, '销售数据');
    assert.equal(p('{"success":true,"activeSheet":"Sheet1","sheetCount":3}').activeSheet, 'Sheet1');
    const arr = p('["A","B"]').data;
    assert.ok(Array.isArray(arr));
    assert.equal(arr.join(','), 'A,B');
    assert.equal(p('{"success":false,"message":"boom"}').message, 'boom');
    assert.equal(p('').success, false);
});

test('pickActivateInfo: 兼容 VSTO 顶层与 Python data 层（含 sheets）', () => {
    const ex = loadHelpers();
    const pick = ex._pickActivateInfo;
    const a = pick({ success: true, activeSheet: '销售数据', sheetCount: 1, sheets: ['销售数据'] });
    assert.equal(a.activeSheet, '销售数据');
    assert.equal(a.sheetCount, 1);
    assert.equal(a.sheets.join(','), '销售数据');
    const b = pick({ success: true, data: { activeSheet: 'S1', sheets: ['S1', 'S2'] } });
    assert.equal(b.activeSheet, 'S1');
    assert.equal(b.sheetCount, 2);
    assert.equal(b.sheets.join(','), 'S1,S2');
});

test('coerceSheetList: GetSheets 数组 / 失败 / 双编码', () => {
    const ex = loadHelpers();
    const c = ex._coerceSheetList;
    assert.equal(c({ success: true, data: ['A', 'B'] }).join(','), 'A,B');
    assert.equal(c({ success: true, data: '["X","Y"]' }).join(','), 'X,Y');
    assert.equal(c({ success: false, error: 'x' }).length, 0);
    assert.equal(c(['P']).join(','), 'P');
});

test('bridgeErr: 优先 error / message / detail', () => {
    const ex = loadHelpers();
    assert.equal(ex._bridgeErr({ message: 'm1' }), 'm1');
    assert.equal(ex._bridgeErr({ error: 'e1', message: 'm1' }), 'e1');
    assert.equal(ex._bridgeErr({}, 'fb'), 'fb');
});

test('源码: VSTO 隐藏连接/断开；不再写死 Bridge Mode', () => {
    const src = fs.readFileSync(SRC, 'utf8');
    assert.match(src, /isHostEnv\(\)\s*\{[\s\S]*btnDisc\.style\.display='none'/);
    assert.ok(!/Bridge Mode/.test(src), '模板仍写死 Bridge Mode');
    assert.match(src, /conn\.hostHint/);
    assert.match(src, /info-workbook/);
});
