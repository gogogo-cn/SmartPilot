/**
 * ============================================================
 * Tests/pdf-page-render.test.js — pdf.js 页面渲染兜底模块
 * ============================================================
 * 覆盖：
 *   1. pdf.js 缺失（CDN 不通）时全链路安全降级返回 []
 *   2. pdfjsLib 存在时：白底填充、边长缩放、base64 剥离、页数上限
 * 运行：node Tests/pdf-page-render.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

// 先加载被测模块（IIFE，直接挂 window.PdfPageRender）
loadChatV2(['pdf-page-render.js']);
const api = global.window.PdfPageRender;

// ---------- pdf.js 缺失场景（Node 环境默认） ----------

test('模块加载后暴露 API', () => {
    assert.ok(api, 'window.PdfPageRender 应存在');
    assert.strictEqual(typeof api.renderPdfPages, 'function');
    assert.strictEqual(api.isAvailable(), false, 'Node 环境无 pdfjsLib');
});

test('pdf.js 未加载 → renderPdfPages 返回 []（不抛错）', async () => {
    const out = await api.renderPdfPages({ name: 'x.pdf', arrayBuffer: async () => new Uint8Array(10) });
    assert.deepStrictEqual(out, []);
});

test('file 为 null / 无 arrayBuffer → 返回 []', async () => {
    assert.deepStrictEqual(await api.renderPdfPages(null), []);
    assert.deepStrictEqual(await api.renderPdfPages({ name: 'x.pdf' }), []);
});

// ---------- pdfjsLib 存在场景（桩） ----------

function installPdfjsStub(numPages, pageSize) {
    const calls = { getDocument: 0, rendered: 0, destroyed: false, lastFill: null, fillRects: 0 };
    global.pdfjsLib = {
        GlobalWorkerOptions: {},
        getDocument(opts) {
            calls.getDocument += 1;
            assert.ok(opts.data instanceof Uint8Array, '应传 Uint8Array 数据');
            return {
                promise: Promise.resolve({
                    numPages,
                    async getPage(p) {
                        return {
                            getViewport({ scale }) {
                                return { width: pageSize.w * scale, height: pageSize.h * scale };
                            },
                            render({ canvasContext, viewport }) {
                                calls.rendered += 1;
                                assert.ok(viewport.width > 0 && viewport.height > 0);
                                return { promise: Promise.resolve() };
                            },
                            cleanup() {}
                        };
                    },
                    destroy() { calls.destroyed = true; }
                })
            };
        }
    };
    // 桩 document.createElement('canvas')：getContext 记录白底填充
    const realCreate = global.document.createElement.bind(global.document);
    global.document.createElement = (tag) => {
        if (tag !== 'canvas') return realCreate(tag);
        return {
            width: 0, height: 0,
            getContext() {
                return {
                    _fs: '',
                    get fillStyle() { return this._fs; },
                    set fillStyle(v) { this._fs = v; calls.lastFill = v; },
                    fillRect() { calls.fillRects += 1; }
                };
            },
            toDataURL(mime) { return 'data:' + mime + ';base64,QUJD'; }
        };
    };
    return calls;
}

test('pdfjsLib 可用：渲染 N 页，白底填充 + base64 剥离', async () => {
    const calls = installPdfjsStub(5, { w: 612, h: 792 });
    const file = { name: 'scan.pdf', arrayBuffer: async () => new Uint8Array([1, 2, 3]) };
    const out = await api.renderPdfPages(file, { maxPages: 3, maxSidePx: 1600 });

    assert.strictEqual(api.isAvailable(), true);
    assert.strictEqual(out.length, 3, '应受 maxPages=3 限制（PDF 共 5 页）');
    assert.deepStrictEqual(out.map(o => o.pageNumber), [1, 2, 3]);
    out.forEach(o => {
        assert.strictEqual(o.base64, 'QUJD', '应剥离 data URL 前缀');
        assert.strictEqual(o.mime, 'image/png');
        assert.ok(o.width >= 1 && o.height >= 1);
    });
    assert.strictEqual(calls.lastFill, '#ffffff', '渲染前应填白底（防透明 PDF 出黑图）');
    assert.strictEqual(calls.fillRects, 3, '每页一次白底填充');
    assert.ok(calls.destroyed, '结束后应 destroy pdfDoc 释放 worker');
});

test('小页不放大：scale 上限 2.0', async () => {
    installPdfjsStub(1, { w: 400, h: 600 }); // 最长边 600 < maxSidePx 1600 → scale 封顶 2
    const file = { name: 'small.pdf', arrayBuffer: async () => new Uint8Array([1]) };
    const out = await api.renderPdfPages(file, { maxSidePx: 1600 });
    assert.strictEqual(out.length, 1);
    assert.strictEqual(out[0].width, 800, '400*2.0=800（scale 封顶）');
    assert.strictEqual(out[0].height, 1200);
});

test('大页缩小：最长边不超过 maxSidePx', async () => {
    installPdfjsStub(1, { w: 3000, h: 4000 }); // 最长边 4000 → scale=1600/4000=0.4
    const file = { name: 'big.pdf', arrayBuffer: async () => new Uint8Array([1]) };
    const out = await api.renderPdfPages(file, { maxSidePx: 1600 });
    assert.strictEqual(out[0].width, 1200, '3000*0.4');
    assert.strictEqual(out[0].height, 1600, '4000*0.4=1600（封顶值）');
});

test('渲染中途异常 → 返回已渲染页面（不整批失败）', async () => {
    // 第 2 页 getPage 抛错：整体 try 捕获后返回 []（保守策略，调用方有文字兜底）
    global.pdfjsLib = {
        GlobalWorkerOptions: {},
        getDocument() {
            return {
                promise: Promise.resolve({
                    numPages: 3,
                    async getPage(p) {
                        if (p === 2) throw new Error('page 2 broken');
                        return {
                            getViewport: ({ scale }) => ({ width: 100 * scale, height: 100 * scale }),
                            async render() { return { promise: Promise.resolve() }; },
                            cleanup() {}
                        };
                    },
                    destroy() {}
                })
            };
        }
    };
    const realCreate = global.document.createElement.bind(global.document);
    global.document.createElement = (tag) => {
        if (tag !== 'canvas') return realCreate(tag);
        return {
            width: 0, height: 0,
            getContext() {
                return { fillStyle: '', fillRect() {} };
            },
            toDataURL(mime) { return 'data:' + mime + ';base64,QUJD'; }
        };
    };
    const file = { name: 'half.pdf', arrayBuffer: async () => new Uint8Array([1]) };
    const out = await api.renderPdfPages(file, { maxPages: 3 });
    assert.deepStrictEqual(out, [], '保守策略：任何一页失败整体返回 []，由调用方文字兜底');
});

test('逐页回调 onProgress，便于进度条显示当前页', async () => {
    installPdfjsStub(4, { w: 200, h: 200 });
    const seen = [];
    const file = { name: 'scan.pdf', arrayBuffer: async () => new Uint8Array([1]) };
    const out = await api.renderPdfPages(file, {
        maxPages: 2,
        onProgress: (info) => seen.push(info.page + '/' + info.total)
    });
    assert.strictEqual(out.length, 2);
    assert.deepStrictEqual(seen, ['1/2', '2/2']);
});

test('workerSrc 自动配置为 CDN 地址', async () => {
    installPdfjsStub(0, { w: 100, h: 100 });
    await api.renderPdfPages({ name: 'x.pdf', arrayBuffer: async () => new Uint8Array([1]) }, { maxPages: 1 });
    assert.strictEqual(global.pdfjsLib.GlobalWorkerOptions.workerSrc,
        'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js');
});
