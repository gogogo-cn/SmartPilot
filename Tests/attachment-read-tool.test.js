/**
 * read_attachment：附件正文按文件名交回，和读 Excel 一样是可规划的一步。
 * 运行：node Tests/attachment-read-tool.test.js
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

const { App } = loadChatV2(['services/attachment-read-tool.js']);
const api = App.ChatV2.attachmentReadTool;

const pdf = {
    type: 'doc',
    fileName: 'plier.pdf',
    mime: 'application/pdf',
    content: '[PDF文档] plier.pdf\nDIM A 10.0\nTOLERANCE ±0.1'
};

test('未指定文件名时返回全部文档附件正文', () => {
    const out = api.readContent([pdf], {});
    assert.strictEqual(out.ok, true);
    assert.strictEqual(out.result.fileCount, 1);
    assert.match(out.result.message, /DIM A 10\.0/);
    assert.match(out.result.message, /TOLERANCE/);
});

test('fileName 只返回点名的那一份', () => {
    const other = { type: 'doc', fileName: 'notes.txt', content: 'hello' };
    const out = api.readContent([pdf, other], { fileName: 'notes' });
    assert.strictEqual(out.result.fileCount, 1);
    assert.strictEqual(out.result.files[0].fileName, 'notes.txt');
    assert.ok(!out.result.message.includes('DIM A'));
});

test('没有附件时说明去读 Excel', () => {
    const out = api.readContent([], {});
    assert.strictEqual(out.ok, false);
    assert.match(out.error, /excel_read_range/);
});

test('图片不当成文档正文', () => {
    const out = api.readContent([{ type: 'image', fileName: 'p1.png', base64: 'AAA' }], {});
    assert.strictEqual(out.ok, true);
    assert.strictEqual(out.result.fileCount, 0);
    assert.strictEqual(out.result.imageCount, 1);
});

test('注册后规划工具表里附件读取排在读表前面', () => {
    const stored = new Map();
    const reg = {
        register(tool) { stored.set(tool.name, tool); return true; },
        get(name) { return stored.get(name) || null; }
    };
    App.ChatV2.toolRegistry = reg;
    reg.register({ name: 'excel_read_range', displayName: '读取Excel数据' });
    const tools = api.planningTools(reg);
    assert.deepStrictEqual(tools.map(t => t.name), ['read_attachment', 'excel_read_range']);
    assert.strictEqual(tools[0].displayName, '读取附件');
});

const scanned = {
    type: 'doc',
    fileName: 'plier.pdf',
    content: '[扫描版PDF] plier.pdf\n注：该 PDF 没有可提取的文本层（扫描件或纯图形）。'
};

test('没有文字层时标记自动识别，页面图可以还没渲', () => {
    const out = api.readContent([scanned], {});
    assert.strictEqual(out.result.needsRecognition, true);
    assert.match(out.result.message, /系统会接着识别前 3 页/);
    assert.ok(!out.result.message.includes('已随消息附上'));
});

test('有文字层时不要求识别页面', () => {
    const out = api.readContent([
        pdf,
        { type: 'image', fileName: 'extra.png', base64: 'IMG' }
    ], {});
    assert.strictEqual(out.result.needsRecognition, false);
    assert.ok(!out.result.message.includes('recognize_pages'));
});

test('有页面图时规划表在读附件和读表之间插入识别', () => {
    const stored = new Map();
    const reg = {
        register(tool) { stored.set(tool.name, tool); return true; },
        get(name) { return stored.get(name) || null; }
    };
    App.ChatV2.toolRegistry = reg;
    reg.register({ name: 'excel_read_range', displayName: '读取Excel数据' });
    const tools = api.planningTools(reg, {
        attachments: [scanned, { type: 'image', fileName: 'plier.pdf 第1页', base64: 'PAGE1' }]
    });
    assert.deepStrictEqual(tools.map(t => t.name), ['read_attachment', 'recognize_pages', 'excel_read_range']);
    assert.strictEqual(tools[1].displayName, '识别页面');
});

test('识别页面最多取前 3 页，点名按文件名过滤', () => {
    const pages = [1, 2, 3, 4].map(n => ({
        type: 'image', fileName: 'drawing.pdf 第' + n + '页', base64: 'P' + n
    }));
    const other = { type: 'image', fileName: 'other.pdf 第1页', base64: 'OX' };
    assert.strictEqual(api.selectPages(pages.concat([other]), {}).length, 3);
    const named = api.selectPages(pages.concat([other]), { fileName: 'other.pdf' });
    assert.strictEqual(named.length, 1);
    assert.strictEqual(named[0].fileName, 'other.pdf 第1页');
});

test('普通模型识别走 OCR 原图，不把图交给视觉模型', async () => {
    const seen = [];
    let visionCalled = false;
    const out = await api.recognizeContent([
        { type: 'image', fileName: 'plier.pdf 第1页', base64: 'ORIGINAL' }
    ], {}, {
        isVision: false,
        askVision: async () => { visionCalled = true; return 'should-not-run'; },
        recognizeWithOcr: async (images) => {
            seen.push(images[0].base64);
            return { pages: [{ name: images[0].fileName, text: 'DIM A 12.5', usedVision: false }], usedVision: false };
        }
    });
    assert.strictEqual(visionCalled, false);
    assert.deepStrictEqual(seen, ['ORIGINAL']);
    assert.strictEqual(out.result.mode, 'ocr');
    assert.match(out.result.message, /DIM A 12\.5/);
});

test('视觉模型只交出原图像素，不在工具内再调模型，也不压缩', async () => {
    let ocrCalled = false;
    const out = await api.recognizeContent([
        { type: 'image', fileName: 'plier.pdf 第1页', base64: 'ORIGINAL', mime: 'image/png' }
    ], {}, {
        isVision: true,
        recognizeWithOcr: async () => { ocrCalled = true; return { pages: [] }; }
    });
    assert.strictEqual(ocrCalled, false);
    assert.strictEqual(out.result.mode, 'vision');
    assert.strictEqual(out.result.visionImages.length, 1);
    assert.strictEqual(out.result.visionImages[0].base64, 'ORIGINAL');
    assert.match(out.result.message, /1600/);
});

test('尺寸行排到正文前面', () => {
    const text = '封面说明\n很长的工艺要求\nDIM A 12.50 mm\n标题栏 材料 钢';
    const out = api.promoteDrawingText(text);
    assert.match(out, /^DIM A 12\.50 mm\n标题栏/);
});

test('问图纸尺寸但文字层没有尺寸时，要接着识别 PDF 页面', () => {
    assert.strictEqual(api.drawingAskNeedsOcr('读取图纸中的尺寸', '封面说明\n没有数字', 'arm.pdf'), true);
    assert.strictEqual(api.drawingAskNeedsOcr('读取图纸中的尺寸', 'DIM A 12.50 mm', 'arm.pdf'), false);
    assert.strictEqual(api.drawingAskNeedsOcr('概述这份附件', '封面说明', 'novel.txt'), false);
});

test('无文字层且本轮还没调用识别时，应自动接上识别', () => {
    const read = api.readContent([scanned], {});
    assert.strictEqual(api.shouldAutoRecognize('read_attachment', read, []), true);
    assert.strictEqual(api.shouldAutoRecognize('read_attachment', read, ['recognize_pages']), false);
    assert.strictEqual(api.shouldAutoRecognize('excel_read_range', read, []), false);
});

test('识别页面时才渲染，已有页面图则不再渲', async () => {
    let calls = 0;
    App.ChatV2.pdfPageRender = {
        renderPdfPages: async (file, opts) => {
            calls += 1;
            assert.strictEqual(file.name, 'plier.pdf');
            assert.strictEqual(opts.maxPages, 3);
            return [{ pageNumber: 1, base64: 'PAGE1', mime: 'image/png' }];
        }
    };
    const turn = { attachments: [{ ...scanned, sourceFile: { name: 'plier.pdf' } }] };
    const first = await api.ensurePageImages(turn);
    assert.strictEqual(calls, 1);
    assert.strictEqual(first.filter(a => a.type === 'image').length, 1);
    await api.ensurePageImages(turn);
    assert.strictEqual(calls, 1);
});

test('没有页面图时识别步骤说明去用正文', async () => {
    const out = await api.recognizeContent([pdf], {}, { isVision: false });
    assert.strictEqual(out.ok, false);
    assert.match(out.error, /read_attachment/);
});
