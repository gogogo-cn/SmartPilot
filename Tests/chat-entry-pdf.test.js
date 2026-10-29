/**
 * ============================================================
 * Tests/chat-entry-pdf.test.js — PDF/Word 附件走后端解析回归
 * ============================================================
 * 背景（用户截图 2026-10-18）：
 *   上传 CAD 图纸 PDF（00600009946X3-03_PLIER ARM...pdf）后 AI 回复
 *   "已收到PDF文件，但需转换为文本/表格后才能解析" —— fileToText 从不把
 *   PDF/Word 路由到已存在的后端 ParseDocumentAsync（只路由 Excel），
 *   直接落入占位文案分支。
 * 修复：
 *   1. fileToText 对 Excel/PDF/Word 统一走后端 ParseDocumentAsync
 *      （DocumentRouter：MinerU/PdfSharp/DocxParser）
 *   2. JSON 解析容错 COM 斜杠丢失（_safeJsonLoose）
 *   3. 扫描版 PDF（success 但 0 字符）明确提示，不静默丢内容
 * 配套后端：Services/Parsers/PdfTextExtractor.vb（ContentReader + 字体/ToUnicode 解码），
 *          已用 4 类真实形态 PDF 样本 vbc 离线验证（见 pdfprobe harness）。
 * 运行：node Tests/chat-entry-pdf.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

const { window, App } = loadChatV2(['chat-entry.js']);
const chatEntry = App.ChatV2.chatEntry;

// ---------- 环境桩 ----------

// 真实桥失败形态：WebView2 COM hostObjects 代理会吃掉反斜杠（项目头号坑）
const com = s => s.replace(/\\/g, '/');

// Node 无 FileReader：最小桩（异步回调，与浏览器一致）
class FileReaderStub {
    readAsDataURL(file) {
        setTimeout(() => {
            this.result = 'data:application/pdf;base64,' + (file.__b64 || '');
            if (this.onload) this.onload({ target: { result: this.result } });
        }, 0);
    }
    readAsText(file) {
        setTimeout(() => {
            this.result = file.__text || '';
            if (this.onload) this.onload({ target: { result: this.result } });
        }, 0);
    }
}
global.FileReader = FileReaderStub;

function installBridge(parseImpl) {
    window.chrome.webview.hostObjects.excelBridge = {
        ParseDocumentAsync: parseImpl
    };
    return window.chrome.webview.hostObjects.excelBridge;
}

function okResult(extra) {
    return Object.assign({
        success: true,
        totalPages: 2,
        totalChars: 30,
        fullText: 'DIM A 10.0\nTOLERANCE ±0.1\nTITLE BLOCK',
        pages: [
            { index: 1, title: '第 1 页', text: 'DIM A 10.0', charCount: 10 },
            { index: 2, title: '第 2 页', text: 'TOLERANCE ±0.1', charCount: 13 }
        ]
    }, extra || {});
}

const pdfFile = { name: '00600009946X3-03_PLIER ARM, 5INCH PLIERS WRENCH.pdf', type: 'application/pdf', size: 2048576, __b64: 'JVBERi0xLjQK' };
const wordFile = { name: 'spec.docx', type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', size: 51200, __b64: 'UEsDBBQA' };
const excelFile = { name: 'data.xlsx', type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', size: 30720, __b64: 'UEsDBBQA' };
const textFile = { name: 'notes.txt', type: 'text/plain', size: 120, __text: 'hello 图纸' };
const pptFile = { name: 'deck.ppt', type: 'application/vnd.ms-powerpoint', size: 40960, __b64: 'xxx' };

// ---------- 用例 ----------

test('PDF 上传走后端 ParseDocumentAsync，不再落入占位文案', async () => {
    let calledWith = null;
    installBridge(async (fileName, b64) => {
        calledWith = { fileName, b64 };
        return com(JSON.stringify(okResult()));
    });

    const text = await chatEntry.fileToText(pdfFile);
    assert.ok(calledWith, '应调用 ParseDocumentAsync');
    assert.strictEqual(calledWith.fileName, pdfFile.name, '应传真实文件名（后端按扩展名路由）');
    assert.strictEqual(calledWith.b64, 'JVBERi0xLjQK', '应传剥离 DataURL 前缀后的纯 base64');
    assert.match(text, /\[PDF文档\]/, '应标注 PDF 文档');
    assert.match(text, /DIM A 10\.0/, '应包含提取到的图纸文字');
    assert.match(text, /TOLERANCE ±0\.1/, '应包含公差标注');
    assert.ok(!text.includes('当前不支持直接解析'), '不得再出现占位文案');
});

test('Word 上传同样走后端解析', async () => {
    let called = false;
    installBridge(async () => { called = true; return JSON.stringify(okResult({ totalPages: 1 })); });
    const text = await chatEntry.fileToText(wordFile);
    assert.ok(called, 'docx 应走后端');
    assert.match(text, /\[Word文档\]/);
});

test('Excel 上传保持既有后端解析路径（回归）', async () => {
    let called = false;
    installBridge(async () => { called = true; return JSON.stringify(okResult()); });
    const text = await chatEntry.fileToText(excelFile);
    assert.ok(called, 'xlsx 应继续走后端');
    assert.match(text, /\[Excel文件\]/);
});

test('COM 斜杠丢失：合法但被吃斜杠的 JSON 仍能解析（parse 不抛错）', async () => {
    // fullText 里的 Windows 路径 C:\Users\a\b 经 COM 变成 C://Users//a//b
    // （注意：formatParsedDocument 有 pages 时优先用 pages，故此处用无 pages 形态验证 fullText 通路）
    installBridge(async () => com(JSON.stringify(okResult({
        pages: [],
        fullText: '导出目录 C:\\Users\\a\\b\\drawing.pdf 已就绪',
        totalChars: 20
    }))));
    const text = await chatEntry.fileToText(pdfFile);
    assert.match(text, /C:\/\/Users\/\/a\/\/b/, '斜杠被吃是已知 COM 行为，内容仍应注入');
    assert.ok(!text.includes('解析失败'), '不得因斜杠丢失判失败');
});

test('COM 斜杠丢失：手拼 JSON 的单反斜杠（非法转义）经 _safeJsonLoose 修复', () => {
    // 后端手拼 JSON 未双写反斜杠：原文 JSON.parse 抛 SyntaxError
    const raw = '{"success":true,"totalChars":5,"fullText":"C:\\Users\\x","totalPages":1}';
    assert.throws(() => JSON.parse(raw), '原文应无法直接 parse');
    const obj = chatEntry._safeJsonLoose(raw);
    assert.ok(obj && obj.success, '修复后应解析成功');
    assert.strictEqual(obj.fullText, 'C:\\Users\\x', '反斜杠内容应保留');
});

test('_safeJsonLoose 对合法 JSON 是恒等变换', () => {
    const raw = JSON.stringify(okResult());
    assert.deepStrictEqual(chatEntry._safeJsonLoose(raw), JSON.parse(raw));
    assert.strictEqual(chatEntry._safeJsonLoose(null), null);
    assert.strictEqual(chatEntry._safeJsonLoose(123), null);
    // 彻底坏掉的串修不好也应返回 null 而不是抛
    assert.strictEqual(chatEntry._safeJsonLoose('{{{ not json at all'), null);
});

test('扫描版 PDF（success 但 0 字符）→ 明确提示，不静默丢内容', async () => {
    installBridge(async () => JSON.stringify(okResult({
        totalChars: 0,
        fullText: '',
        pages: [],
        isScanned: true,
        errorMessage: '检测到该 PDF 可能为扫描版图片文档'
    })));
    const text = await chatEntry.fileToText(pdfFile);
    assert.match(text, /\[扫描版PDF\]/, '应标注扫描版');
    assert.match(text, /没有可提取的文本层/, '应说明原因');
    assert.ok(!text.includes('当前不支持直接解析'), '不再是"格式不支持"的误导文案');
});

test('后端解析失败（success=false）→ 回落占位文案', async () => {
    installBridge(async () => JSON.stringify({ success: false, errorMessage: 'PDF 文件已损坏' }));
    const text = await chatEntry.fileToText(pdfFile);
    assert.ok(text.includes('当前不支持直接解析'), '解析失败应回落占位说明');
    assert.ok(text.includes(pdfFile.name));
});

test('宿主桥抛异常 → 不炸出 fileToText，回落"后端解析失败"占位文案', async () => {
    // 2026-10-18 起异常/超时路径给准确文案（"后端解析失败（原因）"），
    // 不再落到"当前不支持直接解析此格式"的误导文案——格式是支持的，只是这次没解析出来。
    installBridge(async () => { throw new Error('COM 调用超时'); });
    const text = await chatEntry.fileToText(pdfFile);
    assert.match(text, /后端解析失败/, '应说明是后端解析失败');
    assert.match(text, /COM 调用超时/, '应带上失败原因');
    assert.ok(text.includes(pdfFile.name));
    assert.ok(!text.includes('当前不支持直接解析'), '不得再用"格式不支持"误导用户');
});

test('宿主无 ParseDocumentAsync → 回落占位文案', async () => {
    window.chrome.webview.hostObjects.excelBridge = {}; // 旧宿主
    const text = await chatEntry.fileToText(pdfFile);
    assert.ok(text.includes('当前不支持直接解析'));
});

test('纯文本文件仍直接读取，不走后端', async () => {
    let called = false;
    installBridge(async () => { called = true; return JSON.stringify(okResult()); });
    const text = await chatEntry.fileToText(textFile);
    assert.ok(!called, 'txt 不应调后端');
    assert.strictEqual(text, 'hello 图纸');
});

test('未知二进制格式（ppt）保持占位说明（不误路由）', async () => {
    let called = false;
    installBridge(async () => { called = true; return JSON.stringify(okResult()); });
    const text = await chatEntry.fileToText(pptFile);
    assert.ok(!called, 'ppt 不应走后端（DocumentRouter 未注册）');
    assert.ok(text.includes('当前不支持直接解析'));
});

test('filesToAttachments：PDF → doc 附件且带提取内容', async () => {
    installBridge(async () => com(JSON.stringify(okResult())));
    const atts = await chatEntry.filesToAttachments([pdfFile]);
    assert.strictEqual(atts.length, 1);
    assert.strictEqual(atts[0].type, 'doc');
    assert.strictEqual(atts[0].fileName, pdfFile.name);
    assert.match(atts[0].content, /DIM A 10\.0/);
});

test('filesToAttachments：图片走 image 附件（回归，不受 PDF 路由影响）', async () => {
    const img = { name: 'shot.png', type: 'image/png', size: 1000, __b64: 'iVBORw0KGgo=' };
    const atts = await chatEntry.filesToAttachments([img]);
    assert.strictEqual(atts.length, 1);
    assert.strictEqual(atts[0].type, 'image');
    assert.strictEqual(atts[0].base64, 'iVBORw0KGgo=');
});

test('formatParsedDocument：内容超 12000 字符截断', () => {
    const big = okResult({ fullText: 'x'.repeat(20000), totalChars: 20000, pages: [] });
    const text = chatEntry.formatParsedDocument(big, 'big.pdf');
    assert.ok(text.length < 12500, '应截断');
    assert.match(text, /内容已截断/);
});

test('扫描版 PDF → 发送时不渲染，把原文件留给识别页面', async () => {
    installBridge(async () => JSON.stringify(okResult({ totalChars: 0, fullText: '', pages: [], isScanned: true })));
    let renderCalled = false;
    window.App.ChatV2.pdfPageRender = {
        renderPdfPages: async () => { renderCalled = true; return []; }
    };
    const atts = await chatEntry.filesToAttachments([pdfFile]);
    assert.strictEqual(renderCalled, false, '发送阶段不应渲染页面');
    assert.strictEqual(atts.length, 1, '只保留 doc 附件');
    assert.strictEqual(atts[0].type, 'doc');
    assert.strictEqual(atts[0].sourceFile, pdfFile);
    assert.match(atts[0].content, /\[扫描版PDF\]/);
    assert.match(atts[0].content, /识别页面时会渲染前 3 页/);
    assert.ok(!atts[0].content.includes('已仅向 AI 提供文件基本信息'));
});

test('扫描版 PDF 渲染器缺失 → 发送阶段仍只产出 doc', async () => {
    installBridge(async () => JSON.stringify(okResult({ totalChars: 0, fullText: '', pages: [], isScanned: true })));
    window.App.ChatV2.pdfPageRender = null;
    const atts = await chatEntry.filesToAttachments([pdfFile]);
    assert.strictEqual(atts.length, 1);
    assert.strictEqual(atts[0].type, 'doc');
    assert.strictEqual(atts[0].sourceFile, pdfFile);
});

test('有文本层的 PDF → 不渲染图片（控制上下文体积）', async () => {
    installBridge(async () => JSON.stringify(okResult()));
    let renderCalled = false;
    window.App.ChatV2.pdfPageRender = { renderPdfPages: async () => { renderCalled = true; return []; } };
    const atts = await chatEntry.filesToAttachments([pdfFile]);
    assert.ok(!renderCalled, '有文本层时不应触渲染');
    assert.strictEqual(atts.length, 1);
});

test('渲染器抛异常 → 被吞掉，doc 附件仍产出', async () => {
    installBridge(async () => JSON.stringify(okResult({ totalChars: 0, fullText: '', pages: [], isScanned: true })));
    window.App.ChatV2.pdfPageRender = { renderPdfPages: async () => { throw new Error('canvas 爆炸'); } };
    const atts = await chatEntry.filesToAttachments([pdfFile]);
    assert.strictEqual(atts.length, 1);
    assert.strictEqual(atts[0].type, 'doc');
});

// ---------- 2026-10-18 发送链路挂死事故的回退防御 ----------
// 事故：真实 CAD 图纸让后端 PdfTextExtractor.CleanupText 死循环（文化敏感 IndexOf
// 与序号 Replace 不一致，见 PdfTextExtractor.vb 历史坑注释），ParseDocumentAsync
// 永不返回 → 前端 await 无限挂起 → UI 永久停在"正在准备…"、附件发不出去。
// 后端已修；前端再加超时兜底：任何后端卡死都必须在时限内降级，恢复 UI 可操作。

test('withCallTimeout：正常 resolve 原值', async () => {
    const v = await chatEntry.withCallTimeout(Promise.resolve('ok'), 1000, 'x');
    assert.strictEqual(v, 'ok');
});

test('withCallTimeout：原 promise reject 时透传错误', async () => {
    await assert.rejects(
        chatEntry.withCallTimeout(Promise.reject(new Error('boom')), 1000, 'x'),
        /boom/
    );
});

test('withCallTimeout：超时 reject 且带自定义消息（真实短定时器）', async () => {
    const never = new Promise(() => {}); // 永不 settle
    await assert.rejects(
        chatEntry.withCallTimeout(never, 20, '文档解析超时（>90s），已放弃等待后端'),
        /已放弃等待后端/
    );
});

// Node v22 的 MockTimers 没有 tickAsync（v23+ 才提供）：用 tick(step) + 真实
// setImmediate 交替推进——tick 触发到期 mock 定时器，setImmediate 让出事件循环
// 排空微任务链（FileReader.onload → fileToBase64 resolve → 调桥 → withCallTimeout
// 装定时器），分步复现 tickAsync 的语义。setImmediate 不在 mock 列表里，是真实的。
async function advanceMockTimers(t, isSettled, maxMs) {
    const STEP = 500;
    for (let elapsed = 0; elapsed < maxMs; elapsed += STEP) {
        if (isSettled()) return;
        t.mock.timers.tick(STEP);
        await new Promise(resolve => setImmediate(resolve));
    }
}

test('fileToText：后端解析永不返回 → 超时降级，不挂死发送链路（mock timers）', async (t) => {
    installBridge(() => new Promise(() => {})); // 模拟后端死循环：永不 settle
    t.mock.timers.enable({ apis: ['setTimeout'] });
    const p = chatEntry.fileToText(pdfFile);
    let settled = false;
    const tracked = p.then(
        v => { settled = true; return v; },
        e => { settled = true; return { __rejected: e }; }
    );
    await advanceMockTimers(t, () => settled, chatEntry.DOC_PARSE_TIMEOUT_MS + 5000);
    t.mock.timers.reset();
    assert.ok(settled, `必须在 ${chatEntry.DOC_PARSE_TIMEOUT_MS}ms 内 settle，不得挂死"正在准备…"`);
    const outcome = await tracked;
    assert.ok(!outcome || !outcome.__rejected, '超时应被内部 catch 转为占位，不得抛给调用方');
    assert.match(outcome, /后端解析失败/, '应降级为占位说明');
    assert.match(outcome, /文档解析超时/, '占位应带上超时原因');
    assert.ok(outcome.includes(pdfFile.name));
    assert.ok(!outcome.includes('当前不支持直接解析'), '不得用"格式不支持"误导用户');
});

test('扫描版 PDF 渲染器卡死 → 发送不等待渲染', async () => {
    installBridge(async () => JSON.stringify(okResult({ totalChars: 0, fullText: '', pages: [], isScanned: true })));
    window.App.ChatV2.pdfPageRender = { renderPdfPages: () => new Promise(() => {}) };
    const atts = await chatEntry.filesToAttachments([pdfFile]);
    assert.strictEqual(atts.length, 1);
    assert.strictEqual(atts[0].sourceFile, pdfFile);
});
