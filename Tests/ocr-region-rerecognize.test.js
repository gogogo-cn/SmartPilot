/**
 * ============================================================
 * Tests/ocr-region-rerecognize.test.js — OCR 局部重新识别回归
 * ============================================================
 * 背景（2026-10-14 用户反馈「OCR 列表里调整局部重新识别后出现了 JSON」）：
 *   后端 OcrFormatFormatter.FormatAsRawOcrData 把 plaintext 格式输出为
 *   raw_ocr_data 调试 JSON：[{ "text": "销售二部", "confidence": 0.9982,
 *   "box": [[19,19],[180,19],[180,60],[19,60]] }]（原文视图有意用它看坐标）。
 *   但局部区域重新识别（recognizeRegionAndUpdate → enqueueOcrRecognition →
 *   processOcrQueue → recognizeImage）把这个字符串直接当成列表项文本，
 *   于是调整框选后重新识别，列表里显示一整段 JSON。
 *
 * 本测试锁死四件事：
 *   1) parseRawOcrDataItems 存在，且只把"元素均为含 text 字符串字段的对象"的
 *      JSON 数组识别为 raw_ocr_data（纯文本/ Markdown/ 坏 JSON 一律放行，防误伤）；
 *   2) processOcrQueue 在写 item.text / item.originalText 之前先提取纯文本，
 *      并用 raw_ocr_data 的 confidence 刷新列表置信度；
 *   3) recognizeImage 支持 options.isRegion：区域识别固定请求 plaintext
 *      （避免 Markdown/Excel 表格语法进单元格），且不覆盖整图识别的格式缓存；
 *   4) 4 处运行时目录的 ocr-panel.html 都含该修复（防"改了源没部署"）。
 * 运行：node Tests/ocr-region-rerecognize.test.js
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
// 统一成 LF：仓库 HTML 是 CRLF，正则里按 \n 锚定行尾
const read = p => fs.readFileSync(p, 'utf8').replace(/\r\n/g, '\n');
const exists = p => fs.existsSync(p);

const OCR_PANEL = path.join(ROOT, 'UI', 'ocr-panel.html');
const RUNTIME_DIRS = [
    path.join(ROOT, 'bin', 'Debug', 'UI'),
    path.join(ROOT, 'bin', 'Release', 'UI'),
    path.join(ROOT, 'Hosts', 'PptAIVSTOAddIn', 'bin', 'Debug', 'UI'),
    path.join(ROOT, 'Hosts', 'WordAIVSTOAddIn', 'bin', 'Debug', 'UI')
];

/** 从 HTML 中取出 parseRawOcrDataItems 函数源码并实例化（无 DOM 依赖，可直接 eval） */
function loadParseRawOcrDataItems(html) {
    const m = html.match(/function parseRawOcrDataItems\(result\) \{[\s\S]*?\n        \}\n/);
    assert.ok(m, 'ocr-panel.html 中找不到 parseRawOcrDataItems 函数');
    // eslint-disable-next-line no-new-func
    return new Function(m[0] + '; return parseRawOcrDataItems;')();
}

test('parseRawOcrDataItems 定义带防护条件（只认 raw_ocr_data 数组）', () => {
    assert.ok(exists(OCR_PANEL), '缺少 UI/ocr-panel.html');
    const html = read(OCR_PANEL);
    assert.match(html, /function parseRawOcrDataItems\(result\) \{/, '缺少 parseRawOcrDataItems');
    assert.match(html, /trimmed\.startsWith\('\['\)/, '必须以 [ 开头才尝试解析（避免误伤普通文本）');
    assert.match(html, /Array\.isArray\(parsed\)/, '必须校验是数组');
    assert.match(html, /typeof it\.text === 'string'/, '必须校验每个元素都含 text 字符串字段');
});

test('parseRawOcrDataItems 行为：JSON 提取 / 普通文本放行', () => {
    const parse = loadParseRawOcrDataItems(read(OCR_PANEL));

    // 用户截图中的真实返回（Newtonsoft 缩进 JSON）
    const shot = '[\r\n  {\r\n    "text": "销售二部",\r\n    "confidence": 0.9982,\r\n    "box": [[19, 19], [180, 19], [180, 60], [19, 60]]\r\n  }\r\n]';
    const items = parse(shot);
    assert.ok(Array.isArray(items), 'raw_ocr_data JSON 必须被识别为数组');
    assert.strictEqual(items.length, 1);
    assert.strictEqual(items[0].text, '销售二部');
    assert.strictEqual(items[0].confidence, 0.9982);

    // 普通识别文本 / Markdown / 非 raw_ocr_data 数组 / 坏 JSON 一律返回 null（原样使用）
    assert.strictEqual(parse('华南'), null);
    assert.strictEqual(parse('| 销售二部 |'), null);
    assert.strictEqual(parse('[]'), null);
    assert.strictEqual(parse('["a","b"]'), null);
    assert.strictEqual(parse('[1,2]'), null);
    assert.strictEqual(parse('[{ "text": '), null);
    assert.strictEqual(parse(42), null);
    assert.strictEqual(parse(null), null);

    // 多行区域：按行拼接
    const multi = parse('[{"text":"甲","confidence":0.9,"box":[]},{"text":"乙","confidence":0.8,"box":[]}]');
    assert.deepStrictEqual(multi.map(i => i.text), ['甲', '乙']);
});

test('区域识别队列：先提取纯文本再写 item.text / originalText，并刷新置信度', () => {
    const html = read(OCR_PANEL);
    const queue = html.match(/async function processOcrQueue\(\) \{[\s\S]*?\n        \}\n/);
    assert.ok(queue, '找不到 processOcrQueue');
    const src = queue[0];

    // 必须带 isRegion 调用 recognizeImage（固定 plaintext，避免表格语法进单元格）
    assert.match(src, /recognizeImage\(task\.regionBase64, task\.fileName, task\.recognitionMode, \{ isRegion: true \}\)/,
        '区域识别必须传 { isRegion: true }');

    // 解析提取必须发生在保存 originalResult / 写 text 之前（否则 JSON 会进列表）
    const parseIdx = src.indexOf('parseRawOcrDataItems(result)');
    const origIdx = src.indexOf('const originalResult = result;');
    const textIdx = src.indexOf('.text = result.trim()');
    assert.ok(parseIdx > 0 && origIdx > 0 && textIdx > 0, '队列缺少解析/赋值关键行');
    assert.ok(parseIdx < origIdx, '必须先解析 raw_ocr_data 再保存 originalResult');
    assert.ok(parseIdx < textIdx, '必须先解析 raw_ocr_data 再写 item.text');

    // 多行拼接 + 置信度刷新
    assert.match(src, /rawItems\.map\(it => \(it && typeof it\.text === 'string'\) \? it\.text : ''\)\.join\('\\n'\)/,
        '必须把 raw_ocr_data 的 text 按行拼接');
    assert.match(src, /rawItems\[0\]\.confidence|it\.confidence/, '必须用本次识别的 confidence 刷新列表显示');
});

test('recognizeImage：isRegion 固定 plaintext 且不覆盖整图格式缓存', () => {
    const html = read(OCR_PANEL);
    const fn = html.match(/async function recognizeImage\(base64Image, fileName, recognitionMode = 'auto', options = \{\}\) \{[\s\S]*?\n        \}\n/);
    assert.ok(fn, 'recognizeImage 缺少 options 参数（区域识别标记）');
    const src = fn[0];
    assert.match(src, /const isRegion = !!options\.isRegion;/, '缺少 isRegion 标记');
    assert.match(src, /isRegion \? 'plaintext' :/, '区域识别必须固定请求 plaintext');
    assert.match(src, /if \(isRegion\) \{[\s\S]*?return pollObj\.text;/, '区域识别必须提前返回，不能覆盖 rawRecognitionResult / formattedResults');
});

test('识别选中：占位文字不进入解析全文，结束后改掉「正在加载 OCR 引擎」', () => {
    const html = read(OCR_PANEL);
    const queue = html.match(/async function processOcrQueue\(\) \{[\s\S]*?\n        \}\n/);
    assert.ok(queue, '找不到 processOcrQueue');
    assert.match(queue[0], /syncRecognitionTextsFromAnnotations\(\)/, '每块识别完要写回解析用的全文');
    assert.match(queue[0], /updateAnnotationViews\(\)/, '队列结束要刷新列表');
    assert.match(queue[0], /已重新识别 /, '队列结束要把状态从「正在加载 OCR 引擎」改掉');

    const recognize = html.match(/async function recognizeImage\([\s\S]*?\n        \}\n/);
    assert.ok(recognize, '找不到 recognizeImage');
    assert.match(recognize[0], /if \(!isRegion\) updateStatusMessage\(stage, 'info'\)/,
        '区域重识别不能把状态写成「正在加载 OCR 引擎」');

    const plain = html.match(/function annotationPlainText\(item\) \{[\s\S]*?\n        \}\n/);
    assert.ok(plain, '缺少 annotationPlainText');
    const fn = new Function(html.match(/function isPlaceholderOcrText\(text\) \{[\s\S]*?\n        \}\n/)[0]
        + plain[0] + '; return annotationPlainText;')();
    assert.strictEqual(fn({ text: '识别中', _lastRecognizedText: '姓名 韦小宝' }), '姓名 韦小宝');
    assert.strictEqual(fn({ text: '识别失败: 超时', originalText: '公民身份号码 110101199001011234' }), '公民身份号码 110101199001011234');
    assert.strictEqual(fn({ text: '性别 男' }), '性别 男');
});

test('4 处运行时目录的 ocr-panel.html 都含该修复（防改了源没部署）', () => {
    for (const dir of RUNTIME_DIRS) {
        const p = path.join(dir, 'ocr-panel.html');
        assert.ok(exists(p), '运行时副本缺失: ' + p);
        const c = read(p);
        assert.ok(c.includes('function parseRawOcrDataItems'), '运行时副本缺修复: ' + p);
        assert.ok(c.includes('{ isRegion: true }'), '运行时副本缺 isRegion 调用: ' + p);
        assert.ok(c.includes('function annotationPlainText'), '运行时副本缺解析占位过滤: ' + p);
        assert.ok(c.includes('已重新识别 '), '运行时副本缺重识别完成状态: ' + p);
    }
});
