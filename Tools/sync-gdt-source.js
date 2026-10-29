/**
 * 一次性同步：把 Services/GdtSymbolRecognizer.vb 写进
 *   1) UI/ocr-panel.html 的 builtin-recognizer-source 内嵌块（domain-packs.test.js 逐字节比对）
 *   2) UI/gdt-symbol-recognizer.vb（面板加载的展示副本）
 * 用法：node Tools/sync-gdt-source.js
 */
'use strict';
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const SRC_VB = path.join(ROOT, 'Services', 'GdtSymbolRecognizer.vb');
const OPEN = '<script type="text/plain" id="builtin-recognizer-source">';

// 与 Tests/domain-packs.test.js 的归一化保持一致：去 BOM、CRLF→LF、去尾部空白
const norm = (s) => s.replace(/^\uFEFF/, '').replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/\s+$/, '');
const source = norm(fs.readFileSync(SRC_VB, 'utf8'));

// 1) UI/gdt-symbol-recognizer.vb 整份覆盖
const gdtCopy = path.join(ROOT, 'UI', 'gdt-symbol-recognizer.vb');
const before = fs.existsSync(gdtCopy) ? norm(fs.readFileSync(gdtCopy, 'utf8')) : '';
fs.writeFileSync(gdtCopy, source + '\n', 'utf8');
console.log((before === source ? '[=] ' : '[+] ') + 'UI/gdt-symbol-recognizer.vb');

// 2) UI/ocr-panel.html 内嵌块替换
const htmlPath = path.join(ROOT, 'UI', 'ocr-panel.html');
let html = fs.readFileSync(htmlPath, 'utf8');
const start = html.indexOf(OPEN);
if (start < 0) throw new Error('ocr-panel.html 找不到内嵌源码块');
const bodyStart = start + OPEN.length;
const close = html.indexOf('</script>', bodyStart); 
if (close < 0) throw new Error('ocr-panel.html 内嵌源码块没有闭合');
// 行尾约定（Tests/domain-packs.test.js 逐字节比对逼出来的）：
    // 开标签后紧跟一个独立 LF，块内用 LF，闭合前再一个独立 LF。外层 HTML 是 CRLF，
    // 混排是历史约定——测试只 replace(/^\n/) 剥一个 LF，用 CRLF 开会多个 \r 导致比对失败。
const oldBlock = html.slice(bodyStart, close);
const trailing = oldBlock.endsWith('\n') ? '\n' : '';
html = html.slice(0, bodyStart) + '\n' + source + trailing + html.slice(close);
fs.writeFileSync(htmlPath, html, 'utf8');
console.log('[+] UI/ocr-panel.html embedded recognizer source');
