/**
 * ============================================================
 * Tests/host-export-pdf-dispatch.test.js — PPT/Word 导出 PDF 确定性硬短路回归
 * ============================================================
 * 背景（实测 2026-10-06 00:48）：
 *   Word 宿主里用户说"将当前文档导出为 PDF"，dispatcher 没有对应规则 →
 *   落到 agent-loop 后小模型认为"当前文档中没有任何可读取的内容"直接反问用户
 *   （用户看到的提示就是"导出 PDF 要什么可读取的文字内容"），导出工具一次都没调。
 * 修复：非 Excel 宿主命中"导出/另存 + PDF"→ 直调 <fam>_export_pdf（filePath 留空，
 *       app-mcp 的 _fillOptionalExportPathIfNeeded 会补全真实桌面路径，后端落盘）。
 * 运行：node Tests/host-export-pdf-dispatch.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

// dispatcher 的关键词常量来自 config/routing-constants.js（单一事实源），必须一并加载
const { window, App } = loadChatV2(['config/routing-constants.js', 'services/mcp-dispatcher.js']);
const dispatcher = App.ChatV2.mcpDispatcher;

const opts = (message) => ({ message, toolRegistry: null, sessionStore: null, turn: { context: {} } });
const setHost = (host) => { window.SmartPilotHost = host; };

test('Word：将当前文档导出为PDF → 硬短路直调 word_export_pdf', async () => {
    setHost({ isExcel: false, isPowerPoint: false, kind: 'word' });
    const d = await dispatcher.tryDispatch(opts('将当前文档导出为PDF'));
    assert.strictEqual(d.hardMatched, true, '应走确定性硬短路（不进 LLM，避免反问"文档没有内容"）');
    assert.strictEqual(d.toolName, 'word_export_pdf');
});

test('Word：另存为 PDF 文件 → word_export_pdf', async () => {
    setHost({ isExcel: false, isPowerPoint: false, kind: 'word' });
    const d = await dispatcher.tryDispatch(opts('把这份文档另存为 PDF 文件'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(d.toolName, 'word_export_pdf');
});

test('PPT：将演示文稿导出为 PDF → ppt_export_pdf', async () => {
    setHost({ isExcel: false, isPowerPoint: true, kind: 'powerpoint' });
    const d = await dispatcher.tryDispatch(opts('将当前演示文稿导出为 PDF'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(d.toolName, 'ppt_export_pdf');
});

test('PPT：导出为图片 → ppt_export_image', async () => {
    setHost({ isExcel: false, isPowerPoint: true, kind: 'powerpoint' });
    const d = await dispatcher.tryDispatch(opts('把当前幻灯片导出为图片'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(d.toolName, 'ppt_export_image');
});

test('Word：总结文档内容 不得被导出规则劫持', async () => {
    setHost({ isExcel: false, isPowerPoint: false, kind: 'word' });
    const d = await dispatcher.tryDispatch(opts('帮我总结这份文档的内容'));
    assert.notStrictEqual(d.toolName, 'word_export_pdf', '非导出请求不得命中导出硬短路');
});

test('Excel 宿主：导出 PDF 规则不得命中（Excel 无导出 PDF 工具）', async () => {
    setHost({ isExcel: true, isPowerPoint: false, kind: 'excel' });
    const d = await dispatcher.tryDispatch(opts('将当前工作簿导出为PDF'));
    assert.notStrictEqual(d.toolName, 'word_export_pdf');
    assert.notStrictEqual(d.toolName, 'ppt_export_pdf');
});
