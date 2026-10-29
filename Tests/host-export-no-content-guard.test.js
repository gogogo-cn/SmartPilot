/**
 * ============================================================
 * Tests/host-export-no-content-guard.test.js — 导出类请求不得被"文档没有可读内容"误导
 * ============================================================
 * 背景（实测 2026-10-06 00:55，Word 宿主）：
 *   用户："将当前文档导出为 PDF。"
 *   模型回答："当前当前文档中没有可读取的内容。因此无法直接导出为 PDF…"
 * 根因：prompt-builder 非 Excel 分支里，凡是消息命中 _refersDocNonExcel（含"文档"二字）
 *   且 documentText 为空，就注入
 *     （当前当前文档没有可读取的文字内容）+ 必须如实告知"…没有可读取的内容"
 *   ——弱模型（qwen2.5）原样照抄；而"导出 PDF"根本不需要任何可读正文（空白文档也应能导出）。
 *   附带 bug：label 已是"当前文档"，模板又拼一次"当前" → "当前当前文档"。
 * 修复：文件产出类请求（detectFileOutputIntent）完全豁免文档素材/无内容警告；
 *       "没有可读取的内容"只在真·内容处理任务（总结/润色/改写/翻译/提炼）里声明。
 * 运行：node Tests/host-export-no-content-guard.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

const { window, App } = loadChatV2(['services/prompt-builder.js']);
const pb = App.ChatV2.promptBuilder;

const setHost = (host) => { window.SmartPilotHost = host; };
const WORD = { isExcel: false, isWord: true, isPowerPoint: false, kind: 'word', label: 'Word' };
const PPT = { isExcel: false, isWord: false, isPowerPoint: true, kind: 'powerpoint', label: 'PowerPoint' };

test('Word 导出 PDF：不得注入"没有可读取的内容"警告', () => {
    setHost(WORD);
    const out = pb.buildUserMessage({ kind: 'text', userMessage: '将当前文档导出为 PDF。', context: {} });
    assert.ok(!out.includes('没有可读取的文字内容'),
        '"导出 PDF"不需要正文，绝不能给模型"读不到内容"的结论（那会被照抄成"无法导出"）');
    assert.ok(!out.includes('无法直接导出'), '不得注入任何"无法导出"引导');
});

test('PPT 导出 PDF：同样不注入无内容警告', () => {
    setHost(PPT);
    const out = pb.buildUserMessage({ kind: 'text', userMessage: '将当前演示文稿导出为 PDF', context: {} });
    assert.ok(!out.includes('没有可读取的文字内容'));
});

test('真·内容任务且正文为空：仍必须声明"没有可读取的内容"（防编造）', () => {
    setHost(WORD);
    const out = pb.buildUserMessage({ kind: 'text', userMessage: '帮我总结这份文档的内容', context: {} });
    assert.ok(out.includes('没有可读取的文字内容'),
        '总结类任务在无正文时仍要明确告知读不到内容，避免模型凭标题编造正文');
});

test('文案不得出现重复的"当前当前"（label 已含"当前"）', () => {
    setHost(WORD);
    const out = pb.buildUserMessage({ kind: 'text', userMessage: '帮我润色这份文档的文字', context: {} });
    assert.ok(!/当前当前/.test(out), `出现重复"当前"：${out.slice(0, 200)}`);
});
