/**
 * ============================================================
 * Tests/host-export-write-gate.test.js — 写类工具裁决：文档级导出必须是"执行命令"
 * ============================================================
 * 背景（实测 2026-10-06 00:59，Word 宿主）：
 *   用户："将当前文档导出为 PDF。"
 *   实际：导出成功，但模型先跑了一次 word_format_selected_table_cell（美化表头）才导出。
 * 根因：word_export_pdf 含 'export' → 判定为写类工具 → 走"写类工具裁决闸门"。
 *   而 EXEC_COMMAND_PATTERNS 的产物关键词白名单只有 文件|csv|xlsx|txt|图片|图表|桌面|…
 *   —— **没有 PDF/文档/幻灯片 **，于是"导出为 PDF"三条规则都不命中 → 落到 LLM 二次裁决
 *   → 小模型判成 other → 放行通用路由 → agent-loop 自由发挥（顺手美化表格）。
 * 修复：新增"命令式 + 导出动词 + 文档级产物（pdf/文档/幻灯片/演示文稿/图片…）"的确定性规则。
 * 运行：node Tests/host-export-write-gate.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

const { App } = loadChatV2();
const orch = App.ChatV2.orchestrator;

test('word_export_pdf 确实是写类工具（会被裁决闸门拦）', () => {
    assert.strictEqual(orch.isSideEffectTool('word_export_pdf'), true);
    assert.strictEqual(orch.isSideEffectTool('ppt_export_pdf'), true);
});

test('"将当前文档导出为 PDF。" → 规则判定 command（不走 LLM 二次裁决）', async () => {
    const v = await orch.adjudicateWriteHit({ message: '将当前文档导出为 PDF。', toolName: 'word_export_pdf' });
    assert.strictEqual(v.verdict, 'command');
    assert.strictEqual(v.source, 'rule', `应由规则闸门直接判定，实际 source=${v.source}（落到 LLM 裁决会让导出前多跑格式化）`);
});

test('"把演示文稿另存为 PDF" / "导出为图片" 同样判定 command', async () => {
    for (const msg of ['把演示文稿另存为 PDF', '将当前幻灯片导出为图片', '把当前文档输出为 PDF 文件']) {
        const v = await orch.adjudicateWriteHit({ message: msg, toolName: 'ppt_export_pdf' });
        assert.strictEqual(v.verdict, 'command', `${msg} 应判为 command`);
    }
});

test('要代码时仍判 code（"写一段 VBA 把文档导出为 PDF"）', async () => {
    const v = await orch.adjudicateWriteHit({ message: '写一段 VBA 把文档导出为 PDF', toolName: 'word_export_pdf' });
    assert.strictEqual(v.verdict, 'code', '含"写…VBA"必须只出代码，绝不执行导出');
});

test('疑问语气仍判 other（"导出PDF怎么做？"）', async () => {
    const v = await orch.adjudicateWriteHit({ message: '导出PDF怎么做？', toolName: 'word_export_pdf' });
    assert.strictEqual(v.verdict, 'other');
});

// ══ 副作用分级：导出 = 只产出文件，不等于"改文档" ══
test('产出文件类：导出/保存工具免 LLM 裁决，随意措辞都判 command', async () => {
    assert.strictEqual(orch.isArtifactOnlyTool('word_export_pdf'), true);
    assert.strictEqual(orch.isArtifactOnlyTool('ppt_export_image'), true);
    assert.strictEqual(orch.isArtifactOnlyTool('excel_export_csv'), true);
    for (const msg of ['搞份 PDF 出来看看', '弄一个 pdf 文件给我', '输出 pdf 版本']) {
        const v = await orch.adjudicateWriteHit({ message: msg, toolName: 'word_export_pdf' });
        assert.strictEqual(v.verdict, 'command', `"${msg}" 应判 command`);
        assert.strictEqual(v.source, 'rule', `"${msg}" 不得落到 LLM 二次裁决`);
    }
});

test('改动类工具仍受严格裁决保护（不因含 save/export 就放松）', () => {
    assert.strictEqual(orch.isArtifactOnlyTool('excel_write_range'), false);
    assert.strictEqual(orch.isArtifactOnlyTool('word_format_selected_table_cell'), false);
    assert.strictEqual(orch.isArtifactOnlyTool('excel_delete_row'), false);
    assert.strictEqual(orch.isArtifactOnlyTool('word_save'), true, '纯保存工具属于产出文件类');
    assert.strictEqual(orch.isArtifactOnlyTool('word_read_text'), false, '读取类不属于任何写类');
});
