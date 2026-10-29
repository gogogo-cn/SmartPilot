/**
 * 工具信封、Turn 完成条件、缺参补槽。
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2, DEFAULT_ORDER } = require('./helpers/harness');

const { App } = loadChatV2(DEFAULT_ORDER.concat(['services/mcp-dispatcher.js']));
const env = App.ChatV2.toolEnvelope;
const goalApi = App.ChatV2.turnGoal;
const disp = App.ChatV2.mcpDispatcher;

test('信封：MCP content 收成 value / context，失败单独放 error', () => {
    const ok = env.fromInvoke('excel_create_chart', {
        ok: true,
        result: {
            content: [{ type: 'text', text: JSON.stringify({ success: true, message: '已创建图表', chartName: 'Chart 1', anchor: 'E2' }) }]
        }
    });
    assert.strictEqual(ok.ok, true);
    assert.strictEqual(ok.value, '已创建图表');
    assert.strictEqual(ok.context.chartName, 'Chart 1');
    assert.strictEqual(ok.context.anchor, 'E2');
    assert.strictEqual(ok.error, '');

    const bad = env.fromRaw('excel_apply_filter', { success: false, error: '范围不能为空' }, true);
    assert.strictEqual(bad.ok, false);
    assert.strictEqual(bad.error, '范围不能为空');
});

test('回喂带验收头；JSON 正文和假成功被验收结论替换', () => {
    const envelope = { ok: false, value: '', context: { sheet: 'Sheet1' }, error: '范围不能为空', tool: 'excel_apply_filter' };
    const feed = env.prefixFeed(envelope, '{"success":false}');
    assert.ok(feed.startsWith('[验收] 失败'));
    assert.ok(feed.includes('[错误] 范围不能为空'));
    assert.ok(feed.includes('Sheet1'));

    const turn = {
        goal: goalApi.build('把表头加粗；再画一张柱状图'),
        _envelopes: [{ ok: true, tool: 'excel_format_range', error: '', value: '已加粗', context: {} }]
    };
    const verdict = env.buildVerdict(turn);
    assert.strictEqual(verdict.ok, false);
    assert.ok(verdict.unmet.some(u => u.kind === 'chart'));
    const text = env.applyVerdict('已成功完成全部操作。', verdict);
    assert.ok(!/已成功完成/.test(text));
    assert.ok(/尚未完成/.test(text));
    assert.strictEqual(env.looksLikeRawJson('{"success":true,"mergedRows":1}'), true);
    assert.strictEqual(env.applyVerdict('{"success":true,"mergedRows":1}', { ok: true, failed: [], unmet: [], text: '已完成：excel_merge_sheets', failureNote: '', unmetNote: '' }), '已完成：excel_merge_sheets');
});

test('目标：短命令写完成条件，成功的工具不算缺项', () => {
    const g = goalApi.build('不要改原始数据，把 A1:D1 加粗；再画一张柱状图');
    assert.ok(g.constraints.some(c => c.includes('不要改原始数据')));
    assert.ok(g.doneWhen.some(d => d.kind === 'format' && d.label.includes('A1:D1')));
    assert.ok(g.doneWhen.some(d => d.kind === 'chart'));
    const multi = goalApi.markMultiStep(g, g.text);
    assert.strictEqual(multi.multiStep, true);
    assert.ok(multi.constraints.some(c => c.includes('不得声称已完成')));

    const open = goalApi.unmet({ goal: g, _envelopes: [] });
    assert.ok(open.some(u => u.kind === 'chart'));
    assert.ok(open.some(u => u.kind === 'format'));
    const left = goalApi.unmet({
        goal: g,
        _chartDone: true,
        _envelopes: [{ ok: true, tool: 'excel_format_range' }]
    });
    assert.deepStrictEqual(left, []);
    assert.deepStrictEqual(goalApi.unmet({ goal: goalApi.build('今天天气怎么样'), _envelopes: [] }), []);
});

test('缺参只补必填槽；缺键或没有模型时不执行', async () => {
    const registry = {
        get() {
            return {
                name: 'excel_clean_data',
                inputSchema: {
                    required: ['operations'],
                    properties: {
                        operations: { description: '清洗操作' },
                        range: { description: '区域' }
                    }
                }
            };
        }
    };
    App.ChatV2.llmIo = {
        async askJson() {
            return { json: { operations: ['trim'], range: 'A1:B2', extra: 'no' } };
        }
    };
    const filled = await disp.fillRequiredSlots({ toolName: 'excel_clean_data', message: '清洗这份数据', toolRegistry: registry });
    assert.deepStrictEqual(filled.operations, ['trim']);
    assert.strictEqual(filled.range, 'A1:B2');
    assert.strictEqual(filled.extra, undefined);

    App.ChatV2.llmIo = { async askJson() { return { json: { range: 'A1' } }; } };
    assert.strictEqual(await disp.fillRequiredSlots({ toolName: 'excel_clean_data', message: '清洗', toolRegistry: registry }), null);

    App.ChatV2.llmIo = null;
    assert.strictEqual(await disp.fillRequiredSlots({ toolName: 'excel_clean_data', message: '清洗', toolRegistry: registry }), null);
});

test('invoke 出口带信封', async () => {
    const tr = App.ChatV2.toolRegistry;
    tr.register({
        name: 'probe_envelope_tool',
        displayName: '探针',
        description: 'd',
        source: 'builtin',
        async execute() {
            return { success: true, message: '已加粗', address: 'A1:D1' };
        }
    });
    const res = await tr.invoke('probe_envelope_tool', {}, {});
    assert.strictEqual(res.ok, true);
    assert.ok(res.envelope);
    assert.strictEqual(res.envelope.ok, true);
    assert.strictEqual(res.envelope.value, '已加粗');
    assert.strictEqual(res.envelope.context.address, 'A1:D1');
    assert.strictEqual(res.envelope.error, '');
});
