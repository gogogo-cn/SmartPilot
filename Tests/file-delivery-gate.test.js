/**
 * ============================================================
 * Tests/file-delivery-gate.test.js — 文件产物交付闸门
 * ============================================================
 * 实测失败（截图 2026-10-28 19:04 / 19:19，连续两次）：
 *   用户："将当前工作表导出为 CSV 并保存到桌面"
 *   实际：只执行了 excel_save_workbook（存了 .xlsx）；
 *         正文却写"工作表已导出为 CSV 文件并保存到桌面，文件路径为：C:\...\xxx.csv"
 *         —— CSV 从未生成，路径是模型编的（顺序错 + 假成功）。
 *
 * 闸门逻辑：产物类工具是否真的执行过 = 唯一可信证据。
 *   没执行 → 先补做引导（重新下发 tools）；引导用尽 → 当场确定性补执行；
 *   补执行失败 → 删掉编造的路径并如实说明。
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

const EXCEL = { isExcel: true, isPowerPoint: false, kind: 'excel' };
const WORD = { isExcel: false, isPowerPoint: false, kind: 'word' };
const MSG = '将当前工作表导出为 CSV 并保存到桌面';

function gateOf() {
    const { window, App } = loadChatV2(['agent-loop/file-delivery-gate.js']);
    window.SmartPilotHost = EXCEL;
    return { window, App, gate: App.ChatV2.agentLoopFileDeliveryGate };
}

// ───────────────────────── 纯函数 ─────────────────────────
test('detect：导出 CSV → 识别为 csv 产物（工具 excel_export_csv）', () => {
    const { gate } = gateOf();
    const spec = gate.detect(MSG, EXCEL);
    assert.ok(spec, '必须识别为产物型请求');
    assert.strictEqual(spec.tool, 'excel_export_csv');
});

test('detect："导出为 CSV 格式" 里的"格式"不得被当成排版意图排除', () => {
    const { gate } = gateOf();
    assert.ok(gate.detect('把这张表导出为 CSV 格式', EXCEL));
});

test('detect：询问语气（怎么导出 CSV？）不是命令，不触发', () => {
    const { gate } = gateOf();
    assert.strictEqual(gate.detect('怎么导出 CSV？', EXCEL), null);
});

test('detect：多表批量导出交规划，不由闸门代劳', () => {
    const { gate } = gateOf();
    assert.strictEqual(gate.detect('把所有工作表都导出成 CSV', EXCEL), null);
});

test('detect：Word 宿主不出现在 Excel 产物闸门里', () => {
    const { gate } = gateOf();
    assert.strictEqual(gate.detect(MSG, WORD), null);
});

test('isDelivered：只执行过 excel_save_workbook 不算已交付', () => {
    const { gate } = gateOf();
    const spec = gate.detect(MSG, EXCEL);
    const turn = { _executedTools: ['excel_save_workbook'] };
    assert.strictEqual(gate.isDelivered(turn, spec), false, '存工作簿 ≠ 导出 CSV');
    turn._executedTools.push('excel_export_csv');
    assert.strictEqual(gate.isDelivered(turn, spec), true);
});

test('isDelivered：_executedToolCalls 里记录过也算已交付', () => {
    const { gate } = gateOf();
    const spec = gate.detect(MSG, EXCEL);
    const turn = { _executedTools: [], _executedToolCalls: [{ name: 'excel_export_csv', args: {} }] };
    assert.strictEqual(gate.isDelivered(turn, spec), true);
});

test('scrubClaims：删掉编造的"已导出 + csv 路径"，保留其它内容', () => {
    const { gate } = gateOf();
    const spec = gate.detect(MSG, EXCEL);
    const fake = '已先保存工作簿以便后续处理。\n'
        + '工作表已导出为 CSV 文件并保存到桌面，文件路径为：C:\\Users\\Administrator\\Desktop\\demo_test_data_销售数据.csv\n'
        + '如需继续整理请告诉我。';
    const out = gate.scrubClaims(fake, spec);
    assert.ok(!/\.csv/i.test(out), `不得残留编造的 csv 路径，实际: ${out}`);
    assert.ok(out.includes('如需继续整理'), '无关内容不能被误删');
    assert.ok(out.includes('已先保存工作簿'), '不涉及产物扩展名的句子应保留');
});

test('buildNudge：点名产物工具、禁止用保存工作簿冒充、禁止编造结论', () => {
    const { gate } = gateOf();
    const spec = gate.detect(MSG, EXCEL);
    const { text } = gate.buildNudge(spec, { _executedTools: ['excel_save_workbook'] });
    assert.ok(text.includes('excel_export_csv'), '必须点名要调的工具');
    assert.ok(/excel_save_workbook/.test(text), '必须明确 save_workbook 不能代替导出');
    assert.ok(/不许写/.test(text), '必须禁止编造"已导出/文件路径"');
});

// ───────────────────── agent-loop 集成 ─────────────────────
function stubLLM(scn, captured) {
    return {
        stream(opts) {
            if (captured) captured.push({ messages: opts.messages, tools: opts.tools || [] });
            return (async function* () {
                yield { type: 'model', model: 'test' };
                const step = scn[Math.min(captured ? captured.length - 1 : 0, scn.length - 1)] || scn[scn.length - 1];
                if (Array.isArray(step)) for (const e of step) yield e;
                yield { type: 'done', finishReason: 'stop' };
            })();
        }
    };
}
function toolCallsEvt(calls) {
    return [{ type: 'tool_calls', toolCalls: calls.map(c => ({ id: 'c_' + c.name, function: { name: c.name, arguments: JSON.stringify(c.args || {}) } })) }];
}
function textEvt(t) { return [{ type: 'token', delta: t }]; }
async function collectEvents(al, runOpts) {
    const out = [];
    for await (const e of al.run(runOpts)) out.push(e);
    return out;
}

const FAKE_CLAIM = '工作表已导出为 CSV 文件并保存到桌面，文件路径为：C:\\Users\\Administrator\\Desktop\\demo_test_data_销售数据.csv';
const REAL_PATH = 'C:\\Users\\test\\Desktop\\demo_test_data_销售数据.csv';

function setupTools(App, rec) {
    const tr = App.ChatV2.toolRegistry;
    tr.register({
        name: 'excel_save_workbook', displayName: '保存工作簿', description: '保存工作簿', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute() { rec.saved++; return { ok: true, result: { success: true, path: 'C:\\Users\\test\\Desktop\\demo.xlsx' } }; }
    });
    tr.register({
        name: 'excel_export_csv', displayName: '导出CSV', description: '将工作表导出为 CSV', source: 'mcp',
        inputSchema: { type: 'object', properties: { filePath: { type: 'string' }, sheet: { type: 'string' } } },
        async execute(args) { rec.csv++; rec.csvArgs = args; return { ok: true, result: { success: true, path: REAL_PATH } }; }
    });
    App.excelBridge = { GetDesktopPath: async () => 'C:\\Users\\test\\Desktop' };
    return tr;
}

test('集成：模型只存了 xlsx 却宣称 CSV 已导出 → 闸门补做引导后真导出', async () => {
    const { window, App } = loadChatV2();
    window.SmartPilotHost = EXCEL;
    const al = App.ChatV2.agentLoop;
    const rec = { saved: 0, csv: 0, csvArgs: null };
    const tr = setupTools(App, rec);

    const scn = [
        toolCallsEvt([{ name: 'excel_save_workbook' }]),  // 第1轮：顺序错（先存工作簿）
        textEvt(FAKE_CLAIM),                               // 第2轮：假结论（没有点名工具，旧守卫抓不到）
        toolCallsEvt([{ name: 'excel_export_csv' }]),      // 第3轮：被闸门纠正后真导出
        textEvt('已导出完成。')
    ];
    const captured = [];
    App.ChatV2.llmClient = stubLLM(scn, captured);
    const stTurn = App.ChatV2.sessionStore.startTurn({ message: MSG, kind: 'text' });
    const events = await collectEvents(al, {
        turn: stTurn,
        messages: [{ role: 'user', content: MSG }],
        tools: [tr.get('excel_save_workbook'), tr.get('excel_export_csv')],
        signal: new AbortController().signal
    });

    const retry = events.find(e => e.type === 'force_retry' && e.reason === 'file_product_not_delivered');
    assert.ok(retry, `假结论必须被闸门拦下并要求补做，实际: ${events.map(e => e.type).join('|')}`);
    assert.strictEqual(rec.csv, 1, '最终必须真正执行一次 excel_export_csv');
    assert.ok(events.some(e => e.type === 'tool_call_result' && e.toolName === 'excel_export_csv'));
});

test('集成：模型始终不导出 → 闸门确定性补执行，正文用真实路径覆盖假路径', async () => {
    const { window, App } = loadChatV2();
    window.SmartPilotHost = EXCEL;
    const al = App.ChatV2.agentLoop;
    const rec = { saved: 0, csv: 0, csvArgs: null };
    const tr = setupTools(App, rec);

    // 引导一轮后模型仍然只写字 → 闸门兜底补执行
    const scn = [
        toolCallsEvt([{ name: 'excel_save_workbook' }]),
        textEvt(FAKE_CLAIM),
        textEvt(FAKE_CLAIM)
    ];
    const captured = [];
    App.ChatV2.llmClient = stubLLM(scn, captured);
    const stTurn = App.ChatV2.sessionStore.startTurn({ message: MSG, kind: 'text' });
    const events = await collectEvents(al, {
        turn: stTurn,
        messages: [{ role: 'user', content: MSG }],
        tools: [tr.get('excel_save_workbook'), tr.get('excel_export_csv')],
        signal: new AbortController().signal
    });

    assert.strictEqual(rec.csv, 1, '模型不听话时框架必须自己把 CSV 导出来');
    const resultEvt = events.find(e => e.type === 'tool_call_result' && e.toolName === 'excel_export_csv');
    assert.ok(resultEvt && resultEvt.ok, '补执行的工具结果必须是成功');

    const finalText = String(stTurn.content || '');
    assert.ok(finalText.includes(REAL_PATH), `正文必须给出真实路径，实际: ${finalText}`);
    assert.ok(!finalText.includes('C:\\Users\\Administrator\\Desktop\\demo_test_data_销售数据.csv'),
        `编造的路径必须被删除，实际: ${finalText}`);
});
