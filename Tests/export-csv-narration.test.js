/**
 * ============================================================
 * Tests/export-csv-narration.test.js — "承诺重试死锁"回归
 * ============================================================
 * 实测失败（截图 2026-10-27 19:34）：用户请求"将当前工作表导出为 CSV 并保存到桌面"，
 * 模型第 1 轮调了 excel_save_workbook（成功），第 2 轮（总结轮，模型手里已没有工具）
 * 输出"我将直接调用 excel_export_csv 工具…"的文字计划，任务就此结束——
 * CSV 永远没导出，UI 显示"任务完成 执行 1 个工具"。
 *
 * 死锁链（修复前）：
 *  ① 第 1 轮工具成功 → toolsExecuted=true → 第 2 轮 shouldProvideTools=false（总结轮）
 *  ② 模型在总结轮点名 excel_export_csv → 承诺未执行守卫命中 → 引导重试
 *  ③ 但重试轮 toolsExecuted 仍为 true → 模型手里依然没有工具 → 只能继续干说
 *  ④ 2 次重试全部空转 → 把"永不执行的计划"展示给用户
 *
 * 修复：a) 三处重试分支 continue 前重置 toolsExecuted=false（重试轮重新下发 tools）；
 *      b) 点名工具本轮不可见但从注册表能找到 → 注入本轮工具列表再重试；
 *      c) 强制兜底优先执行模型点名的工具；缺参但有安全默认生成器（filePath→桌面）
 *         不再放弃兜底。
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

/** LLM 桩：捕获每轮请求的 messages 与 tools */
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

async function collectEvents(al, runOpts) {
    const out = [];
    for await (const e of al.run(runOpts)) out.push(e);
    return out;
}

function toolCallsEvt(calls) {
    return [
        { type: 'tool_calls', toolCalls: calls.map(c => ({ id: 'c_' + c.name, function: { name: c.name, arguments: JSON.stringify(c.args || {}) } })) }
    ];
}
function textEvt(t) {
    return [{ type: 'token', delta: t }];
}

test('死锁回归: 已执行过工具后模型点名 excel_export_csv → 重试轮 tools 必须重新下发且真执行', async () => {
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const tr = App.ChatV2.toolRegistry;

    let saveInvoked = 0, csvArgs = null;
    tr.register({
        name: 'excel_save_workbook', displayName: '保存工作簿', description: '保存', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute() { saveInvoked++; return { ok: true, result: { success: true } }; }
    });
    tr.register({
        name: 'excel_export_csv', displayName: '导出CSV', description: '将工作表导出为 CSV', source: 'mcp',
        inputSchema: { type: 'object', required: ['filePath'], properties: { filePath: { type: 'string' } } },
        async execute(args) { csvArgs = args; return { ok: true, result: { success: true } }; }
    });
    // inferExportFilePath 的桌面路径桥桩（与 tool-required-args.test.js 同款）
    App.excelBridge = { GetDesktopPath: async () => 'C:\\Users\\test\\Desktop' };

    const NARRATION = '工具未能准确完成预期任务。根据提供的信息，我将直接调用 excel_export_csv 工具'
        + '将当前工作表导出为 CSV 文件并保存到桌面。';
    const scn = [
        toolCallsEvt([{ name: 'excel_save_workbook' }]),          // 第1轮：先保存工作簿
        textEvt(NARRATION),                                        // 第2轮：点名 export_csv 但没调
        toolCallsEvt([{ name: 'excel_export_csv' }]),             // 第3轮：被纠正后真调（不带 filePath）
        textEvt('已导出完成。')                                     // 第4轮：收尾
    ];
    const captured = [];
    App.ChatV2.llmClient = stubLLM(scn, captured);

    const saveTool = tr.get('excel_save_workbook');
    const csvTool = tr.get('excel_export_csv');
    const stTurn = App.ChatV2.sessionStore.startTurn({
        message: '将当前工作表导出为 CSV 并保存到桌面', kind: 'text'
    });
    const events = await collectEvents(al, {
        turn: stTurn,
        messages: [{ role: 'user', content: '将当前工作表导出为 CSV 并保存到桌面' }],
        tools: [saveTool, csvTool],
        signal: new AbortController().signal
    });

    // ① 承诺未执行守卫命中
    const retry = events.find(e => e.type === 'force_retry' && e.reason === 'named_tool_commitment_without_call');
    assert.ok(retry, `应触发承诺未执行重试，实际: ${events.map(e => e.type).join('|')}`);

    // ② 死锁修复核心：重试轮必须重新下发 tools（修复前 toolsExecuted 未重置 → tools=[]）
    // captured[0]=第1轮（调 save）、captured[1]=第2轮总结轮（本来就不下发 tools）、
    // captured[2]=第3轮重试轮（修复后必须重新带上 tools）
    assert.ok(captured.length >= 3, `至少 3 次 LLM 请求，实际 ${captured.length}`);
    assert.ok(captured[2].tools.length > 0,
        `重试轮必须重新下发 tools（修复前为 0），实际 ${captured[2].tools.length}`);
    assert.ok(captured[2].tools.some(t => String(t.name || '').toLowerCase() === 'excel_export_csv'),
        '重试轮 tools 里必须包含被点名的 excel_export_csv');

    // ③ 重试后 export_csv 真正执行，且缺省 filePath 自动落桌面
    const resultEvt = events.find(e => e.type === 'tool_call_result' && e.toolName === 'excel_export_csv');
    assert.ok(resultEvt, '重试后必须真正执行 excel_export_csv');
    assert.strictEqual(saveInvoked, 1, 'excel_save_workbook 恰好执行一次');
    assert.ok(csvArgs && typeof csvArgs.filePath === 'string' && csvArgs.filePath.startsWith('C:\\Users\\test\\Desktop\\'),
        `filePath 应自动补全到桌面，实际: ${csvArgs && JSON.stringify(csvArgs)}`);
});

test('死锁回归(复数句式): "现在我们将…调用 `excel_export_csv`" 同样要触发重试并真导出', async () => {
    // 实测截图 2026-10-28 19:04：模型总结轮写的是"现在我们将当前工作表导出为 CSV 文件，
    // 并保存到桌面。\n\n调用 `excel_export_csv` 工具导出 CSV 文件。"——没有"我将"，
    // 旧 COMMIT_RE 只认单数"我将/我会" → 整句被判成非承诺 → 静默收尾，CSV 永不导出。
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const tr = App.ChatV2.toolRegistry;

    let csvArgs = null;
    tr.register({
        name: 'excel_save_workbook', displayName: '保存工作簿', description: '保存', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute() { return { ok: true, result: { success: true } }; }
    });
    tr.register({
        name: 'excel_export_csv', displayName: '导出CSV', description: '将工作表导出为 CSV', source: 'mcp',
        inputSchema: { type: 'object', properties: { filePath: { type: 'string' } } },
        async execute(args) { csvArgs = args; return { ok: true, result: { success: true } }; }
    });
    App.excelBridge = { GetDesktopPath: async () => 'C:\\Users\\test\\Desktop' };

    const NARRATION = '现在我们将当前工作表导出为 CSV 文件，并保存到桌面。\n\n'
        + '调用 `excel_export_csv` 工具导出 CSV 文件。';
    const scn = [
        toolCallsEvt([{ name: 'excel_save_workbook' }]),
        textEvt(NARRATION),
        toolCallsEvt([{ name: 'excel_export_csv' }]),
        textEvt('已导出完成。')
    ];
    const captured = [];
    App.ChatV2.llmClient = stubLLM(scn, captured);

    const stTurn = App.ChatV2.sessionStore.startTurn({
        message: '将当前工作表导出为 CSV 并保存到桌面', kind: 'text'
    });
    const events = await collectEvents(al, {
        turn: stTurn,
        messages: [{ role: 'user', content: '将当前工作表导出为 CSV 并保存到桌面' }],
        tools: [tr.get('excel_save_workbook'), tr.get('excel_export_csv')],
        signal: new AbortController().signal
    });

    const retry = events.find(e => e.type === 'force_retry' && e.reason === 'named_tool_commitment_without_call');
    assert.ok(retry, `"我们将…" 句式必须被识别为承诺未执行，实际: ${events.map(e => e.type).join('|')}`);
    const resultEvt = events.find(e => e.type === 'tool_call_result' && e.toolName === 'excel_export_csv');
    assert.ok(resultEvt, '重试后必须真正执行 excel_export_csv（不能只写计划）');
    // filePath 在真实 schema 里是可选参数：留空由后端落盘桌面（工作表名_时间戳.csv），
    // 这里只要求若被补全则必须是真实桌面路径（不得出现幻觉路径）。
    if (csvArgs && csvArgs.filePath) {
        assert.ok(String(csvArgs.filePath).startsWith('C:\\Users\\test\\Desktop\\'),
            `filePath 若被补全必须落在桌面，实际: ${JSON.stringify(csvArgs)}`);
    }
});

test('点名工具本轮不可见但在注册表 → 注入工具列表并重试执行', async () => {
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const tr = App.ChatV2.toolRegistry;

    let csvArgs = null;
    tr.register({
        name: 'excel_export_csv', displayName: '导出CSV', description: '导出 CSV', source: 'mcp',
        inputSchema: { type: 'object', properties: { filePath: { type: 'string' } } },
        async execute(args) { csvArgs = args; return { ok: true, result: { success: true } }; }
    });
    // 注册表里【没有】的工具（衬托"注册表没有→如实告知"的旧路径不受影响）
    const readTool = { name: 'excel_read_range', originalName: 'excel_read_range', description: '读数据', inputSchema: { type: 'object', properties: {} } };

    const scn = [
        textEvt('接下来我将使用 excel_export_csv 工具将数据导出。'),   // 点名（本轮 tools 里没有）
        toolCallsEvt([{ name: 'excel_export_csv', args: { filePath: 'C:\\out\\data.csv' } }]),
        textEvt('完成。')
    ];
    const captured = [];
    App.ChatV2.llmClient = stubLLM(scn, captured);

    const stTurn = App.ChatV2.sessionStore.startTurn({ message: '导出数据', kind: 'text' });
    const events = await collectEvents(al, {
        turn: stTurn,
        messages: [{ role: 'user', content: '导出数据' }],
        tools: [readTool],
        signal: new AbortController().signal
    });

    const retry = events.find(e => e.type === 'force_retry' && e.reason === 'named_tool_commitment_without_call');
    assert.ok(retry, '应触发承诺未执行重试');
    // 重试轮 tools 必须包含被注入的 excel_export_csv
    assert.ok(captured[1] && captured[1].tools.some(t => String(t.name || '').toLowerCase() === 'excel_export_csv'),
        '注册表中存在的点名工具必须被注入重试轮工具列表');
    assert.ok(csvArgs && csvArgs.filePath === 'C:\\out\\data.csv', '注入后工具可被真正调用');
});

test('hasSafeDefault: export_csv.filePath 有安全默认，write_range.values 没有', () => {
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;
    assert.strictEqual(typeof tr.hasSafeDefault, 'function', 'registry.hasSafeDefault 应存在');
    assert.strictEqual(tr.hasSafeDefault('excel_export_csv', 'filePath'), true, 'CSV 导出路径应有安全默认');
    assert.strictEqual(tr.hasSafeDefault('excel_export_csv', 'file_path'), true, '别名键同样应有安全默认');
    assert.strictEqual(tr.hasSafeDefault('excel_write_range', 'values'), false, '数据参数不得盲目造值');
    assert.strictEqual(tr.hasSafeDefault('excel_not_exist_tool', 'x'), false, '未注册工具返回 false');
});
