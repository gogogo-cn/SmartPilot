/**
 * ============================================================
 * Tests/multi-step-tools.test.js — "一次要做很多件事，却只做了一件"的回归
 * ============================================================
 * 实测失败（截图 08:21 / 08:32 两次）：
 *   "美化当前选中的表格区域: 设置表头底色、边框、对齐方式与列宽，使其清晰美观"
 * 只执行 1 个工具 excel_format_range(range=$A$1:$G$1, format={borderStyle,borderColor})，
 * 底色/对齐/列宽一条没做，UI 却显示"任务完成 · 执行 1 个工具"。
 *
 * ── 本文件最重要的一条教训（夹具保真）────────────────────────────
 * 上一轮加了"要求覆盖度自检"，测试全绿，但生产环境**完全没生效**。原因不是逻辑，
 * 而是**测试夹具与生产形状不一致**：
 *   · 测试里手工 `toolRegistry.register({ ..., intentKeywords: [...] })` → 字段在
 *   · 生产里走 `syncFromMcp()`，它只搬 name/displayName/description/inputSchema/isDangerous
 *     → **intentKeywords 被丢弃** → 自检对每个工具都 `kws.length===0` 而 continue
 *     → uncovered 恒为空 → 守卫静默空转；且 ChatLog 只写 console，没有任何日志能暴露。
 * 因此本文件**不再手工注册**，而是走真实链路：
 *   Config/excel_tools.json → app-mcp.js 的 loadMcpToolsForAI（真实 buildMcpToolMeta）
 *   → toolRegistry.syncFromMcp()（真实注册）→ 原样传给 agentLoop.run
 * 并专门断言"注册后 intentKeywords 必须还在"——这条断言就是当初漏掉的那条。
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const { loadChatV2 } = require('./helpers/harness');

const ROOT = path.join(__dirname, '..');
const CFG_PATH = path.join(ROOT, 'Config/excel_tools.json');
const USER_MSG = '美化当前选中的表格区域: 设置表头底色、边框、对齐方式与列宽，使其清晰美观';

const endsWithBase = (n, base) => String(n || '') === base || String(n || '').endsWith('_' + base);

/** 走真实链路装载工具：excel_tools.json → buildMcpToolMeta → syncFromMcp（只把 execute 换成 spy） */
async function setupRealTools(App) {
    const cfg = JSON.parse(fs.readFileSync(CFG_PATH, 'utf8').replace(/^\uFEFF/, ''));
    // 【夹具保真】真实后端（GetAvailableMcpTools）下发的每个工具都带 source:"excel-local"，
    // buildMcpToolMeta 据此拼出 excel-local__xxx 前缀、syncFromMcp 再剥成裸注册名。
    // excel_tools.json 本身无 source 字段，不补的话注册名会变成 undefined_excel_xxx（单下划线），
    // _bareToolName 只剥 '__' 剥不动它，导致 agent-loop 里的工具名比对（含语义等价豁免）全部失配。
    cfg.tools.forEach(t => { if (!t.source) t.source = 'excel-local'; });
    const calls = [];

    // 必须用 app-mcp.js 真身：buildMcpToolMeta 只在这里且未对外暴露。
    // 在测试里手写一遍 meta 形状，等于把这次踩的坑（夹具≠生产）再复制一份。
    if (!App.McpModule) {
        if (typeof globalThis.ChatLog === 'undefined') globalThis.ChatLog = global.window.ChatLog;
        (0, eval)(fs.readFileSync(path.join(ROOT, 'UI/js/app-mcp.js'), 'utf8'));
    }
    const loader = (App.McpModule && App.McpModule.loadMcpToolsForAI) || App.loadMcpToolsForAI;
    assert.strictEqual(typeof loader, 'function', 'app-mcp.js 必须提供 loadMcpToolsForAI');

    const origFetch = globalThis.fetch;
    globalThis.fetch = () => Promise.reject(new Error('test: 无网络'));
    try {
        App.excelBridge = App.excelBridge || {};
        App.excelBridge.GetAvailableMcpTools = async () =>
            JSON.stringify({ success: true, tools: cfg.tools });
        await loader();
        // _readMcpToolSource() 优先读 App.mcpTools；harness 默认给的空数组（真值）会遮住
        // window._mcpTools 导致注册 0 个工具，这里对齐真实来源。
        App.mcpTools = global.window._mcpTools;
        await App.ChatV2.toolRegistry.syncFromMcp();   // ← 此前就是这里把 intentKeywords 丢了
    } finally {
        globalThis.fetch = origFetch;
    }

    assert.ok(Array.isArray(global.window._mcpTools) && global.window._mcpTools.length > 0,
        '真实链路应把 excel_tools.json 的工具装载进 window._mcpTools');

    const all = App.ChatV2.toolRegistry.listBySource('mcp');
    const fmt = all.find(t => endsWithBase(t.name, 'excel_format_range'));
    const fit = all.find(t => endsWithBase(t.name, 'excel_autofit'));
    const tools = [fmt, fit].filter(Boolean);
    tools.forEach(t => {
        t.execute = async (args) => {
            calls.push({ tool: t.name, args });
            return { ok: true, success: true, result: '已处理' };
        };
    });
    return { tools, calls, all, fmtName: fmt && fmt.name, fitName: fit && fit.name };
}

function stubLLM(rounds) {
    let n = 0;
    return {
        stream() {
            return (async function* () {
                const r = rounds[Math.min(n, rounds.length - 1)];
                n++;
                yield { type: 'model', model: 'test' };
                for (const e of r) yield e;
                yield { type: 'done', finishReason: 'stop' };
            })();
        }
    };
}

const ntc = (id, name, args) => ({ type: 'tool_calls', toolCalls: [{ id, function: { name, arguments: JSON.stringify(args) } }] });
const say = (t) => ({ type: 'token', delta: t });

async function run(App, al, tools, msg) {
    const stTurn = App.ChatV2.sessionStore.startTurn({ message: msg, kind: 'text' });
    const events = [];
    for await (const e of al.run({
        turn: stTurn,
        messages: [{ role: 'user', content: msg }],
        tools,
        signal: new AbortController().signal
    })) events.push(e);
    return { turn: stTurn, events };
}

// ── 根因护栏：这条断言就是当初漏掉的那条 ──
test('①根因护栏: 真实注册链路必须把 intentKeywords 保留到工具对象上', async () => {
    const { App } = loadChatV2();
    const { tools, all, fmtName, fitName } = await setupRealTools(App);

    assert.ok(all.length > 5, `真实链路应装载到多个 MCP 工具，实际 ${all.length}`);
    assert.ok(fmtName && fitName, '格式化/自适应列宽两个工具都应装载成功。已注册: ' +
        JSON.stringify(all.map(t => t.name).slice(0, 12)));
    const fmt = tools.find(t => t.name === fmtName);
    const fit = tools.find(t => t.name === fitName);
    assert.ok(Array.isArray(fmt.intentKeywords) && fmt.intentKeywords.length > 0,
        'excel_format_range 注册后必须保留 intentKeywords（丢了 → 覆盖度自检静默空转）');
    assert.ok(Array.isArray(fit.intentKeywords) && fit.intentKeywords.length > 0,
        'excel_autofit 注册后必须保留 intentKeywords（丢了 → 缺项永远不会被发现）');
    assert.ok(fit.intentKeywords.some(k => USER_MSG.includes(k)),
        `autofit 关键词应能命中用户原话: ${JSON.stringify(fit.intentKeywords)}`);
});

// ── 覆盖度自检：完全复刻截图那一轮（真实注册的工具 + 真实 intentKeywords）──
test('②覆盖度自检: 5 条要求只做 1 条并宣称"已完成" → 必须续做并点名缺项', async () => {
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const { tools, calls, fmtName, fitName } = await setupRealTools(App);

    const seen = [];
    let i = 0;
    const rounds = [
        // 第1轮：复刻截图——只调 1 次 format_range，format 里只有边框
        [ntc('n1', fmtName, { range: '$A$1:$G$1', format: { borderStyle: 'thin', borderColor: '000000' } })],
        // 第2轮：完成语气（不含"接下来/我将"）→ 原 INCOMPLETE_SIGNAL 抓不到
        [say('已格式化工作表《测试表》的区域 A1:G1（共 7 个单元格）成功格式化 A1:G1')],
        // 第3轮：被续做引导纠正后补上漏掉的
        [ntc('n2', fitName, { range: '$A$1:$G$201' })],
        [say('底色、边框、对齐、列宽都处理完了。')]
    ];
    App.ChatV2.llmClient = {
        stream({ messages }) {
            seen.push(messages);
            const r = rounds[Math.min(i, rounds.length - 1)];
            i++;
            return (async function* () {
                yield { type: 'model', model: 'test' };
                for (const e of r) yield e;
                yield { type: 'done', finishReason: 'stop' };
            })();
        }
    };

    const { events } = await run(App, al, tools, USER_MSG);

    const nudge = events.find(e => e.type === 'force_retry' && e.reason === 'requirements_not_covered');
    assert.ok(nudge, '只做一条就宣称完成时必须触发覆盖度续做。事件: ' +
        events.map(e => e.type + (e.reason ? ':' + e.reason : '')).join(' | '));
    assert.ok(Array.isArray(nudge.missing) && nudge.missing.some(n => endsWithBase(n, 'excel_autofit')),
        `未完成项应含 excel_autofit，实际 ${JSON.stringify(nudge.missing)}`);

    const hint = (seen[2] || []).filter(m => m.role === 'system').map(m => String(m.content || '')).join('\n');
    assert.ok(/excel_autofit/.test(hint), '续做提示必须点名缺失工具');
    assert.ok(/还没有执行/.test(hint), '续做提示必须说明"还没有执行"');

    assert.ok(calls.some(c => endsWithBase(c.tool, 'excel_autofit')), '续做后 excel_autofit 必须真的被执行');
});

// ── 模型不听引导时：框架不再空耗告警，而是兜底把缺失的安全工具真正执行掉 ──
// 【行为升级】此前 multiStepSafeMatch 因外层门槛 (hasStrongMatch || isFakeExecution) 而永远不可达，
// 导致"模型无视续做引导"时只能零工具收尾；现在弱但安全的匹配也能进入兜底执行，
// 把用户要求里"还没做"的那一项（列宽 → excel_autofit）真正补做掉。
test('③兜底执行: 续做引导后模型仍不做 → 框架强制补做缺失的安全工具', async () => {
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const { tools, calls, fmtName } = await setupRealTools(App);

    App.ChatV2.llmClient = stubLLM([
        [ntc('n1', fmtName, { range: '$A$1:$G$1', format: { borderStyle: 'thin' } })],
        [say('已完成所有美化。')],
        [say('确实都做完了。')],   // ← 无视续做引导
        [say('确实都做完了。')]
    ]);

    await run(App, al, tools, USER_MSG);

    assert.ok(calls.some(c => endsWithBase(c.tool, 'excel_autofit')),
        '模型无视引导时，框架应兜底执行缺失的 excel_autofit。已执行: ' +
        JSON.stringify(calls.map(c => c.tool)));
});

// ── 复刻 23:20 截图：零工具 + "已成功美化…#FFFFCC" → 必须强制调 format_range（空 format→预设）──
test('③b零工具假完成: 已成功美化+色值幻觉 → 强制 excel_format_range 或整段诚实失败', async () => {
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const { tools, calls, fmtName } = await setupRealTools(App);

    const FAKE = '已成功美化选中区域 $A$1:$I$18：\n'
        + '- 背景颜色：设置为浅黄色（#FFFFCC）\n'
        + '- 边框：添加了黑色细线（Border Style: thin）\n'
        + '- 列宽已按表头自动调整';

    App.ChatV2.llmClient = stubLLM([
        [say(FAKE)],
        [say(FAKE)],
        [say(FAKE)]
    ]);

    const stTurn = App.ChatV2.sessionStore.startTurn({ message: USER_MSG, kind: 'text' });
    stTurn.context = { selection: { address: '$A$1:$I$18', range: '$A$1:$I$18' } };
    const events = [];
    for await (const e of al.run({
        turn: stTurn,
        messages: [{ role: 'user', content: USER_MSG }],
        tools,
        signal: new AbortController().signal
    })) events.push(e);

    const forcedFmt = calls.some(c => endsWithBase(c.tool, 'excel_format_range'));
    const honest = /没有调用任何工具|未被修改|实际没有调用/.test(String(stTurn.content || ''));
    assert.ok(forcedFmt || honest,
        `零工具完成腔必须强制执行或诚实失败。forced=${forcedFmt} honest=${honest} calls=${JSON.stringify(calls.map(c => c.tool))} content=${JSON.stringify(stTurn.content)}`);
    if (forcedFmt) {
        const fmtCall = calls.find(c => endsWithBase(c.tool, 'excel_format_range'));
        assert.ok(fmtCall.args && (fmtCall.args.format != null),
            '强制 format_range 必须带 format（空对象即可触发预设）');
    }
    assert.ok(fmtName, 'format_range 应已注册');
});

// ── 无法兜底时（候选为空/必填参数拿不到）：绝不能把"没做"说成"做完了" ──
test('⑤诚实兜底: 零工具执行且无法兜底时 → 正文必须如实告知"未调用任何工具"', async () => {
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const { tools } = await setupRealTools(App);

    // 让后端能力匹配返回空 → 既无引导建议、也无法兜底执行第一名
    if (App.excelBridge) {
        App.excelBridge.MatchCapability = async () => JSON.stringify({ success: true, matches: [] });
    }

    App.ChatV2.llmClient = stubLLM([
        [say('操作完成。结果摘要：改了表头的底色、边框和对齐方式。')],  // 无"已"、无工具名
        [say('操作完成。')],
        [say('操作完成。')]
    ]);

    const { turn } = await run(App, al, tools, '美化当前选中的表格区域：设置表头底色、边框、对齐方式与列宽');

    assert.ok(/没有调用任何工具|未被修改/.test(String(turn.content || '')),
        `零执行且无法兜底时必须如实告知，实际正文: ${JSON.stringify(turn.content)}`);
});

// ── 单步请求不得被误判 ──
test('④不得误伤: 只要求一件事时，不得触发覆盖度续做', async () => {
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const { tools, fmtName } = await setupRealTools(App);

    App.ChatV2.llmClient = stubLLM([
        [ntc('n1', fmtName, { range: '$A$1:$G$1', format: { bold: true } })],
        [say('已把表头加粗。')]
    ]);

    const { events } = await run(App, al, tools, '把表头加粗');

    assert.ok(!events.some(e => e.type === 'force_retry' && e.reason === 'requirements_not_covered'),
        '单步请求不得触发覆盖度续做');
});

// ── 等价完成：实质美化后 format_range 已 AutoFit，不得再误报 excel_autofit 未执行 ──
// 复刻截图：模型调 excel_format_range（底色/边框/字体/对齐齐全），摘要写「列宽自动调整」，
// 但覆盖度仍按工具名要单独的 excel_autofit → 黄条「没有执行：自动调整列宽行高」。
// excel_autofit 工具一直存在；缺的是「format 已等价完成列宽」的豁免。
test('⑥等价完成: format_range 实质美化覆盖列宽 → 不得再报 excel_autofit 未执行', async () => {
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const { tools, calls, fmtName } = await setupRealTools(App);

    App.ChatV2.llmClient = stubLLM([
        [ntc('n1', fmtName, {
            range: '$A$1:$I$18',
            format: {
                bgColor: '#FFC7CE', borderStyle: 'thin', borderColor: '000000',
                fontName: 'Arial', fontSize: 12
            },
            alignment: { horizontal: 'center', vertical: 'middle', wrapText: false }
        })],
        [say('## 格式化结果摘要\n- 列宽自动调整以适应内容\n已完成美化。')]
    ]);

    const { turn, events } = await run(App, al, tools, USER_MSG);

    assert.ok(!events.some(e => e.type === 'force_retry' && e.reason === 'requirements_not_covered'),
        '实质美化后不得因 excel_autofit 触发覆盖度续做。事件: ' +
        events.map(e => e.type + (e.reason ? ':' + e.reason : '')).join(' | '));
    assert.ok(!calls.some(c => endsWithBase(c.tool, 'excel_autofit')),
        '不应再强制补调 excel_autofit。已执行: ' + JSON.stringify(calls.map(c => c.tool)));
    assert.ok(!/excel_autofit|自动调整列宽行高/.test(String(turn.content || '')),
        `正文不得再追加 autofit 未执行告警，实际: ${JSON.stringify(turn.content)}`);
});

// ── 等价完成：写 =SUMIFS 公式一步完成"写入+合计"，不得再报"没有执行" ──
// 复刻 2026-10-02 20:37 截图：模型用 excel_add_formula 在 H2 写入 SUMIFS，
// 覆盖度自检按关键词字面映射（"写入"→write_range、"合计"→aggregate）误判缺项，
// 收尾正文追加"⚠️ 你要求的以下内容本次【没有执行】…请让我继续"——
// 任务明明完成，用户却看到"还有下一步"。
test('⑦等价完成: add_formula 写 SUMIFS 覆盖"写入+合计" → 不得续做、不得追加未执行告警', async () => {
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const { all } = await setupRealTools(App);

    const pick = (base) => all.find(t => endsWithBase(t.name, base));
    const read = pick('excel_read_range');
    const addF = pick('excel_add_formula');
    const write = pick('excel_write_range');
    const agg = pick('excel_aggregate');
    assert.ok(read && addF && write && agg, '四个工具都应已注册');
    const tools = [read, addF, write, agg];
    tools.forEach(t => {
        t.execute = async () => ({ ok: true, success: true, result: '已处理' });
    });

    // 拆分后恰 3 项（≥3 才进自检）；含"写入"与"合计"两个触发词，复刻截图原话结构
    const MSG = '统计销售数据：在H2单元格写入公式，计算商品类别为服饰且区域为华东的订单金额合计';
    let i = 0;
    App.ChatV2.llmClient = {
        stream() {
            const r = [
                // 第1轮：读取 + 写 SUMIFS 公式（复刻截图的两笔工具调用）
                [ntc('n1', read.name, { range: '$A$1:$G$201' }),
                 ntc('n2', addF.name, { cell: 'H2', formula: '=SUMIFS(F:F,E:E,"服饰",C:C,"华东")' })],
                // 第2轮：完成语气收尾
                [say('已成功在 H2 单元格中添加公式 =SUMIFS(F:F,E:E,"服饰",C:C,"华东") 计算订单金额合计。')]
            ][Math.min(i, 1)];
            i++;
            return (async function* () {
                yield { type: 'model', model: 'test' };
                for (const e of r) yield e;
                yield { type: 'done', finishReason: 'stop' };
            })();
        }
    };

    const { turn, events } = await run(App, al, tools, MSG);

    assert.ok(!events.some(e => e.type === 'force_retry' && e.reason === 'requirements_not_covered'),
        '等价完成时不得触发覆盖度续做。事件: ' +
        events.map(e => e.type + (e.reason ? ':' + e.reason : '')).join(' | '));
    assert.ok(!/没有执行/.test(String(turn.content || '')),
        `正文不得再追加"没有执行"告警，实际: ${JSON.stringify(turn.content)}`);
});

// ── 等价完成的边界：多维分组结果一条公式装不下，仍必须点名 excel_aggregate ──
test('⑧不放过真缺项: 要"分组统计"多行结果时，单条公式不得豁免 excel_aggregate', async () => {
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const { all } = await setupRealTools(App);

    const pick = (base) => all.find(t => endsWithBase(t.name, base));
    const read = pick('excel_read_range');
    const addF = pick('excel_add_formula');
    const agg = pick('excel_aggregate');
    const tools = [read, addF, agg].filter(Boolean);
    tools.forEach(t => {
        t.execute = async () => ({ ok: true, success: true, result: '已处理' });
    });

    // 含"分组"→ 多行分组结果不该被一条 SUMIF 豁免
    const MSG = '统计一下：读取表格数据，按部门给出每个部门的销售额合计';
    let i = 0;
    App.ChatV2.llmClient = {
        stream() {
            const r = [
                [ntc('n1', read.name, { range: '$A$1:$G$201' }),
                 ntc('n2', addF.name, { cell: 'I2', formula: '=SUMIF(E:E,"华东",F:F)' })],
                [say('已用 SUMIF 计算销售额合计。')]
            ][Math.min(i, 1)];
            i++;
            return (async function* () {
                yield { type: 'model', model: 'test' };
                for (const e of r) yield e;
                yield { type: 'done', finishReason: 'stop' };
            })();
        }
    };

    const { events } = await run(App, al, tools, MSG);

    const nudge = events.find(e => e.type === 'force_retry' && e.reason === 'requirements_not_covered');
    assert.ok(nudge, '多维分组必须仍触发覆盖度续做。事件: ' +
        events.map(e => e.type + (e.reason ? ':' + e.reason : '')).join(' | '));
    assert.ok(Array.isArray(nudge.missing) && nudge.missing.some(n => endsWithBase(n, 'excel_aggregate')),
        `缺项应含 excel_aggregate，实际 ${JSON.stringify(nudge.missing)}`);
});
