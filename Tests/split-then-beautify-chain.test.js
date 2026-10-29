/**
 * ============================================================
 * Tests/split-then-beautify-chain.test.js —
 *   "按X拆分成不同Sheet，然后分别对工作表美化"必须两项任务都做
 * ============================================================
 * 背景（实测 2026-10-22，用户截图）：
 *   用户："按销售员拆分成不同Sheet,然后分别对工作表美化"
 *   实际：只执行了 excel_split_by_column 一个工具就"任务完成"——
 *         9 个新工作表建出来了，但"然后分别对工作表美化"被整个吞掉。
 *
 * 根因：mcp-dispatcher 的拆分硬短路只判"拆分动词 + 工作表对象 + 按X列名"，
 * 命中后确定性直调拆分并直接收尾；句尾的第二项任务（美化）既没有进链，
 * 也不在 _splitNon 排除表里 → 被硬短路独吞。agent-loop 的要求覆盖度自检
 * （turnGoal.unmet / detectUncoveredRequirements）根本没机会跑——dispatcher
 * 路径不经过 agent-loop。
 *
 * 本测试锁死四件事：
 *   ① 复合消息（拆分+美化）→ 拆分成功后按产物表逐表确定性美化
 *      （每表：get_used_range 取显式范围 → format_range 商务蓝预设）；
 *   ② 纯拆分消息 → 不链美化（旧行为不回退）；
 *   ③ 句尾挂其它任务（排序/筛选/导出…）或美化要求确定性覆盖不全（行高/对齐/底纹/数据条…）
 *      → 不硬短路独吞，放行 agent-loop 多步处理；
 *   ④ 单表失败/取不到范围 → 失败点名（beautifyFailed + steps 注记），不伪造"全部完成"。
 *
 * 运行：node --test Tests/split-then-beautify-chain.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2, DEFAULT_ORDER } = require('./helpers/harness');

const { window, App } = loadChatV2(DEFAULT_ORDER.concat(['services/mcp-dispatcher.js']));
const dispatcher = App.ChatV2.mcpDispatcher;

// 测试环境无 LLM：清空 config 让 refineArgsFromSemantics 立即返回种子参数（不发 fetch、不挂起）
App.config = {};

// ── 桥桩：按工具名回 canned 结果，并记录全部调用 ──
function installBridge(handlers) {
    const calls = [];
    window.excelBridge = {
        CallLocalTool: async (tool, argsJson) => {
            let args = {};
            try { args = typeof argsJson === 'string' ? JSON.parse(argsJson) : (argsJson || {}); } catch (_) { args = {}; }
            calls.push({ tool: tool, args: args });
            const h = handlers[tool];
            const payload = (typeof h === 'function') ? h(args, calls)
                : { success: false, isError: true, result: '测试桥未实现 ' + tool };
            return JSON.stringify(payload);
        }
    };
    return calls;
}

const CREATED = ['赵六', '李四', '张三'];

const splitOk = (created) => ({
    success: true,
    groups: created.length,
    created: created,
    skipped: [],
    replaced: [],
    summary: '新建 ' + created.length + ' 个工作表：' + created.join('、'),
    steps: created.map(n => "新建工作表 '" + n + "'，写入表头 + 20 行数据")
});

/** 默认 handlers：拆分成功 + 每表取范围/美化均成功 */
function defaultHandlers(created) {
    return {
        'excel_split_by_column': () => splitOk(created),
        'excel_get_used_range': (args) => ({ success: true, range: 'A1:G21', sheet: args.sheet }),
        'excel_format_range': () => ({ success: true, isError: false, result: '已格式化' })
    };
}

const opts = (message) => ({ message: message, toolRegistry: null, sessionStore: null, turn: { context: {} }, allowHardDispatch: true });
const callsOf = (calls, tool) => calls.filter(c => c.tool === tool);

// ══ ① 核心用例：拆分 + 逐表美化两项都执行 ══
test('"按销售员拆分成不同Sheet,然后分别对工作表美化" → 拆分与逐表美化都真实执行', async () => {
    const calls = installBridge(defaultHandlers(CREATED));
    const d = await dispatcher.tryDispatch(opts('按销售员拆分成不同Sheet,然后分别对工作表美化'));

    assert.strictEqual(d.hardMatched, true, '应走确定性硬短路');
    assert.strictEqual(d.toolName, 'excel_split_by_column');
    assert.strictEqual(d.args.column, '销售员');
    assert.strictEqual(d.args.outputMode, 'sheets');

    // 拆分只调一次（不重复拆分）
    assert.strictEqual(callsOf(calls, 'excel_split_by_column').length, 1);

    // 每张新表都走了 get_used_range + format_range
    assert.strictEqual(callsOf(calls, 'excel_get_used_range').length, CREATED.length, '每张新表都应先取已用范围');
    const fmts = callsOf(calls, 'excel_format_range');
    assert.strictEqual(fmts.length, CREATED.length, '每张新表都应美化一次（旧版这里为 0）');
    assert.deepStrictEqual(fmts.map(c => c.args.sheet), CREATED);
    for (const f of fmts) {
        assert.strictEqual(f.args.range, 'A1:G21',
            '必须传显式 range：空 range 会走后端"当前选区"兜底，把用户停在别的表上的选区格式化掉');
        assert.strictEqual(f.args.format, '{}', '纯"美化"→ 空 format 走后端商务蓝预设');
    }

    // 链式步骤面板：1 拆分 + N 美化，全部 done
    const steps = d.execResult.steps;
    assert.ok(Array.isArray(steps), 'execResult.steps 应承载链式步骤');
    assert.strictEqual(steps.length, 1 + CREATED.length);
    assert.strictEqual(steps[0].title, '按列拆分工作表');
    assert.strictEqual(steps[1].title, '美化「赵六」');
    assert.ok(steps.every(s => s.status === 'done'),
        '全部步骤应为 done，实际: ' + JSON.stringify(steps.map(s => s.status)));

    // 结果字段：美化成功名单 + steps 末尾注记（orchestrator 的产物摘要会渲染 steps）
    const r = d.execResult.result;
    assert.deepStrictEqual(r.beautified, CREATED);
    assert.strictEqual(r.beautifyFailed, undefined, '无失败时不得有 beautifyFailed');
    assert.ok(/商务蓝美化预设/.test(r.steps[r.steps.length - 1]),
        'steps 末尾应有美化注记: ' + r.steps[r.steps.length - 1]);
});

// ══ ② 纯拆分行为不回退 ══
test('纯拆分消息（无句尾美化）→ 不链美化，与旧行为一致', async () => {
    const calls = installBridge(defaultHandlers(CREATED));
    const d = await dispatcher.tryDispatch(opts('按销售员拆分成不同Sheet'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(d.toolName, 'excel_split_by_column');
    assert.strictEqual(callsOf(calls, 'excel_format_range').length, 0);
    assert.strictEqual(callsOf(calls, 'excel_get_used_range').length, 0);
    assert.strictEqual(d.execResult.result.beautified, undefined);
});

test('"不要美化"否定豁免 → 只拆分，不链美化', async () => {
    const calls = installBridge(defaultHandlers(CREATED));
    const d = await dispatcher.tryDispatch(opts('按销售员拆分成不同Sheet，不要美化'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(callsOf(calls, 'excel_format_range').length, 0);
});

// ══ ③ 具体样式与覆盖度 ══
test('用户点名具体样式（加粗）→ 透传为 format 参数，不吃默认预设', async () => {
    const calls = installBridge(defaultHandlers(['赵六']));
    const d = await dispatcher.tryDispatch(opts('按销售员拆分成不同Sheet，然后把每个sheet加粗'));
    assert.strictEqual(d.hardMatched, true);
    const fmts = callsOf(calls, 'excel_format_range');
    assert.strictEqual(fmts.length, 1);
    assert.strictEqual(JSON.parse(fmts[0].args.format).bold, true);
    assert.deepStrictEqual(d.execResult.result.beautified, ['赵六']);
});

test('美化要求确定性覆盖不全（加粗+行高）→ 放弃硬短路放行 agent-loop，不硬调半套美化', async () => {
    const calls = installBridge(defaultHandlers(CREATED));
    const d = await dispatcher.tryDispatch(opts('按销售员拆分成不同Sheet，然后给每个表加粗并设置行高'));
    assert.strictEqual(d.hit, false, '覆盖不全应放行通用管道（与格式化硬短路"多要求排他"同范式）');
    assert.strictEqual(calls.length, 0, '放行时不应硬调任何工具');
});

test('条件格式类诉求（数据条）确定性链盖不住 → 放行 agent-loop', async () => {
    const calls = installBridge(defaultHandlers(CREATED));
    const d = await dispatcher.tryDispatch(opts('按销售员拆分成不同Sheet，然后给每个表加数据条'));
    assert.strictEqual(d.hit, false);
    assert.strictEqual(calls.length, 0);
});

// ══ ④ 其它复合第二任务不被吞 ══
test('句尾挂其它任务（按金额排序）→ 不放行拆分硬短路独吞，交给 agent-loop 规划', async () => {
    const calls = installBridge(defaultHandlers(CREATED));
    const d = await dispatcher.tryDispatch(opts('按销售员拆分成不同Sheet，然后按金额排序'));
    assert.strictEqual(d.hit, false, '第二项任务无确定性链 → 必须放行');
    assert.strictEqual(callsOf(calls, 'excel_split_by_column').length, 0, '不得只跑拆分');
});

test('句尾挂导出（另存文件）→ 同样放行，不被拆分独吞', async () => {
    const calls = installBridge(defaultHandlers(CREATED));
    const d = await dispatcher.tryDispatch(opts('按销售员拆分成不同Sheet，然后分别导出成文件'));
    assert.strictEqual(d.hit, false);
    assert.strictEqual(calls.length, 0);
});

// ══ ⑤ 边界：无产物表 / 覆盖模式 / 单表失败 / 取不到范围 ══
test('拆分全部跳过（无新建表）→ 无链可执行，不伪造美化', async () => {
    const calls = installBridge({
        'excel_split_by_column': () => ({
            success: true, created: [], skipped: CREATED, replaced: [],
            summary: '拆分表已存在（' + CREATED.join('、') + '），未重复创建', steps: []
        })
    });
    const d = await dispatcher.tryDispatch(opts('按销售员拆分成不同Sheet,然后分别对工作表美化'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(callsOf(calls, 'excel_format_range').length, 0);
    assert.strictEqual(d.execResult.result.beautified, undefined);
});

test('覆盖模式（重新拆分）→ 对 replaced 表逐表美化', async () => {
    const calls = installBridge({
        'excel_split_by_column': () => ({
            success: true, created: [], skipped: [], replaced: CREATED,
            summary: '覆写 ' + CREATED.length + ' 个工作表', steps: []
        }),
        'excel_get_used_range': (args) => ({ success: true, range: 'A1:F30', sheet: args.sheet }),
        'excel_format_range': () => ({ success: true, isError: false })
    });
    const d = await dispatcher.tryDispatch(opts('重新按销售员拆分成不同Sheet,然后分别对工作表美化'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(d.args.onConflict, 'replace');
    assert.strictEqual(callsOf(calls, 'excel_format_range').length, CREATED.length);
    assert.deepStrictEqual(d.execResult.result.beautified, CREATED);
});

test('单表美化失败 → 其余表继续，失败点名进 beautifyFailed 与 steps 注记（不伪造全部完成）', async () => {
    const calls = installBridge({
        'excel_split_by_column': () => splitOk(CREATED),
        'excel_get_used_range': (args) => ({ success: true, range: 'A1:G21', sheet: args.sheet }),
        'excel_format_range': (args) => (args.sheet === '李四'
            ? { success: false, isError: true, result: '找不到工作表' }
            : { success: true, isError: false })
    });
    const d = await dispatcher.tryDispatch(opts('按销售员拆分成不同Sheet,然后分别对工作表美化'));
    const r = d.execResult.result;
    assert.deepStrictEqual(r.beautified, ['赵六', '张三']);
    assert.deepStrictEqual(r.beautifyFailed, ['李四']);
    assert.ok(/失败 1 个（李四:/.test(r.steps[r.steps.length - 1]),
        'steps 注记必须点名失败表: ' + r.steps[r.steps.length - 1]);
    const steps = d.execResult.steps;
    assert.strictEqual(steps[2].status, 'failed', '失败表对应步骤应标红');
    assert.strictEqual(steps[1].status, 'done');
    assert.strictEqual(steps[3].status, 'done', '失败不影响后续表继续');
});

test('get_used_range 拿不到范围 → 该表不美化，也不空 range 硬调（防格式化到别的表）', async () => {
    const calls = installBridge({
        'excel_split_by_column': () => splitOk(['赵六']),
        'excel_get_used_range': () => ({ success: false, isError: true, result: '找不到工作表' }),
        'excel_format_range': () => ({ success: true, isError: false })
    });
    const d = await dispatcher.tryDispatch(opts('按销售员拆分成不同Sheet,然后分别对工作表美化'));
    assert.strictEqual(callsOf(calls, 'excel_format_range').length, 0,
        '取不到显式范围时不得调 format_range（空 range 会走后端选区兜底做错表）');
    assert.deepStrictEqual(d.execResult.result.beautifyFailed, ['赵六']);
});

test('拆分产物表 UsedRange 陈旧（G30 单格）→ 降级 read_range 解析后再美化，不拿单格硬调', async () => {
    const calls = installBridge({
        'excel_split_by_column': () => splitOk(['赵六', '李四']),
        'excel_get_used_range': () => ({ success: true, range: 'G30' }),   // 未激活新表的陈旧单格
        'excel_read_range': (args) => ({
            success: true, isError: false,
            result: JSON.stringify({ worksheet: args.sheet, range: '$A$1:$G$21', rows: 20, columns: 7 })
        }),
        'excel_format_range': () => ({ success: true, isError: false })
    });
    const d = await dispatcher.tryDispatch(opts('按销售员拆分成不同Sheet,然后分别对工作表美化'));
    const fmts = callsOf(calls, 'excel_format_range');
    assert.strictEqual(fmts.length, 2);
    assert.deepStrictEqual(fmts.map(c => c.args.range), ['A1:G21', 'A1:G21'],
        '陈旧单格必须降级解析，不得原样拿去格式化（旧版只美化 1 格还报成功）');
    assert.deepStrictEqual(d.execResult.result.beautified, ['赵六', '李四']);
});

// ══ ⑥ 其它宿主不误伤 ══
test('Word/PPT 宿主不走 Excel 拆分链（宿主守卫不回归）', async () => {
    const prev = window.SmartPilotHost;
    window.SmartPilotHost = { isExcel: false, isWord: true, kind: 'word' };
    try {
        const calls = installBridge(defaultHandlers(CREATED));
        const d = await dispatcher.tryDispatch(opts('按销售员拆分成不同Sheet,然后分别对工作表美化'));
        assert.notStrictEqual(d.toolName, 'excel_split_by_column');
        assert.strictEqual(calls.length, 0);
    } finally {
        window.SmartPilotHost = prev;
    }
});
