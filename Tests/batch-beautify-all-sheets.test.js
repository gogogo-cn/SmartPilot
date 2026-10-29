/**
 * ============================================================
 * Tests/batch-beautify-all-sheets.test.js —
 *   "美化工作簿中的所有工作表"必须真的逐表美化，不得只做 1 张表 + 过程语言收尾
 * ============================================================
 * 背景（实测 2026-10-22，用户截图两连）：
 *   ① "美化工作簿中的所有工作表" → 只执行 1 个工具 excel_format_range（range 还填成 A1，
 *      只美化 1 个单元格），随后"正在依次应用相同预设至其他工作表，请稍候。"，回合以
 *      "任务完成"收场，正文还挂着"接下来，将继续美化工作簿中的其他工作表，请稍候。"
 *   ② "将工作簿中的所有工作表转为商务蓝" → 同样只说不做：模型逐表"汇报"已改商务蓝，
 *      实际 1 个工具都没为其余表调用（主题词"商务蓝"不在判定表内，批量通道没接）。
 *
 * 根因三层（本测试各锁一条）：
 *   ① dispatcher 没有"批量美化"确定性通道——单表"美化表格"硬短路不匹配该措辞，
 *      请求落给 agent-loop；弱模型调 1 次 format_range 就用过程语言充数；
 *   ② turn-goal 的 format 完成条件被 1 次调用即满足（"所有工作表"未分解为多目标），
 *      INCOMPLETE_SIGNAL 也不含"请稍候/正在依次/将继续" → 守卫双双失守；
 *   ③ 收尾把悬挂承诺句当最终总结，过程语言直接进正文。
 *
 * 修复：dispatcher 新增批量美化硬短路（枚举可见表 → 首表主调用 + 其余表进链，全程无 LLM，
 *      支持主题词：商务蓝=后端预设。点名了预设盖不住的外观（中国红、科技绿等）
 *      不在这里硬执行，交给模型按原话规划，避免改了表、丢了颜色）；
 *       turn-goal 多目标按【不同工作表】计数；INCOMPLETE_SIGNAL 扩充 + 收尾清洗悬挂承诺句。
 *
 * 运行：node --test Tests/batch-beautify-all-sheets.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2, DEFAULT_ORDER } = require('./helpers/harness');

// ══ 第一组：dispatcher 批量美化硬短路 ══
const { window, App } = loadChatV2(DEFAULT_ORDER.concat(['services/mcp-dispatcher.js']));
const dispatcher = App.ChatV2.mcpDispatcher;
App.config = {};   // 无 LLM：refineArgsFromSemantics 立即返回种子参数

const SHEETS = [
    { name: '市场部', index: 1, usedRange: 'A1:G21' },
    { name: '销售分析', index: 2, usedRange: 'A1:F15' },
    { name: '库存管理', index: 3, usedRange: 'A1:D40' }
];

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

const sheetsHandler = (sheets) => () => ({
    success: true,
    isError: false,
    result: JSON.stringify(sheets)
});

const fmtOk = () => ({ success: true, isError: false, result: '已格式化' });

const opts = (message) => ({ message: message, toolRegistry: null, sessionStore: null, turn: { context: {} }, allowHardDispatch: true });
const callsOf = (calls, tool) => calls.filter(c => c.tool === tool);

// ── ① 核心：枚举 + 逐表美化 ──
test('"美化工作簿中的所有工作表" → 枚举可见表后逐表确定性美化（不再只做 1 张）', async () => {
    const calls = installBridge({
        'excel_get_sheets': sheetsHandler(SHEETS),
        'excel_format_range': fmtOk
    });
    const d = await dispatcher.tryDispatch(opts('美化工作簿中的所有工作表'));

    assert.strictEqual(d.hardMatched, true, '应走批量美化硬短路');
    assert.strictEqual(d.toolName, 'excel_format_range');

    // 枚举只调一次
    assert.strictEqual(callsOf(calls, 'excel_get_sheets').length, 1);

    // 每张可见表都真实美化了一次（首表走主调用、其余走链），且都带显式 sheet+range
    const fmts = callsOf(calls, 'excel_format_range');
    assert.strictEqual(fmts.length, SHEETS.length, 'N 张可见表就应调 N 次 format_range');
    assert.deepStrictEqual(fmts.map(c => c.args.sheet), SHEETS.map(s => s.name));
    assert.deepStrictEqual(fmts.map(c => c.args.range), SHEETS.map(s => s.usedRange),
        '必须用枚举出的 usedRange：空 range 会走后端"当前选区"兜底做错表');
    for (const f of fmts) assert.strictEqual(f.args.format, '{}', '纯"美化"→ 空 format 走后端商务蓝预设');

    // 步骤面板：首表(主调用 done) + 其余表各一步
    const steps = d.execResult.steps;
    assert.ok(Array.isArray(steps) && steps.length === SHEETS.length, '步骤数应等于表数');
    assert.ok(steps.every(s => s.status === 'done'), '全部步骤应 done: ' + JSON.stringify(steps.map(s => s.status)));

    // 汇总：覆盖全部表、可验收
    const r = d.execResult.result;
    assert.deepStrictEqual(r.beautified, SHEETS.map(s => s.name));
    assert.ok(/已美化 3 个工作表/.test(r.summary), 'summary 应报告全部表: ' + r.summary);
    assert.ok(/商务蓝/.test(r.steps[r.steps.length - 1]), 'steps 末尾应有美化注记');
});

test('"把所有sheet都美化一下" 变体同样命中批量链', async () => {
    const calls = installBridge({
        'excel_get_sheets': sheetsHandler(SHEETS),
        'excel_format_range': fmtOk
    });
    const d = await dispatcher.tryDispatch(opts('把所有sheet都美化一下'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(callsOf(calls, 'excel_format_range').length, SHEETS.length);
});

test('点名具体样式（给所有工作表加边框）→ 透传 format，且仍逐表执行（不被单表块截走）', async () => {
    const calls = installBridge({
        'excel_get_sheets': sheetsHandler(SHEETS),
        'excel_format_range': fmtOk
    });
    const d = await dispatcher.tryDispatch(opts('给所有工作表加边框'));
    assert.strictEqual(d.hardMatched, true);
    const fmts = callsOf(calls, 'excel_format_range');
    assert.strictEqual(fmts.length, SHEETS.length);
    assert.strictEqual(JSON.parse(fmts[0].args.format).borderStyle, 'thin', '点名边框应透传');
});

test('只有 1 张可见表 → 主调用即全部，不画蛇添足补链', async () => {
    const calls = installBridge({
        'excel_get_sheets': sheetsHandler([SHEETS[0]]),
        'excel_format_range': fmtOk
    });
    const d = await dispatcher.tryDispatch(opts('美化所有工作表'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(callsOf(calls, 'excel_format_range').length, 1);
    assert.deepStrictEqual(d.execResult.result.beautified, ['市场部']);
});

test('部分表美化失败 → 成功表继续、失败点名，不伪造全部完成', async () => {
    const calls = installBridge({
        'excel_get_sheets': sheetsHandler(SHEETS),
        'excel_format_range': (args) => (args.sheet === '销售分析'
            ? { success: false, isError: true, result: '找不到工作表' }
            : fmtOk())
    });
    const d = await dispatcher.tryDispatch(opts('美化工作簿中的所有工作表'));
    const r = d.execResult.result;
    assert.deepStrictEqual(r.beautified, ['市场部', '库存管理']);
    assert.deepStrictEqual(r.beautifyFailed, ['销售分析']);
    assert.ok(/失败 1 个（销售分析:/.test(r.steps[r.steps.length - 1]), '失败须点名: ' + r.steps[r.steps.length - 1]);
    const steps = d.execResult.steps;
    assert.strictEqual(steps[1].status, 'failed');
    assert.strictEqual(steps[2].status, 'done', '失败不影响后续表继续');
});

// ── ② 覆盖度与枚举失败：不放行半套美化 ──
test('美化要求覆盖不全（加边框并设置行高）→ 放行 agent-loop，不硬调', async () => {
    const calls = installBridge({
        'excel_get_sheets': sheetsHandler(SHEETS),
        'excel_format_range': fmtOk
    });
    const d = await dispatcher.tryDispatch(opts('给所有工作表加边框并设置行高'));
    assert.strictEqual(d.hit, false, '覆盖不全应放行');
    assert.strictEqual(calls.length, 0, '放行时不应硬调任何工具');
});

test('条件格式类（所有表加数据条）→ 确定性链盖不住，放行', async () => {
    const calls = installBridge({
        'excel_get_sheets': sheetsHandler(SHEETS),
        'excel_format_range': fmtOk
    });
    const d = await dispatcher.tryDispatch(opts('给所有工作表加数据条'));
    assert.strictEqual(d.hit, false);
    assert.strictEqual(calls.length, 0);
});

test('excel_get_sheets 枚举失败 → 确定性诚实报错（不放行，防 JSON 倾倒）', async () => {
    const calls = installBridge({
        'excel_get_sheets': () => { throw new Error('桥不可用'); },
        'excel_format_range': fmtOk
    });
    const d = await dispatcher.tryDispatch(opts('美化工作簿中的所有工作表'));
    // 实测 2026-10-22 22:04 编辑重发：枚举失败放行后，通用兜底把"所有工作表"ngram 匹配到
    // excel_get_sheets，原始 JSON 倾倒进正文冒充完成。现在必须确定性报错。
    assert.strictEqual(d.hardMatched, true, '枚举失败也必须给出确定性结果');
    assert.strictEqual(callsOf(calls, 'excel_format_range').length, 0, '不得在没清单的情况下盲调');
    const r = d.execResult.result;
    assert.strictEqual(r.success, false);
    assert.ok(/未能枚举工作簿的工作表清单/.test(r.error || ''), '报错应说明原因与建议: ' + JSON.stringify(r));
});

test('get_sheets 返回无法解析的形态 → 同样诚实报错（不放行）', async () => {
    const calls = installBridge({
        'excel_get_sheets': () => ({ success: true, isError: false, result: 'garbage-not-json' }),
        'excel_format_range': fmtOk
    });
    const d = await dispatcher.tryDispatch(opts('美化工作簿中的所有工作表'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(callsOf(calls, 'excel_format_range').length, 0);
    assert.ok(/未能枚举/.test((d.execResult.result || {}).error || ''));
});

test('"不要美化所有工作表"否定豁免 → 不枚举不美化', async () => {
    const calls = installBridge({
        'excel_get_sheets': sheetsHandler(SHEETS),
        'excel_format_range': fmtOk
    });
    const d = await dispatcher.tryDispatch(opts('不要美化所有工作表'));
    assert.strictEqual(d.hit, false, '否定语气不应触发批量美化');
    assert.strictEqual(callsOf(calls, 'excel_get_sheets').length, 0);
    assert.strictEqual(callsOf(calls, 'excel_format_range').length, 0);
});

test('非格式化复合请求（统计所有工作表的行数）不误触批量美化', async () => {
    const calls = installBridge({
        'excel_get_sheets': sheetsHandler(SHEETS),
        'excel_format_range': fmtOk
    });
    const d = await dispatcher.tryDispatch(opts('统计所有工作表的行数'));
    assert.strictEqual(d.hit, false);
    assert.strictEqual(callsOf(calls, 'excel_get_sheets').length, 0, '不应枚举');
    assert.strictEqual(callsOf(calls, 'excel_format_range').length, 0);
});

test('Word 宿主不走 Excel 批量美化链', async () => {
    const prev = window.SmartPilotHost;
    window.SmartPilotHost = { isExcel: false, isWord: true, kind: 'word' };
    try {
        const calls = installBridge({
            'excel_get_sheets': sheetsHandler(SHEETS),
            'excel_format_range': fmtOk
        });
        const d = await dispatcher.tryDispatch(opts('美化所有工作表'));
        assert.notStrictEqual(d.toolName, 'excel_format_range');
        assert.strictEqual(calls.length, 0);
    } finally {
        window.SmartPilotHost = prev;
    }
});

// ── ②b 主题类请求（实测 2026-10-22："转为商务蓝"曾因主题词缺失被漏判，又变成只说不做）──
test('"将工作簿中的所有工作表转为商务蓝" → 逐表走后端商务蓝预设', async () => {
    const calls = installBridge({
        'excel_get_sheets': sheetsHandler(SHEETS),
        'excel_format_range': fmtOk
    });
    const d = await dispatcher.tryDispatch(opts('将工作簿中的所有工作表转为商务蓝'));
    assert.strictEqual(d.hardMatched, true, '主题词必须被识别为美化诉求');
    const fmts = callsOf(calls, 'excel_format_range');
    assert.strictEqual(fmts.length, SHEETS.length, 'N 张表就应调 N 次');
    assert.strictEqual(fmts[0].args.format, '{}', '商务蓝 = 后端默认预设（空 format）');
    assert.deepStrictEqual(d.execResult.result.beautified, SHEETS.map(s => s.name));
    assert.ok(/商务蓝/.test(d.execResult.result.steps.slice(-1)[0]), '注记应点名商务蓝');
});

test('"将所有工作表改成科技绿" → 不关键词硬涂，交给模型规划', async () => {
    const calls = installBridge({
        'excel_get_sheets': sheetsHandler(SHEETS),
        'excel_format_range': fmtOk
    });
    const d = await dispatcher.tryDispatch(opts('将所有工作表改成科技绿'));
    assert.strictEqual(d.hit, false, '点名的外观不能被关键词预设吞掉');
    assert.strictEqual(d.deferToPlanner, true);
    assert.ok(/科技绿/.test(d.styleTarget), '应把外观原话交给规划: ' + d.styleTarget);
    assert.strictEqual(callsOf(calls, 'excel_format_range').length, 0, '硬短路不得先涂一层别的颜色');
});

test('"美化工作簿中的所有工作表，统一为中国红" → 不换成商务蓝，交给模型规划', async () => {
    const calls = installBridge({
        'excel_get_sheets': sheetsHandler(SHEETS),
        'excel_format_range': fmtOk
    });
    const d = await dispatcher.tryDispatch(opts('美化工作簿中的所有工作表，统一为中国红'));
    assert.strictEqual(d.hit, false);
    assert.strictEqual(d.deferToPlanner, true);
    assert.ok(/中国红/.test(d.styleTarget), d.styleTarget);
    assert.strictEqual(callsOf(calls, 'excel_format_range').length, 0, '不得用空 format 走商务蓝预设');
    const goal = App.ChatV2.turnGoal.build('美化工作簿中的所有工作表，统一为中国红');
    assert.ok(/中国红/.test(goal.styleTarget || ''));
    const dropped = {
        goal,
        _executedToolCalls: SHEETS.map(s => ({ name: 'excel_format_range', args: { sheet: s.name, format: '{}' } }))
    };
    assert.ok(App.ChatV2.turnGoal.unmet(dropped).length > 0, '逐表涂了空 format（商务蓝）仍算没按语意做完');
    const kept = {
        goal,
        _executedToolCalls: SHEETS.map(s => ({
            name: 'excel_format_range',
            args: { sheet: s.name, format: JSON.stringify({ bgColor: 'DE2910', bgScope: 'header', headerFontColor: 'FFFFFF' }) }
        }))
    };
    assert.strictEqual(App.ChatV2.turnGoal.unmet(kept).length, 0, '每张表都写上了指定颜色才算完成');
    const pale = {
        goal,
        _executedToolCalls: SHEETS.map(s => ({
            name: 'excel_format_range',
            args: { sheet: s.name, format: JSON.stringify({ bgColor: 'F4C2C4', fontColor: 'FFFFFF' }) }
        }))
    };
    assert.ok(App.ChatV2.turnGoal.unmet(pale).length > 0, '浅粉 F4C2C4 不是中国红');
    const fixed = App.ChatV2.turnGoal.prepareFormatArgs(
        { sheet: '市场部', range: 'A1:I21', format: JSON.stringify({ bgColor: 'F4C2C4', bgScope: 'header', fontColor: 'FFFFFF' }) },
        '中国红'
    );
    const fixedFmt = JSON.parse(fixed.format);
    assert.strictEqual(fixedFmt.bgColor, 'DE2910', '浅粉要纠正成中国红');
    assert.strictEqual(fixedFmt.bgScope, 'header');
    assert.strictEqual(fixedFmt.headerFontColor, 'FFFFFF', '白字只留给表头');
    assert.strictEqual(fixedFmt.fontColor, '000000', '表身字必须保持黑色，不能整段涂白');
});

test('"换个主题美化所有工作表"（未点名主题）→ 默认商务蓝预设', async () => {
    const calls = installBridge({
        'excel_get_sheets': sheetsHandler(SHEETS),
        'excel_format_range': fmtOk
    });
    const d = await dispatcher.tryDispatch(opts('换个主题美化所有工作表'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(callsOf(calls, 'excel_format_range').length, SHEETS.length);
    assert.strictEqual(callsOf(calls, 'excel_format_range')[0].args.format, '{}');
});

// ── ②c 范围可靠性判断与降级（实测 2026-10-22：拆分新表 UsedRange 返回 G30 单格/空串）──
const staleSheets = [
    { name: '市场部', index: 1, usedRange: 'A1:I21' },   // 激活过 → 可靠
    { name: '销售三部', index: 4, usedRange: 'G30' },     // 从未激活 → 陈旧单格（Excel 维度缓存不刷新）
    { name: '大客户部', index: 3, usedRange: '' }         // 从未激活 → 空
];
const staleHandlers = {
    'excel_get_sheets': sheetsHandler(staleSheets),
    'excel_get_used_range': (args) => ({ success: true, range: args.sheet === '市场部' ? 'A1:I21' : 'G30', sheet: args.sheet }),
    'excel_format_range': fmtOk
};

test('UsedRange 返回单格/空（未激活表）→ 降级 read_range CurrentRegion 解析，不拿 G30 硬调', async () => {
    const calls = installBridge(Object.assign({}, staleHandlers, {
        // read_range：CurrentRegion 扩展 + maxRows=1 截断注记（JSON 尾部有垃圾，必须仍能解析出 range）
        'excel_read_range': (args) => ({
            success: true, isError: false,
            result: JSON.stringify({ worksheet: args.sheet, range: '$A$1:$G$21', rows: 0, columns: 7, headers: [], result: [] })
                + '\n/* ⚠️ 数据已截断：行数超过 maxRows=1。请缩小 range 范围后重新读取 */'
        })
    }));
    const d = await dispatcher.tryDispatch(opts('美化工作簿中的所有工作表'));
    assert.strictEqual(d.hardMatched, true);
    const fmts = callsOf(calls, 'excel_format_range');
    assert.strictEqual(fmts.length, 3, '三张表都应美化');
    const bySheet = {};
    fmts.forEach(c => { bySheet[c.args.sheet] = c.args.range; });
    assert.strictEqual(bySheet['市场部'], 'A1:I21', '可靠种子直接用');
    assert.strictEqual(bySheet['销售三部'], 'A1:G21', '单格种子不可靠 → 降级解析');
    assert.strictEqual(bySheet['大客户部'], 'A1:G21', '空种子不可靠 → 降级解析');
    assert.ok(!/G30/.test(JSON.stringify(bySheet)), '绝不拿单格 G30 当表格范围（旧版事故）');
    assert.ok(callsOf(calls, 'excel_read_range').length >= 2, '两张不可靠表都应走 read_range 降级');
});

test('三级解析都失败 → 该表点名跳过（不硬调、不静默丢），其余表照常', async () => {
    const calls = installBridge(Object.assign({}, staleHandlers, {
        'excel_read_range': () => ({ success: true, isError: false, result: JSON.stringify({ worksheet: 'x', range: 'G30', rows: 0, columns: 1, headers: [], result: [] }) })
    }));
    const d = await dispatcher.tryDispatch(opts('美化工作簿中的所有工作表'));
    const fmts = callsOf(calls, 'excel_format_range');
    assert.strictEqual(fmts.length, 1, '只有市场部范围可靠可调');
    assert.strictEqual(fmts[0].args.sheet, '市场部');
    const r = d.execResult.result;
    assert.deepStrictEqual(r.beautified, ['市场部']);
    assert.deepStrictEqual(r.beautifySkipped, ['销售三部', '大客户部']);
    assert.ok(/跳过 2 个/.test(r.summary) && /销售三部/.test(r.summary) && /大客户部/.test(r.summary),
        'summary 应点名跳过: ' + r.summary);
    const lastStep = r.steps[r.steps.length - 1];
    assert.ok(/未能定位数据范围/.test(lastStep) && /请点击这些工作表后重试/.test(lastStep),
        '跳过须带原因和行动建议: ' + lastStep);
});

test('全部表范围都不可靠 → 确定性逐表跳过汇报（不放行 LLM、不调 format_range）', async () => {
    const calls = installBridge({
        'excel_get_sheets': sheetsHandler([
            { name: '表A', index: 1, usedRange: 'G30' },
            { name: '表B', index: 2, usedRange: '' }
        ]),
        'excel_get_used_range': () => ({ success: true, range: 'G30' }),
        'excel_read_range': () => ({ success: true, isError: false, result: JSON.stringify({ range: 'G30' }) }),
        'excel_format_range': fmtOk
    });
    const d = await dispatcher.tryDispatch(opts('美化工作簿中的所有工作表'));
    // 实测 2026-10-22 21:40：此场景旧版静默放行 agent-loop → 弱模型编造"已美化"被零工具
    // 守卫道歉（gaslight）。修复后必须确定性产出诚实汇报，绝不经过 LLM。
    assert.strictEqual(d.hardMatched, true, '应命中批量美化块（跳过汇报也是确定性结果）');
    assert.strictEqual(callsOf(calls, 'excel_format_range').length, 0, '没有可靠范围就绝不调 format_range');
    const r = d.execResult.result;
    assert.deepStrictEqual(r.beautified, [], '不得伪造任何美化成功');
    assert.deepStrictEqual(r.beautifySkipped, ['表A', '表B'], '逐表点名');
    assert.ok(/均未能定位到可美化的数据范围/.test(r.summary), 'summary 说明真相: ' + r.summary);
    assert.ok(/可能为空表/.test(r.steps[0]) && /重试/.test(r.steps[0]), '步骤注明原因与行动建议: ' + r.steps[0]);
});

// ── ②d "其他/其余"目标词 + 清除类排除（实测 2026-10-22 晚：'美化…其他工作表'漏接落入 LLM，
//        被执行成 7 次"清除现有格式"，第一个表字色与底色一致文字隐形）──
test('"美化其他工作表"（拆分/合并后对产物表美化的自然说法）→ 命中批量链', async () => {
    const calls = installBridge({
        'excel_get_sheets': sheetsHandler(SHEETS),
        'excel_format_range': () => ({ success: true, isError: false })
    });
    const d = await dispatcher.tryDispatch(opts('美化其他工作表'));
    assert.strictEqual(d.hardMatched, true, '"其他工作表"必须被当作批量目标');
    assert.strictEqual(callsOf(calls, 'excel_format_range').length, SHEETS.length);
    assert.deepStrictEqual(d.execResult.result.beautified, SHEETS.map(s => s.name));
});

test('"清除所有工作表的样式"是删除格式诉求 → 绝不反过来涂美化预设', async () => {
    const calls = installBridge({
        'excel_get_sheets': sheetsHandler(SHEETS),
        'excel_format_range': () => ({ success: true, isError: false })
    });
    const d = await dispatcher.tryDispatch(opts('清除所有工作表的样式'));
    assert.strictEqual(callsOf(calls, 'excel_get_sheets').length, 0, '不得枚举（不该走美化链）');
    assert.strictEqual(callsOf(calls, 'excel_format_range').length, 0, '不得涂预设');
    assert.notStrictEqual(d.toolName, 'excel_format_range');
});

test('"去掉其他工作表的背景色"同为清除诉求 → 不接手美化', async () => {
    const calls = installBridge({
        'excel_get_sheets': sheetsHandler(SHEETS),
        'excel_format_range': () => ({ success: true, isError: false })
    });
    await dispatcher.tryDispatch(opts('去掉其他工作表的背景色'));
    assert.strictEqual(callsOf(calls, 'excel_format_range').length, 0);
});

test('"把结果写到其他工作表"（无格式词）不误触批量美化', async () => {
    const calls = installBridge({
        'excel_get_sheets': sheetsHandler(SHEETS),
        'excel_format_range': () => ({ success: true, isError: false })
    });
    const d = await dispatcher.tryDispatch(opts('把结果写到其他工作表'));
    assert.strictEqual(callsOf(calls, 'excel_get_sheets').length, 0);
    assert.strictEqual(d.hit, false);
});

// ── ②e 执行登记（实测 2026-10-22 22:16：硬短路直调不进 turn 执行记录 → 覆盖度层误判
//        "没做完"，追加一轮空参 format_range 1ms 失败，用户看到莫名 ❌ 步骤）──
test('链执行登记进 turn 执行记录（覆盖度层可见，不再追加空参调用）', async () => {
    const turn = { context: {} };
    const calls = installBridge({
        'excel_get_sheets': sheetsHandler(SHEETS),
        'excel_format_range': () => ({ success: true, isError: false })
    });
    await dispatcher.tryDispatch({ message: '美化工作簿中的所有工作表', toolRegistry: null, sessionStore: null, turn: turn, allowHardDispatch: true });
    assert.ok(Array.isArray(turn._executedToolCalls), '_executedToolCalls 应存在');
    const fmtCalls = turn._executedToolCalls.filter(c => c.name === 'excel_format_range');
    const sheets = new Set(fmtCalls.map(c => c.args && c.args.sheet).filter(Boolean));
    assert.ok(sheets.size >= SHEETS.length, '每张成功表都应登记（多目标计数依据）: ' + JSON.stringify([...sheets]));
    assert.ok(turn._executedTools.indexOf('excel_format_range') >= 0, '_executedTools 应含 excel_format_range');
    assert.strictEqual(callsOf(calls, 'excel_format_range').length, SHEETS.length);
});

// ── ⑤ 主调用不再补裸工具卡（22:31 "没有任何改善" 的真身：`excel_format_range {} · 表名 · 1ms`）──
test('批量链：execResult 置 suppressToolCard，且首步用具体表名（不再叫泛泛的"美化表格"）', async () => {
    const calls = installBridge({
        'excel_get_sheets': sheetsHandler(SHEETS),
        'excel_format_range': fmtOk
    });
    const d = await dispatcher.tryDispatch(opts('美化工作簿中的所有工作表'));
    assert.strictEqual(d.execResult.suppressToolCard, true, '链已覆盖主调用 → orchestrator 不应再补裸卡');
    const steps = d.execResult.steps;
    assert.strictEqual(steps[0].title, '美化「' + SHEETS[0].name + '」',
        '首步必须点名首张工作表（主调用本来就是它）: ' + JSON.stringify(steps[0]));
    assert.ok(/excel_format_range/.test(steps[0].detail || ''), '首步 detail 保留工具与范围');
    assert.strictEqual(steps.length, SHEETS.length, '首步 + 其余表，不重复');
    assert.deepStrictEqual(steps.map(s => s.title), SHEETS.map(s => '美化「' + s.name + '」'));
});

test('单表美化（无链）不受影响：仍渲染正常工具卡', async () => {
    installBridge({
        'excel_get_sheets': sheetsHandler(SHEETS),
        'excel_format_range': fmtOk
    });
    const d = await dispatcher.tryDispatch(opts('美化表格'));
    assert.notStrictEqual(d.execResult.suppressToolCard, true, '单表无步骤面板，工具卡是唯一展示');
});

test('orchestrator 尊重 suppressToolCard（仅在步骤存在时跳过补卡）', () => {
    const fs = require('node:fs');
    const path = require('node:path');
    const src = fs.readFileSync(path.join(__dirname, '..', 'UI', 'js', 'chat-v2', 'chat-orchestrator.js'), 'utf8');
    assert.ok(src.includes('suppressToolCard'), 'orchestrator 未接入 suppressToolCard');
    const idx = src.indexOf('const _suppressCard');
    assert.ok(idx > 0, '缺少 _suppressCard 判定');
    const slice = src.slice(idx, idx + 500);
    assert.ok(/steps\.length > 0/.test(slice), '跳过补卡必须以"步骤确实存在"为前提，避免整轮无过程展示');
    assert.ok(/if \(_suppressCard\)/.test(src), '判定未被用于分支');
});

// ══ 第三组：turn-goal 多目标分解 ══
test('turn-goal："美化工作簿中的所有工作表" 完成条件带 multiTarget，1 次调用不算完成', () => {
    const goal = App.ChatV2.turnGoal.build('美化工作簿中的所有工作表');
    const fmt = (goal.doneWhen || []).find(d => d.kind === 'format');
    assert.ok(fmt, '应提取出 format 完成条件');
    assert.strictEqual(fmt.multiTarget, true, '多目标措辞应打 multiTarget 标记');

    const turn1 = { goal: goal, _executedToolCalls: [{ name: 'excel_format_range', args: { sheet: '市场部' } }] };
    assert.strictEqual(App.ChatV2.turnGoal.unmet(turn1).length, 1, '只美化 1 张表 → 未完成');

    const turn2 = { goal: goal, _executedToolCalls: [
        { name: 'excel_format_range', args: { sheet: '市场部' } },
        { name: 'excel_format_range', args: { sheet: '销售分析' } }
    ] };
    assert.strictEqual(App.ChatV2.turnGoal.unmet(turn2).length, 0, '没有表清单时，≥2 张不同表 → 视为已覆盖');
});

test('turn-goal：清单里有 5 张表、只格式化了 4 张 → 未完成，并点名缺的表', () => {
    const goal = App.ChatV2.turnGoal.build('美化工作簿中的所有工作表，统一为中国红');
    const fmt = { bgColor: 'DE2910', bgScope: 'header', headerFontColor: 'FFFFFF' };
    const sheets = [
        { name: '销售二部', index: 2, usedRange: 'A1:I27' },
        { name: '销售一部', index: 3, usedRange: 'A1:I20' },
        { name: '大客户部', index: 4, usedRange: 'A1:I18' },
        { name: '销售三部', index: 5, usedRange: 'A1:I19' },
        { name: '市场部', index: 6, usedRange: 'A1:I21' }
    ];
    const done = ['销售二部', '销售一部', '大客户部', '销售三部'].map(function (name, i) {
        return {
            name: 'excel_format_range', status: 'completed',
            args: { sheet: name, range: sheets[i].usedRange, format: fmt }
        };
    });
    const turn = {
        goal: goal,
        toolCalls: [{ name: 'excel_get_sheets', status: 'completed', result: JSON.stringify(sheets) }].concat(done),
        _executedToolCalls: done.map(function (c) { return { name: c.name, args: c.args }; })
    };
    const missing = App.ChatV2.turnGoal.missingFormatSheets(turn);
    assert.deepStrictEqual(missing.map(function (s) { return s.name; }), ['市场部']);
    assert.strictEqual(missing[0].usedRange, 'A1:I21');
    const gap = App.ChatV2.turnGoal.unmet(turn);
    assert.strictEqual(gap.length, 1);
    assert.ok(/市场部/.test(gap[0].label), gap[0].label);
    done.push({
        name: 'excel_format_range', status: 'completed',
        args: { sheet: '市场部', range: 'A1:I21', format: fmt }
    });
    turn.toolCalls.push(done[done.length - 1]);
    assert.strictEqual(App.ChatV2.turnGoal.missingFormatSheets(turn).length, 0);
    assert.strictEqual(App.ChatV2.turnGoal.unmet(turn).length, 0);
    assert.ok(/市场部/.test(App.ChatV2.turnGoal.formatCoverageText(turn)));
    assert.ok(!/没有格式化/.test(App.ChatV2.turnGoal.formatCoverageText(turn)));
});

test('turn-goal：普通单表美化不打 multiTarget（行为不回退）', () => {
    const goal = App.ChatV2.turnGoal.build('美化表格');
    const fmt = (goal.doneWhen || []).find(d => d.kind === 'format');
    assert.ok(fmt);
    assert.notStrictEqual(fmt.multiTarget, true);
    const turn = { goal: goal, _executedToolCalls: [{ name: 'excel_format_range', args: { sheet: 'Sheet1' } }] };
    assert.strictEqual(App.ChatV2.turnGoal.unmet(turn).length, 0, '单表请求 1 次调用即满足');
});

// ══ 第四组：零工具守卫（未完成信号 + 悬挂承诺清洗）══
test('守卫：过程语言"正在依次应用相同预设至其他工作表，请稍候。"判为未完成', () => {
    const ztg = App.ChatV2.agentLoopZeroToolGuards;
    assert.ok(ztg, '零工具守卫应已加载');
    assert.strictEqual(
        ztg.isIncompleteReply('正在依次应用相同预设至其他工作表，请稍候。', { keepToolsEveryRound: true, userOpIntent: true }),
        true, '过程语言+续做措辞必须拦下，不得当总结收尾');
    assert.strictEqual(
        ztg.isIncompleteReply('接下来，将继续美化工作簿中的其他工作表，请稍候。', { keepToolsEveryRound: true, userOpIntent: true }),
        true);
});

test('守卫：正常完成总结不被误判为未完成', () => {
    const ztg = App.ChatV2.agentLoopZeroToolGuards;
    assert.strictEqual(
        ztg.isIncompleteReply('已美化全部 3 个工作表：市场部、销售分析、库存管理，均应用商务蓝预设。', { keepToolsEveryRound: true, userOpIntent: true }),
        false, '完成语气不得误伤');
    assert.strictEqual(
        ztg.isIncompleteReply('正在依次应用相同预设至其他工作表，请稍候。', { keepToolsEveryRound: false, userOpIntent: true }),
        false, '非多步骤通道不启用该判定（保持旧行为）');
});

test('守卫：收尾清洗悬挂承诺句，过程语言不进最终总结', () => {
    const ztg = App.ChatV2.agentLoopZeroToolGuards;
    const dirty = '已成功应用【商务蓝专业美化预设】至工作表【市场部】的 A1 单元格。\n'
        + '接下来，将继续美化工作簿中的其他工作表，请稍候。';
    const clean = ztg.stripDanglingPromises(dirty);
    assert.ok(!/请稍候/.test(clean), '承诺句应被剔除: ' + clean);
    assert.ok(!/接下来/.test(clean), '"接下来将继续"应被剔除: ' + clean);
    assert.ok(/市场部/.test(clean), '已完成的事实描述应保留');
    // 无承诺句的正文原样返回
    const ok = '已美化 3 个工作表，任务完成。';
    assert.strictEqual(ztg.stripDanglingPromises(ok), ok);
});
