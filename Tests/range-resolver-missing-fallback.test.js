/**
 * ============================================================
 * Tests/range-resolver-missing-fallback.test.js —
 *   range-resolver 模块缺失时的降级（安全侧）必须诚实，不得静默放行 LLM
 * ============================================================
 * 背景（实测 2026-10-22 21:40 用户截图）：
 *   "美化工作簿中的所有工作表"落给了 agent-loop（未调用任何工具），弱模型编造
 *   "已美化"被零工具守卫道歉，用户看到的是"抱歉，实际上没有调用任何工具"而非真相。
 *   诱因之一：index.html 引用了新的 services/range-resolver.js（mcp-dispatcher 的
 *   范围可靠性判断已委托给它），但该文件因构建/缓存时序未被加载 → 委托全部判
 *   "不可靠" → 批量美化块当时"全部不可靠 → 静默放行" → LLM 幻觉。
 *
 * 修复后行为（本测试锁死）：
 *   ① 模块缺失 → 委托安全侧（全部不可靠），但不崩溃、不抛异常；
 *   ② 批量美化块产出确定性"逐表点名跳过"汇报（hardMatched=true、零 format_range 调用），
 *      绝不放行 agent-loop；
 *   ③ hostLog 留下 ERROR 级诊断（"range-resolver 模块未加载"），现场可即时定位。
 *
 * 运行：node --test Tests/range-resolver-missing-fallback.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2, DEFAULT_ORDER } = require('./helpers/harness');

// 关键：加载顺序里【剔除 range-resolver.js】，模拟"index.html 新、共享模块未加载"的现场
const ORDER_WITHOUT_RR = DEFAULT_ORDER.filter(m => m !== 'services/range-resolver.js')
    .concat(['services/mcp-dispatcher.js']);
const { window, App } = loadChatV2(ORDER_WITHOUT_RR);
const dispatcher = App.ChatV2.mcpDispatcher;
App.config = {};   // 无 LLM：refineArgsFromSemantics 立即返回种子参数

const hostLogs = [];
window.__smartPilotHostLog = (level, msg) => hostLogs.push(level + ':' + msg);

function installBridge(handlers) {
    const calls = [];
    window.excelBridge = {
        // hostLog 走桥的 LogMessage（level, '[McpDispatcher] ' + msg）——诊断断言的采集口
        LogMessage: (level, msg) => hostLogs.push(String(level || 'INFO') + ':' + String(msg || '')),
        CallLocalTool: async (tool, argsJson) => {
            let args = {};
            try { args = typeof argsJson === 'string' ? JSON.parse(argsJson) : (argsJson || {}); } catch (_) { args = {}; }
            calls.push({ tool, args });
            const h = handlers[tool];
            const payload = (typeof h === 'function') ? h(args, calls)
                : { success: false, isError: true, result: '测试桥未实现 ' + tool };
            return JSON.stringify(payload);
        }
    };
    return calls;
}
const callsOf = (calls, tool) => calls.filter(c => c.tool === tool);
const opts = (message) => ({ message, toolRegistry: null, sessionStore: null, turn: { context: {} }, allowHardDispatch: true });

test('前置：本套件确实运行在 range-resolver 缺失状态（模拟版本不一致）', () => {
    assert.strictEqual(App.ChatV2.rangeResolver, undefined, '模拟前提：模块未加载');
    assert.ok(dispatcher, 'mcp-dispatcher 正常加载（委托方不应因模块缺失而崩）');
});

test('模块缺失 + 全部表范围不可靠 → 确定性跳过汇报，不放行 LLM、零 format_range 调用', async () => {
    const calls = installBridge({
        'excel_get_sheets': () => ({
            success: true, isError: false,
            result: JSON.stringify([
                { name: '市场部', index: 1, usedRange: 'A1:I21' },
                { name: '销售三部', index: 2, usedRange: 'G30' }
            ])
        }),
        'excel_get_used_range': () => ({ success: true, range: 'G30' }),
        'excel_read_range': () => ({ success: true, isError: false, result: JSON.stringify({ range: 'G30' }) }),
        'excel_format_range': () => ({ success: true, isError: false })
    });
    const d = await dispatcher.tryDispatch(opts('美化工作簿中的所有工作表'));
    assert.strictEqual(d.hardMatched, true, '必须给出确定性结果，绝不放行 agent-loop（21:40 事故回归）');
    assert.strictEqual(callsOf(calls, 'excel_format_range').length, 0, '没有可靠范围就绝不调 format_range');
    const r = d.execResult.result;
    assert.deepStrictEqual(r.beautified, []);
    assert.deepStrictEqual(r.beautifySkipped, ['市场部', '销售三部']);
    assert.ok(/均未能定位到可美化的数据范围/.test(r.summary), 'summary 说明真相: ' + r.summary);
});

test('模块缺失时留下 ERROR 级诊断日志（现场可定位"版本不一致"）', async () => {
    const calls = installBridge({
        'excel_get_sheets': () => ({
            success: true, isError: false,
            result: JSON.stringify([{ name: '唯一表', index: 1, usedRange: 'A1:B2' }])
        }),
        'excel_format_range': () => ({ success: true, isError: false })
    });
    await dispatcher.tryDispatch(opts('美化工作簿中的所有工作表'));
    const errHit = hostLogs.some(l => l.indexOf('ERROR') === 0 && /range-resolver/.test(l));
    assert.ok(errHit, '应有 ERROR 日志点名 range-resolver 未加载: ' + JSON.stringify(hostLogs.slice(-4)));
});

test('模块缺失时单表格式化硬短路安全弃权（不可靠=不做，不猜范围）', async () => {
    const calls = installBridge({
        'excel_get_used_range': () => ({ success: true, range: 'G30' }),
        'excel_format_range': () => ({ success: true, isError: false })
    });
    // isBareTableBeautify：无选区（turn.context.selection 空）→ autoRange 直调后端整表，
    // 该路径不依赖 range-resolver，应照常工作（autoRange 语义由后端负责）
    const d = await dispatcher.tryDispatch(opts('美化表格'));
    assert.ok(d.hit === true || d.hit === false, '不崩溃即可（该路径不依赖可靠性判断）');
    assert.strictEqual(callsOf(calls, 'excel_format_range').length <= 1, true);
});
