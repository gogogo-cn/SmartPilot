/**
 * ============================================================
 * Tests/mcp-tool-audit-fixes.test.js — MCP 审核三件套回归
 * ============================================================
 * 1) filterToolsForMessage 孤儿词表（check_errors / layout / protect / shape / copy_range）
 * 2) excel_create_chart 不再是全局终态（persistTools 下多图可继续拿 tools）
 * 3) 终态名表使用真实 excel_export_csv（源码守门）
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const { loadChatV2 } = require('./helpers/harness');

const ROOT = path.join(__dirname, '..');

function util() {
    const { App } = loadChatV2();
    return App.ChatV2.toolRegistryUtil;
}

const ORPHAN_TOOLS = [
    'excel_check_errors',
    'excel_get_layout_snapshot',
    'excel_protect_sheet',
    'excel_unprotect_sheet',
    'excel_add_shape',
    'excel_get_cell_value',
    'excel_copy_range',
    'excel_set_cell_value',
    'excel_read_range',
    'excel_create_chart'
];

test('瘦身: 检查错误 → 保留 excel_check_errors', () => {
    const kept = util().filterToolsForMessage(
        ORPHAN_TOOLS.map(name => ({ name })),
        '检查表格有没有公式错误'
    ).map(t => t.name);
    assert.ok(kept.includes('excel_check_errors'), '实际=' + kept.join(','));
});

test('瘦身: 布局快照 / 看板 → 保留 get_layout_snapshot 与 add_shape', () => {
    const kept = util().filterToolsForMessage(
        ORPHAN_TOOLS.map(name => ({ name })),
        '做个数据看板，看一下当前布局还能放哪里'
    ).map(t => t.name);
    assert.ok(kept.includes('excel_get_layout_snapshot'), '实际=' + kept.join(','));
    assert.ok(kept.includes('excel_add_shape'), '看板应带上形状工具，实际=' + kept.join(','));
});

test('瘦身: 保护工作表 → 保留 protect/unprotect', () => {
    const kept = util().filterToolsForMessage(
        ORPHAN_TOOLS.map(name => ({ name })),
        '给当前工作表加保护锁定'
    ).map(t => t.name);
    assert.ok(kept.includes('excel_protect_sheet'), '实际=' + kept.join(','));
});

test('瘦身: 复制区域 → 保留 excel_copy_range', () => {
    const kept = util().filterToolsForMessage(
        ORPHAN_TOOLS.map(name => ({ name })),
        '把选中区域复制到右边'
    ).map(t => t.name);
    assert.ok(kept.includes('excel_copy_range'), '实际=' + kept.join(','));
});

test('瘦身: 设置/读取单元格 → 保留 set/get_cell_value', () => {
    const u = util();
    const setKept = u.filterToolsForMessage(
        ORPHAN_TOOLS.map(name => ({ name })),
        '设置单元格 A1 的值'
    ).map(t => t.name);
    assert.ok(setKept.includes('excel_set_cell_value'), '实际=' + setKept.join(','));
    const getKept = u.filterToolsForMessage(
        ORPHAN_TOOLS.map(name => ({ name })),
        '读取单元格 B2 的内容'
    ).map(t => t.name);
    assert.ok(getKept.includes('excel_get_cell_value'), '实际=' + getKept.join(','));
});

test('终态: excel_create_chart 成功后 persistTools 下一轮仍下发 tools', async () => {
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;
    tr.register({
        name: 'excel_create_chart', displayName: '创建图表', description: '画图', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute() { return { ok: true, result: { chartName: 'AIChart_1' } }; }
    });

    const toolsOffered = [];
    let n = 0;
    App.ChatV2.llmClient = {
        stream({ tools }) {
            toolsOffered.push(Array.isArray(tools) ? tools.length : -1);
            return (async function* () {
                n++;
                yield { type: 'model', model: 'test' };
                if (n === 1) {
                    yield {
                        type: 'tool_calls',
                        toolCalls: [{
                            id: 'c1',
                            function: {
                                name: 'excel_create_chart',
                                arguments: JSON.stringify({ chartType: 'xlPie', dataRange: 'A1:B3' })
                            }
                        }]
                    };
                    yield { type: 'done', finishReason: 'tool_calls' };
                } else {
                    yield { type: 'token', delta: '已生成第一张图，接下来再画一张。' };
                    yield { type: 'done', finishReason: 'stop' };
                }
            })();
        }
    };

    const turn = App.ChatV2.sessionStore.startTurn({
        message: '按部门画饼图，再画一张柱状图',
        kind: 'text',
        context: { _forceChart: true }
    });
    for await (const _ of App.ChatV2.agentLoop.run({
        turn,
        messages: [{ role: 'user', content: '按部门画饼图，再画一张柱状图' }],
        tools: [{ name: 'excel_create_chart', description: '画图', inputSchema: { type: 'object', properties: {} } }],
        signal: new AbortController().signal,
        persistTools: true,
        maxIterations: 5
    })) { /* drain */ }

    assert.ok(toolsOffered.length >= 2, '至少两轮 LLM 调用，实际=' + toolsOffered.length);
    assert.ok(toolsOffered[1] > 0,
        'create_chart 成功后下一轮仍应下发 tools（不再是全局终态），序列=' + JSON.stringify(toolsOffered));
});

test('终态名表: 源码使用 excel_export_csv，不含幽灵 export_to_* / create_chart', () => {
    const src = fs.readFileSync(path.join(ROOT, 'UI/js/chat-v2/agent-loop.js'), 'utf8');
    const m = src.match(/const BUILTIN_TERMINAL_TOOLS = new Set\(\[([\s\S]*?)\]\);/);
    assert.ok(m, '应能定位 BUILTIN_TERMINAL_TOOLS');
    const block = m[1];
    assert.ok(block.includes("'excel_export_csv'"), '应包含真实工具 excel_export_csv');
    assert.ok(!block.includes('excel_export_to_csv'), '不得再含幽灵名 excel_export_to_csv');
    assert.ok(!block.includes('excel_export_to_pdf'), '不得再含幽灵名 excel_export_to_pdf');
    assert.ok(!block.includes("'excel_create_chart'"), 'create_chart 不得作为全局终态');
});

test('clear_range schema: clearFormat 默认文案为 false', () => {
    const cfg = JSON.parse(fs.readFileSync(path.join(ROOT, 'Config/excel_tools.json'), 'utf8').replace(/^\uFEFF/, ''));
    const t = cfg.tools.find(x => x.name === 'excel_clear_range');
    assert.ok(t, '应有 excel_clear_range');
    const p = (t.parameters || []).find(x => x.name === 'clearFormat');
    assert.ok(p && /默认\s*false/.test(String(p.description || '')),
        'clearFormat 参数描述应写默认 false，实际=' + (p && p.description));
    assert.ok(/clearFormat=false/.test(String(t.description || '')),
        '工具描述应声明默认 clearFormat=false');
});

if (require.main === module) { setTimeout(() => process.exit(0), 4000).unref(); }
