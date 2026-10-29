/**
 * ============================================================
 * Tests/chart-command-selfcontained.test.js — 自带维度的图表命令不得继承上一轮分析
 * ============================================================
 * 回归背景（2026-10-10 00:32 日志取证）：
 *   第 1 轮「使用饼图按部门展示销售额」→ excel_aggregate + excel_create_chart(pie)（正确）
 *   第 2 轮「分析选中数据的描述统计」→ python-pipeline（分析）
 *   第 3 轮「使用饼图按部门展示销售额」重发 → **又进 python-pipeline 重跑整段分析**
 *     （日志：excel_run_python_script；气泡：分组统计结果/关键发现/结论 + ✅ 分析完成！共完成 1 项分析）
 *   根因：intent-analyzer.referencesPreviousTask 把 REFERENCE_PATTERNS.chart（含"饼图"）
 *   当作"引用上一轮任务"的明确信号 —— 于是任何含图表类型词的**独立新命令**
 *   （自带数据维度"按部门展示销售额"）只要上一轮是 python/text，就被判为后续消息，
 *   targetPipeline 继承上一轮 python → 分析重跑，饼图一次没画。
 *
 * 判据：图表词只是"产物类型"，不等于"引用上文"。
 *   自带数据维度/范围的图表命令 → 独立新任务（不继承）；
 *   纯换图型（换成柱状图）/无数据线索的图表诉求 → 仍继承（保留原有跨管道能力）。
 *
 * 运行：node Tests/chart-command-selfcontained.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const path = require('path');
const fs = require('fs');

const BASE = path.join(__dirname, '..', 'UI', 'js', 'chat-v2');

function loadAnalyzer() {
    const window = {};
    window.App = { ChatV2: {} };
    global.window = window;
    globalThis.App = window.App;
    // routing-constants 提供 CHART_TYPE_SWITCH_RE（跨管道图表继承的判定依据），
    // 不加载会让 harness 与生产不一致（"换成柱状图" 在生产命中的是 cross_chart 分支）。
    (0, eval)(fs.readFileSync(path.join(BASE, 'config/routing-constants.js'), 'utf8'));
    (0, eval)(fs.readFileSync(path.join(BASE, 'services/intent-analyzer.js'), 'utf8'));
    return App.ChatV2.intentAnalyzer;
}

/** 上一轮：python 分析管道（含任务结构与数据列） */
function prevPythonTurn() {
    return {
        turnId: 'turn_prev_python',
        spec: {
            pipeline: 'python',
            taskStructure: { tasks: [{ name: '分组聚合', detail: '按部门求和' }] },
            data: {
                selection: { headers: ['部门', '销售额'], rows: [{ 部门: '销售一部', 销售额: 480788 }] }
            },
            result: { summary: '销售二部销售额最高' }
        }
    };
}

/** 当前上下文：有选区（带表头） */
const ctxWithSelection = {
    selection: { address: 'A1:B10', sheetName: '销售数据', headers: ['部门', '销售额'] }
};

test('自带维度的图表命令（有选区）→ 不得继承上一轮分析', () => {
    const ia = loadAnalyzer();
    const r = ia.analyzeFollowUpIntent({
        message: '使用饼图按部门展示销售额',
        prevSpec: prevPythonTurn(),
        context: ctxWithSelection
    });
    assert.strictEqual(r.isFollowUp, false,
        '「使用饼图按部门展示销售额」是独立新命令（自带维度"按部门"），不应继承上一轮 ' +
        'python 分析；实际=' + JSON.stringify(r));
});

test('纯换图型「换成柱状图」→ 仍继承（跨管道生成图表）', () => {
    const ia = loadAnalyzer();
    const r = ia.analyzeFollowUpIntent({
        message: '换成柱状图',
        prevSpec: prevPythonTurn(),
        context: ctxWithSelection
    });
    assert.strictEqual(r.isFollowUp, true, '明确"仅切图表类型"必须继承上一轮数据；实际=' + JSON.stringify(r));
    assert.strictEqual(r.followUpType, 'cross_chart', '应判为 cross_chart；实际=' + JSON.stringify(r));
    assert.strictEqual(r.targetPipeline, 'skill', '应路由 skill-pipeline；实际=' + JSON.stringify(r));
});

test('无数据线索的图表诉求「用饼图展示一下」→ 仍继承', () => {
    const ia = loadAnalyzer();
    const r = ia.analyzeFollowUpIntent({
        message: '用饼图展示一下',
        prevSpec: prevPythonTurn(),
        context: { selection: { address: 'B2', sheetName: 'Sheet1' } }
    });
    assert.strictEqual(r.isFollowUp, true, '无自带维度的图表诉求应继承上一轮结果；实际=' + JSON.stringify(r));
});

test('无选区的自带维度图表命令 → 也不继承（不复活上一轮分析管道）', () => {
    const ia = loadAnalyzer();
    const r = ia.analyzeFollowUpIntent({
        message: '使用饼图按部门展示销售额',
        prevSpec: prevPythonTurn(),
        context: { selection: { address: 'B2', sheetName: 'Sheet1' } }
    });
    assert.strictEqual(r.isFollowUp, false,
        '无选区时也不应把自带维度的图表命令判成"引用上一轮分析"；实际=' + JSON.stringify(r));
});

test('明确引用上文的措辞优先于自带维度（「把刚才的分析按部门做成饼图」仍继承）', () => {
    const ia = loadAnalyzer();
    const r = ia.analyzeFollowUpIntent({
        message: '把刚才的分析按部门做成饼图',
        prevSpec: prevPythonTurn(),
        context: ctxWithSelection
    });
    assert.strictEqual(r.isFollowUp, true, '「刚才的分析」是明确引用信号，应继承；实际=' + JSON.stringify(r));
});

test('无图表词的自带维度命令「按部门展示销售额」→ 不继承', () => {
    const ia = loadAnalyzer();
    const r = ia.analyzeFollowUpIntent({
        message: '按部门展示销售额',
        prevSpec: prevPythonTurn(),
        context: ctxWithSelection
    });
    assert.strictEqual(r.isFollowUp, false, '不继承；实际=' + JSON.stringify(r));
});
