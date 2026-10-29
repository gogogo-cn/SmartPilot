/**
 * ============================================================
 * Tests/ppt-requirement-parser.test.js — PPT 需求解析器（规则层）测试
 * ============================================================
 * 验证核心能力：
 *   1. 图表类型识别（折线/柱状/饼图）
 *   2. 页面类型识别（对比/时间线/明细/KPI/总结）
 *   3. 强调重点提取
 *   4. 自定义章节提取
 *   5. 复杂度判定（叙事结构 → 降级 LLM）
 *   6. 需求应用到推荐 slides
 * 运行：node Tests/ppt-requirement-parser.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const path = require('path');
const fs = require('fs');

const BASE = path.join(__dirname, '..', 'UI', 'js', 'chat-v2');

function loadParser() {
    const window = {};
    window.App = { ChatV2: { services: {} }, config: {} };
    global.window = window;
    globalThis.App = window.App;
    (0, eval)(fs.readFileSync(path.join(BASE, 'services/ppt-requirement-parser.js'), 'utf8'));
    return App.ChatV2.services.pptRequirementParser;
}

test('需求解析: 图表类型识别（折线/柱状/饼图）', () => {
    const p = loadParser();
    assert.strictEqual(p.parsePptRequirements('用折线图展示趋势').chartType, 'line');
    assert.strictEqual(p.parsePptRequirements('做柱状图对比').chartType, 'column');
    assert.strictEqual(p.parsePptRequirements('看下占比用饼图').chartType, 'pie');
    assert.strictEqual(p.parsePptRequirements('生成一份报告').chartType, null);
    console.log('  [ok] 图表类型识别正确');
});

test('需求解析: 页面类型识别（对比/时间线/明细/KPI/总结）', () => {
    const p = loadParser();
    assert.ok(p.parsePptRequirements('加一页竞品对比').extraPages.includes('comparison'));
    assert.ok(p.parsePptRequirements('加时间线').extraPages.includes('timeline'));
    assert.ok(p.parsePptRequirements('要数据明细').extraPages.includes('table'));
    assert.ok(p.parsePptRequirements('放关键指标').extraPages.includes('kpi'));
    assert.ok(p.parsePptRequirements('最后总结一下').extraPages.includes('text'));
    console.log('  [ok] 页面类型识别正确');
});

test('需求解析: 强调重点提取', () => {
    const p = loadParser();
    const req = p.parsePptRequirements('重点突出销售增长，用折线图展示');
    assert.ok(req.emphasis.includes('销售增长'), `应提取到"销售增长"，实际: ${req.emphasis.join(',')}`);
    assert.strictEqual(req.chartType, 'line');
    console.log('  [ok] 强调重点提取正确');
});

test('需求解析: 自定义章节提取', () => {
    const p = loadParser();
    const req = p.parsePptRequirements('包括背景、市场分析、总结建议');
    assert.ok(req.sections.length > 0, '应提取到章节');
    console.log(`  [ok] 章节提取: ${req.sections.join('/')}`);
});

test('需求解析: 复杂度判定（叙事结构 → 降级 LLM）', () => {
    const p = loadParser();
    assert.strictEqual(p.parsePptRequirements('先讲背景，再讲问题，最后给建议').needsLlm, true,
        '叙事结构应判定为复杂需求');
    assert.strictEqual(p.parsePptRequirements('第一部分介绍背景，第二部分分析数据').needsLlm, true,
        '分部分结构应判定为复杂需求');
    assert.strictEqual(p.parsePptRequirements('生成一份商务蓝的 PPT').needsLlm, false,
        '简单需求不应判定为复杂');
    console.log('  [ok] 复杂度判定正确');
});

test('需求解析: 应用到推荐 slides（图表类型覆盖 + 页面增加 + 强调注入）', () => {
    const p = loadParser();
    const slides = [
        { type: 'title', title: '报告' },
        { type: 'agenda', title: '目录', items: ['关键指标'] },
        { type: 'chart', title: '趋势变化', chartType: 'line', xField: '月份', series: ['销售额'] },
        { type: 'kpi', title: '关键指标', items: [{ label: '销售额', value: '100' }] }
    ];
    const req = p.parsePptRequirements('用柱状图，加一页竞品对比，重点突出销售增长');
    const result = p.applyRequirementsToDesign(slides, req);

    const chart = result.find(s => s.type === 'chart');
    assert.strictEqual(chart.chartType, 'column', '图表类型应被覆盖为柱状图');
    assert.ok(result.some(s => s.type === 'comparison'), '应新增对比页');
    const kpi = result.find(s => s.type === 'kpi');
    assert.ok(kpi.items.some(it => it.label === '销售增长'), 'KPI 应注入强调重点');
    console.log('  [ok] 需求应用到推荐 slides 正确');
});

test('需求解析: 无需求时返回空结构且不改变 slides', () => {
    const p = loadParser();
    const req = p.parsePptRequirements('生成一份 PPT');
    assert.strictEqual(req.chartType, null);
    assert.strictEqual(req.extraPages.length, 0);
    assert.strictEqual(req.emphasis.length, 0);
    assert.strictEqual(req.needsLlm, false);

    const slides = [{ type: 'title', title: '报告' }];
    const result = p.applyRequirementsToDesign(slides, req);
    assert.strictEqual(result.length, 1);
    assert.strictEqual(result[0].type, 'title');
    console.log('  [ok] 无需求时行为正确');
});

// 测试结束后强制退出（与其它测试文件一致，防止定时器挂起进程）
if (require.main === module) { setTimeout(() => process.exit(0), 3000).unref(); } // 仅独立运行时强制退出；聚合运行（run-tests.js require）不注册，避免进程提前退出
