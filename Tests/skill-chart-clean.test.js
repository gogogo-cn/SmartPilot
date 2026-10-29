/**
 * ============================================================
 * Tests/skill-chart-clean.test.js — Chart/Cleaner SkillSpec 模式测试
 * ============================================================
 * 验证：SmartChartWizard（图表）和 SmartDataCleaner（清洗）
 * 走"LLM 输出操作 spec → 工具链/确定性 VBA"路径。
 * 运行：node Tests/skill-chart-clean.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const path = require('path');
const fs = require('fs');

const BASE = path.join(__dirname, '..', 'UI', 'js', 'chat-v2');

function loadEngine() {
    const window = {};
    window.App = { ChatV2: { services: {} }, config: {} };
    global.window = window;
    globalThis.App = window.App;
    // 配置驱动：须先加载 intent-config.js（skill-spec-engine 经 defs() 取图表映射）
    (0, eval)(fs.readFileSync(path.join(BASE, 'services/intent-config.js'), 'utf8'));
    (0, eval)(fs.readFileSync(path.join(BASE, 'services/skill-spec-engine.js'), 'utf8'));
    return App.ChatV2.services.skillSpecEngine;
}

test('chart spec: 中文图表类型映射 + 规范化', () => {
    const engine = loadEngine();
    // 用户说"画个饼图看占比" → LLM 输出中文 → 映射 xlPie
    const spec = engine.normalizeChartSpec({ chartType: '饼图', title: '销售占比', dataRange: 'A1:B10' });
    assert.strictEqual(spec.chartType, 'xlPie', '饼图 → xlPie');
    assert.strictEqual(spec.title, '销售占比');
    // 柱状图 → xlColumnClustered
    const spec2 = engine.normalizeChartSpec({ chartType: '柱状图' });
    assert.strictEqual(spec2.chartType, 'xlColumnClustered', '柱状图 → xlColumnClustered');
    // 折线/趋势 → xlLineMarkers
    const spec3 = engine.normalizeChartSpec({ chartType: '折线图' });
    assert.strictEqual(spec3.chartType, 'xlLineMarkers', '折线图 → xlLineMarkers');
    // 非法类型回退默认
    const spec4 = engine.normalizeChartSpec({ chartType: 'evil' });
    assert.strictEqual(spec4.chartType, 'xlColumnClustered', '非法类型回退默认');
    console.log('✅ 图表类型映射：饼图/柱状图/折线图 → Excel 枚举');
});

test('clean spec: 操作白名单 + 规范化', () => {
    const engine = loadEngine();
    // 用户"去重并填充空值"
    const spec = engine.normalizeCleanSpec({ operations: ['removeDuplicates', 'fillEmpty'], fillMethod: 'zero' });
    assert.deepStrictEqual(spec.operations, ['removeDuplicates', 'fillEmpty']);
    assert.strictEqual(spec.fillMethod, 'zero');
    // 非法操作被过滤
    const spec2 = engine.normalizeCleanSpec({ operations: ['removeDuplicates', 'DELETE ALL', 'hack'] });
    assert.deepStrictEqual(spec2.operations, ['removeDuplicates'], '非法操作过滤');
    // 字符串逗号分隔
    const spec3 = engine.normalizeCleanSpec({ operations: 'trimSpaces,normalizeDates' });
    assert.deepStrictEqual(spec3.operations, ['trimSpaces', 'normalizeDates']);
    console.log('✅ 清洗操作白名单 + 规范化');
});

test('clean 执行: 清洗不再走 VBA 渲染（迁移到 MCP excel_clean_data）', () => {
    const engine = loadEngine();
    // 清洗已迁移到 MCP excel_clean_data 原生引擎，引擎不再提供 renderCleanVBA
    assert.strictEqual(typeof engine.renderCleanVBA, 'undefined', 'renderCleanVBA 已移除（清洗走 MCP excel_clean_data）');
    // 清洗 spec 规范化仍可用
    const spec = engine.normalizeCleanSpec({
        operations: ['removeDuplicates', 'fillEmpty', 'trimSpaces', 'normalizeDates', 'normalizeNumbers'],
        fillMethod: 'zero'
    });
    assert.deepStrictEqual(spec.operations, ['removeDuplicates', 'fillEmpty', 'trimSpaces', 'normalizeDates', 'normalizeNumbers']);
    assert.strictEqual(spec.fillMethod, 'zero');
    console.log('✅ 清洗执行路径：MCP excel_clean_data（无 VBA 渲染），spec 规范化正常');
});

test('chart spec: 组合图（主柱+次折）', () => {
    const engine = loadEngine();
    const spec = engine.normalizeChartSpec({
        chartType: 'xlColumnClustered', comboChart: true, secondaryValueColumn: '同比'
    });
    assert.strictEqual(spec.comboChart, true);
    assert.strictEqual(spec.secondaryValueColumn, '同比');
    console.log('✅ 组合图 spec');
});

test('clean prompt: 引导 LLM 输出清洗操作 spec', () => {
    const engine = loadEngine();
    const prompt = engine.buildCleanPlanningPrompt({ displayName: '智能数据清洗大师' });
    assert.ok(prompt.includes('绝不输出代码'), '禁止写代码');
    assert.ok(prompt.includes('removeDuplicates'), '含去重操作');
    assert.ok(prompt.includes('fillEmpty'), '含填充操作');
    console.log('✅ 清洗规划 prompt');
});

test('chart prompt: 引导 LLM 输出图表 spec', () => {
    const engine = loadEngine();
    const prompt = engine.buildChartPlanningPrompt({ displayName: '智能图表向导' });
    assert.ok(prompt.includes('chartType'), '含图表类型');
    assert.ok(prompt.includes('xlColumnClustered'), '含 Excel 枚举');
    console.log('✅ 图表规划 prompt');
});

if (require.main === module) { setTimeout(() => process.exit(0), 4000).unref(); } // 仅独立运行时强制退出；聚合运行（run-tests.js require）不注册，避免进程提前退出
