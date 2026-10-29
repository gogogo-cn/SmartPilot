/**
 * ============================================================
 * Tests/intent-detector-dimension.test.js — 维度列推导回归测试
 * ============================================================
 * 回归背景（2026-08-29）：
 *   用户输入「按区域展示销售额」时，聚合工具报「未找到维度列 '地区'」。
 *   根因：detectDimensionColumns 命中 DIM_COL_PATTERNS 后返回规范化标签（「地区」），
 *   而非用户实际输入的关键词（「区域」），导致 dimColumn 无法精确命中表头「区域」。
 *   修复：返回实际匹配文本；后端 SemanticMatch 另增「区域↔地区」等同义词兜底。
 * 运行：node Tests/intent-detector-dimension.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const path = require('path');
const fs = require('fs');

const BASE = path.join(__dirname, '..', 'UI', 'js', 'chat-v2');

function loadDetector() {
    const window = {};
    window.App = { ChatV2: {}, config: {} };
    global.window = window;
    globalThis.App = window.App;
    // 配置驱动：须先加载 intent-config.js（intent-detector 经 defs() 取配置）
    (0, eval)(fs.readFileSync(path.join(BASE, 'services/intent-config.js'), 'utf8'));
    (0, eval)(fs.readFileSync(path.join(BASE, 'services/intent-detector.js'), 'utf8'));
    return App.ChatV2.intentDetector;
}

test('dimension: 「按区域」返回实际关键词「区域」而非规范化标签「地区」', () => {
    const det = loadDetector();
    const dims = det.detectDimensionColumns('使用饼图按区域展示销售额');
    assert.ok(dims.includes('区域'), '应命中用户实际输入的「区域」，实际=' + JSON.stringify(dims));
    assert.ok(!dims.includes('地区'), '不应返回规范化标签「地区」，实际=' + JSON.stringify(dims));
});

test('dimension: 「按地区」仍返回「地区」', () => {
    const det = loadDetector();
    const dims = det.detectDimensionColumns('按地区统计销售额');
    assert.ok(dims.includes('地区'), '实际=' + JSON.stringify(dims));
});

test('dimension: 其他维度词返回实际输入', () => {
    const det = loadDetector();
    const cases = [
        ['按部门汇总利润', '部门'],
        ['按月展示销售额', '月'],
        ['按城市分析收入', '城市'],
        ['按产品类别汇总销售额', '产品'],
        ['按渠道对比利润', '渠道']
    ];
    for (const [input, expected] of cases) {
        const dims = det.detectDimensionColumns(input);
        assert.ok(dims.includes(expected), `输入「${input}」应包含「${expected}」，实际=${JSON.stringify(dims)}`);
    }
});

test('dimension: 未命中词表时回退到「按X」短语捕获', () => {
    const det = loadDetector();
    const dims = det.detectDimensionColumns('按大区排名销售额');
    assert.ok(dims.includes('大区'), '回退捕获应含「大区」，实际=' + JSON.stringify(dims));
});
