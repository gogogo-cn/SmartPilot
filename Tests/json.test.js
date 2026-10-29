/**
 * json.test.js — renderers/json 纯函数子模块单测
 *   JSON 语法高亮 / 可折叠面板 / 结构化数据转表。
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

function load() {
    const ctx = loadChatV2();
    const j = ctx.App.ChatV2.renderers.json;
    assert.ok(j, 'renderers.json 应已挂载');
    return j;
}

test('json renderJsonHtml：语法高亮分类标签', () => {
    const { renderJsonHtml } = load();
    const html = renderJsonHtml('{"name": "Alice", "age": 30, "active": true}');
    assert.ok(html.includes('json-key'), '键应有 json-key 类');
    assert.ok(html.includes('json-string'), '字符串应有 json-string 类');
    assert.ok(html.includes('json-number'), '数字应有 json-number 类');
    assert.ok(html.includes('json-bool'), '布尔应有 json-bool 类');
    assert.ok(html.startsWith('<pre class="json-block">'), '应包裹在 json-block 中');
});

test('json renderCollapsibleJson：生成 details/summary 结构', () => {
    const { renderCollapsibleJson } = load();
    const html = renderCollapsibleJson('{"a":1}');
    assert.ok(html.includes('<details class="json-collapsible"'), '应生成 details');
    assert.ok(html.includes('查看原始结果（JSON）'), '应有 summary 文案');
    assert.ok(html.includes('json-block'), '应内嵌高亮 JSON');
});

test('json tryConvertToTable：描述统计 → 统计表', () => {
    const { tryConvertToTable } = load();
    const data = {
        '销量': { count: 12, mean: 150.5, median: 148, stdev: 12.3, min: 130, max: 180, q1: 140, q3: 160 },
        '利润': { count: 12, mean: 50.2, median: 48, stdev: 5.1, min: 40, max: 60, q1: 46, q3: 55 }
    };
    const html = tryConvertToTable(data);
    assert.ok(html, '应生成表格');
    assert.ok(html.includes('py-stats-table'), '应带统计表样式');
    assert.ok(html.includes('样本数') && html.includes('均值'), '应含统计量中文标签');
    assert.ok(html.includes('销量') && html.includes('利润'), '应含列名');
});

test('json tryConvertToTable：分组聚合 → 聚合表', () => {
    const { tryConvertToTable } = load();
    const data = { '华东': { value: 100, count: 5, agg: 'sum' }, '华北': { value: 80, count: 4, agg: 'sum' } };
    const html = tryConvertToTable(data);
    assert.ok(html && html.includes('py-agg-table'), '应生成聚合表');
    assert.ok(html.includes('求和'), '应把 agg=sum 映射为"求和"');
});

test('json tryConvertToTable：相关性矩阵 → 矩阵表', () => {
    const { tryConvertToTable } = load();
    const data = { A: { A: 1, B: 0.8 }, B: { A: 0.8, B: 1 } };
    const html = tryConvertToTable(data);
    assert.ok(html && html.includes('py-corr-table'), '应生成相关矩阵表');
    assert.ok(html.includes('py-corr-strong'), '|r|>=0.7 应为强相关');
    assert.ok(html.includes('0.800'), '数值应保留三位小数');
});

test('json tryConvertToTable：异常检测 → 异常表', () => {
    const { tryConvertToTable } = load();
    const data = { '销量': { total: 12, anomaly_count: 2, anomaly_rate: 16.7, anomalies: [{ value: 300, z_score: 3.2 }] } };
    const html = tryConvertToTable(data);
    assert.ok(html && html.includes('py-anomaly-table'), '应生成异常表');
    assert.ok(html.includes('异常数') && html.includes('Z=3.2'), '应含异常详情');
});

test('json tryConvertToTable：数据分布 → 分布表', () => {
    const { tryConvertToTable } = load();
    const data = { '销量': { min: 100, max: 200, bins: 5, distribution: [{ range: '100-120', percentage: 20 }] } };
    const html = tryConvertToTable(data);
    assert.ok(html && html.includes('py-dist-table'), '应生成分布表');
    assert.ok(html.includes('width:20%'), '分布条宽度应为 20%');
});

test('json renderTrendTable：方向+预测值', () => {
    const { renderTrendTable } = load();
    const html = renderTrendTable({ direction: 'up', forecast: [1, 2, 3] });
    assert.ok(html && html.includes('py-trend-table'), '应生成趋势表');
    assert.ok(html.includes('预测值') && html.includes('1, 2, 3'), '应含预测值行');
});

test('json tryConvertToTable：数组对象 → 标准表格并截断', () => {
    const { tryConvertToTable } = load();
    // 展示上限已从 20 上调到 200（见 renderers/json.js 的 DISPLAY_MAX），
    // 断言跟随真实上限：25 条不应截断，超过 200 条才提示。
    const rows = Array.from({ length: 25 }, (_, i) => ({ id: i + 1, name: 'x' + i }));
    const html = tryConvertToTable(rows);
    assert.ok(html && html.includes('py-data-table'), '应生成标准表');
    assert.ok(!html.includes('仅显示前'), '25 条未超过 200 条上限，不应出现截断提示');

    const many = Array.from({ length: 260 }, (_, i) => ({ id: i + 1, name: 'x' + i }));
    const html2 = tryConvertToTable(many);
    assert.ok(html2.includes('共 260 条，仅显示前 200 条'), '超过 200 条必须提示截断（全量仍可通过下载/插入 Excel 获取）');
});

test('json tryConvertToTable：无法识别返回 null', () => {
    const { tryConvertToTable } = load();
    assert.strictEqual(tryConvertToTable(null), null);
    assert.strictEqual(tryConvertToTable('str'), null);
    assert.strictEqual(tryConvertToTable({}), null);
});

test('json tryRowsMatrixTable：{success,data:[[…]]} 操作类结果（合并工作表）→ 表格', () => {
    const { tryRowsMatrixTable } = load();
    // excel_merge_sheets 的真实回传形态：data 是二维矩阵，首行即表头
    const merge = {
        success: true,
        sourceSheets: ['大客户部', '销售三部'],
        mergedRows: 3,
        dataRows: 3,
        headerPreserved: true,
        data: [['日期', '地区', '部门'], ['45353', '华南', '大客户部'], ['45354', '华东', '销售三部']]
    };
    const html = tryRowsMatrixTable(merge);
    assert.ok(html && html.includes('py-range-table'), '应生成行矩阵表');
    assert.ok(html.includes('<th>日期</th>') && html.includes('<th>地区</th>'), '首行应作为表头');
    assert.ok(html.includes('大客户部') && html.includes('销售三部'), '应含数据行');
    assert.ok(!html.includes('<th>日期</th><th>地区</th><th>部门</th></tr></thead><tbody><tr><td>日期</td>'),
        '表头行不应重复出现在表体');

    // 对象数组形态的 data（聚合/统计结果）仍不归这里管 → 返回 null，交回 tryConvertToTable
    assert.strictEqual(tryRowsMatrixTable({ success: true, data: [{ a: 1 }, { a: 2 }] }), null);
});

test('json 通过 chat-renderer.renderer 委托暴露', () => {
    const ctx = loadChatV2();
    assert.strictEqual(typeof ctx.App.ChatV2.renderer.renderJsonHtml, 'function');
    assert.strictEqual(typeof ctx.App.ChatV2.renderer.tryConvertToTable, 'function');
});