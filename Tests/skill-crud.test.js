'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2, DEFAULT_ORDER } = require('./helpers/harness');
const ORDER = DEFAULT_ORDER.slice(0, DEFAULT_ORDER.indexOf('ui/chat-renderer.js'));

function load() {
    const { App } = loadChatV2(ORDER);
    return App.ChatV2;
}

test('routing-constants: isCrudTool 识别增删改查工具', () => {
    const rc = load().routingConstants;
    assert.ok(Array.isArray([...rc.CRUD_TOOL_NAMES]) || rc.CRUD_TOOL_NAMES instanceof Set);
    // 读/写/增/删/清五族各抽查一个
    assert.equal(rc.isCrudTool('excel_read_range'), true);
    assert.equal(rc.isCrudTool('excel_set_cell_value'), true);
    assert.equal(rc.isCrudTool('excel_insert_row'), true);
    assert.equal(rc.isCrudTool('excel_delete_column'), true);
    assert.equal(rc.isCrudTool('excel_clear_range'), true);
    assert.equal(rc.isCrudTool('excel_clear_format'), true);
    // 非 CRUD 工具不受影响
    assert.equal(rc.isCrudTool('excel_create_chart'), false);
    assert.equal(rc.isCrudTool('excel_merge_cells'), false);
    assert.equal(rc.isCrudTool(''), false);
});

test('crudSpecFromMessage: 读取范围', () => {
    const se = load().services.skillSpecEngine;
    const spec = se.crudSpecFromMessage('读取A1:B10的数据');
    assert.equal(spec && spec.operation, 'read');
    assert.equal(spec.range, 'A1:B10');
});

test('crudSpecFromMessage: 写入单格带值', () => {
    const se = load().services.skillSpecEngine;
    const spec = se.crudSpecFromMessage('把A1改成张三');
    assert.equal(spec && spec.operation, 'write');
    assert.equal(spec.range, 'A1');
    assert.equal(spec.value, '张三');
});

test('crudSpecFromMessage: 清空仅格式', () => {
    const se = load().services.skillSpecEngine;
    const spec = se.crudSpecFromMessage('清除A1:B10的格式');
    assert.equal(spec && spec.operation, 'clear');
    assert.equal(spec.clearContent, false);
    assert.equal(spec.clearFormat, true);
    assert.equal(spec.range, 'A1:B10');
});

test('crudSpecFromMessage: 插入行', () => {
    const se = load().services.skillSpecEngine;
    const spec = se.crudSpecFromMessage('在第3行插入两行');
    assert.equal(spec && spec.operation, 'insert');
    assert.equal(spec.row, 3);
    assert.equal(spec.count, 2);
});

test('crudSpecFromMessage: 排除其它 Skill 专属意图（清洗/图表/合并/排序）', () => {
    const se = load().services.skillSpecEngine;
    assert.equal(se.crudSpecFromMessage('清洗这份数据去重'), null);
    assert.equal(se.crudSpecFromMessage('画个柱状图'), null);
    assert.equal(se.crudSpecFromMessage('合并单元格'), null);
    assert.equal(se.crudSpecFromMessage('按销售额排序'), null);
    assert.equal(se.crudSpecFromMessage('新建工作表'), null);
});

test('crudSpecFromMessage: normalizeCrudSpec 归一操作族英文说法', () => {
    const se = load().services.skillSpecEngine;
    const spec = se.normalizeCrudSpec({ operation: 'update', range: 'A1', value: 'x' });
    assert.equal(spec.operation, 'write');
    const addSpec = se.normalizeCrudSpec({ operation: 'add', row: 2 });
    assert.equal(addSpec.operation, 'insert');
});