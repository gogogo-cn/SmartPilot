'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');

const BASE = path.join(__dirname, '..', 'UI', 'js', 'chat-v2');

function loadUtil() {
    const window = {};
    window.App = { ChatV2: { skillPipelineModules: {} }, config: {} };
    global.window = window;
    globalThis.App = window.App;
    (0, eval)(fs.readFileSync(path.join(BASE, 'pipelines/skill/skill-shared-util.js'), 'utf8'));
    return App.ChatV2.skillPipelineModules.util;
}

test('util: detectHeaderRows — 内容识别多行表头（合并标题行+列头行）', () => {
    const util = loadUtil();
    // 场景①：合并标题行 + 列头行 + 数据 → 2
    const rows = [
        ['员工月度提成计算表', '', '', '', '', '', ''],
        ['员工编号', '员工姓名', '所属部门', '月度业绩(元)', '基础提成(元)', '额外奖惩(元)', '已验证2'],
        ['EMP001', '员工1', '销售一部', '120000', '3600', '0', '3600'],
        ['EMP002', '员工2', '销售二部', '98000', '2940', '-500', '2440']
    ];
    assert.strictEqual(util.detectHeaderRows(rows), 2, '合并标题+列头 → 表头 2 行');

    // 场景②：仅单行列头 → 1
    const single = [
        ['员工编号', '员工姓名', '所属部门', '月度业绩(元)'],
        ['EMP001', '员工1', '销售一部', '120000']
    ];
    assert.strictEqual(util.detectHeaderRows(single), 1, '单行列头 → 表头 1 行');

    // 场景③：数据行不前导表头（如只给数据）→ 兜底 1
    const dataOnly = [['EMP001', '员工1', '销售一部', '120000'], ['EMP002', '员工2', '销售二部', '98000']];
    assert.strictEqual(util.detectHeaderRows(dataOnly), 1, '纯数据 → 兜底 1');

    // 场景④：两级纯文本列头（罕见层级标题+子列头）→ 文本全占多数仍计为表头
    const twoText = [
        ['一级类别', '一级类别', '二级明细', '二级明细'],
        ['员工编号', '员工姓名', '月度业绩(元)', '基础提成(元)'],
        ['EMP001', '员工1', '120000', '3600']
    ];
    assert.strictEqual(util.detectHeaderRows(twoText), 2, '两级文本表头 → 2');
    console.log('✅ detectHeaderRows：合并标题+列头→2 / 单行→1 / 纯数据兜底1 / 两级文本→2');
});