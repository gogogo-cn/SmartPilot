/**
 * steps-plan-outline.test.js — renderStepsPlanHtml 步骤渲染契约
 *
 * 【契约演进 v20261003】原「顶部『规划步骤』路线图」绿面板（exec-plan-card，含
 * "规划步骤：1. …　2. …" 顶栏 + "已规划 N 步 · 待执行" 摘要）已**有意废弃**：
 *   - chat-renderer.js L252：旧版绿面板（exec-plan-card / _planHtml / _steps）全量废弃，
 *     thoughtList 的迭代替换模型统一承载所有步骤展示；
 *   - chat-renderer.js L4179-4182：renderStepsPlanHtml 改为输出 cbx 扁平列表，
 *     与 thoughtList 完全同一套视觉。
 * 因此本套件不再断言顶栏路线图，改为守护 **当前 cbx 扁平列表契约**：
 *   - 容器 cbx-list cbx-list--process，每步一个 cbx-item
 *   - 标题落在 cbx-label（整句步骤展示完整句子，不再做顶栏截短）
 *   - 状态：done/success/ok → `cbx-status ok` + ✓；pending → `○ 等待`
 *   - detail 落在 cbx-detail--text
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

function render() {
    const ctx = loadChatV2();
    assert.ok(ctx.App.ChatV2.renderer, 'renderer 应已挂载');
    assert.strictEqual(typeof ctx.App.ChatV2.renderer.renderStepsPlanHtml, 'function', '应导出 renderStepsPlanHtml');
    return ctx.App.ChatV2.renderer.renderStepsPlanHtml;
}

/** 抽取所有步骤行标题（cbx-label 文本） */
function labelsOf(html) {
    return Array.from(html.matchAll(/class="cbx-label">([^<]*)</g)).map(m => m[1]);
}

test('P1-6 后端整句步骤：渲染为 cbx 分步清单（每步一行）', () => {
    const html = render()(['读取数据', '清洗空值', '生成图表']);
    assert.ok(html.includes('cbx-list cbx-list--process'), '应渲染 cbx 步骤列表容器');
    assert.deepStrictEqual(labelsOf(html), ['读取数据', '清洗空值', '生成图表'],
        '每个整句步骤都应成为一行标题');
    assert.strictEqual((html.match(/cbx-item/g) || []).length, 3, '应有 3 个步骤项');
});

test('P1-6 长整句步骤：展示完整句子（顶栏截短已随旧面板废弃）', () => {
    const html = render()(['分析各地区的销售数据走势并给出未来三个月的预测，同时标注异常波动点']);
    const labels = labelsOf(html);
    assert.strictEqual(labels.length, 1, '应有 1 个步骤项');
    assert.ok(labels[0].includes('分析各地区的销售数据走势并给出未来三个月的预测，同时标注异常波动点'),
        'cbx 列表展示完整句子（不再为顶栏做逗号分段截短）');
});

test('P1-6 结构化步骤（title+detail）：title 作标题、detail 进详情区', () => {
    const html = render()([
        { title: '数据清洗', detail: '填充 12 个空值', status: 'done' },
        { title: '聚合统计', detail: '按部门求平均', status: 'done' }
    ]);
    assert.deepStrictEqual(labelsOf(html), ['数据清洗', '聚合统计'], '结构化步骤用 title 作行标题');
    assert.ok(html.includes('填充 12 个空值'), 'detail 应在详情区');
    assert.ok(html.includes('按部门求平均'), 'detail 应在详情区');
    assert.ok(html.includes('cbx-status ok'), 'done 应显示已完成标记');
});

test('P1-6 空 steps：返回空字符串（安全）', () => {
    assert.strictEqual(render()([]), '');
    assert.strictEqual(render()(null), '');
});

test('P7 规划态步骤（通用对话解析的【执行】标注）：pending 显示等待态而非已完成', () => {
    const html = render()([
        { title: '按部门拆分数据到不同Sheet', status: 'pending' },
        { title: '合并选中工作表的数据', status: 'pending' }
    ]);
    assert.deepStrictEqual(labelsOf(html), ['按部门拆分数据到不同Sheet', '合并选中工作表的数据'],
        '步骤标题应出现');
    assert.ok(html.includes('○ 等待'), '待执行步骤应显示等待态');
    assert.ok(!html.includes('cbx-status ok'), '全 pending 不应出现已完成标记（避免"已完成 0"式误导）');
});
