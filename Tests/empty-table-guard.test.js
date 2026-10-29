/**
 * ============================================================
 * Tests/empty-table-guard.test.js — 空表格 / 占位表格清洗回归
 * ============================================================
 * 背景（2026-10-16 实测截图）：
 *   用户说"按销售额从高到低排序"，qwen2.5 执行 excel_sort_range 后总结时
 *   "承诺画表格但填不出数据"——正文输出：
 *       | 地区 | 部门 | 产品 | 销售额 |
 *       | --- | --- | --- | --- |
 *       |  |  |  | $ |
 *   渲染层照单全收 → 用户看到一张空的"表格卡片"（截图右侧 SmartPilot 面板），
 *   下载 CSV / 插入 Excel 也只能得到一张空表，比没有表格更误导。
 *
 * 修复（三层）：
 *   ① 渲染层 markdown.js：sanitizePlaceholderTables 删掉"整行全是占位符"的数据行；
 *      表头+分隔行孤存（整张表无真实数据）时整表删除，只留一行中性提示。
 *   ② 提示词层：prompts.md 通用 prompt / app-bootstrap DEFAULT_PROMPTS.general /
 *      prompt-builder ANALYSIS_TASK_PROMPT 均新增"禁止空表格/占位表格"规则。
 *   ③ 本测试锁死两层行为，防止回归。
 *
 * 运行：node Tests/empty-table-guard.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const path = require('path');
const fs = require('fs');

const ROOT = path.resolve(__dirname, '..');

/** 独立加载 renderers/markdown.js（harness 默认不加载它，且它不依赖重模块） */
function loadMarkdown() {
    const window = {};
    window.ChatLog = { info() {}, warn() {}, error() {}, debug() {} };
    window.App = { ChatV2: {} };
    // 测试环境不加载 render-html.js；给一个最小转义桩，
    // 否则降级渲染路径的 escapeHtml 会整体返回空串（与生产行为不符）
    window.App.ChatV2.renderHtml = {
        escapeHtml: (s) => String(s == null ? '' : s)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    };
    global.window = window;
    globalThis.App = window.App;
    (0, eval)(fs.readFileSync(path.join(ROOT, 'UI', 'js', 'chat-v2', 'ui', 'renderers', 'markdown.js'), 'utf8'));
    const m = window.App.ChatV2.renderers.markdown;
    assert.ok(m && typeof m.sanitizePlaceholderTables === 'function', 'markdown.js 应导出 sanitizePlaceholderTables');
    return m;
}

// ── 截图复现：只有表头 + "$" 占位数据行 ──────────────────────────
const FAKE_SORT_TABLE = [
    '已经按销售额从高到低对数据进行了排序。',
    '',
    '| 地区 | 部门 | 产品 | 销售额 |',
    '| --- | --- | --- | --- |',
    '|  |  |  | $ |',
    '',
    '### 总体情况',
    '已对数据进行了排序。'
].join('\n');

test('截图复现：表头 + "$"占位数据行的假表格被整表清除', () => {
    const { sanitizePlaceholderTables } = loadMarkdown();
    const out = sanitizePlaceholderTables(FAKE_SORT_TABLE);
    assert.ok(!/\|\s*地区\s*\|\s*部门/.test(out), '占位表头不应残留；实际=' + JSON.stringify(out));
    assert.ok(!/\$\s*\|/.test(out), '"$" 占位单元格不应残留；实际=' + JSON.stringify(out));
    assert.ok(!/\|\s*---/.test(out), '分隔行不应残留；实际=' + JSON.stringify(out));
    assert.ok(out.includes('（未生成有效的表格数据）'), '应留一行中性提示说明表格数据缺失');
    // 正文文字必须原样保留
    assert.ok(out.includes('已经按销售额从高到低对数据进行了排序。'), '正文不能被动到');
    assert.ok(out.includes('### 总体情况'), '后续段落不能丢');
});

test('表头+分隔行孤存（无任何数据行）同样被整表清除', () => {
    const { sanitizePlaceholderTables } = loadMarkdown();
    const src = [
        '以下是销售额排名：',
        '',
        '| 排名 | 产品 | 销售额 |',
        '| --- | --- | --- |',
        '',
        '如需明细请让我读取数据。'
    ].join('\n');
    const out = sanitizePlaceholderTables(src);
    assert.ok(!/\|\s*排名\s*\|/.test(out), '只有表头的表格应被删除；实际=' + JSON.stringify(out));
    assert.ok(out.includes('（未生成有效的表格数据）'), '应留中性提示');
    assert.ok(out.includes('如需明细请让我读取数据。'), '周边文字必须保留');
});

test('混合行：只删占位行，真实数据行保留', () => {
    const { sanitizePlaceholderTables } = loadMarkdown();
    const src = [
        '| 地区 | 部门 | 产品 | 销售额 |',
        '| --- | --- | --- | --- |',
        '| 华东 | 销售一部 | 智能台灯 | 480788 |',
        '|  |  |  | $ |',
        '| - | - | - | - |',
        '| 华北 | 销售二部 | 降噪耳机 | 392341 |',
        '| 待填充 | 待填充 | 待填充 | 待填充 |'
    ].join('\n');
    const out = sanitizePlaceholderTables(src);
    assert.ok(out.includes('| 华东 | 销售一部 | 智能台灯 | 480788 |'), '真实数据行必须保留');
    assert.ok(out.includes('| 华北 | 销售二部 | 降噪耳机 | 392341 |'), '真实数据行必须保留');
    assert.ok(!out.includes('$ |'), '占位行必须删除');
    assert.ok(!/\|\s*-\s*\|\s*-\s*\|/.test(out), '全 "-" 占位行必须删除');
    assert.ok(!out.includes('待填充'), '"待填充" 占位行必须删除');
    assert.ok(!out.includes('（未生成有效的表格数据）'), '还有真实行时不应加空表提示');
    // 表头与分隔行保留
    assert.ok(out.includes('| 地区 | 部门 | 产品 | 销售额 |'), '表头应保留');
    assert.ok(out.includes('| --- | --- | --- | --- |'), '分隔行应保留');
});

test('真实表格（含部分空单元格）一字不动', () => {
    const { sanitizePlaceholderTables } = loadMarkdown();
    const src = [
        '### 分组统计',
        '| 分组 | 求和 | 环比 |',
        '| --- | ---: | ---: |',
        '| 销售一部 | 480788 | +12% |',
        '| 销售二部 | 392341 |  |',
        '',
        '结论：销售一部贡献最大。'
    ].join('\n');
    const out = sanitizePlaceholderTables(src);
    assert.strictEqual(out, src, '真实表格不应被改动（含空单元格/对齐分隔符/无首尾竖线变体之外的形态）');
});

test('无首尾竖线的 GFM 表格同样支持', () => {
    const { sanitizePlaceholderTables } = loadMarkdown();
    const src = [
        '地区 | 部门 | 产品 | 销售额',
        '--- | --- | --- | ---',
        '华东 | 销售一部 | 智能台灯 | 480788',
        ' |  |  | $'
    ].join('\n');
    const out = sanitizePlaceholderTables(src);
    assert.ok(out.includes('华东 | 销售一部 | 智能台灯 | 480788'), '真实行保留');
    assert.ok(!/\$\s*$/.test(out.split('\n').pop()), '末行占位行应删除');
});

test('代码围栏内的竖线/伪表格不受影响', () => {
    const { sanitizePlaceholderTables } = loadMarkdown();
    const src = [
        '```python',
        'import pandas as pd',
        '# 伪表格 |',
        'df = pd.DataFrame({"a": [1, 2]})',
        'print(df | df)',
        '```',
        '',
        '| 地区 | 部门 | 产品 | 销售额 |',
        '| --- | --- | --- | --- |',
        '|  |  |  | $ |'
    ].join('\n');
    const out = sanitizePlaceholderTables(src);
    assert.ok(out.includes('```python'), '代码围栏必须保留');
    assert.ok(out.includes('print(df | df)'), '围栏内竖线代码不能被动');
    assert.ok(out.includes('# 伪表格 |'), '围栏内注释不能被动');
    assert.ok(!/\$\s*\|/.test(out), '围栏外的占位表仍应被清除');
});

test('连续多张空表只留一条提示（不刷屏）', () => {
    const { sanitizePlaceholderTables } = loadMarkdown();
    const src = [
        '| A | B |',
        '| --- | --- |',
        '|  |  |',
        '',
        '| C | D |',
        '| --- | --- |',
        '| $ | $ |'
    ].join('\n');
    const out = sanitizePlaceholderTables(src);
    const noteCount = (out.match(/（未生成有效的表格数据）/g) || []).length;
    assert.strictEqual(noteCount, 1, '连续空表只应留一条提示；实际=' + JSON.stringify(out));
});

test('转义竖线 \\| 是单元格内容，不当列分隔', () => {
    const { sanitizePlaceholderTables } = loadMarkdown();
    const src = [
        '| 公式 | 说明 |',
        '| --- | --- |',
        '| =A1\\|B1 | 拼接标识 |'
    ].join('\n');
    const out = sanitizePlaceholderTables(src);
    assert.ok(out.includes('=A1\\|B1'), '含转义竖线的真实行必须保留；实际=' + JSON.stringify(out));
});

test('renderMarkdown 集成：空表格不进最终 HTML', () => {
    const md = loadMarkdown();
    // 测试环境无 marked → 走降级渲染；占位表在 sanitize 阶段就应被清掉
    const html = md.renderMarkdown(FAKE_SORT_TABLE);
    assert.ok(!html.includes('销售额 |'), '表头不应出现在渲染结果；实际=' + JSON.stringify(html));
    assert.ok(!/\$\s*\|/.test(html), '占位单元格不应出现在渲染结果');
    assert.ok(html.includes('（未生成有效的表格数据）'), '提示应出现在渲染结果中');
    // 真实表格的渲染路径不受影响（降级渲染至少保留原文）
    const real = md.renderMarkdown('| 地区 | 销售额 |\n| --- | --- |\n| 华东 | 100 |');
    assert.ok(real.includes('华东'), '真实表格内容应保留');
});

// ── 提示词层：三层规则必须都在 ──────────────────────────────
test('提示词层：prompts.md / app-bootstrap / prompt-builder 均禁止空表格', () => {
    const promptsMd = fs.readFileSync(path.join(ROOT, 'UI', 'js', 'prompts.md'), 'utf8');
    assert.ok(/【表格输出规范 - 强制】/.test(promptsMd), 'prompts.md 应有【表格输出规范 - 强制】章节');
    assert.ok(promptsMd.includes('占位'), 'prompts.md 规则应点名占位表格');
    assert.ok(promptsMd.includes('表格数据必须来自工具返回的真实结果，禁止编造行列数值'), 'prompts.md 应禁止编造表格数据');

    const bootstrap = fs.readFileSync(path.join(ROOT, 'UI', 'js', 'app-bootstrap.js'), 'utf8');
    assert.ok(bootstrap.includes('禁止输出空表格/占位表格'), 'DEFAULT_PROMPTS.general 兜底提示词应含空表格禁令');

    const pb = fs.readFileSync(path.join(ROOT, 'UI', 'js', 'chat-v2', 'services', 'prompt-builder.js'), 'utf8');
    assert.ok(pb.includes('禁止 输出空表格/占位表格'), 'ANALYSIS_TASK_PROMPT 应含空表格禁令');
});

test('工具推荐：excel_sort_range 成功后给"查看前10行"入口且不重复触发排序', () => {
    const { loadChatV2 } = require('./helpers/harness');
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;
    assert.ok(tr && typeof tr.getRecommendedOps === 'function', 'toolRegistry 应暴露 getRecommendedOps');

    const ops = tr.getRecommendedOps('excel-local__excel_sort_range', {
        range: '$A$1:$I$101',
        sortBy: JSON.stringify([{ column: '销售额', order: 'descending' }]),
        hasHeader: true
    }, '成功排序范围: $A$1:$I$101');
    const view = ops.find(o => o.label === '查看前10行');
    assert.ok(view, '应提供"查看前10行"推荐；实际=' + JSON.stringify(ops));
    assert.ok(view.send.includes('前10行'), '推荐消息应明确前10行；实际=' + view.send);
    assert.ok(view.send.includes('销售额'), '中文列名应带入推荐消息；实际=' + view.send);
    // 文案必须避开"排序/按/排名"等词：避免再次触发排序意图，并命中"仅展示"分支
    assert.ok(!/排序|升序|降序|排名/.test(view.send), '推荐消息不应含排序类动词（防重复执行）；实际=' + view.send);
    assert.ok(!/(分析|总结|汇总|统计|对比|排名|趋势|占比|分组|聚合)/.test(view.send),
        '推荐消息不应含分析类关键词（保证走"仅展示数据"分支）；实际=' + view.send);
    assert.ok(ops.some(o => o.label === '恢复原顺序'), '原有"恢复原顺序"推荐应保留');

    // 字母列号 → 通用文案，不出现裸列号
    const ops2 = tr.getRecommendedOps('excel_sort_range', {
        sortBy: JSON.stringify([{ column: 'E', order: 'descending' }])
    }, 'ok');
    const view2 = ops2.find(o => o.label === '查看前10行');
    assert.ok(view2 && !view2.send.includes(' E '), '字母列号应走通用文案；实际=' + JSON.stringify(ops2));
});
