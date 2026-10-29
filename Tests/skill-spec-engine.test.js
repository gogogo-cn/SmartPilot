/**
 * ============================================================
 * Tests/skill-spec-engine.test.js — Skill Spec 规划引擎测试
 * ============================================================
 * 验证核心能力：
 *   1. LLM 输出不同 spec → 渲染出不同但确定的 VBA（灵活性 + 确定性）
 *   2. spec 白名单校验（非法字段/值被忽略，注入防护）
 *   3. 无 LLM 时消息关键词兜底
 *   4. 确定性：相同 spec 永远产出相同代码
 * 运行：node Tests/skill-spec-engine.test.js
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
    (0, eval)(fs.readFileSync(path.join(BASE, 'services/skill-spec-engine.js'), 'utf8'));
    return App.ChatV2.services.skillSpecEngine;
}

test('skill-spec: 不同 spec 渲染出不同 VBA（LLM 规划灵活性）', () => {
    const engine = loadEngine();

    // LLM 规划 A：商务蓝 + 数据条 + 冻结表头
    const specA = {
        theme: 'business_blue', headerStyle: 'bold_colored', alignment: 'center',
        autofit: true, freezeHeader: true, addBorders: true,
        bandedRows: false, conditionalFormat: 'dataBar', numberFormat: null, title: '2023年度销售报表'
    };
    // LLM 规划 B：科技绿 + 斑马纹 + 色阶（headerStyle=bold_only 时表头只加粗不着色，
    // 主题色通过斑马纹带 RGB(0,128,128) 系列体现 —— 用 tech_green 的 band 色验证）
    const specB = {
        theme: 'tech_green', headerStyle: 'bold_only', alignment: 'left',
        autofit: false, freezeHeader: false, addBorders: false,
        bandedRows: true, conditionalFormat: 'colorScale', numberFormat: '#,##0', title: null
    };

    const rA = engine.planAndRender(specA, {});
    const rB = engine.planAndRender(specB, {});

    assert.ok(rA.ok && rB.ok);
    // 代码不同（样式参数化）
    assert.notStrictEqual(rA.code, rB.code, '不同 spec 应渲染不同代码');
    // 主题颜色不同（A 商务蓝表头着色，B 斑马纹带用科技绿）
    assert.ok(rA.code.includes('RGB(0, 90, 156)'), 'A 应为商务蓝表头');
    // 期望值从引擎主题表派生，避免主题配色调整后测试出现"陈旧硬编码"失败
    const greenBand = engine.THEMES.tech_green.band;
    assert.ok(rB.code.includes(`RGB(${greenBand.join(', ')})`), 'B 应为科技绿斑马纹带');
    // 标题不同
    assert.ok(rA.code.includes('2023年度销售报表'), 'A 应有标题行');
    assert.ok(!rB.code.includes('2023年度销售报表'), 'B 无标题行');
    // 条件格式不同
    assert.ok(rA.code.includes('AddDatabar'), 'A 应有数据条');
    assert.ok(rB.code.includes('AddColorScale'), 'B 应有色阶');
    // 冻结表头
    assert.ok(rA.code.includes('FreezePanes'), 'A 应冻结');
    assert.ok(!rB.code.includes('FreezePanes'), 'B 不冻结');
    // 斑马纹
    assert.ok(rB.code.includes('MOD(ROW(),2)=0'), 'B 应有斑马纹');
    console.log('✅ 不同 spec → 不同且正确的 VBA（灵活性）');
});

test('skill-spec: 确定性 — 相同 spec 永远相同代码', () => {
    const engine = loadEngine();
    const spec = { theme: 'elegant_purple', headerStyle: 'italic', alignment: 'right', bandedRows: true };
    const r1 = engine.planAndRender(spec, {});
    const r2 = engine.planAndRender(spec, {});
    assert.strictEqual(r1.code, r2.code, '确定性渲染');
    console.log('✅ 确定性：相同 spec 相同代码');
});

test('skill-spec: 白名单校验 — 非法字段/值被忽略（注入防护）', () => {
    const engine = loadEngine();
    // 恶意/非法 spec：未知字段、非法主题、非法数字格式、注入脚本
    const malicious = {
        theme: 'evil_theme',                    // 非法主题
        headerStyle: 'DROP TABLE',              // 非法表头样式
        numberFormat: '"; MsgBox("hacked", "evil", "rmdir C:\\", ' +  // 注入尝试
            ' # 非法格式串',
        evilField: 'xxx',                       // 未知字段
        conditionalFormat: 'delete_sheet',      // 非法条件格式
        title: 'a'.repeat(500)                  // 超长标题
    };
    const r = engine.planAndRender(malicious, {});
    assert.ok(r.ok, '非法 spec 不应导致渲染失败');
    // theme 不再"非法即回退默认"——任何词都解析为颜色（哈希兜底），
    // 但"evil_theme"这类词渲染的是稳定哈希色而非默认蓝，注入防护仍然生效。
    assert.notStrictEqual(r.spec.theme, 'business_blue', 'evil_theme 应有自己的颜色（非默认）');
    assert.ok(r.spec.headerColor, 'evil_theme 应解析出 headerColor');
    assert.strictEqual(r.spec.headerStyle, 'bold_colored', '非法表头回退默认');
    assert.strictEqual(r.spec.numberFormat, null, '非法数字格式被忽略');
    assert.ok(!('evilField' in r.spec), '未知字段被过滤');
    assert.strictEqual(r.spec.conditionalFormat, null, '非法条件格式被忽略');
    assert.ok(!r.code.includes('MsgBox("hacked"'), '无注入脚本');
    console.log('✅ 白名单校验：非法值回退/过滤，无注入（theme 走哈希兜底）');
});

test('skill-spec: 无 LLM 时消息颜色词兜底（动态颜色）', () => {
    const engine = loadEngine();
    // "科技感的绿色 + 数据条" → 绿色动态映射 + dataBar
    const spec1 = engine.specFromMessage('用科技感的绿色，给销售额加数据条');
    assert.strictEqual(spec1.theme, 'custom', '颜色词映射为 custom + 动态色');
    assert.deepStrictEqual(spec1.headerColor, [60, 150, 90], '绿色 → RGB(60,150,90)');
    assert.strictEqual(spec1.conditionalFormat, 'dataBar');
    // 默认兜底（无颜色词）
    const spec2 = engine.specFromMessage('美化表格');
    assert.strictEqual(spec2.theme, 'business_blue');
    assert.strictEqual(spec2.conditionalFormat, null);
    assert.strictEqual(spec2.bodyColor, null, '只说美化时不给表身派生底色');
    const body = engine.specFromMessage('美化表格，表身浅黄，隔行变色');
    assert.ok(body.bodyColor, '点名表身才写 bodyColor');
    assert.strictEqual(body.bandedRows, true);
    const rendered = engine.planAndRender(spec2, {});
    assert.ok(rendered.code.includes('.Interior.Pattern = xlNone'), '默认数据区不上底色');
    assert.ok(!/dataRng[\s\S]{0,400}\.Interior\.Color =/.test(rendered.code), '默认数据区不写 Interior.Color');
    // 不同颜色词 → 不同颜色
    const red = engine.specFromMessage('红色主题');
    const gold = engine.specFromMessage('金色风格');
    assert.notDeepStrictEqual(red.headerColor, gold.headerColor, '不同颜色词应映射不同颜色');
    console.log('✅ 消息颜色兜底：绿→' + JSON.stringify(spec1.headerColor) + ' 红→' + JSON.stringify(red.headerColor) + ' 金→' + JSON.stringify(gold.headerColor));
});

test('skill-spec: buildPlanningPrompt 引导 LLM 输出 JSON 而非代码', () => {
    const engine = loadEngine();
    const prompt = engine.buildPlanningPrompt({ displayName: '智能表格美化' });
    assert.ok(prompt.includes('绝不输出代码'), 'prompt 应禁止写代码');
    assert.ok(prompt.includes('theme'), 'prompt 应包含 spec 字段');
    assert.ok(prompt.includes('conditionalFormat'), 'prompt 应包含条件格式字段');
    assert.ok(prompt.includes('只输出一个 JSON'), 'prompt 应要求 JSON');
    console.log('✅ 规划 prompt：引导 LLM 输出结构化 spec');
});

test('skill-spec: 渲染代码结构完整可执行', () => {
    const engine = loadEngine();
    const r = engine.planAndRender({}, {}); // 默认 spec
    assert.ok(r.code.startsWith(`' ====`), '注释头');
    assert.ok(r.code.includes('Sub SmartBeautifyTable()'), 'Sub 入口');
    assert.ok(r.code.includes('End Sub'), 'Sub 结尾');
    assert.ok(r.code.includes('On Error GoTo ErrHandler'), '错误处理');
    assert.ok(r.code.includes('Application.ScreenUpdating = scrUpd'), '恢复 ScreenUpdating');
    console.log('✅ 渲染代码结构完整（Sub/错误处理/资源恢复）');
});

test('skill-spec: 多行表头 — 表头内容识别 + 向上扩展，整段套表头样式', () => {
    const engine = loadEngine();
    // 默认 spec：渲染含内容识别（IsHeaderRow）与向上扩展（extRows）
    const dflt = engine.planAndRender({}, {});
    assert.ok(dflt.spec.headerRows === 1, '默认 headerRows=1');
    assert.ok(dflt.code.includes('IsHeaderRow'), '渲染含内容识别辅助函数');
    assert.ok(dflt.code.includes('extRows'), '渲染含向上扩展逻辑');
    assert.ok(dflt.code.includes('selRng.Rows("1:" & hr)'), '表头 = 前 hr 行');
    assert.ok(dflt.code.includes('selRng.Offset(hr, 0)'), '数据区自表头行之后起始');
    assert.ok(dflt.code.includes('MergeCells'), '内容识别含合并标题判定');
    // 多行表头（headerRows=2）spec 被保留；且无 title 时不插标题行
    const multi = engine.planAndRender({ headerRows: 2 }, {});
    assert.strictEqual(multi.spec.headerRows, 2, 'headerRows=2 被保留');
    assert.ok(!multi.code.includes('Rows(selRng.Row).Insert'), '多行表头不重复插入标题行');
    // 非法 headerRows 被抑制（>5 / 非数字 → 回退默认 1）
    const tooBig = engine.planAndRender({ headerRows: 99 }, {});
    assert.strictEqual(tooBig.spec.headerRows, 1, 'headerRows 超过上限回退 1');
    const bad = engine.planAndRender({ headerRows: 'abc' }, {});
    assert.strictEqual(bad.spec.headerRows, 1, 'headerRows 非数字回退 1');
    // 兜底消息关键词：多行表头 → 2
    const guess = engine.specFromMessage('多行表头美化');
    assert.strictEqual(guess.headerRows, 2, '消息含"多行表头"→ headerRows=2');
    // 标题插行仅在"单行表头且非合并标题"时发生
    assert.ok(dflt.code.includes('MergeCells'), '渲染含首行合并标题判定');
    const withTitle = engine.planAndRender({ title: '员工月度提成计算表' }, {});
    assert.ok(withTitle.code.includes('Rows(selRng.Row).Insert'), '单行表头+title 才插标题行');
    assert.ok(withTitle.code.includes('MergeCells'), '插标题行也走首行合并检测');
    console.log('✅ 多行表头：内容识别 + 向上扩展 + 整段表头样式');
});

test('skill-spec: 动态颜色 — LLM 输出颜色词/RGB → 渲染真实颜色（换色词有变化）', () => {
    const engine = loadEngine();
    // 用户换颜色词 → 渲染不同颜色（本次修复核心）
    const colorCases = {
        'red': 'RGB(220, 60, 60)',
        'gold': 'RGB(212, 175, 55)',
        'pink': 'RGB(230, 130, 170)',
        '紫色': 'RGB(130, 90, 180)',
        '[255,200,0]': 'RGB(255, 200, 0)',
        'green': 'RGB(60, 150, 90)',
        'emerald': 'RGB(16, 140, 80)',
        'lavender': 'RGB(140, 110, 200)'
    };
    const renderedColors = new Set();
    for (const [color, expectRgb] of Object.entries(colorCases)) {
        const r = engine.planAndRender({ theme: color }, {});
        const m = r.code.match(/\.Interior\.Color = (RGB\([^)]+\))/);
        assert.ok(m, '应渲染表头颜色: ' + color);
        assert.strictEqual(m[1], expectRgb, color + ' 应映射为 ' + expectRgb);
        renderedColors.add(m[1]);
    }
    // 每个颜色词渲染出不同颜色
    assert.strictEqual(renderedColors.size, Object.keys(colorCases).length, '不同颜色词必须渲染不同颜色');
    console.log('✅ 动态颜色：' + renderedColors.size + ' 种颜色词 → ' + [...renderedColors].join(', '));
});

test('skill-spec: 任意颜色词（含词表外）都有颜色，绝不回退默认', () => {
    const engine = loadEngine();
    // 词表外/任意描述的颜色词 → 哈希兜底生成稳定颜色（不是默认蓝）
    const arbitrary = ['任意自定义色', 'custom_palette', 'my_favorite_shade', 'sunset', 'random color xyz'];
    const DEFAULT_BLUE = 'RGB(0, 90, 156)';
    for (const color of arbitrary) {
        const r = engine.planAndRender({ theme: color }, {});
        const m = r.code.match(/\.Interior\.Color = (RGB\([^)]+\))/);
        assert.ok(m, color + ' 应有颜色');
        assert.notStrictEqual(m[1], DEFAULT_BLUE, color + ' 不应回退默认蓝');
        console.log('  ' + color + ' → ' + m[1]);
    }
    // 确定性：同一词两次渲染同色
    const r1 = engine.planAndRender({ theme: '任意自定义色' }, {});
    const r2 = engine.planAndRender({ theme: '任意自定义色' }, {});
    assert.strictEqual(r1.code, r2.code, '同一颜色词确定性');
    console.log('✅ 任意颜色词哈希兜底：' + arbitrary.length + ' 个词全部有颜色且非默认');
});

test('skill-spec: normalizeCleanSpec 参数容错（dedupeColumns 单值、range 数字）', () => {
    const engine = loadEngine();
    // dedupeColumns 单值字符串 → 包装成数组（含去空格）
    const single = engine.normalizeCleanSpec({ operations: ['removeDuplicates'], dedupeColumns: ' 部门 ' });
    assert.deepStrictEqual(single.dedupeColumns, ['部门'], '单值字符串应包装成数组并去空格');
    // dedupeColumns 数组 → 过滤非法 + 截断
    const arr = engine.normalizeCleanSpec({ operations: ['removeDuplicates'], dedupeColumns: ['部门', '地区', 42, null, ''] });
    assert.deepStrictEqual(arr.dedupeColumns, ['部门', '地区'], '数组应过滤非字符串');
    // range 数字 → 转字符串
    const numRange = engine.normalizeCleanSpec({ operations: ['fillEmpty'], range: 42 });
    assert.strictEqual(numRange.range, '42', 'range 数字应转字符串');
    // range 字符串保留
    const strRange = engine.normalizeCleanSpec({ operations: ['fillEmpty'], range: 'Sheet1!A1:B10' });
    assert.strictEqual(strRange.range, 'Sheet1!A1:B10', 'range 字符串应保留');
    // 非法 range（对象/超大数字）→ 丢弃保持默认 null
    const badRange = engine.normalizeCleanSpec({ operations: ['fillEmpty'], range: { a: 1 } });
    assert.strictEqual(badRange.range, null, '对象 range 应丢弃');
    const hugeRange = engine.normalizeCleanSpec({ operations: ['fillEmpty'], range: Infinity });
    assert.strictEqual(hugeRange.range, null, 'Infinity range 应丢弃');
    console.log('✅ 清洗 spec 参数容错：dedupeColumns 单值/数组、range 数字/字符串/非法值');
});

if (require.main === module) { setTimeout(() => process.exit(0), 4000).unref(); } // 仅独立运行时强制退出；聚合运行（run-tests.js require）不注册，避免进程提前退出
