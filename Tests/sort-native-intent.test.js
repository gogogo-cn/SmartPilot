/**
 * ============================================================
 * Tests/sort-native-intent.test.js — 排序工具"还原原生对话框"回归
 * ============================================================
 * 背景（2026-10-16 用户反馈 + 实测）：
 *   excel_sort_range 没有还原 Excel 原生「排序」对话框的作业方法。实测发现两件事：
 *   ① 旧实现把 SortFields.Add 的 SortOn 硬编码为 1 —— 本机 Excel 16.0 b20430 上
 *      1 = xlSortOnCellColor（按单元格颜色）且未指定颜色 → **排序完全空操作**
 *      （独立 Excel 实例复现：排序后数据一字未动；SortOn=0 才真正按值排序）。
 *   ② 原生对话框的「排序依据」支持 单元格值/单元格颜色/字体颜色/条件格式图标，
 *      「选项」支持 区分大小写/方向（按列|按行）/排序方法（拼音|笔划），工具全都不暴露。
 *
 * 本测试锁死：
 *   A. VB 实现：SortOn 用枚举（xlSortOnValues=0）而非硬编码 1；暴露 sortOn/color/
 *      iconSet/icon/matchCase/orientation/sortMethod/previewRows 参数；颜色经
 *      SortField.SortOnValue 的 "Color" 成员反射设置；排序后回传 previewRows 预览。
 *   B. 工具定义：Config/excel_tools.json 与 VB ToolDefinitions 双方都声明新参数。
 *   C. 前端意图解析：resolveIntentFlagsFromMessage 能从自然语言解析
 *      "把红色的排前面/按字体颜色排序/区分大小写/按行排序/按笔划" 等意图。
 *   D. 工具卡片：describeSortKeys 能展示排序依据（颜色/图标）。
 *
 * 运行：node Tests/sort-native-intent.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const path = require('path');
const fs = require('fs');

const ROOT = path.resolve(__dirname, '..');
const VB_DATAOPS = path.join(ROOT, 'Services', 'Mcp', 'ExcelMcpTools.DataOps.vb');
const VB_TOOLDEF = path.join(ROOT, 'Services', 'Mcp', 'ExcelMcpTools.ToolDefinitions.Data.vb');
const JSON_TOOLS = path.join(ROOT, 'Config', 'excel_tools.json');

function read(p) { return fs.readFileSync(p, 'utf8'); }

/** 取 ExecuteSortRange 函数体及其辅助函数（到下一个工具实现 ExecuteApplyFilter 之前） */
function sortFnBody() {
    const src = read(VB_DATAOPS);
    const start = src.indexOf('Private Function ExecuteSortRange');
    assert.ok(start >= 0, '找不到 ExecuteSortRange');
    const end = src.indexOf('Private Function ExecuteApplyFilter', start);
    assert.ok(end > start, '找不到 ExecuteSortRange 辅助区的结束位置');
    return src.slice(start, end);
}

test('VB：SortOn 必须用枚举，禁止再硬编码 1（=xlSortOnCellColor 空操作）', () => {
    const body = sortFnBody();
    assert.ok(/CInt\(sortOn\)/.test(body), 'SortFields.Add 必须传 CInt(sortOn) 枚举值');
    assert.ok(!/SortFields\.Add\([^)]*,\s*1\s*,/.test(body),
        '不得再硬编码 SortOn=1（本机实测=按单元格颜色且不带颜色→空操作）');
    assert.ok(/XlSortOn\.xlSortOnValues/.test(body), '应引用 xlSortOnValues 枚举');
    assert.ok(/XlSortOn\.xlSortOnCellColor/.test(body), '应支持 xlSortOnCellColor');
    assert.ok(/XlSortOn\.xlSortOnFontColor/.test(body), '应支持 xlSortOnFontColor');
    assert.ok(/XlSortOn\.xlSortOnIcon/.test(body), '应支持 xlSortOnIcon');
});

test('VB：颜色排序经 SortOnValue 的 Color 成员设置（非 ColorFormat.RGB / 非整体赋值）', () => {
    const body = sortFnBody();
    assert.ok(/InvokeMember\("Color"/.test(body) && /SetProperty/.test(body),
        '应反射后期绑定设置 SortOnValue 的 "Color"（本机实测唯一可用路径）');
    assert.ok(/SortOnValue/.test(body), '应读取 SortField.SortOnValue');
});

test('VB：排序后回传预览数据（headers/previewRows），避免 Agent 只能输出占位表格', () => {
    const body = sortFnBody();
    assert.ok(/\.previewRows\s*=/.test(body), '返回载荷应含 previewRows');
    assert.ok(/\.headers\s*=/.test(body), '返回载荷应含 headers');
    assert.ok(/previewRows\b/.test(body) && /GetArg\(Of Integer\)\(args, "previewRows"/.test(body),
        '应读取 previewRows 参数（默认 5）');
});

test('VB：原生"选项"对话框三要素（matchCase/orientation/sortMethod）每次显式重置', () => {
    const body = sortFnBody();
    assert.ok(/sortObj\.MatchCase\s*=\s*matchCase/.test(body), '应设置 MatchCase');
    assert.ok(/sortObj\.Orientation\s*=/.test(body), '应显式设置 Orientation（防上一次按行排序残留）');
    assert.ok(/sortObj\.SortMethod\s*=/.test(body), '应支持 SortMethod（拼音/笔划）');
    assert.ok(/xlSortRows/.test(body) && /xlSortColumns/.test(body), '应支持按行/按列两种方向');
});

test('VB：图标排序用 SetIcon，图标集序号由 XlIconSet 枚举解析', () => {
    const body = sortFnBody();
    assert.ok(/SetIcon\(/.test(body), '应调用 SortField.SetIcon');
    assert.ok(/XlIconSet/.test(body), '图标集名称应经 XlIconSet 枚举解析');
});

test('Python 桥：xlYes/xlNo 常量必须正确（旧值 -4163/-4152 会让 Sort.Header 直接 E_INVALIDARG）', () => {
    const src = read(path.join(ROOT, 'Tools', 'py-excel-bridge.py'));
    const yes = src.match(/^XL_YES\s*=\s*(-?\d+)/m);
    const no = src.match(/^XL_NO\s*=\s*(-?\d+)/m);
    assert.ok(yes && no, '桥应定义 XL_YES / XL_NO');
    assert.strictEqual(yes[1], '1', 'XL_YES 必须为 1（xlYes）；旧值 -4163 是 xlUp，实测令 Sort.Header 抛 E_INVALIDARG');
    assert.strictEqual(no[1], '2', 'XL_NO 必须为 2（xlNo）');
});

test('Python 桥：m_excel_sort_range 还原原生对话框并回传预览', () => {
    const src = read(path.join(ROOT, 'Tools', 'py-excel-bridge.py'));
    const start = src.indexOf('def m_excel_sort_range');
    assert.ok(start >= 0, '桥应有 m_excel_sort_range');
    const end = src.indexOf('\ndef ', start + 10);
    const fn = src.slice(start, end > 0 ? end : start + 8000);
    assert.ok(/ws\.Sort/.test(fn), '应走现代 Sort 对象路径（SortFields/SetRange/Apply）');
    assert.ok(/SortFields\.Add\(/.test(fn), '应使用 SortFields.Add 传 SortOn 枚举');
    assert.ok(/SortOnValue\.Color\s*=/.test(fn), '颜色应经 SortOnValue.Color 设置');
    assert.ok(/SetIcon\(/.test(fn), '应支持 SetIcon 图标排序');
    assert.ok(/previewRows/.test(fn), '应回传 previewRows 预览数据');
    assert.ok(/headers/.test(fn), '应回传 headers');
    assert.ok(/MatchCase/.test(fn) && /Orientation/.test(fn) && /SortMethod/.test(fn),
        '应支持原生选项三要素');
    // 纯按值请求保留 legacy 回退（本环境长期验证过的路径）
    assert.ok(/_apply_legacy/.test(fn), '纯按值场景应保留 legacy 回退');
});

test('工具定义：JSON 与 VB 双方都声明 sortOn/color/iconSet/icon/选项参数', () => {
    const root = JSON.parse(read(JSON_TOOLS).replace(/^\uFEFF/, ''));
    const tools = Array.isArray(root) ? root : (root.tools || []);
    const t = tools.find(x => x.name === 'excel_sort_range');
    assert.ok(t, 'excel_tools.json 应有 excel_sort_range');
    const names = t.parameters.map(p => p.name);
    for (const p of ['sortBy', 'hasHeader', 'matchCase', 'orientation', 'sortMethod', 'previewRows']) {
        assert.ok(names.includes(p), `excel_tools.json 缺少参数 ${p}`);
    }
    assert.ok(/sortOn/.test(t.description), 'description 应说明 sortOn');
    assert.ok(/单元格颜色/.test(t.description) && /字体颜色/.test(t.description) && /条件格式图标/.test(t.description),
        'description 应列出原生对话框的四种排序依据');
    const sortByParam = t.parameters.find(p => p.name === 'sortBy');
    assert.ok(/iconSet/.test(sortByParam.description) && /color/.test(sortByParam.description),
        'sortBy 参数说明应包含 color/iconSet');

    const vbDef = read(VB_TOOLDEF);
    const seg = vbDef.slice(vbDef.indexOf('New McpTool("excel_sort_range"'), vbDef.indexOf('New McpTool("excel_apply_filter"'));
    for (const p of ['sortBy', 'matchCase', 'orientation', 'sortMethod', 'previewRows']) {
        assert.ok(seg.includes(`{"${p}"`), `VB 工具定义缺少参数 ${p}`);
    }
    assert.ok(/xlSortOnValues=0|单元格值/.test(seg) || /单元格值/.test(seg), 'VB 定义应说明排序依据');
});

test('前端：意图解析——把红色的排前面 → cellColor + color', () => {
    const { loadChatV2 } = require('./helpers/harness');
    const { App } = loadChatV2();
    const r = App.ChatV2.toolRegistryUtil.resolveIntentFlagsFromMessage(
        { userMessage: '把红色的行排到前面' }, 'excel_sort_range');
    assert.ok(r && r.sortBy, '应解析出 sortBy；实际=' + JSON.stringify(r));
    const key = JSON.parse(r.sortBy)[0];
    assert.strictEqual(key.sortOn, 'cellColor');
    assert.strictEqual(key.color, '#FF0000');
});

test('前端：意图解析——按字体颜色排序 → fontColor', () => {
    const { loadChatV2 } = require('./helpers/harness');
    const { App } = loadChatV2();
    const r = App.ChatV2.toolRegistryUtil.resolveIntentFlagsFromMessage(
        { userMessage: '按字体颜色排序，蓝色在前' }, 'excel_sort_range');
    const key = JSON.parse(r.sortBy)[0];
    assert.strictEqual(key.sortOn, 'fontColor');
    assert.strictEqual(key.color, '#0000FF');
});

test('前端：意图解析——hex 颜色与"按行排序/区分大小写/按笔划"选项', () => {
    const { loadChatV2 } = require('./helpers/harness');
    const { App } = loadChatV2();
    const r = App.ChatV2.toolRegistryUtil.resolveIntentFlagsFromMessage(
        { userMessage: '按单元格颜色排序 #00FF00，区分大小写，按行排序，按笔划' }, 'excel_sort_range');
    const key = JSON.parse(r.sortBy)[0];
    assert.strictEqual(key.sortOn, 'cellColor');
    assert.strictEqual(key.color, '#00FF00');
    assert.strictEqual(r.matchCase, true);
    assert.strictEqual(r.orientation, 'rows');
    assert.strictEqual(r.sortMethod, 'stroke');
});

test('前端：普通"按销售额降序"把列名交给后端，而不是默认成 A 列', () => {
    const { loadChatV2 } = require('./helpers/harness');
    const { App } = loadChatV2();
    const util = App.ChatV2.toolRegistryUtil;
    const r = util.resolveIntentFlagsFromMessage(
        { userMessage: '按销售额从高到低排序' }, 'excel_sort_range');
    const key = JSON.parse(r.sortBy)[0];
    // 实测：未认出「销售额」时 column 落成 A，整表按第一列（日期）降序。
    // 列名原样下发，VB/Python 按表头解析成真实列。
    assert.strictEqual(key.column, '销售额');
    assert.strictEqual(key.order, 'descending');
    assert.ok(!key.sortOn, '普通排序不应误加 sortOn');
    assert.ok(!r.matchCase && !r.orientation && !r.sortMethod, '不应误加选项参数');

    const quoted = util.resolveIntentFlagsFromMessage(
        { userMessage: '按「利润」升序' }, 'excel_sort_range');
    const qk = JSON.parse(quoted.sortBy)[0];
    assert.strictEqual(qk.column, '利润');
    assert.strictEqual(qk.order, 'ascending');
});

test('前端：模型把 sortBy 写成示例里的 A 列时，用用户点名的列覆盖', () => {
    const { loadChatV2 } = require('./helpers/harness');
    const { App } = loadChatV2();
    const util = App.ChatV2.toolRegistryUtil;
    const args = { range: '$A$1:$I$101', sortBy: JSON.stringify([{ column: 'A', order: 'descending' }]) };
    const changed = util.applyExplicitSortColumn(args, { userMessage: '按销售额从高到低排序' });
    assert.strictEqual(changed, true);
    const key = JSON.parse(args.sortBy)[0];
    assert.strictEqual(key.column, '销售额');
    assert.strictEqual(key.order, 'descending');

    // 已经是对的列名时不改写
    const ok = { sortBy: JSON.stringify([{ column: '销售额', order: 'descending' }]) };
    assert.strictEqual(util.applyExplicitSortColumn(ok, { userMessage: '按销售额从高到低排序' }), false);

    // 没点名列（只说降序）时，不把模型选的列改掉
    const bare = { sortBy: JSON.stringify([{ column: 'C', order: 'descending' }]) };
    assert.strictEqual(util.applyExplicitSortColumn(bare, { userMessage: '降序排列' }), false);

    // 执行层必须接到这个纠正（漏传和写错 A 两条路都覆盖）
    const reg = fs.readFileSync(path.join(ROOT, 'UI', 'js', 'chat-v2', 'services', 'tool-registry.js'), 'utf8');
    assert.ok(reg.includes('applyExplicitSortColumn(effectiveArgs, ctxTurn)'),
        'invokeTracked 应在执行前用用户点名的列纠正 sortBy');
});

test('前端：工具卡片展示排序依据（颜色/图标/次要关键字）', () => {
    const { loadChatV2 } = require('./helpers/harness');
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;
    const sortBy = JSON.stringify([
        { column: '状态', sortOn: 'cellColor', color: '#FF0000', order: 'ascending' },
        { column: '部门', order: 'descending' }
    ]);
    const call = tr.renderCardIntent('excel_sort_range', 'call', { sortBy });
    assert.ok(call && call.content, 'renderCardIntent(call) 应返回 content；实际=' + JSON.stringify(call));
    assert.ok(call.content.includes('单元格颜色 #FF0000'), call.content);
    assert.ok(call.content.includes('次要关键字'), call.content);
    assert.ok(call.content.includes('部门 降序'), call.content);
    // 图标
    const call2 = tr.renderCardIntent('excel_sort_range', 'call', { sortBy: JSON.stringify([{ column: 'A', sortOn: 'icon', order: 'ascending' }]) });
    assert.ok(call2.content.includes('条件格式图标'), call2.content);
});
