/**
 * tool-card.test.js — renderers/tool-card 纯函数子模块单测
 *   解包 / 摘要 / 参数截断 / 错误格式化 / 技能摘要提取。
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

function load() {
    const ctx = loadChatV2();
    const t = ctx.App.ChatV2.renderers.toolCard;
    assert.ok(t, 'renderers.toolCard 应已挂载');
    return t;
}

test('tool-card unwrapToolResult：多层 result 包裹解包', () => {
    const { unwrapToolResult } = load();
    const raw = { result: { result: { success: true, rows: [1, 2] }, taskCompleted: true }, taskCompleted: true };
    const out = unwrapToolResult(raw);
    assert.strictEqual(out.success, true);
    assert.deepStrictEqual(out.rows, [1, 2]);
});

test('tool-card unwrapToolResult：result 内嵌 JSON 字符串', () => {
    const { unwrapToolResult } = load();
    const raw = { success: true, result: '[{"a":1},{"a":2}]' };
    const out = unwrapToolResult(raw);
    assert.deepStrictEqual(out, [{ a: 1 }, { a: 2 }]);
});

test('tool-card unwrapToolResult：MCP content 文本', () => {
    const { unwrapToolResult } = load();
    const raw = { content: [{ type: 'text', text: '{"ok": true}' }] };
    assert.deepStrictEqual(unwrapToolResult(raw), { ok: true });
});

test('tool-card unwrapToolResult：纯字符串原样返回', () => {
    const { unwrapToolResult } = load();
    assert.strictEqual(unwrapToolResult('hello'), 'hello');
});

test('tool-card autoSummarizeToolResult：read_range 数组摘要', () => {
    const { autoSummarizeToolResult } = load();
    const raw = [{ a: 1, b: 2 }, { a: 3, b: 4 }];
    assert.strictEqual(autoSummarizeToolResult('excel-local__read_range', raw), '已读取 2 行 × 2 列数据');
});

test('tool-card autoSummarizeToolResult：read_range 带地址摘要', () => {
    const { autoSummarizeToolResult } = load();
    const raw = { headers: ['a'], rows: [{ a: 1 }], address: 'Sheet1!A1:B2' };
    assert.strictEqual(autoSummarizeToolResult('read_range', raw), '已读取 1 行 × 1 列数据（Sheet1!A1:B2）');
});

test('tool-card autoSummarizeToolResult：工作表/合并/通用成功', () => {
    const { autoSummarizeToolResult } = load();
    assert.strictEqual(autoSummarizeToolResult('create_sheet', { sheetName: 'Sheet2' }), '已创建工作表：Sheet2');
    assert.strictEqual(autoSummarizeToolResult('merge_cells', {}), '单元格已合并');
    assert.strictEqual(autoSummarizeToolResult('foo', { success: true, message: '完成' }), '完成');
});

test('tool-card autoSummarizeToolResult：合并工作表 → 统计摘要（不回退原始 JSON）', () => {
    const { autoSummarizeToolResult } = load();
    const raw = {
        success: true,
        sourceSheets: ['大客户部', '销售三部', '市场部', '销售二部', '销售一部'],
        mergedRows: 100,
        dataRows: 100,
        headerPreserved: true,
        usedSelectedSheets: false,
        data: [['日期', '地区', '部门'], ['45353', '华南', '大客户部']]
    };
    assert.strictEqual(autoSummarizeToolResult('excel_merge_sheets', raw),
        '已合并 5 张工作表 → 共 100 行（含表头）');
    // 工具名带 excel-local__ 前缀同样要命中（executeLocalExcelTool 的注册名形态）
    assert.strictEqual(autoSummarizeToolResult('excel-local__merge_sheets', raw),
        '已合并 5 张工作表 → 共 100 行（含表头）');
});

test('tool-card summarizeMergeResult：统计字段成条 + 非合并结果返回 null', () => {
    const { summarizeMergeResult } = load();
    const out = summarizeMergeResult({
        success: true, sourceSheets: ['A', 'B', 'C'], mergedRows: 42,
        headerPreserved: false, targetSheet: '汇总数据'
    });
    assert.strictEqual(out.text, '已合并 3 张工作表 → 共 42 行');
    const map = {};
    out.chips.forEach(c => { map[c.label] = c.value; });
    assert.strictEqual(map['来源工作表'], '3 张');
    assert.strictEqual(map['合并行数'], '42 行');
    assert.strictEqual(map['表头'], '未保留（整表当数据）');
    assert.strictEqual(map['目标工作表'], '汇总数据');

    // 不是合并结果 → 不接管（读取类/清洗类结果仍走各自既有渲染）
    assert.strictEqual(summarizeMergeResult({ success: true, headers: ['a'], result: [[1]] }), null);
    assert.strictEqual(summarizeMergeResult([{ a: 1 }]), null);
    assert.strictEqual(summarizeMergeResult(null), null);
    // 泛化字段（rowCount/sheetCount）不能被误判成"已合并"：清洗/拆分结果也常带这些字段
    assert.strictEqual(summarizeMergeResult({ success: true, rowCount: 25, sheetCount: 3 }), null);
});

test('chat-renderer 渲染：合并结果 → 统计条 + 数据表 + 折叠 JSON（不铺原始 JSON）', () => {
    const ctx = loadChatV2();
    const fmt = ctx.App.ChatV2.renderer.formatToolResultHtml;
    assert.strictEqual(typeof fmt, 'function', 'formatToolResultHtml 应暴露供测试');
    const raw = {
        success: true,
        sourceSheets: ['大客户部', '销售三部', '市场部', '销售二部', '销售一部'],
        mergedRows: 100,
        dataRows: 100,
        headerPreserved: true,
        usedSelectedSheets: false,
        data: [['日期', '地区', '部门'], ['45353', '华南', '大客户部'], ['45354', '华东', '销售三部']]
    };
    const html = fmt(raw, {});
    assert.ok(html.includes('exec-op-summary'), '应有统计条容器');
    assert.ok(html.includes('已合并 5 张工作表 → 共 100 行（含表头）'), '应显示合并统计文案');
    assert.ok(html.includes('来源工作表') && html.includes('5 张'), '应有来源工作表 chip');
    assert.ok(html.includes('<th>日期</th>') && html.includes('<th>地区</th>'), '应把合并后的数据渲染成表格');
    assert.ok(html.includes('查看原始结果（JSON）'), '原始 JSON 保留但收进折叠面板');
    assert.ok(!/<details class="json-collapsible"\s+open/.test(html), '折叠面板默认收起，不抢视觉主位');
    // 原始 JSON 只允许出现在折叠面板内部
    const jsonBlockAt = html.indexOf('<pre class="json-block">');
    const detailsAt = html.indexOf('<details class="json-collapsible"');
    assert.ok(jsonBlockAt > detailsAt && detailsAt >= 0, '原始 JSON 必须收在折叠面板里，不做主展示');

    // 非合并结果不受影响（回落既有渲染路径）
    const other = fmt({ success: true, message: '已保存' }, {});
    assert.ok(!other.includes('exec-op-summary'), '普通结果不应套用合并渲染');
});

test('tool-card summarizeMergeFromRaw：穿透 {result:{…}} 等包裹识别合并结果', () => {
    const { summarizeMergeFromRaw } = load();
    const merge = {
        success: true, sourceSheets: ['A', 'B', 'C'], mergedRows: 30,
        headerPreserved: true, data: [['日期'], ['1']]
    };
    // 实测形态：app-mcp 把本地工具结果再包一层 {result: …}
    const r1 = summarizeMergeFromRaw({ result: merge });
    assert.ok(r1, '应穿透 result 包裹');
    assert.strictEqual(r1.text, '已合并 3 张工作表 → 共 30 行（含表头）');
    assert.strictEqual(r1.source, merge, 'source 应指向含统计字段的那一层（供表格/折叠 JSON 使用）');
    // 其它包裹形态
    assert.ok(summarizeMergeFromRaw({ success: true, data: merge }), '{success,data:{…}} 包裹');
    assert.ok(summarizeMergeFromRaw(JSON.stringify({ result: merge })), 'JSON 字符串');
    assert.ok(summarizeMergeFromRaw({ content: [{ type: 'text', text: JSON.stringify(merge) }] }), 'MCP content 包裹');
    assert.ok(summarizeMergeFromRaw({ result: { result: merge } }), '两层 result');
    // 行矩阵本身不该被误判（100 行读取结果）
    assert.strictEqual(summarizeMergeFromRaw({ headers: ['a'], result: [[1, 2], [3, 4]] }), null);
    assert.strictEqual(summarizeMergeFromRaw(Array.from({ length: 100 }, () => [1, 2, 3])), null);
});

test('chat-renderer 渲染：{result:{合并结果}} 包裹形态也识别（不再摊裸矩阵）', () => {
    const ctx = loadChatV2();
    const fmt = ctx.App.ChatV2.renderer.formatToolResultHtml;
    const merge = {
        success: true,
        sourceSheets: ['大客户部', '销售三部', '市场部', '销售二部', '销售一部'],
        mergedRows: 100, dataRows: 100, headerPreserved: true, usedSelectedSheets: false,
        data: [['日期', '地区', '部门'], ['45353', '华南', '大客户部']]
    };
    const html = fmt({ result: merge }, {});
    assert.ok(html.includes('exec-op-summary'), '包裹形态也应命中操作结果渲染');
    assert.ok(html.includes('已合并 5 张工作表 → 共 100 行（含表头）'), '应显示合并统计文案');
    assert.ok(html.includes('<th>日期</th>'), '应渲染数据表（不是裸 JSON 矩阵）');
    assert.ok(html.includes('sourceSheets'), '折叠 JSON 应保留含统计字段的那一层对象');
});

test('chat-renderer 渲染：合并结果即使带 _view:json 提示也必须渲染摘要卡（不可被 view 分支抢走）', () => {
    const ctx = loadChatV2();
    const fmt = ctx.App.ChatV2.renderer.formatToolResultHtml;
    const merge = {
        success: true,
        sourceSheets: ['大客户部', '销售三部', '市场部', '销售二部', '销售一部'],
        mergedRows: 100, dataRows: 100, headerPreserved: true, usedSelectedSheets: false,
        data: [['日期', '地区', '部门'], ['45353', '华南', '大客户部']]
    };
    // LLM 在工具调用参数里显式要 json / code 视图（真实链路：_vt 分支曾抢在合并摘要之前）
    const htmlViewJson = fmt({ result: merge }, { _view: 'json' });
    assert.ok(htmlViewJson.includes('exec-op-summary'), '_view:json 仍应命中操作结果渲染（而非裸 JSON）');
    assert.ok(htmlViewJson.includes('已合并 5 张工作表 → 共 100 行（含表头）'), '_view:json 仍应显示合并统计文案');
    assert.ok(htmlViewJson.includes('<th>日期</th>'), '_view:json 仍应渲染数据表');
    assert.ok(!htmlViewJson.includes('"mergedRows": 100'), '_view:json 下不应直接摊出原始 JSON 矩阵');

    const htmlViewCode = fmt({ result: merge }, { view: { type: 'code' } });
    assert.ok(htmlViewCode.includes('exec-op-summary'), '_view:code 仍应命中操作结果渲染（而非 code 块）');

    // 后端在结果体里塞 view.type 的极端情况也要覆盖
    const htmlResultView = fmt({ result: { ...merge, view: { type: 'json' } } }, {});
    assert.ok(htmlResultView.includes('exec-op-summary'), '结果内嵌 view.type 仍应命中操作结果渲染');
});

test('chat-renderer 渲染：真实入参是"merge JSON 字符串"（VB McpManager 把 Content[0].Text 序列化为 result.result）', () => {
    const ctx = loadChatV2();
    const fmt = ctx.App.ChatV2.renderer.formatToolResultHtml;
    const merge = {
        success: true,
        sourceSheets: ['大客户部', '销售三部', '市场部', '销售二部', '销售一部'],
        mergedRows: 100, dataRows: 100, headerPreserved: true, usedSelectedSheets: false,
        data: [['日期', '地区', '部门'], ['45353', '华南', '大客户部']]
    };
    // 真实链路：executeLocalExcelTool → parseMcpResult → data.result 是「merge 对象的 JSON 字符串」，
    // 最终 tc.result 就是这个字符串，而不是 {result:{…}} 包裹对象。
    const html = fmt(JSON.stringify(merge), {});
    assert.ok(html.includes('exec-op-summary'), '字符串入参也应识别为合并结果');
    assert.ok(html.includes('已合并 5 张工作表 → 共 100 行（含表头）'), '应显示合并统计文案');
    assert.ok(html.includes('<th>日期</th>'), '应渲染数据表（不是裸 JSON 矩阵）');
    assert.ok(html.includes('sourceSheets'), '折叠 JSON 应保留含统计字段的外层对象');
    // unwrapToolResult 不应把 matrix 当包裹层剥掉（回归：曾导致只剩裸矩阵）
    const dataMatrixOnly = fmt(JSON.stringify([['日期', '地区'], ['45353', '华南']]), {});
    assert.ok(!dataMatrixOnly.includes('exec-op-summary'), '纯数据矩阵不应被误判为合并结果');
});

test('参数展示：bgScope=header 时从整段 range 析出表头范围', () => {
    const ctx = loadChatV2();
    const present = ctx.App.ChatV2.renderer.presentFormatArgs;
    assert.strictEqual(typeof present, 'function');
    const raw = {
        range: 'A1:I27',
        format: { bgColor: 'DE2910', bgScope: 'header', headerFontColor: 'FFFFFF' },
        sheet: '销售二部'
    };
    const shown = present(raw);
    assert.strictEqual(shown['表头范围'], 'A1:I1');
    assert.strictEqual(shown.range, 'A1:I27');
    assert.strictEqual(shown.format.bgColor, 'DE2910');
    assert.strictEqual(raw['表头范围'], undefined, '展示析出不能写回真正下发的参数');
    const asString = present({ range: '销售二部!A1:I21', format: '{"bgColor":"DE2910","bgScope":"header"}' });
    assert.strictEqual(asString['表头范围'], 'A1:I1');
    const whole = present({ range: 'A1:I27', format: { bgColor: 'DE2910', bgScope: 'all' } });
    assert.strictEqual(whole['表头范围'], undefined);
    const alreadyHeader = present({ range: 'A1:I1', format: { bgScope: 'header' } });
    assert.strictEqual(alreadyHeader['表头范围'], undefined);
});

test('过程列表：同名工具行已经在时，图钉步骤行不再显示', () => {
    const ctx = loadChatV2();
    const covered = ctx.App.ChatV2.renderer.stepCoveredByToolRow;
    assert.strictEqual(typeof covered, 'function');
    const tool = { kind: 'tool', toolName: 'excel_format_range', title: 'excel_format_range' };
    const pin = { kind: 'step', toolName: 'excel_format_range', title: 'excel_format_range', content: 'excel_format_range ✓（303ms）' };
    const plan = { kind: 'step', title: '读取数据并出图', content: '先读再画' };
    assert.strictEqual(covered(pin, [tool, pin]), true);
    assert.strictEqual(covered(plan, [tool, plan]), false);
    assert.strictEqual(covered(pin, [pin]), false);
});

test('tool-card 参数摘要：JSON 数组参数（sourceSheets）显示为 [a／b／c 等 N 项]', () => {
    const ctx = loadChatV2();
    const sum = ctx.App.ChatV2.renderer.cbxArgsSummary;
    assert.strictEqual(typeof sum, 'function', 'cbxArgsSummary 应暴露供测试');
    assert.strictEqual(
        sum({ args: { sourceSheets: '["大客户部","销售三部","市场部","销售二部","销售一部"]', mergeAll: true } }),
        '[大客户部／销售三部／市场部 等 5 项]');
    assert.strictEqual(sum({ args: { sourceSheets: ['A', 'B'], mergeAll: false } }), '[A／B]');
    assert.strictEqual(sum({ args: { range: 'A1:C10' } }), 'A1:C10');
});

test('tool-card sanitizeArgsForDisplay：长字符串截断', () => {
    const { sanitizeArgsForDisplay } = load();
    const long = 'x'.repeat(500);
    const out = sanitizeArgsForDisplay({ code: long });
    assert.ok(out.code.startsWith('x'.repeat(200)), '应保留前 200 字符');
    assert.ok(out.code.includes('(500 chars)'), '应标注原始长度');
});

test('tool-card sanitizeArgsForDisplay：深度截断与透传', () => {
    const { sanitizeArgsForDisplay } = load();
    assert.strictEqual(sanitizeArgsForDisplay({ a: { b: { c: { d: { e: 1 } } } } }).a.b.c.d, '...');
    assert.strictEqual(sanitizeArgsForDisplay({ a: 1, b: 'ok' }).b, 'ok');
});

test('tool-card formatToolErrorHtml：Python 堆栈识别', () => {
    const { formatToolErrorHtml } = load();
    const err = 'Traceback (most recent call last):\n  File "test.py", line 3, in <module>\nKeyError: "x"';
    const html = formatToolErrorHtml(err);
    assert.ok(html.includes('python-error'), '应识别为 python 错误');
    assert.ok(html.includes('KeyError'), '应显示错误类型');
});

test('tool-card formatToolErrorHtml：MCP 标准错误', () => {
    const { formatToolErrorHtml } = load();
    const html = formatToolErrorHtml('{"success": false, "message": "写入失败"}');
    assert.ok(html.includes('err-raw-text'), '应展示原始错误文本');
    assert.ok(html.includes('写入失败'), '应包含错误消息');
});

test('tool-card formatToolErrorHtml：通用错误带兜底建议', () => {
    const { formatToolErrorHtml } = load();
    const html = formatToolErrorHtml('Tool not found: xxx');
    assert.ok(html.includes('err-suggestion'), '应给出建议');
    assert.ok(html.includes('该工具未注册或已被移除'), '应命中工具未注册建议');
});

test('tool-card extractSkillResultSummary：剥离前缀与步骤列表', () => {
    const { extractSkillResultSummary } = load();
    const raw = {
        outputs: {
            message: 'Skill "合并" Hybrid 执行成功：✓ 合并所有Sheet完成\n执行步骤：\n1. 读取\n2. 合并'
        }
    };
    const out = extractSkillResultSummary(raw);
    assert.ok(out.includes('✓ 合并所有Sheet完成'), '应保留结果摘要');
    assert.ok(!out.includes('Skill "合并"'), '应剥离前缀');
    assert.ok(!out.includes('执行步骤'), '应剥离步骤列表');
});

test('tool-card truncateForDisplay：长结果截断', () => {
    const { truncateForDisplay } = load();
    const long = 'y'.repeat(3000);
    const out = truncateForDisplay(long);
    assert.ok(out.includes('...(结果已截断'), '应标注截断');
    assert.ok(out.includes('3000'), '应标注总长度');
});

test('tool-card extractExportFilePath：从结果取路径并归一化斜杠', () => {
    const { extractExportFilePath } = load();
    assert.ok(typeof extractExportFilePath === 'function', '应导出 extractExportFilePath');
    const info = extractExportFilePath(
        'ppt_export_pdf',
        { success: true, filePath: 'C:/Users/Demo/Desktop/演示_20261011.pdf' },
        null
    );
    assert.ok(info && info.path, '应解析出路径');
    assert.ok(info.path.includes('演示_20261011.pdf'), '应保留文件名');
    assert.ok(!info.path.includes('////'), '不应残留连续斜杠异常');
});

test('tool-card autoSummarizeToolResult：导出 PDF 摘要含文件名', () => {
    const { autoSummarizeToolResult } = load();
    const summary = autoSummarizeToolResult('ppt_export_pdf', {
        success: true,
        filePath: 'C:\\Users\\Demo\\Desktop\\演示_20261011.pdf'
    });
    assert.ok(summary && summary.includes('已导出 PDF'), '应标注导出 PDF');
    assert.ok(summary.includes('演示_20261011.pdf'), '应含文件名');
});

test('tool-card 通过 chat-renderer.renderer 委托暴露', () => {
    const ctx = loadChatV2();
    assert.strictEqual(typeof ctx.App.ChatV2.renderer.unwrapToolResult, 'function');
    assert.strictEqual(typeof ctx.App.ChatV2.renderer.autoSummarizeToolResult, 'function');
    assert.strictEqual(typeof ctx.App.ChatV2.renderer.formatToolErrorHtml, 'function');
});