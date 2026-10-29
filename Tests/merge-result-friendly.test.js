/**
 * ============================================================
 * Tests/merge-result-friendly.test.js — 合并工作表成功结果必须渲染成可读摘要，不能是裸 JSON
 * ============================================================
 * 背景（2026-10-09 用户反馈）：
 *   "将选中工作表进行数据合并" → 确定性硬短路直调 excel_merge_sheets，合并成功，
 *   但气泡里直接摊了一坨 {"success":true,"sourceSheets":[…],"mergedRows":…,"data":[[…]],"steps":[…]}。
 *
 * 根因（Node 内实测复现）：
 *   chat-orchestrator.excelFriendlyText 只认"产物类"字段（summary / files_created / steps），
 *   取内层时又只看 result 为**字符串**的形态。合并结果经部分通道（桥/通道预解析）回传时
 *   result 是**对象** → 停在信封层 → 三项判据全空 → 返回 null
 *   → 兜底 safeDisplayText 把整份 JSON（含 data 矩阵）倒进正文。
 *
 * 修复：① 新增 excel_merge_sheets 专属分支（按 mergedRows/dataRows/sourceSheets 逐层下钻，
 *      字符串与对象两种信封都能命中）；② 通用下钻允许 result/data 为对象。
 *
 * 运行：node --test Tests/merge-result-friendly.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

const { App } = loadChatV2(['services/tool-registry.util.js', 'chat-orchestrator.js', 'ui/renderers/tool-card.js']);
const friendly = App.ChatV2.orchestrator.excelFriendlyText;
const replaceEcho = App.ChatV2.toolRegistryUtil.replaceJsonEchoWithFriendly;
const toolCard = App.ChatV2.renderers.toolCard;

/** 后端 ExecuteMergeSheets 的真实返回体（Services/Mcp/ExcelMcpTools.SheetOps.SplitMerge.vb） */
function mergePayload() {
    return {
        success: true,
        sourceSheets: ['产品A', '产品B'],
        mergedRows: 2,
        dataRows: 2,
        headerPreserved: true,
        usedSelectedSheets: true,
        data: [
            ['日期', '地区', '部门', '产品', '销售额', '成本', '利润', '销售员', '客户类型'],
            ['2024-01', '华南', '销售一部', '产品A', 45300, 24500, 20800, '孙七', '政府客户'],
            ['2024-02', '华南', '销售一部', '产品B', 31500, 19138, 12364, '孙七', '政府客户']
        ],
        steps: [
            '读取并校对源表 2 张：产品A、产品B',
            '准备目标表：合并结果（已存在则清空重用，不会留空白表）',
            { writtenTo: '合并结果', targetSheet: '合并结果', cleanedEmptySheets: 0 }
        ],
        writtenTo: '合并结果',
        targetSheet: '合并结果',
        cleanedEmptySheets: 0
    };
}

const notRawJson = (t) => {
    assert.ok(t, '合并结果必须被友好化，不能返回 null（否则兜底会倾倒裸 JSON）');
    assert.ok(!/^\s*[{[]/.test(String(t)), `结果是裸 JSON：${String(t).slice(0, 120)}`);
    assert.ok(!/mergedRows|sourceSheets/.test(String(t)), `结果里出现了原始字段名：${String(t).slice(0, 120)}`);
};

test('result 为 JSON 字符串（宿主 CallLocalTool 原样）：输出可读摘要', () => {
    const env = { success: true, result: JSON.stringify(mergePayload(), null, 2), isError: false, error: '', taskCompleted: true };
    const t = friendly('excel_merge_sheets', env);
    notRawJson(t);
    assert.ok(t.includes('已合并 2 张工作表'), t);
    assert.ok(t.includes('产品A'), `应点名源工作表：${t}`);
    assert.ok(t.includes('2 行数据'), `应报告合并行数：${t}`);
    assert.ok(t.includes('合并结果'), `应报告落地表名：${t}`);
});

test('result 为已解析对象（桥/通道预解析）：同样输出可读摘要（回归根因）', () => {
    const env = { success: true, result: mergePayload(), isError: false, error: '', taskCompleted: true };
    const t = friendly('excel_merge_sheets', env);
    notRawJson(t);
    assert.ok(t.includes('已合并 2 张工作表'), t);
    assert.ok(t.includes('合并结果'), t);
});

test('直接给到合并 payload（无信封层）：仍输出可读摘要', () => {
    const t = friendly('excel_merge_sheets', mergePayload());
    notRawJson(t);
    assert.ok(t.includes('已合并 2 张工作表'), t);
});

test('双层包裹（result 里再包一层 result 字符串）：仍输出可读摘要', () => {
    const env = { success: true, result: JSON.stringify({ success: true, result: mergePayload() }) };
    const t = friendly('excel_merge_sheets', env);
    notRawJson(t);
    assert.ok(t.includes('已合并 2 张工作表'), t);
});

test('合并失败：如实报失败，不伪装成功', () => {
    const env = { success: false, result: JSON.stringify({ success: false, error: '未选择源工作表' }), isError: true };
    const t = friendly('excel_merge_sheets', env);
    assert.ok(t && t.includes('合并失败'), `失败分支应如实报失败：${t}`);
    assert.ok(t.includes('未选择源工作表'), t);
});

test('源表过多时只点名前 6 张，标题行不被表名撑爆', () => {
    const p = mergePayload();
    p.sourceSheets = ['S1', 'S2', 'S3', 'S4', 'S5', 'S6', 'S7', 'S8', 'S9'];
    const t = friendly('excel_merge_sheets', { success: true, result: p });
    assert.ok(t.includes('9 张工作表'), t);
    assert.ok(t.includes('等 9 张'), t);
    assert.ok(!t.includes('S7'), `第 7 张之后不应点名：${t}`);
});

test('非合并工具行为不变：清洗摘要照旧', () => {
    const env = { success: true, result: JSON.stringify({ success: true, cleanedRows: 3, trimmedCells: 5 }) };
    const t = friendly('excel_clean_data', env);
    assert.ok(t && t.includes('清洗完成'), `清洗摘要应保留：${t}`);
});

test('清洗 dryRun 不得说成已完成', () => {
    const env = { success: true, result: JSON.stringify({
        success: true, cleanedRows: 100, removedEmptyRows: 1, removedRows: 1,
        mode: 'dryRun_preview', writtenTo: '（dryRun 预览，未修改工作表）'
    }) };
    const t = friendly('excel_clean_data', env);
    assert.ok(t && t.includes('工作表未修改'), `预览不得伪装成已清洗：${t}`);
    assert.ok(!t.includes('清洗完成'), t);
});

test('标记某个值的摘要要写出标黄了谁', () => {
    const t = friendly('excel_apply_conditional_format', {
        success: true, markLiteral: '王五', sheet: '销售数据', range: 'A1:I100'
    });
    assert.ok(t && t.includes('已把「王五」标成黄色'), t);
    assert.ok(t.includes('销售数据'), t);
});

test('标记重复行的摘要要写标黄且未删除', () => {
    const env = { success: true, result: JSON.stringify({
        success: true, worksheet: '销售数据', range: 'A1:I100',
        duplicateRows: [17], highlightedRows: [17], markOnly: true,
        writtenTo: '（已标黄重复行，未删除数据）'
    }) };
    const t = friendly('excel_remove_duplicates', env);
    assert.ok(t && t.includes('已标黄重复行'), t);
    assert.ok(t.includes('第 17 行'), t);
    assert.ok(t.includes('数据未删除'), t);
    assert.ok(!t.includes('已写回工作表'), t);
});

test('清洗结果要写出发现的重复行号', () => {
    const env = { success: true, result: JSON.stringify({
        success: true, worksheet: '销售数据', range: 'A1:I102',
        cleanedRows: 99, removedDuplicateRows: 1, duplicateRows: [17],
        mode: 'writeBack_dirty', writtenTo: '销售数据（已直接修改，工作簿待保存）'
    }) };
    const t = friendly('excel_clean_data', env);
    assert.ok(t && t.includes('发现重复行 1 行'), t);
    assert.ok(t.includes('第 17 行'), t);
});

test('清洗写回后摘要标明已改表', () => {
    const env = { success: true, result: JSON.stringify({
        success: true, worksheet: '销售数据', range: 'A1:I102',
        cleanedRows: 100, removedEmptyRows: 1, removedRows: 1,
        mode: 'writeBack_dirty', writtenTo: '销售数据（已直接修改，工作簿待保存）'
    }) };
    const t = friendly('excel_clean_data', env);
    assert.ok(t && t.includes('已写回工作表'), `写回应可见：${t}`);
});

test('读取类结果仍返回 null（交回元数据卡渲染，不被摘要抢走）', () => {
    const env = { success: true, result: JSON.stringify({ success: true, headers: ['A', 'B'], rows: [[1, 2]] }) };
    assert.strictEqual(friendly('excel_read_range', env), null);
});

// ── excel_get_sheets：表清单必须转人话（实测 2026-10-22 22:04 编辑重发事故：
//    "美化工作簿中的所有工作表"被通用兜底误匹配到本工具，原始 JSON 数组整段倾倒进正文）──
test('get_sheets：result 为 JSON 字符串数组 → 转表清单人话，不摊裸 JSON', () => {
    const env = { success: true, result: JSON.stringify([
        { name: 'Fat data', index: 1, usedRange: 'A1:J35' },
        { name: '销售三部', index: 2, usedRange: 'A1:I27' },
        { name: '市场部', index: 6, usedRange: 'A1:I21' }
    ]) };
    const t = friendly('excel_get_sheets', env);
    assert.ok(t && t.includes('3 张可见工作表'), t);
    assert.ok(t.includes('Fat data') && t.includes('市场部'), t);
    notRawJson(t);
});

test('get_sheets：result 为已解析数组（桥预解析形态）→ 同样转人话', () => {
    const env = { success: true, result: [{ name: 'A' }, { name: 'B' }] };
    const t = friendly('excel_get_sheets', env);
    assert.ok(t && t.includes('2 张可见工作表') && t.includes('A、B'), t);
});

test('get_sheets：空清单 → 明确说没有可见工作表；解析不出结构 → 仍返回 null 走原路径', () => {
    assert.ok(/没有可见的工作表/.test(friendly('excel_get_sheets', { success: true, result: '[]' })));
    assert.strictEqual(friendly('excel_get_sheets', { success: true, result: 'garbage' }), null);
});

// ── JSON 回声替换守卫（skill-pipeline.formatSkillResult / agent-loop 收尾共用）──

test('回声守卫：合并 payload JSON 字符串（技能管道直通形态）→ 可读摘要', () => {
    const t = replaceEcho(JSON.stringify(mergePayload(), null, 2));
    notRawJson(t);
    assert.ok(t.includes('已合并 2 张工作表'), t);
});

test('回声守卫：信封字符串 {"success":true,"result":"<payload>"} → 可读摘要', () => {
    const t = replaceEcho(JSON.stringify({ success: true, result: JSON.stringify(mergePayload()) }));
    notRawJson(t);
    assert.ok(t.includes('已合并 2 张工作表'), t);
});

test('回声守卫：带 summary 的产物 JSON → ✅ 摘要 + 文件列表', () => {
    const t = replaceEcho(JSON.stringify({ success: true, summary: '已拆分为 3 张工作表', files_created: ['D:\\out\\a.xlsx', 'D:\\out\\b.xlsx'] }));
    assert.ok(t && t.includes('已拆分为 3 张工作表'), t);
    assert.ok(t.includes('D:/out/a.xlsx'), t);
});

test('回声守卫：正文带说明文字的 JSON 回声（模型复读）→ 只替换 JSON 块，保留说明文字', () => {
    const t = replaceEcho('好的，已为您合并：' + JSON.stringify(mergePayload(), null, 2));
    assert.ok(t && t.includes('好的，已为您合并：'), `说明文字必须保留：${t}`);
    assert.ok(t.includes('已合并 2 张工作表'), t);
    assert.ok(!t.includes('mergedRows'), `合并 JSON 块必须被替换掉：${t.slice(0, 120)}`);
});

test('回声守卫：双层 JSON 编码（字符串里再套字符串）→ 仍命中', () => {
    const t = replaceEcho(JSON.stringify(JSON.stringify(mergePayload())));
    notRawJson(t);
    assert.ok(t.includes('已合并 2 张工作表'), t);
});

test('形状无关探测：未知包装键 payload/response 也能挖出合并块', () => {
    const wrapped = { ok: true, payload: JSON.stringify(mergePayload()) };
    const hit = App.ChatV2.toolRegistryUtil.findMergePayloadInText(JSON.stringify(wrapped));
    assert.ok(hit && hit.obj && Array.isArray(hit.obj.sourceSheets), '未知包装键必须被文本级探测命中');
    const t = App.ChatV2.toolRegistryUtil.replaceMergeJsonInText(JSON.stringify(wrapped));
    assert.ok(t && t.includes('已合并 2 张工作表'), t);
});

test('形状无关探测：普通 JSON（用户点名要的输出）不被动', () => {
    assert.strictEqual(App.ChatV2.toolRegistryUtil.findMergePayloadInText(JSON.stringify([{ a: 1 }, { a: 2 }])), null);
    assert.strictEqual(App.ChatV2.toolRegistryUtil.replaceMergeJsonInText('这是普通文本 {"a":1}'), null);
});

test('工具卡兜底：双层编码的结果也能渲染成摘要（不再摊裸 JSON）', () => {
    const double = JSON.stringify(JSON.stringify(mergePayload()));
    const sum = toolCard.summarizeMergeFromRaw(double);
    assert.ok(sum && sum.text && sum.text.includes('已合并'), `双层编码必须命中：${JSON.stringify(sum)}`);
});

test('回声守卫：自然语言总结不动、用户要的 JSON 不动、非 JSON 不动', () => {
    assert.strictEqual(replaceEcho('✅ 已合并 2 张工作表，共 2 行数据。'), null, '纯文本绝不能被改写');
    assert.strictEqual(replaceEcho('结果如下：{"success":true}'), null, '带说明文字的不动');
    assert.strictEqual(replaceEcho(JSON.stringify([{ a: 1 }, { a: 2 }])), null, '普通对象数组不动');
    assert.strictEqual(replaceEcho('不是json'), null);
    assert.strictEqual(replaceEcho(null), null);
});

// ── 非标准 JSON（无逗号 pprint 风格，运行时真实形态 v202610092157）──
// 这种字符串 JSON.parse 必然失败，旧逻辑会整坨倾倒；新逻辑用正则抽取字段生成摘要。
test('回声守卫：无逗号 pprint 风格合并块 → 可读摘要（不依赖 JSON.parse）', () => {
    const pprint = [
        '{',
        '"success": true',
        '"sourceSheets": [',
        '"产品A"',
        '"产品B"',
        ']',
        '"mergedRows": 2',
        '"dataRows": 2',
        '"headerPreserved": true',
        '"writtenTo": "合并结果"',
        '"targetSheet": "合并结果"',
        '"cleanedEmptySheets": 0',
        '}',
    ].join('\n');
    const t = replaceEcho(pprint);
    notRawJson(t);
    assert.ok(t.includes('已合并'), t);
    assert.ok(t.includes('产品A') && t.includes('产品B'), '源表名应被抽出: ' + t);
    assert.ok(t.includes('合并结果'), '目标表名应被抽出: ' + t);
});

test('回声守卫：无逗号格式 + 大量源表（只点名前 6 张）', () => {
    const sheets = Array.from({ length: 9 }, (_, i) => '"表' + (i + 1) + '"').join('\n');
    const pprint = ['{', '"mergedRows": 30', '"sourceSheets": [', sheets, ']', '"writtenTo": "合并结果"', '}'].join('\n');
    const t = replaceEcho(pprint);
    assert.ok(t && t.includes('已合并 9 张工作表'), t);
    assert.ok(t.includes('等 9 张'), '应显示"等 N 张"折叠: ' + t);
});

test('_extractMergeFields 精准抽取（含布尔/数字/表名数组）', () => {
    const t = '{ "success": true\n "sourceSheets": [ "A" "B" ]\n "mergedRows": 2 "headerPreserved": false "writtenTo": "X" }';
    const out = replaceEcho(t);
    assert.ok(out.includes('已合并'), out);
    assert.ok(out.includes('表头未保留'), 'headerPreserved=false 应提示未保留: ' + out);
});

test('编辑重发：工作表清单 JSON 不再当收尾，改写成合并结果', () => {
    const dump = [
        '[{"name":"合并结果","index":2,"usedRange":"A1:I101"},',
        '{"name":"销售三部","index":6,"usedRange":"A1:I19"}]',
        '[{"日期":"2024-12-25","部门":"市场部","销售额":21475}]'
    ].join('\n');
    const turn = {
        _executedToolCalls: [
            { name: 'excel_merge_sheets', args: { sourceSheets: ['销售三部', '市场部'], targetSheet: '合并结果' } },
            { name: 'excel_format_range', args: { sheet: '合并结果', format: { bgColor: 'DE2910' } } }
        ]
    };
    const t = App.ChatV2.toolRegistryUtil.replaceToolDumpEcho(dump, turn);
    assert.ok(t && t.includes('销售三部') && t.includes('市场部'), t);
    assert.ok(t.includes('合并结果'), t);
    assert.ok(!/"usedRange"/.test(t));
    assert.ok(t.includes('合并结果'), t);
});
