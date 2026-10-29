/**
 * ============================================================
 * Tests/remove-duplicates.test.js — excel_remove_duplicates 去重工具专项测试
 * ============================================================
 * 验证：
 *   1. 后端 VB 注册 & excel_tools.json 中 columns 参数类型为 array（支持多列）
 *   2. 三处描述中均明确提及"多列"能力
 *   3. SmartDataCleaner skill 的 dedupeColumns 字段支持数组
 *   4. skill-pipeline 的 MCP 去重调用正确传递 columns 数组
 *
 * 运行：node Tests/remove-duplicates.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const VB_FILE = path.join(ROOT, 'Services', 'Mcp', 'ExcelMcpTools.vb');
const JSON_FILE = path.join(ROOT, 'Config', 'excel_tools.json');
const SKILL_FILE = path.join(ROOT, 'Skills', 'SmartDataCleaner.md');
const PIPELINE_FILE = path.join(ROOT, 'UI', 'js', 'chat-v2', 'pipelines', 'skill-pipeline.js');
// 【skill-pipeline 拆分适配】清洗/CRUD 执行器已迁至 skill-clean-crud.js，源码断言需拼接读取
const SKILL_CLEAN_CRUD_FILE = path.join(ROOT, 'UI', 'js', 'chat-v2', 'pipelines', 'skill', 'skill-clean-crud.js');
const SPEC_ENGINE_FILE = path.join(ROOT, 'UI', 'js', 'chat-v2', 'services', 'skill-spec-engine.js');

function readFile(p) {
    const buf = fs.readFileSync(p);
    // 判断编码
    if (buf[0] === 0xEF && buf[1] === 0xBB && buf[2] === 0xBF) {
        return buf.toString('utf8', 3);
    }
    if (buf[0] === 0xFF && buf[1] === 0xFE) {
        return buf.toString('utf16le', 2);
    }
    if (buf[0] === 0xFE && buf[1] === 0xFF) {
        return buf.toString('utf16be', 2);
    }
    return buf.toString('utf8');
}

/** 从 excel_tools.json 提取指定工具的定义 */
function getToolFromJson(toolName) {
    const raw = readFile(JSON_FILE);
    const obj = JSON.parse(raw.replace(/^\uFEFF/, ''));
    return (obj.tools || []).find(t => t.name === toolName) || null;
}

/** 从 VB 文件提取指定工具的描述和参数（定义已拆到 ExcelMcpTools.*.vb partial，须一并扫描） */
function loadVbAll() {
    const dir = path.join(ROOT, 'Services', 'Mcp');
    const files = fs.readdirSync(dir)
        .filter(f => /^ExcelMcpTools(\.\w+)*\.vb$/.test(f))
        .map(f => path.join(dir, f))
        .sort();
    return files.map(readFile).join('\n');
}

/** 从 VB 文件提取指定工具的描述和参数 */
function getToolFromVb(toolName) {
    const src = loadVbAll();
    const pattern = new RegExp(
        `tools\\.Add\\(New McpTool\\("${toolName.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&')}",\\s*` +
        `"([^"]*(?:""[^"]*)*)"`,  // description (VB 中双引号转义为 "")
        's'
    );
    const m = src.match(pattern);
    if (!m) return null;
    const desc = m[1].replace(/""/g, '"');  // 还原双引号
    return { description: desc };
}

// ========== 测试 1：excel_tools.json 中 columns 参数类型为 array ==========
test('去重工具: excel_tools.json 中 columns 类型为 array 且描述提及多列', () => {
    const tool = getToolFromJson('excel_remove_duplicates');
    assert.ok(tool, 'excel_tools.json 应包含 excel_remove_duplicates');

    const colsParam = (tool.parameters || []).find(p => p.name === 'columns');
    assert.ok(colsParam, 'columns 参数应存在');
    assert.strictEqual(colsParam.type, 'array', 'columns 类型应为 array（支持多列）');
    assert.ok(
        /多列|数组|列字母.*B|['""]A['""].*['""]B['""]/i.test(colsParam.description),
        `columns 描述应明确支持多列，实际: ${colsParam.description}`
    );
    console.log('  [ok] columns 参数类型=array，描述明确支持多列');
});

// ========== 测试 2：VB 后端描述提及多列 ==========
test('去重工具: VB 后端描述明确支持多列', () => {
    const info = getToolFromVb('excel_remove_duplicates');
    assert.ok(info, 'VB 应能提取到 excel_remove_duplicates 定义');
    assert.ok(
        /多列/.test(info.description),
        `VB 描述应提及"多列"能力，实际: ${info.description.substring(0, 100)}...`
    );
    console.log('  [ok] VB 描述明确支持多列');
});

// ========== 测试 3：SmartDataCleaner skill 的 dedupeColumns 支持数组 ==========
test('去重工具: SmartDataCleaner skill 中 dedupeColumns 声明为数组类型', () => {
    const src = readFile(SKILL_FILE);
    assert.ok(
        /"dedupeColumns"\s*:\s*null.*?去重依据列数组/.test(src),
        'SmartDataCleaner.md 中 dedupeColumns 应声明为数组类型'
    );
    assert.ok(
        /null=整行去重/.test(src),
        'SmartDataCleaner.md 中应说明 null=整行去重（多列全匹配）'
    );
    console.log('  [ok] SmartDataCleaner dedupeColumns 为数组类型');
});

// ========== 测试 4：skill-spec-engine normalizeCleanSpec 保留 dedupeColumns 数组 ==========
test('去重工具: skill-spec-engine normalizeCleanSpec 正确处理多列 dedupeColumns', () => {
    const src = readFile(SPEC_ENGINE_FILE);
    // 验证 normalizeCleanSpec 中有 dedupeColumns 的数组过滤逻辑
    assert.ok(
        /Array\.isArray\(raw\.dedupeColumns\)/.test(src),
        'skill-spec-engine 中应检查 dedupeColumns 是否为数组'
    );
    assert.ok(
        /raw\.dedupeColumns\.filter\(c => typeof c === 'string'/.test(src),
        'skill-spec-engine 中应过滤 dedupeColumns 为字符串数组'
    );
    console.log('  [ok] skill-spec-engine 正确处理 dedupeColumns 数组');
});

// ========== 测试 5：skill-pipeline 去重调用传递 columns 数组 ==========
test('去重工具: skill-pipeline 中去重调用正确传递 columns 数组', () => {
    const src = readFile(PIPELINE_FILE) + readFile(SKILL_CLEAN_CRUD_FILE);
    // 验证 executeCleanWithMcp 中去重调用传递了 columns
    assert.ok(
        /(?:dedupeArgs|cleanArgs)\.columns\s*=\s*plan\.dedupeColumns/.test(src),
        'skill-pipeline 中应将 plan.dedupeColumns 传给 columns 参数'
    );
    assert.ok(
        /Array\.isArray\(plan\.dedupeColumns\)/.test(src),
        'skill-pipeline 中应检查 dedupeColumns 是否为数组后再传递'
    );
    console.log('  [ok] skill-pipeline 正确传递 columns 数组');
});

// ========== 测试 6：后端 VB 代码中 columns 参数解析兼容数组 ==========
test('去重工具: 后端 VB 解析兼容 JArray / IEnumerable 数组输入', () => {
    const src = loadVbAll();
    // 在 ExecuteRemoveDuplicates 函数范围内查找
    const fnStart = src.indexOf('Private Function ExecuteRemoveDuplicates');
    assert.ok(fnStart !== -1, '应存在 ExecuteRemoveDuplicates 函数');
    // 找函数结束位置（End Function 后面跟着下一个函数或 Region）
    const nextFn = src.indexOf("Private Function ExecuteCleanData", fnStart);
    const fnEnd = nextFn !== -1 ? nextFn : src.length;
    const fnBody = src.substring(fnStart, fnEnd);

    // 验证 IEnumerable 处理（JArray 会走这条路径）
    assert.ok(
        /TypeOf columnsArg Is System\.Collections\.IEnumerable/.test(fnBody),
        'VB 解析应兼容 IEnumerable（JArray 数组输入）'
    );
    // 验证 colList 是 List(Of Integer)，支持多列
    assert.ok(
        /colList\.Add\(resolved\)/.test(fnBody),
        'VB 应将每列解析结果加入 colList（支持多列）'
    );
    console.log('  [ok] 后端 VB 兼容数组输入，支持多列解析');
});

// ========== 测试 7：后端 VB 去重列号使用范围相对列号 ==========
test('去重工具: 后端 VB 去重使用范围相对列号（内存去重路径）', () => {
    const src = loadVbAll();
    const fnStart = src.indexOf('Private Function ExecuteRemoveDuplicates');
    const nextFn = src.indexOf("Private Function ExecuteCleanData", fnStart);
    const fnEnd = nextFn !== -1 ? nextFn : src.length;
    const fnBody = src.substring(fnStart, fnEnd);

    // 验证存在"全局列号 → 相对列号"的转换逻辑（gc - firstCol + 1）
    assert.ok(
        /gc - firstCol \+ 1|dedupRelativeCols|相对列号/.test(fnBody),
        'VB 中应存在全局列号转相对列号的逻辑（修复列错位 bug）'
    );
    // 【实现演进适配】现实现为内存去重（读 Value2 → HashSet 去重 → 写回），
    // 不再调用 Interop Range.RemoveDuplicates，列错位/类型不匹配类 COM bug 结构性消失。
    assert.ok(
        /TryCast\(rng\.Value2, Object\(,\)\)/.test(fnBody),
        '去重应读取范围数据到内存二维数组'
    );
    assert.ok(
        !/\.RemoveDuplicates\(/.test(fnBody),
        '不应再调用 Interop Range.RemoveDuplicates（已由内存去重取代）'
    );
    console.log('  [ok] 去重使用相对列号且走内存去重，无列错位');
});

// ========== 测试 8：内存去重路径不再经 COM 传参（避免 DISP_E_TYPEMISMATCH） ==========
test('去重工具: 内存去重不经 COM 传列数组（结构性规避类型不匹配）', () => {
    const src = loadVbAll();
    const fnStart = src.indexOf('Private Function ExecuteRemoveDuplicates');
    const nextFn = src.indexOf("Private Function ExecuteCleanData", fnStart);
    const fnEnd = nextFn !== -1 ? nextFn : src.length;
    const fnBody = src.substring(fnStart, fnEnd);

    // 【实现演进适配】旧实现经 COM 传 Object() 列数组易触发 DISP_E_TYPEMISMATCH；
    // 现实现整块读入内存（Value2 → HashSet Ordinal 去重 → 写回），不再有任何 COM 列数组传参。
    assert.ok(
        /New HashSet\(Of String\)\(StringComparer\.Ordinal\)/.test(fnBody),
        '应存在内存 HashSet 去重（Ordinal 精确比较）'
    );
    assert.ok(
        /\.RemoveDuplicates\(/.test(fnBody) === false,
        '不应存在 Interop RemoveDuplicates COM 调用（类型不匹配无从发生）'
    );
    console.log('  [ok] 内存去重路径，无 COM 列数组传参');
});

// ========== 测试 9：normalizeNumbers 跳过日期/时间列 ==========
test('去重工具: normalizeNumbers 检测日期格式列并跳过（修复日期被变数值）', () => {
    const src = loadVbAll();

    // 1) 存在 IsDateTimeFormat 辅助方法
    assert.ok(
        /Private Shared Function IsDateTimeFormat\(fmt As String\) As Boolean/.test(src),
        '应存在 IsDateTimeFormat 日期格式检测方法'
    );
    // 2) 中文日期格式检测（年月日）
    assert.ok(
        /f\.IndexOf\("日"\) >= 0|f\.IndexOf\("月"\) >= 0|IndexOf\("日"c\) >= 0/.test(src),
        'IsDateTimeFormat 应识别中文年月日格式'
    );
    // 3) normalizeNumbers 中使用 nonNumericCols 跳过非数字列
    const fnStart = src.indexOf('Private Function ExecuteCleanData');
    // 注意：ExecuteCleanData 内部含局部函数（其首个 End Function 是局部函数结尾），
    // 须以下一个顶层成员（EscapeJson）作为截断边界，方能覆盖 normalizeNumbers 主体。
    const fnEnd = src.indexOf('Private Shared Function EscapeJson', fnStart);
    const fnBody = src.substring(fnStart, fnEnd !== -1 ? fnEnd : fnStart + 10000);
    assert.ok(
        /normalizenumbers/i.test(fnBody),
        '应存在 normalizeNumbers 操作'
    );
    assert.ok(
        /nonNumericCols\.Contains\(c\)/.test(fnBody),
        'normalizeNumbers 中应通过 nonNumericCols 跳过日期/文本列'
    );
    assert.ok(
        /IsDateTimeFormat\(fmt\)/.test(fnBody),
        '应调用 IsDateTimeFormat 判断列格式'
    );
    // 4) DateTime 值本身不被当作数字
    assert.ok(
        /TypeOf v Is DateTime/.test(fnBody),
        'normalizeNumbers 应显式跳过 DateTime 类型值'
    );
    // 5) 三重防线：格式检测 + DateTime 值 + 日期序列号区间/小数判断
    assert.ok(
        /IsDateTimeFormat\(fmt\)/.test(fnBody),
        '应调用 IsDateTimeFormat 判断列格式'
    );
    assert.ok(
        /allInDateRange\(totalCols\)/.test(fnBody),
        'normalizeNumbers 应检测日期序列号区间'
    );
    assert.ok(
        /hasFractional\(totalCols\)/.test(fnBody),
        'normalizeNumbers 应检测带小数（时间部分）的日期序列号'
    );
    // 6) values 用 rng.Value 读取（使日期单元格返回 DateTime，而非 Value2 的 Double）
    const cleanStart = src.indexOf('Private Function ExecuteCleanData');
    const cleanBody = src.substring(cleanStart, src.indexOf('Private Shared Function EscapeJson', cleanStart));
    assert.ok(
        /rng\.Value{0,2}/.test(cleanBody) && !/rng\.Value2/.test(cleanBody),
        '读取值应改用 rng.Value（而非 Value2），使日期单元格返回 DateTime 类型'
    );
    console.log('  [ok] normalizeNumbers 跳过日期/时间列，避免日期被变数值');
});

// ========== 测试 10：全部列去重回退 + 空行跳过 ==========
test('去重工具: 未解析到列时回退全列，空行不参与去重', () => {
    const src = loadVbAll();
    const fnStart = src.indexOf('Private Function ExecuteRemoveDuplicates');
    const nextFn = src.indexOf("Private Function ExecuteCleanData", fnStart);
    const fnEnd = nextFn !== -1 ? nextFn : src.length;
    const fnBody = src.substring(fnStart, fnEnd);

    // 1) 相对列越界/解析失败时回退全部列（等价旧 Missing.Value 整行去重语义）
    assert.ok(
        /If valueCols\.Count = 0 Then/.test(fnBody) && /For vc = 1 To values\.GetLength\(1\) : valueCols\.Add\(vc\)/.test(fnBody),
        '未解析到有效列时应回退为全列去重（整行比较语义）'
    );
    // 2) 全空行跳过（不参与去重，交给删除空行处理）
    assert.ok(
        /isBlankRow/.test(fnBody) && /skippedBlankRows/.test(fnBody),
        '全空行应跳过去重（空行不误判为重复行）'
    );
    console.log('  [ok] 全列回退 + 空行跳过，无 1004 类索引错误风险');
});

// 测试结束后强制退出
if (require.main === module) { setTimeout(() => process.exit(0), 3000).unref(); } // 仅独立运行时强制退出；聚合运行（run-tests.js require）不注册，避免进程提前退出
