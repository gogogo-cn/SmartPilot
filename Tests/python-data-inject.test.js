/**
 * ============================================================
 * Tests/python-data-inject.test.js — Python 数据注入保护测试
 * ============================================================
 * 覆盖：LLM 生成的脚本中伪造的注入变量赋值（input_data/selection_data/
 *       reference_data/input_headers/_sel）在执行前被移除，使用逻辑保留
 * 运行：node Tests/python-data-inject.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const path = require('path');
const fs = require('fs');

/** 从源码读取 buildExecArgs 的数据注入保护逻辑（提取 INJECTED_VARS + 正则） */
function getInjectedVars() {
    const src = fs.readFileSync(path.join(__dirname, '..', 'UI/js/chat-v2/pipelines/python-pipeline.js'), 'utf8');
    const m = src.match(/(?:var|let|const)\s+INJECTED_VARS = \[([^\]]+)\]/);
    assert.ok(m, '源码应包含 INJECTED_VARS');
    return m[1].split(',').map(s => s.trim().replace(/['"]/g, ''));
}

/** 复现 buildExecArgs 的清理逻辑 */
function sanitizeScript(script, injectedVars) {
    let cleaned = script;
    for (const varname of injectedVars) {
        const multiRe = new RegExp(
            '^\\s*' + varname + '\\s*=\\s*(?:\\[[\\s\\S]*?\\]|\\{[\\s\\S]*?\\}|json\\.loads?\\([^)]*\\)|None|""|\'\')\\s*(?:#[^\\n]*)?\\n?', 'gm');
        cleaned = cleaned.replace(multiRe, '# [自动] ' + varname + ' 由系统注入真实数据，移除 LLM 占位符赋值\n');
        const singleRe = new RegExp('^\\s*' + varname + '\\s*=\\s*\\[[^\\n]*\\]\\s*$', 'gm');
        cleaned = cleaned.replace(singleRe, '# [自动] ' + varname + ' 由系统注入真实数据，移除 LLM 占位符赋值');
    }
    return cleaned;
}

test('python-pipeline: 移除全部注入变量的伪造赋值', () => {
    const vars = getInjectedVars();
    assert.deepStrictEqual(vars, ['input_data', 'selection_data', 'reference_data', 'input_headers', '_sel']);
    console.log('✅ 覆盖注入变量:', vars.join(', '));

    // 模拟 LLM 生成的全套伪造脚本（用户报告的场景 + 其他变量）
    const fakeScript = `import pandas as pd

# 假设 input_data 已经定义
input_data = [
    {"日期": "2023-01-01", "部门": "销售部", "地区": "北京", "销售额": 1000},
    # ... 其他数据
]

selection_data = {"workbookName": "Book1", "sheetName": "Sheet1", "address": "A1:D10", "headers": ["日期","地区","销售额"], "rows": []}
reference_data = [{"workbook": "Book1", "worksheet": "Sheet1", "address": "A1:D10", "headers": ["日期","地区","销售额"], "rows": [{"日期":"2023-01-01"}]}]
input_headers = ["日期", "地区", "销售额"]
_sel = {"a": 1}

df = pd.DataFrame(input_data)
print(df.groupby("地区").agg({"销售额": "sum"}))
output = {"rows": len(df), "sheet": _sel("sheetName")}
`;

    const cleaned = sanitizeScript(fakeScript, vars);

    // 所有伪造赋值被移除
    for (const v of vars) {
        const re = new RegExp('^\\s*' + v + '\\s*=\\s*', 'm');
        assert.ok(!re.test(cleaned), v + ' 的赋值应被移除');
    }
    // 使用逻辑保留
    assert.ok(cleaned.includes('pd.DataFrame(input_data)'), 'input_data 使用保留');
    assert.ok(cleaned.includes('_sel("sheetName")'), '_sel 使用保留');
    assert.ok(cleaned.includes('output = {"rows": len(df)'), 'output 保留');
    console.log('✅ 全部伪造赋值移除，使用逻辑保留');
});

test('python-pipeline: 多行/单行/字典/json.loads 赋值形态全覆盖', () => {
    const vars = getInjectedVars();
    const cases = [
        'input_data = [{"a": 1}]',                    // 单行数组
        'selection_data = {"a": 1}',                  // 单行字典
        'input_data = []',                            // 空数组
        'reference_data = json.loads(x)',             // json.loads
        'input_data = None',                          // None
        'input_headers = ["a", "b"]',                 // 数组
        '_sel = {}'                                   // 空字典
    ];
    for (const c of cases) {
        const cleaned = sanitizeScript(c, vars);
        assert.ok(!/^(input_data|selection_data|reference_data|input_headers|_sel)\s*=/.test(cleaned), '应移除: ' + c);
    }
    // 使用语句不被误伤
    const useCases = [
        'df = pd.DataFrame(input_data)',
        'print(selection_data["address"])',
        'for r in reference_data:',
        'cols = input_headers'
    ];
    for (const u of useCases) {
        const cleaned = sanitizeScript(u, vars);
        assert.strictEqual(cleaned, u, '使用语句应保留: ' + u);
    }
    console.log('✅ 6 种赋值形态覆盖，4 类使用语句不误伤');
});

// 兜底退出
if (require.main === module) { setTimeout(() => process.exit(0), 4000).unref(); } // 仅独立运行时强制退出；聚合运行（run-tests.js require）不注册，避免进程提前退出
