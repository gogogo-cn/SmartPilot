/**
 * ============================================================
 * Tests/parse-text-toolcalls.test.js — 文本工具调用解析（含方括号外参数修复）
 * 运行：node Tests/parse-text-toolcalls.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

/** 最小工具注册表 stub：只提供 names() 与 get() */
function makeRegistry(names) {
    return {
        names: () => names,
        get: () => null
    };
}

const { App } = loadChatV2();
const parse = App.ChatV2.agentLoop.parseTextToolCalls;

test('参数写在方括号外(JSON 形式)能被解析', () => {
    const reg = makeRegistry(['excel_format_range', 'excel_autofit']);
    const text = '请格式化：[excel_format_range] {"range":"A1:C3","bold":true}';
    const { valid } = parse(text, reg, { userMessage: 'x' });
    assert.strictEqual(valid.length, 1);
    assert.strictEqual(valid[0].name, 'excel_format_range');
    assert.deepStrictEqual(valid[0].args, { range: 'A1:C3', bold: true });
});

test('参数写在方括号外(key:val 形式)能被解析', () => {
    const reg = makeRegistry(['excel_format_range']);
    const text = '[excel_format_range] range:A1:C3 bold:true';
    const { valid } = parse(text, reg, { userMessage: 'x' });
    assert.strictEqual(valid.length, 1);
    assert.strictEqual(valid[0].name, 'excel_format_range');
    assert.strictEqual(valid[0].args.range, 'A1:C3');
    assert.strictEqual(valid[0].args.bold, true);
});

test('] 外侧接自然语言时不被误吞为参数', () => {
    const reg = makeRegistry(['excel_format_range']);
    const text = '[excel_format_range] 然后请把标题加粗';
    const { valid } = parse(text, reg, { userMessage: 'x' });
    assert.strictEqual(valid.length, 1);
    assert.deepStrictEqual(valid[0].args, {}); // 自然语言不是参数形态 → args 为空
});

test('标准 [tool]{json} 方括号内形式仍正常', () => {
    const reg = makeRegistry(['excel_format_range', 'excel_autofit']);
    const text = '先做 [excel_format_range]{"range":"A1:C3"} 再 [excel_autofit]{"columns":"A:F"}';
    const { valid } = parse(text, reg, { userMessage: 'x' });
    assert.strictEqual(valid.length, 2);
    assert.strictEqual(valid[0].args.range, 'A1:C3');
    assert.strictEqual(valid[1].name, 'excel_autofit');
});

test('尾随 JSON 不会越界吞掉下一个工具调用', () => {
    const reg = makeRegistry(['excel_format_range', 'excel_autofit']);
    const text = '[excel_format_range] {"range":"A1:C3"} [excel_autofit]{"columns":"A:F"}';
    const { valid } = parse(text, reg, { userMessage: 'x' });
    assert.strictEqual(valid.length, 2);
    assert.strictEqual(valid[0].args.range, 'A1:C3');
    assert.strictEqual(valid[1].args.columns, 'A:F');
});

test('方括号外参数 + 真实多参数场景(模拟用户"底色/边框/对齐/列宽")', () => {
    const reg = makeRegistry(['excel_format_range', 'excel_autofit']);
    // 模型常把多个要求拆成多个 [tool]，每个 tool 的参数写在 ] 外侧
    const text = [
        '设置底色：[excel_format_range] {"range":"A1:G1","bgColor":"FFF2CC"}',
        '加边框：[excel_format_range] {"range":"A1:G10","borderStyle":"thin"}',
        '自动列宽：[excel_autofit] {"columns":"A:G"}'
    ].join('\n');
    const { valid } = parse(text, reg, { userMessage: 'x' });
    assert.strictEqual(valid.length, 3);
    assert.ok(valid.find(v => v.name === 'excel_format_range' && v.args.bgColor === 'FFF2CC'));
    assert.ok(valid.find(v => v.name === 'excel_format_range' && v.args.borderStyle === 'thin'));
    assert.ok(valid.find(v => v.name === 'excel_autofit' && v.args.columns === 'A:G'));
});
