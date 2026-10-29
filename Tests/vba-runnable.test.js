/**
 * 生成的 VBA 必须能原样粘进 VBA 编辑器：过程成对、引号成对、同一过程里不重复 Dim。
 * 运行：node --test Tests/vba-runnable.test.js
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

const { App } = loadChatV2(['services/vba-template-engine.js']);
const engine = App.ChatV2.services.vbaTemplateEngine;

function stripForScan(code) {
    let out = '';
    let inString = false;
    for (let i = 0; i < code.length; i++) {
        const ch = code[i];
        if (inString) {
            if (ch === '"') {
                if (code[i + 1] === '"') { i++; continue; }
                inString = false;
            }
            continue;
        }
        if (ch === '"') { inString = true; continue; }
        if (ch === "'") {
            while (i < code.length && code[i] !== '\n') i++;
            out += '\n';
            continue;
        }
        out += ch;
    }
    return out;
}

function assertPastable(code, label) {
    assert.ok(code && code.trim(), label + ' 没有生成代码');
    const plain = stripForScan(code);
    const subs = (plain.match(/^\s*(?:(?:Public|Private)\s+)?Sub\s+\S+/gim) || []).length;
    const ends = (plain.match(/^\s*End\s+Sub\b/gim) || []).length;
    assert.strictEqual(subs, ends, label + ` Sub/End Sub 不成对 (${subs}/${ends})`);
    assert.ok(subs >= 1, label + ' 没有入口 Sub');

    let inString = false;
    for (let i = 0; i < code.length; i++) {
        const ch = code[i];
        if (ch === "'") {
            if (!inString) { while (i < code.length && code[i] !== '\n') i++; continue; }
        }
        if (ch !== '"') continue;
        if (inString && code[i + 1] === '"') { i++; continue; }
        inString = !inString;
    }
    assert.strictEqual(inString, false, label + ' 有未闭合的双引号');

    const procs = plain.split(/^\s*End\s+Sub\b/gim);
    for (const proc of procs) {
        const seen = new Set();
        const re = /\bDim\s+([^\n]+)/gi;
        let m;
        while ((m = re.exec(proc))) {
            for (const part of m[1].split(',')) {
                const name = (part.trim().split(/\s+/)[0] || '').replace(/[()]/g, '').toLowerCase();
                if (!name) continue;
                assert.ok(!seen.has(name), label + ' 重复声明 ' + name);
                seen.add(name);
            }
        }
    }
    assert.ok(!/Attribute\s+VB_/i.test(plain), label + ' 含有不能粘贴的 Attribute');
    assert.ok(!/\.Text\s*=\s*"\s*&\s*ChrW/i.test(code), label + ' 要点字符串仍然是坏的');
}

test('无关的“工作表”请求不会套用重命名模板', () => {
    const result = engine.render({ scenario: '' }, '把工作表 A1 加粗', {});
    assert.strictEqual(result.matched, false);
});

test('重命名模板去掉非法表名字符，并且可以粘贴', () => {
    const result = engine.render(
        { scenario: 'rename_sheets_batch', pattern: 'Q{i}' },
        '批量重命名工作表',
        {}
    );
    assert.strictEqual(result.matched, true);
    assert.strictEqual(result.template, 'rename_sheets_batch');
    assertPastable(result.code, 'rename');
    assert.ok(result.code.includes('Replace(newName, "\\", "")') || result.code.includes('Replace(newName, "\\\\", "")'),
        '应清理反斜杠。实际片段：\n' + result.code.split('\n').filter(l => l.includes('Replace(newName')).join('\n'));
});

test('按列拆分使用真实表名，并清理新表名', () => {
    const result = engine.render({
        scenario: 'split_sheets_by_column',
        dataSheet: '销售明细',
        columnIndex: 3,
        headerRow: 2
    }, '按地区列拆分', {});
    assert.strictEqual(result.template, 'split_sheets_by_column');
    assertPastable(result.code, 'split');
    assert.ok(result.code.includes('Worksheets("销售明细")'), result.code);
    assert.ok(result.code.includes('headerRowNum = 2'), result.code);
    assert.ok(/sheetName = Replace\(sheetName, "\\", ""\)/.test(result.code) || /sheetName = Replace\(sheetName, "\\\\", ""\)/.test(result.code));
});

test('PPT 文本页和图表页可以编译：要点是表达式，图表参数顺序正确，变量不重复声明', () => {
    const result = engine.render({
        scenario: 'ppt_from_data',
        reportTitle: '周报',
        dataSheet: 'Sheet1',
        dataRange: 'A1:C10',
        slides: [
            { type: 'text', title: '结论', bullets: ['增长 12%', '甲"乙'] },
            { type: 'table', title: '明细', columns: ['地区', '金额'], rowCount: 5, startRow: 2 },
            { type: 'chart', title: '趋势', chartType: 'bar' },
            { type: 'chart', title: '占比', chartType: 'pie' }
        ]
    }, '把数据做成幻灯片', {});
    assert.strictEqual(result.template, 'ppt_from_data');
    assertPastable(result.code, 'ppt');
    assert.ok(result.code.includes('ChrW(8226) & " " & "增长 12%"'), result.code);
    assert.ok(result.code.includes('ChrW(8226) & " " & "甲""乙"'), result.code);
    assert.ok(result.code.includes('AddChart2(-1, 51,'), result.code);
    assert.ok(result.code.includes('AddChart2(-1, 5,'), result.code);
    assert.ok(result.code.includes('startedPpt'), result.code);
    assert.ok(result.code.includes('Environ("USERPROFILE")'), result.code);
    const dimNRows = (result.code.match(/\bDim\s+[^\n]*\bnRows\b/g) || []).length;
    const dimChart = (result.code.match(/\bDim\s+[^\n]*\bchartShape\b/g) || []).length;
    assert.strictEqual(dimNRows, 1, 'nRows 应只声明一次，实际 ' + dimNRows);
    assert.strictEqual(dimChart, 1, 'chartShape 应只声明一次，实际 ' + dimChart);
});

test('整理模型直接吐出的 VBA：去掉说明、Attribute、弯引号和重复 Dim', () => {
    const raw = [
        '下面是代码，复制后运行：',
        '```vba',
        'Attribute VB_Name = "Module1"',
        'Sub Main()',
        '    Dim total As Long',
        '    Dim total As Long',
        '    Range（"A1"）.Value = “完成”',
        '    GoTo ErrHandler',
        'End Sub',
        '```',
        '运行前请备份。'
    ].join('\n');
    const code = engine.toRunnableVba(raw);
    assertPastable(code, 'sanitize');
    assert.ok(!code.includes('下面是代码'), code);
    assert.ok(!code.includes('运行前请备份'), code);
    assert.ok(!/Attribute\s+VB_/i.test(code), code);
    assert.ok(code.includes('Range("A1").Value = "完成"'), code);
    assert.strictEqual((code.match(/\bDim\s+total\b/g) || []).length, 1);
    assert.ok(/^\s*ErrHandler\s*:/m.test(code), code);
});

test('纯说明文字不会被包成一个假模块', () => {
    assert.strictEqual(engine.toRunnableVba('这个需求我没法写成宏。'), '');
});

test('缺 End Sub 的语句片段会补上过程外壳', () => {
    const code = engine.toRunnableVba('Range("A1").Value = 1\nMsgBox "ok"');
    assertPastable(code, 'wrap');
    assert.ok(code.includes('Sub GeneratedMacro()'), code);
});
