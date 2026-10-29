/**
 * 默认机械制图解析脚本：同一行标题，以及 OCR 把标题和内容拆成两行，都要能取出。
 * 运行：node Tests/mechanical-drawing-parse.test.js
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const mdPath = path.join(__dirname, '..', 'Config', 'DefaultParseScripts', 'mechanical_drawing.md');
const md = fs.readFileSync(mdPath, 'utf8').replace(/\r\n/g, '\n');

function cellsOf(line) {
    let s = line.trim();
    if (s.startsWith('|')) s = s.slice(1);
    if (s.endsWith('|')) s = s.slice(0, -1);
    const cells = [];
    let cur = '';
    let tick = false;
    for (const ch of s) {
        if (ch === '`') tick = !tick;
        else if (ch === '|' && !tick) {
            cells.push(cur.trim());
            cur = '';
        } else cur += ch;
    }
    if (cur.trim()) cells.push(cur.trim());
    return cells;
}

const fields = [];
for (const line of md.split('\n')) {
    if (!line.trim().startsWith('|') || line.includes('---') || line.includes('字段名')) continue;
    const cells = cellsOf(line);
    if (cells.length < 4) continue;
    fields.push({
        name: cells[0],
        key: cells[1],
        pattern: cells[2].replace(/^`|`$/g, ''),
        required: cells[3].includes('是'),
        visible: cells.length < 5 || !cells[4].includes('否')
    });
}

const script = md.match(/```javascript\n([\s\S]*?)\n```/)[1];
const parse = new Function('FIELDS', script + '\nreturn parse;')(fields);

const inline = `
技术要求
1.未注倒角 C0.5
2.去除毛刺
零件名称：法兰盘
图样代号：HT-01-002
材料：45
比例：1:2
重量：0.86kg
共 2 张  第 1 张
设计：张三 2024.03.15
校对：李四
审核：王五
某某机械有限公司
`;

const broken = `
技术要求
1.锐边倒钝
零件名称
轴套
图号
JD.12.003
材料
Q235
比例
2:1
重量
1.25kg
共 3 张
第 2 张
设计
赵六
2024.06.01
校对
钱七
阶段标记
S
`;

test('同一行标题栏能取出图号、名称、比例和设计人', () => {
    const r = parse(inline);
    assert.strictEqual(r.part_name, '法兰盘');
    assert.strictEqual(r.drawing_no, 'HT-01-002');
    assert.strictEqual(r.material, '45');
    assert.strictEqual(r.scale, '1:2');
    assert.strictEqual(r.weight, '0.86kg');
    assert.strictEqual(r.sheet_total, '2');
    assert.strictEqual(r.sheet_index, '1');
    assert.strictEqual(r.designer, '张三');
    assert.strictEqual(r.design_date, '2024.03.15');
    assert.strictEqual(r.checker, '李四');
    assert.strictEqual(r.reviewer, '王五');
    assert.strictEqual(r.company, '某某机械有限公司');
    assert.match(r.tech_requirements, /未注倒角/);
    assert.match(r.tech_requirements, /去除毛刺/);
});

test('标签和内容各占一行时也能取出', () => {
    const r = parse(broken);
    assert.strictEqual(r.part_name, '轴套');
    assert.strictEqual(r.drawing_no, 'JD.12.003');
    assert.strictEqual(r.material, 'Q235');
    assert.strictEqual(r.scale, '2:1');
    assert.strictEqual(r.weight, '1.25kg');
    assert.strictEqual(r.sheet_total, '3');
    assert.strictEqual(r.sheet_index, '2');
    assert.strictEqual(r.designer, '赵六');
    assert.strictEqual(r.design_date, '2024.06.01');
    assert.strictEqual(r.checker, '钱七');
    assert.strictEqual(r.phase, 'S');
    assert.match(r.tech_requirements, /锐边倒钝/);
});
