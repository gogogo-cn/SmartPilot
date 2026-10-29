/**
 * 球标尺寸：圆球标为非重要，六边形球标为重要，并拆出标称值与上下公差。
 * 运行：node Tests/dimension-balloon-parse.test.js
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const md = fs.readFileSync(path.join(__dirname, '..', 'Config', 'DefaultParseScripts', 'dimension_balloon.md'), 'utf8').replace(/\r\n/g, '\n');

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

const sample = `
① Ø25±0.02
⬡3
50
+0.02
-0.01
○4 R3
⬢5 位置度 Ø0.05 A B
(6) C1
六边形7 M8×1.25-6H
⊥ 0.03 A
○ 0.01
比例
1:2
图号
HT-01-002
2024.03.15
`;

function find(r, pred) {
    return r.items.filter(pred);
}

test('圆球标是非重要，六边形球标是重要，并拆出标称值和上下公差', () => {
    const r = parse(sample);
    assert.strictEqual(r.items.length, Number(r.count));
    assert.strictEqual(r.tableData.header.join(','), '尺寸序号,重要性,类型,标称值,下公差,上公差');

    const dia = find(r, function (it) { return it.dim_type === '直径'; })[0];
    assert.strictEqual(dia.seq, '1');
    assert.strictEqual(dia.importance, '非重要');
    assert.strictEqual(dia.nominal, '25');
    assert.strictEqual(dia.lower_tol, '-0.02');
    assert.strictEqual(dia.upper_tol, '+0.02');

    const linear = find(r, function (it) { return it.dim_type === '线性尺寸' && it.nominal === '50'; })[0];
    assert.strictEqual(linear.seq, '3');
    assert.strictEqual(linear.importance, '重要');
    assert.strictEqual(linear.lower_tol, '-0.01');
    assert.strictEqual(linear.upper_tol, '+0.02');

    const radius = find(r, function (it) { return it.dim_type === '半径'; })[0];
    assert.strictEqual(radius.seq, '4');
    assert.strictEqual(radius.importance, '非重要');
    assert.strictEqual(radius.nominal, '3');

    const pos = find(r, function (it) { return it.dim_type === '位置度'; })[0];
    assert.strictEqual(pos.seq, '5');
    assert.strictEqual(pos.importance, '重要');
    assert.strictEqual(pos.nominal, '');
    assert.strictEqual(pos.lower_tol, '');
    assert.strictEqual(pos.upper_tol, '0.05');

    const chamfer = find(r, function (it) { return it.dim_type === '倒角'; })[0];
    assert.strictEqual(chamfer.seq, '6');
    assert.strictEqual(chamfer.importance, '非重要');
    assert.strictEqual(chamfer.nominal, '1');

    const thread = find(r, function (it) { return it.dim_type === '螺纹'; })[0];
    assert.strictEqual(thread.seq, '7');
    assert.strictEqual(thread.importance, '重要');
    assert.ok(thread.nominal.indexOf('M8') >= 0);
});

test('没有球标的形位公差仍保留，圆度小数不会被当成圆球标', () => {
    const r = parse(sample);
    const perp = find(r, function (it) { return it.dim_type === '垂直度'; })[0];
    assert.strictEqual(perp.seq, '');
    assert.strictEqual(perp.importance, '');
    assert.strictEqual(perp.nominal, '');
    assert.strictEqual(perp.lower_tol, '');
    assert.strictEqual(perp.upper_tol, '0.03');

    const round = find(r, function (it) { return it.dim_type === '圆度'; })[0];
    assert.strictEqual(round.seq, '');
    assert.strictEqual(round.nominal, '');
    assert.strictEqual(round.lower_tol, '');
    assert.strictEqual(round.upper_tol, '0.01');

    const blob = r.items.map(function (it) { return it.nominal; }).join('\n');
    assert.ok(blob.indexOf('HT-01-002') < 0);
    assert.ok(blob.indexOf('1:2') < 0);
    assert.ok(blob.indexOf('2024.03.15') < 0);
});

test('改字段名、隐藏列后，写入表格跟着字段配置走', () => {
    const edited = fields.map(function (field) {
        const copy = Object.assign({}, field);
        if (copy.key === 'seq') copy.name = '球标号';
        if (copy.key === 'upper_tol') copy.visible = false;
        return copy;
    });
    const parseEdited = new Function('FIELDS', script + '\nreturn parse;')(edited);
    const r = parseEdited(sample);
    assert.strictEqual(r.tableData.header.join(','), '球标号,重要性,类型,标称值,下公差');
    assert.strictEqual(r.tableData.rows[0][0], '1');
    assert.strictEqual(r.tableData.rows[0].length, 5);
    assert.strictEqual(r.items[0].seq, '1');
    assert.strictEqual(r.items[0].upper_tol, '+0.02');
});

test('表面粗糙度识别 Ra、粗糙度符号，以及符号和数值拆成两行', () => {
    const r = parse(`
① Ra 1.6
√3.2
表面粗糙度
6.3
Rz
12.5
R a 0.8
Ra≤3.2
R3
○ 0.01
`);
    const ra = find(r, function (it) { return it.dim_type === 'Ra' && it.nominal === '1.6'; })[0];
    assert.strictEqual(ra.seq, '1');
    assert.strictEqual(ra.importance, '非重要');
    assert.strictEqual(ra.lower_tol, '');
    assert.strictEqual(ra.upper_tol, '');

    const symbol = find(r, function (it) { return it.dim_type === 'Ra' && it.nominal === '3.2' && it.seq === ''; })[0];
    assert.ok(symbol);

    const titled = find(r, function (it) { return it.nominal === '6.3'; })[0];
    assert.strictEqual(titled.dim_type, '表面粗糙度');

    const rz = find(r, function (it) { return it.dim_type === 'Rz'; })[0];
    assert.strictEqual(rz.nominal, '12.5');

    const spaced = find(r, function (it) { return it.nominal === '0.8'; })[0];
    assert.strictEqual(spaced.dim_type, 'Ra');

    const capped = find(r, function (it) { return it.dim_type === 'Ra' && it.nominal === '3.2'; })[0];
    assert.ok(capped);

    const radius = find(r, function (it) { return it.dim_type === '半径'; })[0];
    assert.strictEqual(radius.nominal, '3');

    const round = find(r, function (it) { return it.dim_type === '圆度'; })[0];
    assert.strictEqual(round.upper_tol, '0.01');
});

test('斜杠公差直径、裸直径、单侧配合能解析', () => {
    const r = parse(`
(52) Ø50 0/-0.025
(53) 32h7
(57) Ø18
(81)
Ø35 +0.012/-0.006
(84) 28H7
(85) 18f7
`);
    const d52 = find(r, function (it) { return it.seq === '52'; })[0];
    assert.ok(d52);
    assert.strictEqual(d52.dim_type, '直径');
    assert.strictEqual(d52.nominal, '50');
    assert.strictEqual(d52.upper_tol, '0');
    assert.strictEqual(d52.lower_tol, '-0.025');

    const fit53 = find(r, function (it) { return it.seq === '53'; })[0];
    assert.strictEqual(fit53.dim_type, '配合');
    assert.strictEqual(fit53.nominal, '32h7');

    const d57 = find(r, function (it) { return it.seq === '57'; })[0];
    assert.strictEqual(d57.dim_type, '直径');
    assert.strictEqual(d57.nominal, '18');
    assert.strictEqual(d57.lower_tol, '');
    assert.strictEqual(d57.upper_tol, '');

    const d81 = find(r, function (it) { return it.seq === '81'; })[0];
    assert.strictEqual(d81.dim_type, '直径');
    assert.strictEqual(d81.nominal, '35');
    assert.strictEqual(d81.upper_tol, '+0.012');
    assert.strictEqual(d81.lower_tol, '-0.006');

    const fit84 = find(r, function (it) { return it.seq === '84'; })[0];
    assert.strictEqual(fit84.dim_type, '配合');
    assert.strictEqual(fit84.nominal, '28H7');

    const fit85 = find(r, function (it) { return it.seq === '85'; })[0];
    assert.strictEqual(fit85.dim_type, '配合');
    assert.strictEqual(fit85.nominal, '18f7');
});

test('OCR 吃掉左括号时（52) Ø50 0/-0.025）仍能识别球标序号', () => {
    const r = parse(`
52) Ø50 0/-0.025
53) 32h7
57) Ø18
81)
Ø35 +0.012/-0.006
84) 28H7
85) 18f7
`);
    assert.strictEqual(r.items.length, 6);
    const d52 = find(r, function (it) { return it.seq === '52'; })[0];
    assert.ok(d52, '52 号球标应被识别');
    assert.strictEqual(d52.dim_type, '直径');
    assert.strictEqual(d52.nominal, '50');
    assert.strictEqual(d52.lower_tol, '-0.025');
    assert.strictEqual(d52.upper_tol, '0');

    const fit84 = find(r, function (it) { return it.seq === '84'; })[0];
    assert.strictEqual(fit84.dim_type, '配合');
    assert.strictEqual(fit84.nominal, '28H7');

    const d81 = find(r, function (it) { return it.seq === '81'; })[0];
    assert.strictEqual(d81.nominal, '35');
    assert.strictEqual(d81.upper_tol, '+0.012');
});

test('形位公差框按球标、公差值和基准整条解析，不拆成线性尺寸', () => {
    const plain = parse('01 2X 0.7 A B C');
    assert.strictEqual(plain.items.length, 1);
    assert.strictEqual(plain.items[0].seq, '01');
    assert.strictEqual(plain.items[0].importance, '非重要');
    assert.strictEqual(plain.items[0].dim_type, '形位公差');
    assert.strictEqual(plain.items[0].nominal, '');
    assert.strictEqual(plain.items[0].lower_tol, '');
    assert.strictEqual(plain.items[0].upper_tol, '0.7');

    const profile = parse('01\t2X\t⌓\t0.7\tA\tB\tC');
    assert.strictEqual(profile.items.length, 1);
    assert.strictEqual(profile.items[0].dim_type, '面轮廓度');
    assert.strictEqual(profile.items[0].upper_tol, '0.7');
    assert.strictEqual(profile.items[0].nominal, '');

    const circled = parse('① 2X ⌒ 0.05 A');
    assert.strictEqual(circled.items.length, 1);
    assert.strictEqual(circled.items[0].seq, '1');
    assert.strictEqual(circled.items[0].importance, '非重要');
    assert.strictEqual(circled.items[0].dim_type, '线轮廓度');
    assert.strictEqual(circled.items[0].upper_tol, '0.05');
});
