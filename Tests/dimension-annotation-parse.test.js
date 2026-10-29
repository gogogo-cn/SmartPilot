/**
 * 尺寸标注脚本要收齐尺寸、形位公差、粗糙度，并且不把图号、比例、日期当成尺寸。
 * 运行：node Tests/dimension-annotation-parse.test.js
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const mdPath = path.join(__dirname, '..', 'Config', 'DefaultParseScripts', 'dimension_annotation.md');
const md = fs.readFileSync(mdPath, 'utf8').replace(/\r\n/g, '\n');
const script = md.match(/```javascript\n([\s\S]*?)\n```/)[1];
const parse = new Function('FIELDS', script + '\nreturn parse;')([]);

const sample = `
Ø25±0.02
Ø12.5
+0.021
0
R3
C1
M8×1.25-6H
45°
120±0.1
50
+0.02
-0.01
位置度 Ø0.05 A B
垂直度 0.03 A
圆度 0.01
同轴度 Ø0.02 A
⊥ 0.03 A
Ra 1.6
√3.2
表面粗糙度
6.3
基准 A
H7/g6
未注公差按 GB/T 1804-m
比例
1:2
图号
HT-01-002
2024.03.15
`;

function ofType(r, type) {
    return r.annotations.filter(function (it) { return it['类型'] === type; }).map(function (it) { return it['标注']; });
}

test('尺寸、形位公差、粗糙度和基准都按类型收进同一张表', () => {
    const r = parse(sample);
    assert.strictEqual(r.annotations.length, Number(r.count));
    assert.strictEqual(r.tableData.header.join(','), '类型,标注,公差,基准');
    assert.strictEqual(r.tableData.rows.length, r.annotations.length);

    assert.ok(ofType(r, '直径').some(function (s) { return s.indexOf('Ø25') >= 0 && s.indexOf('0.02') >= 0; }));
    assert.ok(ofType(r, '直径').some(function (s) { return s.indexOf('Ø12.5') >= 0 && s.indexOf('0.021') >= 0; }));
    assert.deepStrictEqual(ofType(r, '半径'), ['R3']);
    assert.deepStrictEqual(ofType(r, '倒角'), ['C1']);
    assert.ok(ofType(r, '螺纹').some(function (s) { return s.indexOf('M8') >= 0; }));
    assert.deepStrictEqual(ofType(r, '角度'), ['45°']);
    assert.ok(ofType(r, '线性尺寸').indexOf('120±0.1') >= 0);
    assert.ok(ofType(r, '线性尺寸').some(function (s) { return s.indexOf('50') >= 0 && s.indexOf('0.02') >= 0 && s.indexOf('0.01') >= 0; }));

    const geo = r.annotations.filter(function (it) { return it['类型'] === '形位公差'; });
    assert.ok(geo.some(function (it) { return it['标注'].indexOf('位置度') >= 0 && it['基准'] === 'A B'; }));
    assert.ok(geo.some(function (it) { return it['标注'].indexOf('垂直度') >= 0 && it['基准'] === 'A'; }));
    assert.ok(geo.some(function (it) { return it['标注'].indexOf('圆度') >= 0; }));
    assert.ok(geo.some(function (it) { return it['标注'].indexOf('同轴度') >= 0 && it['公差'].indexOf('0.02') >= 0; }));

    assert.ok(ofType(r, '表面粗糙度').indexOf('Ra1.6') >= 0);
    assert.ok(ofType(r, '表面粗糙度').indexOf('3.2') >= 0);
    assert.ok(ofType(r, '表面粗糙度').indexOf('6.3') >= 0);
    assert.ok(ofType(r, '基准').indexOf('基准 A') >= 0);
    assert.ok(ofType(r, '配合').indexOf('H7/g6') >= 0);
    assert.match(r.general_tolerance, /GB\/T 1804-m/);
});

test('图号、比例和日期不会被当成尺寸', () => {
    const r = parse(sample);
    const blob = r.annotations.map(function (it) { return it['标注']; }).join('\n');
    assert.ok(blob.indexOf('HT-01-002') < 0);
    assert.ok(blob.indexOf('1:2') < 0);
    assert.ok(blob.indexOf('2024.03.15') < 0);
});
