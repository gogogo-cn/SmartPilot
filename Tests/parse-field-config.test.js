/**
 * 字段配置要能改结果表：顺序、中文名、显示开关，以及身份证脚本写出的 *_from_id。
 * 运行：node Tests/parse-field-config.test.js
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const file = path.join(__dirname, '..', 'UI', 'js', 'app-ocr-parser.js');
const html = fs.readFileSync(file, 'utf8').replace(/\r\n/g, '\n');
const source = html.match(/function presentByFieldConfig\(result, fields\) \{[\s\S]*?\n    \}\n/);
assert.ok(source, '找不到 presentByFieldConfig');
const presentByFieldConfig = new Function(source[0] + '; return presentByFieldConfig;')();

const idFields = [
    { name: '性别', key: 'gender', visible: true },
    { name: '出生日期', key: 'birth_date', visible: true },
    { name: '住址', key: 'address', visible: true },
    { name: '省份', key: 'province', visible: false },
    { name: '公民身份号码', key: 'id_number', visible: true }
];

test('身份证派生字段并入配置里的性别和出生日期，隐藏字段不再出现', () => {
    const raw = {
        address: '北京市东城区景山前街4号 紫禁城敬事房',
        id_number: '11204416541220243X',
        birth_date_from_id: '1654-12-20',
        gender_from_id: '男',
        province: '北京市',
        city: '北京市'
    };
    const projected = presentByFieldConfig(raw, idFields);
    assert.deepStrictEqual(Object.keys(projected.data), ['gender', 'birth_date', 'address', 'id_number', 'city']);
    assert.strictEqual(projected.data.gender, '男');
    assert.strictEqual(projected.data.birth_date, '1654-12-20');
    assert.strictEqual(projected.names.birth_date, '出生日期');
    assert.strictEqual(projected.names.gender, '性别');
    assert.strictEqual(projected.data.province, undefined);
    assert.strictEqual(projected.data.birth_date_from_id, undefined);
    assert.strictEqual(projected.data.gender_from_id, undefined);
});

test('已识别出的出生日期不被身份证号推算值盖掉', () => {
    const projected = presentByFieldConfig(
        { birth_date: '1654年12月20日', birth_date_from_id: '1654-12-20' },
        [{ name: '出生日期', key: 'birth_date', visible: true }]
    );
    assert.strictEqual(projected.data.birth_date, '1654年12月20日');
    assert.strictEqual(Object.keys(projected.data).length, 1);
});

test('尺寸明细的列按字段配置改名、排序并隐藏', () => {
    const fields = [
        { name: '球标号', key: 'seq', visible: true },
        { name: '标称值', key: 'nominal', visible: true },
        { name: '上公差', key: 'upper_tol', visible: false },
        { name: '类型', key: 'dim_type', visible: true },
        { name: '尺寸明细', key: 'items', visible: true }
    ];
    const projected = presentByFieldConfig({
        count: '1',
        items: [{ seq: '1', importance: '非重要', dim_type: '直径', nominal: '25', lower_tol: '-0.02', upper_tol: '+0.02' }]
    }, fields);
    assert.deepStrictEqual(Object.keys(projected.data.items[0]), ['球标号', '标称值', '类型']);
    assert.strictEqual(projected.data.items[0]['球标号'], '1');
    assert.strictEqual(projected.data.items[0]['上公差'], undefined);
    assert.strictEqual(projected.data.count, '1');
});
