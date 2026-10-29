'use strict';
/**
 * ledger.test.js — CSV 台账读写（含与 gen-key 的格式对齐细节）
 */
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const ledger = require('../lib/ledger');

function tmpFile(name) {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ops-ledger-'));
    return path.join(dir, name || 'issued-keys.csv');
}

test('csvCell：含逗号/引号/换行时加引号并双写引号', () => {
    assert.strictEqual(ledger.csvCell('plain'), 'plain');
    assert.strictEqual(ledger.csvCell('a,b'), '"a,b"');
    assert.strictEqual(ledger.csvCell('say "hi"'), '"say ""hi"""');
    assert.strictEqual(ledger.csvCell('line1\nline2'), '"line1\nline2"');
    assert.strictEqual(ledger.csvCell(null), '');
});

test('parseCsv：引号内逗号/换行/双写引号正确解析', () => {
    const text = 'a,"b,c",d\r\n"say ""hi""","line1\nline2"\r\n';
    const rows = ledger.parseCsv(text);
    assert.deepStrictEqual(rows, [
        ['a', 'b,c', 'd'],
        ['say "hi"', 'line1\nline2']
    ]);
});

test('parseCsv：空文本与尾随换行不产生幻影行', () => {
    assert.deepStrictEqual(ledger.parseCsv(''), []);
    assert.deepStrictEqual(ledger.parseCsv('a,b\r\n'), [['a', 'b']]);
    assert.deepStrictEqual(ledger.parseCsv('a,b\n\nc,d\n'), [['a', 'b'], ['c', 'd']]);
});

test('appendRows：新文件带表头，CRLF，7 列', () => {
    const file = tmpFile();
    ledger.appendRows(file, [
        { key: 'SPK1-aaa.bbb', login: 'u@x.com', plan: 'pro_yearly', seats: 1, expiresAt: '2027-01-01T00:00:00.000Z' }
    ], '订单A');
    const text = fs.readFileSync(file, 'utf8');
    const lines = text.split('\r\n');
    assert.strictEqual(lines[0], ledger.HEADER);
    assert.strictEqual(lines.length, 3); // 表头 + 1 行 + 尾部空串
    assert.strictEqual(lines[1].split(',').length, 7);
    assert.ok(lines[1].endsWith(',订单A'));
    // 再追加不重复表头
    ledger.appendRows(file, [
        { key: 'SPK1-ccc.ddd', login: 'v@x.com', plan: 'team', seats: 3, expiresAt: '2027-02-01T00:00:00.000Z' }
    ], '订单B');
    const text2 = fs.readFileSync(file, 'utf8');
    assert.strictEqual(text2.split('\r\n')[0], ledger.HEADER);
    assert.strictEqual((text2.match(/issuedAt,key/g) || []).length, 1, '表头只允许出现一次');
});

test('readLedger：跳过表头/空行，note 含引号逗号可还原', () => {
    const file = tmpFile();
    ledger.appendRows(file, [
        { key: 'SPK1-aaa.bbb', login: 'u@x.com', plan: 'pro_yearly', seats: 1, expiresAt: '2027-01-01T00:00:00.000Z' },
        { key: 'SPK1-ccc.ddd', login: 'v@x.com', plan: 'team', seats: 2, expiresAt: '2027-02-01T00:00:00.000Z' }
    ], '张三, 微信订单20261021');
    const rows = ledger.readLedger(file);
    assert.strictEqual(rows.length, 2);
    assert.strictEqual(rows[0].key, 'SPK1-aaa.bbb');
    assert.strictEqual(rows[1].seats, '2');
    assert.strictEqual(rows[1].note, '张三, 微信订单20261021', '含逗号的 note 必须完整还原');
});

test('readLedger：文件不存在返回空数组', () => {
    assert.deepStrictEqual(ledger.readLedger('/nonexistent/xyz.csv'), []);
});
