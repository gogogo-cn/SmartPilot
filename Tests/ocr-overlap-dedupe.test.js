/**
 * 同一行两个检测框叠在同一个字上时，这个字只能留下一次。
 * 用户截图：列表里同时出现「出生」和「生1654年12月20日」。
 * 运行：node Tests/ocr-overlap-dedupe.test.js
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const { dedupeOverlappingOcrItems } = require('../UI/js/ocr-box-dedupe.js');

function box(l, t, r, b) {
    return [{ x: l, y: t }, { x: r, y: t }, { x: r, y: b }, { x: l, y: b }];
}

function texts(items) {
    return items.map(it => it.text);
}

test('交界上的「生」只保留在「出生」里', () => {
    // 「出生」宽 64；日期框从「生」的左缘起，重叠正好一个字
    const items = [
        { text: '出生', confidence: 1, box: box(100, 40, 164, 72) },
        { text: '生1654年12月20日', confidence: 0.97, box: box(132, 40, 400.8, 72) }
    ];
    const out = dedupeOverlappingOcrItems(items);
    assert.deepStrictEqual(texts(out), ['出生', '1654年12月20日']);
    assert.strictEqual(items[1].text, '生1654年12月20日', '不改调用方原来的对象');
    const again = dedupeOverlappingOcrItems(out);
    assert.deepStrictEqual(texts(again), ['出生', '1654年12月20日']);
});

test('日期框盖住「出生」两个字宽时，仍只去掉一个「生」', () => {
    const items = [
        { text: '出生', confidence: 1, box: box(100, 40, 164, 72) },
        { text: '生1654年12月20日', confidence: 0.97, box: box(100, 40, 420, 72) }
    ];
    assert.deepStrictEqual(texts(dedupeOverlappingOcrItems(items)), ['出生', '1654年12月20日']);
});

test('两个框只隔一条缝时，交界上的「生」也去掉', () => {
    const items = [
        { text: '出生', confidence: 1, box: box(100, 40, 164, 72) },
        { text: '生1654年12月20日', confidence: 0.97, box: box(172, 40, 440, 72) }
    ];
    assert.deepStrictEqual(texts(dedupeOverlappingOcrItems(items)), ['出生', '1654年12月20日']);
});

test('框之间有空隙时，相同的字都保留', () => {
    const items = [
        { text: '甲乙', confidence: 1, box: box(100, 40, 164, 72) },
        { text: '乙丙丁', confidence: 0.97, box: box(180, 40, 448, 72) }
    ];
    assert.deepStrictEqual(texts(dedupeOverlappingOcrItems(items)), ['甲乙', '乙丙丁']);
});

test('上下两行即使横坐标重叠也不去字', () => {
    const items = [
        { text: '出生', confidence: 1, box: box(100, 40, 164, 72) },
        { text: '生1654年12月20日', confidence: 0.97, box: box(132, 90, 400.8, 122) }
    ];
    assert.deepStrictEqual(texts(dedupeOverlappingOcrItems(items)), ['出生', '生1654年12月20日']);
});

test('重叠但交界文字不同时不去字', () => {
    const items = [
        { text: '出生', confidence: 1, box: box(100, 40, 164, 72) },
        { text: '于1654年12月20日', confidence: 0.97, box: box(132, 40, 400.8, 72) }
    ];
    assert.deepStrictEqual(texts(dedupeOverlappingOcrItems(items)), ['出生', '于1654年12月20日']);
});

test('字更靠近右边的短框时，从左边去掉', () => {
    const items = [
        { text: '甲乙生', confidence: 0.9, box: box(0, 10, 96, 42) },
        { text: '生丙', confidence: 0.9, box: box(64, 10, 128, 42) }
    ];
    assert.deepStrictEqual(texts(dedupeOverlappingOcrItems(items)), ['甲乙', '生丙']);
});

test('两个框读到同一段文字时只留一条', () => {
    const items = [
        { text: '出生', confidence: 0.99, box: box(0, 0, 64, 32) },
        { text: '出生', confidence: 0.91, box: box(0, 0, 64, 32) }
    ];
    const out = dedupeOverlappingOcrItems(items);
    assert.deepStrictEqual(texts(out), ['出生']);
});

test('ocr-panel 会在拿到结果时调用去重', () => {
    const html = fs.readFileSync(path.join(__dirname, '..', 'UI', 'ocr-panel.html'), 'utf8');
    assert.match(html, /js\/ocr-box-dedupe\.js/);
    assert.match(html, /dedupeOverlappingOcrItems\(structObj\.items\)/);
    assert.match(html, /dedupeOverlappingOcrItems\(items\)/);
});
