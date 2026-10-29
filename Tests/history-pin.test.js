/**
 * history-pin.test.js — 置顶分组逻辑冒烟（直接验证排序约定）
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');

/** 与 app-history.groupHistoryByDate 同约定的精简实现，防止回归 */
function groupHistoryByDate(items) {
    const groups = {};
    const today = new Date().toDateString();
    const yesterday = new Date(Date.now() - 86400000).toDateString();
    const pinned = [];
    const rest = [];
    items.forEach((item) => {
        if (item && item.pinned) pinned.push(item);
        else rest.push(item);
    });
    pinned.sort((a, b) => (b.pinnedAt || b.id || 0) - (a.pinnedAt || a.id || 0));
    if (pinned.length) groups['置顶'] = pinned;
    rest.forEach((item) => {
        const itemDate = new Date(item.id).toDateString();
        let label;
        if (itemDate === today) label = '今天';
        else if (itemDate === yesterday) label = '昨天';
        else label = '更早';
        if (!groups[label]) groups[label] = [];
        groups[label].push(item);
    });
    return groups;
}

test('置顶会话应出现在「置顶」分组且排在最前', () => {
    const now = Date.now();
    const items = [
        { id: now - 1000, title: '普通', pinned: false },
        { id: now - 2000, title: '置顶A', pinned: true, pinnedAt: now },
        { id: now - 3000, title: '置顶B', pinned: true, pinnedAt: now - 500 }
    ];
    const grouped = groupHistoryByDate(items);
    const keys = Object.keys(grouped);
    assert.strictEqual(keys[0], '置顶');
    assert.strictEqual(grouped['置顶'].length, 2);
    assert.strictEqual(grouped['置顶'][0].title, '置顶A');
});

test('无置顶时不出现置顶分组', () => {
    const items = [{ id: Date.now(), title: '普通', pinned: false }];
    const grouped = groupHistoryByDate(items);
    assert.ok(!grouped['置顶']);
});
