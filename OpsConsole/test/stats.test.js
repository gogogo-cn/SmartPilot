'use strict';
/**
 * stats.test.js — 运营统计（退款剔除 / 月度窗口 / 临期清单）
 */
const { test } = require('node:test');
const assert = require('node:assert');
const { summarize } = require('../lib/stats');

const NOW = new Date('2026-10-20T12:00:00Z');

function row(over) {
    return Object.assign({
        key: 'SPK1-x.y', login: 'a@b.c', plan: 'pro_yearly', seats: 1,
        expiresAt: '2027-10-20T00:00:00.000Z', note: '', status: 'pending',
        daysLeft: 365, expired: false, expiringSoon: false,
        amount: 199, issuedTime: Date.UTC(2026, 9, 15, 8, 0, 0)
    }, over || {});
}

test('summarize：退款单不计销售额与占比，不计 active', () => {
    const rows = [
        row({ key: 'k1', amount: 199, status: 'sent' }),
        row({ key: 'k2', amount: 599, status: 'refunded' }),
        row({ key: 'k3', amount: 29, plan: 'pro_monthly', status: 'pending' })
    ];
    const s = summarize(rows, NOW);
    assert.strictEqual(s.totalSales, 199 + 29);
    assert.strictEqual(s.refundedCount, 1);
    assert.strictEqual(s.activeCount, 2);
    assert.deepStrictEqual(s.byPlan, { pro_yearly: { count: 1, sales: 199 }, pro_monthly: { count: 1, sales: 29 } });
    assert.ok(!s.byPlan.team, '退款的 team 单不应出现在占比里');
});

test('summarize：本月销售额只算本月签发且未退款的', () => {
    const rows = [
        row({ key: 'k1', amount: 199, issuedTime: Date.UTC(2026, 9, 1, 0, 0, 1) }),   // 10 月 1 日 → 本月
        row({ key: 'k2', amount: 599, issuedTime: Date.UTC(2026, 8, 30, 23, 59, 59) }), // 9 月 → 不计本月
        row({ key: 'k3', amount: 29, issuedTime: Date.UTC(2026, 9, 20, 0, 0, 0), status: 'refunded' }) // 本月但退款
    ];
    const s = summarize(rows, NOW);
    assert.strictEqual(s.monthSales, 199);
    assert.strictEqual(s.totalSales, 199 + 599);
});

test('summarize：临期清单按到期升序、剔退款、过期不计入', () => {
    const rows = [
        row({ key: 'k1', daysLeft: 20, expiringSoon: true, expiresAt: '2026-11-09T00:00:00.000Z' }),
        row({ key: 'k2', daysLeft: 5, expiringSoon: true, expiresAt: '2026-10-25T00:00:00.000Z' }),
        row({ key: 'k3', daysLeft: 2, expiringSoon: true, status: 'refunded' }),
        row({ key: 'k4', daysLeft: -3, expired: true, expiringSoon: false })
    ];
    const s = summarize(rows, NOW);
    assert.deepStrictEqual(s.expiringSoon.map(r => r.key), ['k2', 'k1']);
    assert.strictEqual(s.expiredCount, 1);
});

test('summarize：recent 最多 10 条按签发时间倒序', () => {
    const rows = [];
    for (let i = 0; i < 15; i++) {
        rows.push(row({ key: 'k' + i, issuedTime: Date.UTC(2026, 9, 1) + i * 1000 }));
    }
    const s = summarize(rows, NOW);
    assert.strictEqual(s.recent.length, 10);
    assert.strictEqual(s.recent[0].key, 'k14');
});

test('summarize：空台账不炸', () => {
    const s = summarize([], NOW);
    assert.strictEqual(s.totalSales, 0);
    assert.strictEqual(s.activeCount, 0);
    assert.deepStrictEqual(s.expiringSoon, []);
});
