'use strict';
/**
 * store.test.js — 合并视图 / 生成登记 / 状态层 / 搜索 / 客户聚合
 */
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { OpsStore } = require('../lib/store');
const lk = require('../../LicenseServer/lib/licenseKey.js');

function mkStore() {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ops-store-'));
    const store = new OpsStore({
        ledgerPath: path.join(dir, 'issued-keys.csv'),
        statePath: path.join(dir, 'ops-state.json')
    });
    return { store, dir };
}

test('addKeys：生成的 key 能过验签，且写进台账（可读回）', () => {
    const { store } = mkStore();
    const created = store.addKeys({ login: 'a@b.c', plan: 'pro_yearly', seats: 2, days: 365, note: '测试单' });
    assert.strictEqual(created.length, 1);
    assert.strictEqual(lk.verify(created[0].key).valid, true);
    assert.strictEqual(lk.verify(created[0].key).seats, 2);
    const rows = store.rows();
    assert.strictEqual(rows.length, 1);
    assert.strictEqual(rows[0].login, 'a@b.c');
    assert.strictEqual(rows[0].note, '测试单');
    assert.strictEqual(rows[0].status, 'pending', '新 key 默认待发货');
});

test('addKeys：批量时 login 自动加 +N 后缀（一客一码，与 gen-key 同规则）', () => {
    const { store } = mkStore();
    // 团队版 5 人起订：count=5 刚好踩线（people = max(seats, count)）
    const created = store.addKeys({ login: 'team@corp.com', plan: 'team', days: 365, count: 5 });
    assert.deepStrictEqual(created.map(r => r.login),
        ['team+1@corp.com', 'team+2@corp.com', 'team+3@corp.com', 'team+4@corp.com', 'team+5@corp.com']);
    const keys = new Set(created.map(r => r.key));
    assert.strictEqual(keys.size, 5, '同参数批量 key 必须互异（nonce）');
});

test('addKeys：团队版不满 5 人起订直接拒（OpsError 400）', () => {
    const { store } = mkStore();
    // 2026-10-26 用户拍板：团队版 5 人以上。少于 5 人发码会和价格口径对不上，
    // 这里守住 store 层，server 层把它转成 400（见 server.test.js）。
    assert.throws(() => store.addKeys({ login: 'a@b.c', plan: 'team', seats: 3, days: 365 }), (e) => {
        assert.strictEqual(e.name, 'OpsError');
        assert.strictEqual(e.status, 400);
        assert.ok(/5 人起订/.test(e.message), '错误信息要写清起订量：' + e.message);
        return true;
    });
    // 恰好 5 人放行；一口价方案不受起订量约束
    const ok = store.addKeys({ login: 'a@b.c', plan: 'team', seats: 5, days: 365 });
    assert.strictEqual(ok.length, 1);
    const monthly = store.addKeys({ login: 'a@b.c', plan: 'pro_monthly', days: 31 });
    assert.strictEqual(monthly.length, 1);
});

test('addKeys：不填 login 时台账记空串（与 gen-key CLI 行为一致：verify 把 - 归一成空）', () => {
    const { store } = mkStore();
    const created = store.addKeys({ plan: 'pro_monthly', days: 31 });
    assert.strictEqual(created[0].login, '');
    assert.strictEqual(store.rows()[0].login, '');
});

test('rows 派生字段：到期/临期/金额/过期', () => {
    const { store } = mkStore();
    // 固定"现在"，三行台账日期全部相对它写死——不依赖真实时钟。
    // （旧版用 addKeys 生成真实 key（expiresAt 按真实当前时间算）再对比写死的 now，
    //   UTC 日期每天后滚一天、daysLeft 就 +1，是颗日期炸弹：2026-10-21 已炸出 367 天。
    //   生成链路本身由上面几个 addKeys 用例覆盖，这里只测 rows() 的派生公式。）
    const now = new Date('2026-10-20T00:00:00Z');
    const ledger = require('../lib/ledger');
    ledger.appendRows(store.ledgerPath, [
        { key: 'SPK1-y1.y1', login: 'a@b.c', plan: 'pro_yearly', seats: 1, expiresAt: '2027-10-20T00:00:00.000Z' }, // 365 天后
        { key: 'SPK1-m1.m1', login: 'c@d.e', plan: 'pro_monthly', seats: 1, expiresAt: '2026-10-30T00:00:00.000Z' }, // 10 天后
        { key: 'SPK1-x1.x1', login: 'f@g.h', plan: 'team', seats: 5, expiresAt: '2026-10-15T00:00:00.000Z' }         // 已过期
    ], 'z');
    const rows = store.rows({ now });
    const byLogin = Object.fromEntries(rows.map(r => [r.login, r]));
    // daysLeft = ceil((expiresAt - now) / 1天)；日期都是零点整，结果精确无边界抖动
    assert.strictEqual(byLogin['a@b.c'].daysLeft, 365, '年付 key 应剩 365 天');
    assert.strictEqual(byLogin['a@b.c'].expiringSoon, false);
    assert.strictEqual(byLogin['c@d.e'].daysLeft, 10);
    assert.strictEqual(byLogin['c@d.e'].expiringSoon, true);
    assert.strictEqual(byLogin['f@g.h'].expired, true);
    // 默认定价（2026-10-26 调价后）：pro_yearly 199 / pro_monthly 19 / team 129×席位
    assert.strictEqual(byLogin['a@b.c'].amount, 199);
    assert.strictEqual(byLogin['c@d.e'].amount, 19);
    assert.strictEqual(byLogin['f@g.h'].amount, 645, '团队版按席位：129 × 5');
});

test('setStatus：三态流转 + 时间戳只记第一次 + 非法状态拒绝', () => {
    const { store } = mkStore();
    const created = store.addKeys({ login: 'a@b.c', plan: 'pro_yearly', days: 365 });
    const key = created[0].key;
    assert.throws(() => store.setStatus(key, { status: 'bogus' }), /非法状态/);
    const s1 = store.setStatus(key, { status: 'sent' });
    assert.strictEqual(s1.status, 'sent');
    assert.ok(s1.sentAt, '标记已发货应记录时间');
    const firstSentAt = s1.sentAt;
    store.setStatus(key, { status: 'pending' });
    const s2 = store.setStatus(key, { status: 'sent' });
    assert.strictEqual(s2.sentAt, firstSentAt, '重复标记不应刷新时间戳');
    const s3 = store.setStatus(key, { status: 'refunded', opsNote: '客户不要了' });
    assert.strictEqual(s3.status, 'refunded');
    assert.ok(s3.refundedAt);
    assert.strictEqual(s3.opsNote, '客户不要了');
    // 覆盖层持久化：重读后仍在
    const row = store.rows().find(r => r.key === key);
    assert.strictEqual(row.status, 'refunded');
    assert.strictEqual(row.opsNote, '客户不要了');
});

test('setStatus：状态文件损坏时不炸，回退默认 pending', () => {
    const { store, dir } = mkStore();
    const created = store.addKeys({ login: 'a@b.c', plan: 'pro_yearly', days: 365 });
    fs.writeFileSync(store.statePath, '{ broken json', 'utf8');
    const row = store.rows().find(r => r.key === created[0].key);
    assert.strictEqual(row.status, 'pending');
});

test('search：q 匹配 key/login/备注，plan/status 筛选，排序', () => {
    const { store } = mkStore();
    store.addKeys({ login: 'alice@x.com', plan: 'pro_yearly', days: 365, note: '微信订单001' });
    store.addKeys({ login: 'bob@x.com', plan: 'team', seats: 5, days: 365, note: '支付宝订单002' });
    store.addKeys({ login: 'carol@x.com', plan: 'pro_monthly', days: 31, note: '微信订单003' });
    store.setStatus(store.rows().find(r => r.login === 'bob@x.com').key, { status: 'sent' });

    assert.strictEqual(store.search({ q: 'alice' }).length, 1);
    assert.strictEqual(store.search({ q: '微信' }).length, 2);
    assert.strictEqual(store.search({ q: 'SPK1-' }).length, 3);
    assert.strictEqual(store.search({ plan: 'team' }).length, 1);
    assert.strictEqual(store.search({ status: 'sent' }).length, 1);
    assert.strictEqual(store.search({ status: 'refunded' }).length, 0);
    // 默认按生成时间倒序；expiresAt 排序时临期的排前面
    const byExp = store.search({ sort: 'expiresAt' });
    assert.strictEqual(byExp[0].login, 'carol@x.com', '31 天到期的应排最前');
});

test('customers：按 login 聚合，退款不计金额，未指定 login 归一组', () => {
    const { store } = mkStore();
    store.addKeys({ login: 'alice@x.com', plan: 'pro_yearly', days: 365 });
    store.addKeys({ login: 'alice@x.com', plan: 'pro_monthly', days: 31 });
    store.addKeys({ login: 'bob@x.com', plan: 'team', seats: 5, days: 365 });
    store.addKeys({ plan: 'pro_yearly', days: 365 }); // login 缺省 → -
    store.setStatus(store.rows().find(r => r.login === 'bob@x.com').key, { status: 'refunded' });

    const cs = store.customers();
    const alice = cs.find(c => c.login === 'alice@x.com');
    assert.strictEqual(alice.orders, 2);
    assert.strictEqual(alice.amount, 199 + 19);
    assert.strictEqual(alice.active, 2);
    assert.deepStrictEqual(alice.plans, { pro_yearly: 1, pro_monthly: 1 });
    const bob = cs.find(c => c.login === 'bob@x.com');
    assert.strictEqual(bob.amount, 0, '退款单不计金额');
    assert.strictEqual(bob.active, 0);
    assert.ok(cs.find(c => c.login === '(未指定)'), '未指定 login 也要有档');
    // 按金额降序：alice(218) > 未指定(199) > bob(0)
    assert.strictEqual(cs[0].login, 'alice@x.com');
});

test('pricing：config.json 可覆盖默认定价（perSeat 生效）', () => {
    const { store, dir } = mkStore();
    const cfg = path.join(dir, 'config.json');
    fs.writeFileSync(cfg, JSON.stringify({ pro_yearly: { amount: 299, perSeat: true, label: '年付·按席位' } }), 'utf8');
    store.pricingPath = cfg;
    const created = store.addKeys({ login: 'a@b.c', plan: 'pro_yearly', seats: 3, days: 365 });
    const row = store.rows().find(r => r.key === created[0].key);
    assert.strictEqual(row.amount, 897, 'perSeat=true 时金额 = amount × seats');
    assert.strictEqual(store.pricing().pro_yearly.label, '年付·按席位');
});
