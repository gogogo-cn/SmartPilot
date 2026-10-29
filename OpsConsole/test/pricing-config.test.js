'use strict';
/**
 * 定价口径守栏（2026-10-26 调价，用户拍板）：
 *   - 专业版月付 = ¥19 / 月（原 ¥29 下调）
 *   - 专业版年付 = ¥199 / 年（一口价，含 1 个席位），不变
 *   - 团队版 = ¥129 / 席位 / 年（perSeat=true，minSeats=5），金额 = 129 × seats，
 *     5 人起订（原 ¥99 上调，同时把「5 人以上」写进套餐名并加起订校验）
 *   - 专业版月付/年付 = 一口价，不按席位
 *   - 套餐名不决定有效期（有效期由发码 --days 决定）
 */
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const { DEFAULTS, loadPricing, priceOf, minSeatsOf } = require('../lib/pricing');

const CONFIG = path.join(__dirname, '..', 'config.json');

test('config.json 定义了团队版按席位计费', () => {
    assert.ok(fs.existsSync(CONFIG), 'OpsConsole/config.json 应存在');
    const j = JSON.parse(fs.readFileSync(CONFIG, 'utf8'));
    assert.strictEqual(j.team.perSeat, true, '团队版应为按席位计费');
    assert.strictEqual(j.team.amount, 129, '团队版单价应为 129');
    assert.strictEqual(j.team.minSeats, 5, '团队版应 5 人起订');
    assert.ok(/5人以上/.test(j.team.label || ''), '团队版套餐名应带「5人以上」');
    assert.strictEqual(j.pro_yearly.perSeat, false, '专业版年付应为一口价');
    assert.strictEqual(j.pro_yearly.amount, 199, '专业版年付应为 199');
    assert.strictEqual(j.pro_monthly.perSeat, false, '专业版月付应为一口价');
    assert.strictEqual(j.pro_monthly.amount, 19, '专业版月付应为 19');
});

test('兜底 DEFAULTS 与 config.json 口径一致', () => {
    const j = JSON.parse(fs.readFileSync(CONFIG, 'utf8'));
    for (const p of ['pro_monthly', 'pro_yearly', 'team']) {
        assert.strictEqual(DEFAULTS[p].amount, j[p].amount, p + ' 金额兜底应一致');
        assert.strictEqual(DEFAULTS[p].perSeat, j[p].perSeat, p + ' 计费方式兜底应一致');
    }
    assert.strictEqual(DEFAULTS.team.minSeats, 5, '兜底也要带 5 人起订');
});

test('团队版金额 = 129 × 席位', () => {
    const pricing = loadPricing(CONFIG);
    assert.strictEqual(priceOf(pricing, 'team', 5), 645, '5 席位应为 645');
    assert.strictEqual(priceOf(pricing, 'team', 10), 1290, '10 席位应为 1290');
    assert.strictEqual(priceOf(pricing, 'pro_yearly', 5), 199, '一口价方案不随席位变化');
    assert.strictEqual(priceOf(pricing, 'pro_monthly', 3), 19, '月付也是一口价');
});

test('起订席位数：团队版 5，一口价方案 1', () => {
    const pricing = loadPricing(CONFIG);
    assert.strictEqual(minSeatsOf(pricing, 'team'), 5, '团队版 5 人起订');
    assert.strictEqual(minSeatsOf(pricing, 'pro_yearly'), 1, '年付不限制席位');
    assert.strictEqual(minSeatsOf(pricing, 'pro_monthly'), 1, '月付不限制席位');
    assert.strictEqual(minSeatsOf(pricing, 'nope'), 1, '未知方案按 1，不误伤');
    // 配置里写 0/负数/非数字都按 1（防呆）
    const weird = loadPricing(null);
    weird.team.minSeats = 0;
    assert.strictEqual(minSeatsOf(weird, 'team'), 1, 'minSeats=0 应按 1');
});

test('套餐名不参与有效期（有效期由 --days 决定）', () => {
    const lk = require(path.join(__dirname, '..', '..', 'LicenseServer', 'lib', 'licenseKey.js'));
    const k = lk.createKey({ login: 'a@b.c', plan: 'team', seats: 5, days: 30 });
    const v = lk.verify(k);
    assert.strictEqual(v.valid, true);
    assert.strictEqual(v.plan, 'team');
    // 30 天后到期 → 证明周期来自 days 参数，而不是套餐名
    const days = Math.round((new Date(v.periodEnd).getTime() - Date.now()) / 86400000);
    assert.ok(days >= 29 && days <= 30, 'team 套餐的周期应由 --days 决定，实测 ' + days + ' 天');
});
