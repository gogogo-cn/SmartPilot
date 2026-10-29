'use strict';
/**
 * pricing.js — 方案定价表。
 *
 * 真实定价写在 OpsConsole/config.json（已建，口径见下），本文件的 DEFAULTS 只是
 * config.json 丢失/损坏时的兜底，两者保持一致：
 *   { "team": { "amount": 129, "perSeat": true, "minSeats": 5, "label": "团队版（5人以上）" } }
 * perSeat=true 时按席位计费（amount × seats），否则按订阅计费。
 * minSeats=该方案最少席位数（团队版 5 人起订），发码时校验，默认 1。
 *
 * ⚠️ 套餐名**不决定有效期**：有效期由发码时的 --days 决定（默认 365 天）。
 *    套餐只影响金额统计与界面显示。
 *
 * 2026-10-26 调价（用户拍板）：专业版月付 ¥29 → ¥19；团队版 ¥99 → ¥129/席位/年
 *   并明确「5 人以上」起订；专业版年付 ¥199 不变。
 */

// period：计费周期（月/年），仅用于界面显示清楚，不参与有效期计算。
// note：下拉旁的一句人话说明。
const DEFAULTS = {
    pro_monthly: { amount: 19, perSeat: false, period: '月', label: '专业版 · 月付', note: '一口价，含 1 个席位' },
    pro_yearly: { amount: 199, perSeat: false, period: '年', label: '专业版 · 年付', note: '一口价，含 1 个席位' },
    // 团队版按**席位**计费：¥129 / 席位 / 年，5 人起订（minSeats=5）。
    // 2026-10-26 调价：原 ¥99 上调到 ¥129，同时把「5 人以上」写进套餐名并加起订校验。
    team: {
        amount: 129,
        perSeat: true,
        period: '年',
        minSeats: 5,
        label: '团队版（5人以上）',
        note: '按席位计价：金额 = 129 × 席位数（5 人起订，每人一个独立激活码）'
    }
};

/** 合并 config.json 覆盖项；文件不存在或损坏时用默认值（不炸服务）。 */
function loadPricing(configPath) {
    const merged = JSON.parse(JSON.stringify(DEFAULTS));
    if (!configPath) return merged;
    try {
        if (require('fs').existsSync(configPath)) {
            const j = JSON.parse(require('fs').readFileSync(configPath, 'utf8'));
            for (const k of Object.keys(j || {})) {
                if (!merged[k]) merged[k] = {};
                merged[k].amount = Number(j[k].amount);
                merged[k].perSeat = !!j[k].perSeat;
                merged[k].label = String(j[k].label || merged[k].label || k);
                merged[k].period = String(j[k].period || merged[k].period || '');
                merged[k].note = String(j[k].note || merged[k].note || '');
                // 起订席位数：配置没写按 1（不限制）；写 0/负数/非数字也按 1，防呆
                const ms = parseInt(j[k].minSeats, 10);
                merged[k].minSeats = Number.isFinite(ms) && ms > 0 ? ms : 1;
            }
        }
    } catch (_) { /* 配置损坏时静默回退默认值，运营台不因此挂掉 */ }
    return merged;
}

/** 单个 key 的金额（分币取整到元；退款单在外层剔除，不在这里判）。 */
function priceOf(pricing, plan, seats) {
    const p = pricing && pricing[plan];
    if (!p || !Number.isFinite(p.amount)) return 0;
    return p.perSeat ? p.amount * (parseInt(seats, 10) || 1) : p.amount;
}

/** 该方案的最少席位数（没配 minSeats 一律 1）。 */
function minSeatsOf(pricing, plan) {
    const p = pricing && pricing[plan];
    const ms = parseInt(p && p.minSeats, 10);
    return Number.isFinite(ms) && ms > 0 ? ms : 1;
}

/** ¥ 格式化：整数不带小数点。 */
function formatMoney(n) {
    const v = Number(n) || 0;
    return '¥' + (Number.isInteger(v) ? String(v) : v.toFixed(2));
}

module.exports = { DEFAULTS, loadPricing, priceOf, minSeatsOf, formatMoney };
