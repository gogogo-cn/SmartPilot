'use strict';
/**
 * stats.js — 运营统计（纯函数，输入 store.rows() 的合并视图）。
 * 退款单不计入任何销售额与占比；过期单计入"历史有效"但不计入 activeCount。
 */

/**
 * @param {Array} rows store.rows() 输出（含 amount/issuedTime/expired/status）
 * @param {Date|number|string} [now] 可注入的"当前时间"（测试用）
 */
function summarize(rows, now) {
    rows = Array.isArray(rows) ? rows : [];
    const nowDate = now ? new Date(now) : new Date();
    const nowMs = nowDate.getTime();
    const monthStart = Date.UTC(nowDate.getUTCFullYear(), nowDate.getUTCMonth(), 1);

    let monthSales = 0;
    let totalSales = 0;
    let activeCount = 0;
    let pendingCount = 0;
    let sentCount = 0;
    let refundedCount = 0;
    let expiredCount = 0;
    const byPlan = {};

    for (const r of rows) {
        const refunded = r.status === 'refunded';
        if (!refunded) {
            totalSales += r.amount || 0;
            if (r.issuedTime >= monthStart) monthSales += r.amount || 0;
            const bp = byPlan[r.plan] || (byPlan[r.plan] = { count: 0, sales: 0 });
            bp.count++;
            bp.sales += r.amount || 0;
        }
        if (r.status === 'pending') pendingCount++;
        else if (r.status === 'sent') sentCount++;
        else if (r.status === 'refunded') refundedCount++;
        if (!refunded) {
            if (r.expired) expiredCount++;
            else activeCount++;
        }
    }

    const expiringSoon = rows
        .filter(r => r.expiringSoon && r.status !== 'refunded')
        .sort((a, b) => String(a.expiresAt).localeCompare(String(b.expiresAt)));
    const recent = rows.slice().sort((a, b) => b.issuedTime - a.issuedTime).slice(0, 10);

    return {
        now: nowDate.toISOString(),
        monthStart: new Date(monthStart).toISOString(),
        monthSales, totalSales,
        activeCount, pendingCount, sentCount, refundedCount, expiredCount,
        expiringSoon, byPlan, recent
    };
}

module.exports = { summarize };
