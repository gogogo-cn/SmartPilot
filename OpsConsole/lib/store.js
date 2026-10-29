'use strict';
/**
 * store.js — 运营台数据层：合并视图 + 写操作。
 *
 * 权威与覆盖层分离：
 *   · 权威账本 = LicenseServer/data/issued-keys.csv（gen-key 命令行也在追加），只读+追加；
 *   · 运营状态 = OpsConsole/data/ops-state.json（待发货/已发货/已退款、运营备注），按 key 索引。
 * 每次 rows() 都重读两个文件——外部用命令行补货后，运营台刷新即见，无需重启。
 */
const fs = require('fs');
const path = require('path');
const lk = require(path.join(__dirname, '..', '..', 'LicenseServer', 'lib', 'licenseKey.js'));
const ledger = require('./ledger');
const { loadPricing, priceOf, minSeatsOf } = require('./pricing');

const STATE_VERSION = 1;
const VALID_STATUS = ['pending', 'sent', 'refunded'];
const EXPIRING_DAYS = 30;

/** 带 HTTP 状态码的错误：运营台接口直接透传 status，不糊成 500。 */
class OpsError extends Error {
    constructor(status, message) {
        super(message);
        this.name = 'OpsError';
        this.status = status || 400;
    }
}

class OpsStore {
    /**
     * @param {{ledgerPath:string, statePath:string, pricingPath?:string}} opts
     */
    constructor(opts) {
        this.ledgerPath = opts.ledgerPath;
        this.statePath = opts.statePath;
        this.pricingPath = opts.pricingPath || null;
    }

    _readState() {
        try {
            if (!fs.existsSync(this.statePath)) return { version: STATE_VERSION, keys: {} };
            const j = JSON.parse(fs.readFileSync(this.statePath, 'utf8'));
            return { version: STATE_VERSION, keys: (j && j.keys) || {} };
        } catch (_) {
            return { version: STATE_VERSION, keys: {} }; // 状态文件损坏不拖垮台账展示
        }
    }

    _writeState(state) {
        const dir = path.dirname(this.statePath);
        if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
        const tmp = this.statePath + '.tmp';
        fs.writeFileSync(tmp, JSON.stringify(state, null, 2), 'utf8');
        fs.renameSync(tmp, this.statePath); // 原子替换，防写一半
    }

    pricing() {
        return loadPricing(this.pricingPath);
    }

    /**
     * 合并视图：CSV 行 + 状态覆盖层 + 派生字段（到期/临期/金额）。
     * @param {{now?:Date|number|string}} [opts] 测试可注入当前时间
     */
    rows(opts) {
        const now = opts && opts.now ? new Date(opts.now) : new Date();
        const nowMs = now.getTime();
        const state = this._readState();
        const pricing = this.pricing();
        return ledger.readLedger(this.ledgerPath).map(r => {
            const st = state.keys[r.key] || {};
            const status = VALID_STATUS.indexOf(st.status) >= 0 ? st.status : 'pending';
            const exp = r.expiresAt ? new Date(r.expiresAt) : null;
            const expMs = exp ? exp.getTime() : 0;
            const daysLeft = exp ? Math.ceil((expMs - nowMs) / 86400000) : null;
            const seats = parseInt(r.seats, 10) || 1;
            return {
                issuedAt: r.issuedAt,
                key: r.key,
                login: r.login,
                plan: r.plan,
                seats,
                expiresAt: r.expiresAt,
                note: r.note,
                status,
                sentAt: st.sentAt || '',
                refundedAt: st.refundedAt || '',
                opsNote: st.opsNote || '',
                daysLeft,
                expired: !!(exp && expMs <= nowMs),
                expiringSoon: !!(exp && daysLeft >= 0 && daysLeft <= EXPIRING_DAYS),
                amount: priceOf(pricing, r.plan, seats),
                issuedTime: r.issuedAt ? new Date(r.issuedAt).getTime() : 0
            };
        });
    }

    /**
     * 生成激活码并登记台账（复用 LicenseServer/lib/licenseKey.js，保证与 VB 端同算法）。
     * 批量时 login 自动加 +N 后缀——与 gen-key --batch 同规则，一客一码。
     * @returns {Array<{key,login,plan,seats,expiresAt}>}
     */
    addKeys(opts) {
        opts = opts || {};
        const count = Math.max(1, Math.min(100, parseInt(opts.count, 10) || 1));
        const plan = String(opts.plan || 'pro_yearly');
        const seats = parseInt(opts.seats, 10) || 1;
        const days = parseInt(opts.days, 10) || 365;
        const loginRaw = String(opts.login || '').trim();
        // 起订校验：团队版 5 人以上（定价表 team.minSeats）。按席位方案里 seats 就是人数，
        // 不够起订量直接拒，避免发出 3 人团的"团队版"码，价格口径和承诺对不上。
        const pricing = this.pricing();
        const need = minSeatsOf(pricing, plan);
        const people = pricing[plan] && pricing[plan].perSeat ? Math.max(seats, count) : seats;
        if (people < need) {
            throw new OpsError(400, (pricing[plan] && pricing[plan].label || plan) + ' 需 ' + need + ' 人起订，当前 ' + people + ' 人');
        }
        const rows = [];
        for (let i = 0; i < count; i++) {
            const login = loginRaw
                ? (count > 1
                    ? loginRaw.replace(/@.*$/, '') + '+' + (i + 1) + '@' + (loginRaw.split('@')[1] || 'local')
                    : loginRaw)
                : '-';
            const key = lk.createKey({ login, plan, seats, days });
            const v = lk.verify(key);
            if (!v.valid) throw new Error('自检失败：生成的 key 无法通过验签 -> ' + v.reason);
            rows.push({ key, login: v.login, plan: v.plan, seats: v.seats, expiresAt: v.periodEnd });
        }
        ledger.appendRows(this.ledgerPath, rows, String(opts.note || ''));
        return rows;
    }

    /** 更新运营状态（写覆盖层；CSV 账本不动）。 */
    setStatus(key, patch) {
        patch = patch || {};
        const status = String(patch.status || '');
        if (VALID_STATUS.indexOf(status) < 0) throw new Error('非法状态: ' + status + '（允许 ' + VALID_STATUS.join('/') + '）');
        const state = this._readState();
        const cur = state.keys[key] || {};
        const next = Object.assign({}, cur, { status });
        if (patch.opsNote !== undefined) next.opsNote = String(patch.opsNote || '');
        if (status === 'sent' && !next.sentAt) next.sentAt = new Date().toISOString();
        if (status === 'refunded' && !next.refundedAt) next.refundedAt = new Date().toISOString();
        state.keys[key] = next;
        this._writeState(state);
        return next;
    }

    /** 搜索 / 筛选 / 排序。q 匹配 key、login、plan、备注、运营备注。 */
    search(query) {
        query = query || {};
        let rows = this.rows();
        const q = String(query.q || '').trim().toLowerCase();
        if (q) {
            rows = rows.filter(r =>
                [r.key, r.login, r.plan, r.note, r.opsNote].join(' ').toLowerCase().indexOf(q) >= 0);
        }
        if (query.plan) rows = rows.filter(r => r.plan === query.plan);
        if (query.status) rows = rows.filter(r => r.status === query.status);
        if (query.sort === 'expiresAt') {
            rows.sort((a, b) => String(a.expiresAt).localeCompare(String(b.expiresAt)));
        } else {
            rows.sort((a, b) => b.issuedTime - a.issuedTime);
        }
        return rows;
    }

    /** 按客户（login）聚合：订单数、累计金额（剔退款）、有效 key 数、各 plan 分布、最晚到期。 */
    customers() {
        const map = new Map();
        for (const r of this.rows()) {
            const login = r.login || '(未指定)';
            if (!map.has(login)) {
                map.set(login, { login, orders: 0, amount: 0, active: 0, lastExpiry: '', plans: {} });
            }
            const c = map.get(login);
            c.orders++;
            if (r.status !== 'refunded') c.amount += r.amount;
            if (!r.expired && r.status !== 'refunded') c.active++;
            c.plans[r.plan] = (c.plans[r.plan] || 0) + 1;
            if (String(r.expiresAt) > c.lastExpiry) c.lastExpiry = r.expiresAt;
        }
        return Array.from(map.values()).sort((a, b) => b.amount - a.amount);
    }
}

module.exports = { OpsStore, OpsError, VALID_STATUS, EXPIRING_DAYS };
