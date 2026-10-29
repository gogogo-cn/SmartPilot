'use strict';
/**
 * server.js — SmartPilot 运营台（OpsConsole）。
 *
 * 本地单人工具：只绑 127.0.0.1，零第三方依赖（node 内置模块），
 * 启动：node OpsConsole/server.js  （或 npm start --prefix OpsConsole）
 * 端口：默认 9600，可用环境变量 OPS_PORT 覆盖。
 *
 * 路由：
 *   GET  /                     运营台单页（public/index.html）
 *   GET  /api/overview         统计卡片 + 临期清单 + 最近订单
 *   GET  /api/keys             台账合并视图（?q=&plan=&status=&sort=）
 *   POST /api/keys             生成激活码 {login,plan,seats,days,note,count}
 *   POST /api/keys/status      标记状态 {key,status,opsNote}
 *   GET  /api/customers        客户档案聚合
 *   GET  /api/pricing          当前定价表
 *   GET  /api/export.csv       合并视图导出（UTF-8 BOM，Excel 直接打开不乱码）
 */
const http = require('http');
const fs = require('fs');
const path = require('path');

// ---- 启动预检：LicenseServer 目录必须存在 ----
// ⚠️ 「不需要服务器」= **不用把 LicenseServer 当服务部署/运行**，不等于可以删掉这个目录：
//    lib/licenseKey.js 是 Excel/Word/PPT/运营台三端共用的激活码算法**唯一来源**，
//    data/issued-keys.csv 是激活码**权威台账**。删了它运营台会 MODULE_NOT_FOUND，
//    这里提前给出可操作的提示，避免让人误以为"不用服务器就把它清掉了"。
(function preflight() {
    // 显式指定 LICENSE_SERVER_DIR 时（自定义部署 / 单测隔离）只警告不退出，
    // 因为调用方是有意识地把目录指到别处的，不是误删。
    const explicit = !!process.env.LICENSE_SERVER_DIR;
    const lsDir = process.env.LICENSE_SERVER_DIR || path.join(__dirname, '..', 'LicenseServer');
    const missing = [
        path.join(lsDir, 'lib', 'licenseKey.js'),
        path.join(lsDir, 'lib', 'revocation.js')
    ].filter(function (p) { return !fs.existsSync(p); });
    if (!missing.length) return;
    if (explicit) {
        console.warn('[OpsConsole] 警告：LICENSE_SERVER_DIR 指向的目录缺少激活码算法文件，相关功能可能不可用');
        return;
    }
    console.error('');
    console.error('[OpsConsole] 启动失败：找不到激活码算法文件');
    missing.forEach(function (p) { console.error('   缺失: ' + p); });
    console.error('');
    console.error('  说明：运营台**不需要运行** LicenseServer 服务，但**需要**它的 lib/ 目录 ——');
    console.error('        licenseKey.js 是三端共用算法的唯一来源，data/issued-keys.csv 是权威台账。');
    console.error('        请勿删除该目录（它只是不当服务启动，不是没用）。');
    console.error('');
    process.exit(1);
})();

const { OpsStore } = require('./lib/store');
const { summarize } = require('./lib/stats');
const { customerMessage } = require('./lib/customerText');
const { formatMoney } = require('./lib/pricing');
const { verifyProof, refundByProof } = require('./lib/deactivation');

const ROOT = path.join(__dirname, '..');
const DEFAULT_LEDGER = path.join(ROOT, 'LicenseServer', 'data', 'issued-keys.csv');
const DEFAULT_STATE = path.join(__dirname, 'data', 'ops-state.json');
const DEFAULT_PRICING = path.join(__dirname, 'config.json');
const PUBLIC_INDEX = path.join(__dirname, 'public', 'index.html');

// 支付配置写到 LicenseServer 的 config.json（与 paymentConfig.js 共用同一份），
// 管理平台填配置即上线，无需手改 .env、无需重启（若在跑则热重载）。
const LICENSE_SERVER_DIR = process.env.LICENSE_SERVER_DIR || path.join(ROOT, 'LicenseServer');
const LICENSE_CONFIG = path.join(LICENSE_SERVER_DIR, 'config.json');
const LICENSE_ENV = path.join(LICENSE_SERVER_DIR, '.env');
const LICENSE_ADMIN_API = (process.env.LICENSE_ADMIN_URL || 'http://127.0.0.1:9530').replace(/\/$/, '');

// 与 LicenseServer/lib/paymentConfig.js 的 KEYS 保持一致（仅允许这些键被平台写入）
const PAYMENT_KEYS = [
    'STRIPE_SECRET_KEY', 'STRIPE_WEBHOOK_SECRET',
    'STRIPE_PRICE_PRO_MONTHLY', 'STRIPE_PRICE_PRO_YEARLY',
    'STRIPE_SUCCESS_URL', 'STRIPE_CANCEL_URL',
    // 国内支付
    'CN_PAY_PROVIDER', 'CN_ICP_BEIAN', 'CN_PAY_SUBJECT',
    'CN_NOTIFY_URL', 'CN_RETURN_URL', 'CN_YEARLY_AMOUNT_FEN',
    'WECHAT_APP_ID', 'WECHAT_MCH_ID', 'WECHAT_API_V3_KEY', 'WECHAT_SERIAL_NO',
    'WECHAT_PRIVATE_KEY', 'WECHAT_PLATFORM_PUBLIC_KEY', 'WECHAT_API_KEY',
    'ALIPAY_APP_ID', 'ALIPAY_PRIVATE_KEY', 'ALIPAY_PUBLIC_KEY', 'ALIPAY_GATEWAY',
    'PUBLIC_BASE_URL'
];

const MAX_BODY = 64 * 1024;

/**
 * CORS 白名单。运营台只绑 127.0.0.1（本机单人工具），但用户经常直接双击
 * public/index.html 打开——file:// 的 Origin 是 "null"，不放行就连不上 API。
 * 放行：null 源 + 127.0.0.1/localhost 的 http 源；其余一律不给 CORS 头
 * （本机服务也不该被任意网页读数据，DNS rebinding 之类的旁路挡在这里）。
 */
function corsHeaders(req) {
    const origin = req.headers.origin;
    if (!origin) return {};
    const ok = origin === 'null'
        || /^https?:\/\/127\.0\.0\.1(:\d+)?$/.test(origin)
        || /^https?:\/\/localhost(:\d+)?$/.test(origin);
    if (!ok) return {};
    return {
        'Access-Control-Allow-Origin': origin,
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, X-Admin-Key',
        'Access-Control-Max-Age': '600',
        'Vary': 'Origin'
    };
}

function sendJson(res, code, obj) {
    const body = JSON.stringify(obj);
    res.writeHead(code, {
        'Content-Type': 'application/json; charset=utf-8',
        'Cache-Control': 'no-store'
    });
    res.end(body);
}

function sendError(res, code, msg) {
    sendJson(res, code, { success: false, error: msg });
}

// 读取 LicenseServer/.env（用于取 ADMIN_ACTIVATE_KEY 触发热重载）
function readEnvFile(p) {
    const out = {};
    try {
        const lines = fs.readFileSync(p, 'utf8').split(/\r?\n/);
        lines.forEach((l) => {
            const m = /^([A-Z0-9_]+)\s*=\s*(.*)$/.exec(l);
            if (m) out[m[1]] = m[2].replace(/^["']|["']$/g, '');
        });
    } catch (_) { /* 无 .env 也行 */ }
    return out;
}

// 最小的 http JSON 请求（best-effort，用于热重载到运行中的 LicenseServer）
function httpJson(method, urlStr, bodyObj, headers) {
    return new Promise((resolve) => {
        let u;
        try { u = new URL(urlStr); } catch (_) { return resolve(null); }
        const data = bodyObj ? JSON.stringify(bodyObj) : null;
        const r = http.request({
            hostname: u.hostname,
            port: u.port,
            path: u.pathname,
            method,
            headers: Object.assign({ 'Content-Type': 'application/json' }, headers || {})
        }, (resp) => {
            let b = '';
            resp.on('data', (c) => { b += c; });
            resp.on('end', () => {
                try { resolve({ status: resp.statusCode, json: JSON.parse(b || '{}') }); }
                catch (_) { resolve({ status: resp.statusCode, json: {} }); }
            });
        });
        r.on('error', () => resolve(null));
        r.setTimeout(1500, () => { r.destroy(); resolve(null); });
        if (data) r.write(data);
        r.end();
    });
}

function maskPayment(cfg) {
    const out = {};
    PAYMENT_KEYS.forEach((k) => {
        const v = cfg[k] || '';
        if (/SECRET|KEY|PRIVATE/.test(k)) out[k] = v ? (v.slice(0, 4) + '…' + v.slice(-4)) : '';
        else out[k] = v;
    });
    return out;
}

function readBody(req) {
    return new Promise((resolve, reject) => {
        let size = 0;
        const chunks = [];
        req.on('data', (c) => {
            size += c.length;
            if (size > MAX_BODY) {
                reject(new Error('请求体过大'));
                req.destroy();
                return;
            }
            chunks.push(c);
        });
        req.on('end', () => {
            if (!chunks.length) return resolve({});
            try {
                resolve(JSON.parse(Buffer.concat(chunks).toString('utf8')));
            } catch (_) {
                reject(new Error('JSON 解析失败'));
            }
        });
        req.on('error', reject);
    });
}

/**
 * 创建服务（不 listen，供测试注入临时目录）。
 * @param {{store: OpsStore, indexHtml?: string}} opts
 */
function createServer(opts) {
    const store = opts.store;
    const indexHtml = opts.indexHtml || PUBLIC_INDEX;

    return http.createServer(async (req, res) => {
        let url;
        try {
            url = new URL(req.url, 'http://127.0.0.1');
        } catch (_) {
            return sendError(res, 400, 'bad url');
        }
        const route = url.pathname;
        const method = req.method || 'GET';

        // CORS 头用 setHeader 提前挂上，后续所有 writeHead 自动合并；
        // OPTIONS 预检直接 204（PUT 带 X-Admin-Key、POST 带 JSON 都会触发预检）。
        const cors = corsHeaders(req);
        for (const k of Object.keys(cors)) res.setHeader(k, cors[k]);
        if (method === 'OPTIONS') {
            res.writeHead(204);
            return res.end();
        }

        try {
            if (method === 'GET' && (route === '/' || route === '/index.html')) {
                if (!fs.existsSync(indexHtml)) return sendError(res, 500, '缺少 public/index.html');
                res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' });
                return res.end(fs.readFileSync(indexHtml));
            }

            if (method === 'GET' && route === '/api/overview') {
                const rows = store.rows();
                const s = summarize(rows);
                return sendJson(res, 200, {
                    success: true,
                    stats: {
                        monthSales: s.monthSales,
                        totalSales: s.totalSales,
                        activeCount: s.activeCount,
                        pendingCount: s.pendingCount,
                        sentCount: s.sentCount,
                        refundedCount: s.refundedCount,
                        expiredCount: s.expiredCount,
                        byPlan: s.byPlan
                    },
                    expiringSoon: s.expiringSoon,
                    recent: s.recent,
                    pricing: store.pricing()
                });
            }

            if (method === 'GET' && route === '/api/keys') {
                const rows = store.search({
                    q: url.searchParams.get('q'),
                    plan: url.searchParams.get('plan'),
                    status: url.searchParams.get('status'),
                    sort: url.searchParams.get('sort')
                });
                return sendJson(res, 200, { success: true, rows });
            }

            if (method === 'POST' && route === '/api/keys') {
                const body = await readBody(req);
                if (!body.login || !String(body.login).trim()) {
                    return sendError(res, 400, '请填写客户 login（邮箱或标识）');
                }
                try {
                    const created = store.addKeys(body);
                    return sendJson(res, 200, { success: true, rows: created });
                } catch (e) {
                    // 起订量不足等业务拒绝按 400 回（OpsError 自带状态码）；
                    // 自检失败之类的真异常仍是 500，别把内部错误说成用户填错。
                    const isOps = e && e.name === 'OpsError' && Number.isFinite(e.status);
                    const status = isOps ? e.status : 500;
                    return sendError(res, status, (e && e.message) || '生成失败');
                }
            }

            if (method === 'POST' && route === '/api/keys/status') {
                const body = await readBody(req);
                if (!body.key) return sendError(res, 400, '缺少 key');
                // 先确认 key 在台账里，防止给不存在的 key 写状态
                const exists = store.rows().some(r => r.key === body.key);
                if (!exists) return sendError(res, 404, '台账中不存在该 key');
                const st = store.setStatus(body.key, body);
                return sendJson(res, 200, { success: true, state: st });
            }

            if (method === 'GET' && route === '/api/customers') {
                return sendJson(res, 200, { success: true, customers: store.customers() });
            }

            if (method === 'GET' && route === '/api/pricing') {
                return sendJson(res, 200, { success: true, pricing: store.pricing() });
            }

            if (method === 'GET' && route === '/api/customer-text') {
                const key = url.searchParams.get('key');
                if (!key) return sendError(res, 400, '缺少 key');
                const row = store.rows().find(r => r.key === key);
                if (!row) return sendError(res, 404, '台账中不存在该 key');
                const contact = process.env.OPS_CONTACT || '';
                return sendJson(res, 200, {
                    success: true,
                    text: customerMessage({ key: row.key, plan: row.plan, expiresAt: row.expiresAt, contact })
                });
            }

            if (method === 'GET' && route === '/api/export.csv') {
                const rows = store.search({ sort: url.searchParams.get('sort') });
                const head = 'issuedAt,key,login,plan,seats,expiresAt,status,amount,note,opsNote';
                const cell = (s) => {
                    const t = String(s == null ? '' : s);
                    return /[",\r\n]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t;
                };
                const lines = rows.map(r => [
                    r.issuedAt, r.key, r.login, r.plan, r.seats, r.expiresAt,
                    r.status, r.amount, r.note, r.opsNote
                ].map(cell).join(','));
                const csv = '\uFEFF' + head + '\r\n' + lines.join('\r\n') + '\r\n';
                res.writeHead(200, {
                    'Content-Type': 'text/csv; charset=utf-8',
                    'Content-Disposition': 'attachment; filename="ops-export.csv"'
                });
                return res.end(csv);
            }

            if (method === 'GET' && route === '/api/payment-config') {
                let cfg = {};
                try { if (fs.existsSync(LICENSE_CONFIG)) cfg = JSON.parse(fs.readFileSync(LICENSE_CONFIG, 'utf8')); } catch (_) {}
                return sendJson(res, 200, { success: true, config: maskPayment(cfg), file: LICENSE_CONFIG });
            }

            if (method === 'PUT' && route === '/api/payment-config') {
                const body = await readBody(req);
                const next = {};
                PAYMENT_KEYS.forEach((k) => {
                    if (body[k] !== undefined && body[k] !== null) next[k] = String(body[k]).trim();
                });
                let raw = {};
                try { if (fs.existsSync(LICENSE_CONFIG)) raw = JSON.parse(fs.readFileSync(LICENSE_CONFIG, 'utf8')); } catch (_) {}
                Object.assign(raw, next);
                Object.keys(raw).forEach((k) => { if (raw[k] === '' || raw[k] == null) delete raw[k]; });
                fs.writeFileSync(LICENSE_CONFIG, JSON.stringify(raw, null, 2) + '\n');

                // best-effort 热重载：若 LicenseServer 正在运行，推一次后台接口
                let reloaded = 'written';
                try {
                    const env = readEnvFile(LICENSE_ENV);
                    const adminKey = env.ADMIN_ACTIVATE_KEY || 'change-me-admin-key';
                    const rr = await httpJson('PUT', LICENSE_ADMIN_API + '/api/admin/payment-config',
                        raw, { 'X-Admin-Key': adminKey });
                    if (rr && rr.status === 200) reloaded = 'reloaded';
                } catch (_) { /* 未运行则下次启动生效 */ }
                return sendJson(res, 200, { success: true, reloaded: reloaded, config: maskPayment(next) });
            }

            // ---- 消 key 退款闭环 ----
            // 用户发来「消 key 码」→ 验签 → 定位台账 → 判断能否退款。只读，不改任何状态。
            if (method === 'POST' && route === '/api/deactivation/verify') {
                const body = await readBody(req);
                const v = verifyProof(body.code, store.rows());
                if (!v.ok) return sendJson(res, 200, { success: false, error: v.error });
                return sendJson(res, 200, {
                    success: true,
                    key: v.row.key,
                    login: v.row.login,
                    plan: v.row.plan,
                    expiresAt: v.row.expiresAt,
                    status: v.row.status,
                    destroyedAt: v.destroyedAt,
                    alreadyRefunded: v.alreadyRefunded,
                    // 吊销名单写在哪（供前端提示）
                    revokeTarget: path.join(LICENSE_SERVER_DIR, 'data', 'revoked.json')
                });
            }

            // 确认退款：标记 refunded + 写入吊销名单（换机备份的那份一并作废）
            if (method === 'POST' && route === '/api/deactivation/refund') {
                const body = await readBody(req);
                const r = refundByProof({
                    code: body.code,
                    store: store,
                    dataDir: path.join(LICENSE_SERVER_DIR, 'data')
                });
                if (!r.ok) return sendJson(res, 200, { success: false, error: r.error, row: r.row });
                return sendJson(res, 200, {
                    success: true,
                    key: r.row.key,
                    login: r.row.login,
                    state: r.state,
                    revoked: r.revoked,
                    destroyedAt: r.destroyedAt
                });
            }

            return sendError(res, 404, 'not found: ' + route);
        } catch (err) {
            return sendError(res, 500, err && err.message ? err.message : String(err));
        }
    });
}

function main() {
    const port = parseInt(process.env.OPS_PORT, 10) || 9600;
    const store = new OpsStore({
        ledgerPath: process.env.OPS_LEDGER || DEFAULT_LEDGER,
        statePath: process.env.OPS_STATE || DEFAULT_STATE,
        pricingPath: DEFAULT_PRICING
    });
    const server = createServer({ store });
    server.on('error', (err) => {
        if (err && err.code === 'EADDRINUSE') {
            console.error('端口 ' + port + ' 已被占用。换端口：OPS_PORT=9610 node OpsConsole/server.js');
            process.exit(1);
        }
        throw err;
    });
    server.listen(port, '127.0.0.1', () => {
        console.log('SmartPilot 运营台已启动: http://127.0.0.1:' + port);
        console.log('台账: ' + (process.env.OPS_LEDGER || DEFAULT_LEDGER));
        console.log('（仅监听本机；Ctrl+C 停止）');
    });
}

if (require.main === module) main();

module.exports = { createServer, main, DEFAULT_LEDGER, DEFAULT_STATE };
