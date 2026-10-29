'use strict';
/**
 * SmartPilot License API — seat subscription only (BYOK).
 * Does not meter LLM tokens; model keys never leave the user's PC.
 */
const path = require('path');
const express = require('express');
const {
    Store,
    loadEnvFile,
    hashPassword,
    verifyPassword,
    entitlementView
} = require('./lib/store');
const stripePay = require('./lib/stripePay');
const cnPay = require('./lib/cnPay');
const paymentConfig = require('./lib/paymentConfig');
const { Revocation, hashKey, isHash } = require('./lib/revocation');

const ROOT = __dirname;
loadEnvFile(ROOT);

const PORT = parseInt(process.env.PORT || '9530', 10);
const TRIAL_DAYS = parseInt(process.env.TRIAL_DAYS || '14', 10);
const SESSION_DAYS = parseInt(process.env.SESSION_DAYS || '30', 10);
const PUBLIC_BASE = (process.env.PUBLIC_BASE_URL || `http://localhost:${PORT}`).replace(/\/$/, '');
const DATA_DIR = process.env.DATA_DIR
    ? path.resolve(ROOT, process.env.DATA_DIR)
    : path.join(ROOT, 'data');

const store = new Store(DATA_DIR);
// 吊销名单：退款兜底。用户在本机「消 key」只让本机失效，若他把激活码备份到
// 别处换机激活，仍能用——故退款后必须同时进吊销名单，由联网客户端拉走。
const revocation = new Revocation(DATA_DIR);
const app = express();

// Stripe needs raw body for webhook signature
app.post('/webhooks/stripe', express.raw({ type: 'application/json' }), async (req, res) => {
    try {
        const sig = req.headers['stripe-signature'];
        const event = stripePay.constructWebhookEvent(req.body, sig);
        await handleStripeEvent(event);
        res.json({ received: true });
    } catch (e) {
        console.error('[stripe webhook]', e.message);
        res.status(400).send(`Webhook Error: ${e.message}`);
    }
});

// WebView2 / 任务窗格可能从虚拟主机发起探测；允许跨域读 health（认证仍走 .NET 桥）
app.use((req, res, next) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Session-Token');
    if (req.method === 'OPTIONS') return res.sendStatus(204);
    next();
});

// 支付/Stripe 回调需要原始字节做签名校验，不能先被 JSON 解析吞掉
app.use(express.json({ limit: '256kb', type: (req) => !/^\/webhooks\//.test(req.path) }));
app.use(express.urlencoded({ extended: false }));
app.use('/checkout', express.static(path.join(ROOT, 'public')));

function ok(res, data) {
    res.json(Object.assign({ success: true }, data));
}
function fail(res, status, message, code) {
    res.status(status).json({ success: false, error: message, code: code || 'error' });
}

function authUser(req) {
    const h = req.headers.authorization || '';
    const m = /^Bearer\s+(\S+)/i.exec(h);
    const token = m ? m[1] : (req.headers['x-session-token'] || '');
    const session = store.findSession(token);
    if (!session) return null;
    const user = store.findUserById(session.userId);
    if (!user) return null;
    return { user, session, token };
}

function startTrial(userId) {
    const periodEnd = new Date(Date.now() + TRIAL_DAYS * 86400000).toISOString();
    return store.upsertEntitlement(userId, {
        plan: 'trial',
        seats: 1,
        status: 'active',
        periodEnd
    });
}

function activatePaid(userId, plan, periodEnd, seats) {
    return store.upsertEntitlement(userId, {
        plan: plan || 'pro_yearly',
        seats: seats || 1,
        status: 'active',
        periodEnd
    });
}

function publicUser(user) {
    return {
        id: user.id,
        email: user.email,
        phone: user.phone
    };
}

app.get('/health', (_req, res) => {
    ok(res, {
        service: 'smartpilot-license',
        trialDays: TRIAL_DAYS,
        stripeConfigured: !!paymentConfig.get(paymentConfig.SECRET_KEY),
        stripeReady: stripePay.ready(),
        cnProvider: cnPay.provider(),
        cnReady: cnPay.ready(),
        revokedCount: revocation.count(),
        revokedVersion: revocation.list().version
    });
});

app.post('/auth/register', (req, res) => {
    const email = (req.body && req.body.email) || '';
    const phone = (req.body && req.body.phone) || '';
    const password = (req.body && req.body.password) || '';
    if (!password || String(password).length < 8) {
        return fail(res, 400, 'password must be at least 8 characters', 'weak_password');
    }
    if (!email && !phone) {
        return fail(res, 400, 'email or phone required', 'missing_login');
    }
    const login = email || phone;
    if (store.findUserByLogin(login)) {
        return fail(res, 409, 'account already exists', 'exists');
    }
    const user = store.createUser({
        email: email || null,
        phone: phone || null,
        passwordHash: hashPassword(password)
    });
    startTrial(user.id);
    const session = store.createSession(user.id, SESSION_DAYS);
    const ent = entitlementView(store.getEntitlement(user.id));
    ok(res, { user: publicUser(user), token: session.token, entitlement: ent });
});

app.post('/auth/login', (req, res) => {
    const login = (req.body && (req.body.login || req.body.email || req.body.phone)) || '';
    const password = (req.body && req.body.password) || '';
    const user = store.findUserByLogin(login);
    if (!user || !verifyPassword(password, user.passwordHash)) {
        return fail(res, 401, 'invalid credentials', 'auth_failed');
    }
    const session = store.createSession(user.id, SESSION_DAYS);
    let ent = store.getEntitlement(user.id);
    if (!ent) ent = startTrial(user.id);
    ok(res, {
        user: publicUser(user),
        token: session.token,
        entitlement: entitlementView(ent)
    });
});

app.post('/auth/logout', (req, res) => {
    const a = authUser(req);
    if (a) store.revokeSessionsForUser(a.user.id);
    ok(res, {});
});

app.get('/me/entitlement', (req, res) => {
    const a = authUser(req);
    if (!a) return fail(res, 401, 'unauthorized', 'unauthorized');
    let ent = store.getEntitlement(a.user.id);
    if (!ent) ent = startTrial(a.user.id);
    ok(res, {
        user: publicUser(a.user),
        entitlement: entitlementView(ent)
    });
});

app.post('/checkout/stripe', async (req, res) => {
    const a = authUser(req);
    if (!a) return fail(res, 401, 'unauthorized', 'unauthorized');
    const plan = (req.body && req.body.plan) === 'pro_yearly' ? 'pro_yearly' : 'pro_monthly';
    try {
        const session = await stripePay.createCheckoutSession({
            user: a.user,
            plan,
            successUrl: req.body && req.body.successUrl,
            cancelUrl: req.body && req.body.cancelUrl
        });
        store.createSubscription({
            userId: a.user.id,
            provider: 'stripe',
            providerSubId: null,
            checkoutSessionId: session.id,
            plan,
            status: 'checkout_pending'
        });
        ok(res, { url: session.url, sessionId: session.id, plan });
    } catch (e) {
        fail(res, 400, e.message, e.code || 'stripe_error');
    }
});

app.post('/checkout/stripe/portal', async (req, res) => {
    const a = authUser(req);
    if (!a) return fail(res, 401, 'unauthorized', 'unauthorized');
    const subs = store._db.subscriptions
        .filter(s => s.userId === a.user.id && s.provider === 'stripe' && s.stripeCustomerId)
        .sort((x, y) => String(y.updatedAt).localeCompare(String(x.updatedAt)));
    const customerId = subs[0] && subs[0].stripeCustomerId;
    try {
        const portal = await stripePay.createPortalSession({
            customerId,
            returnUrl: req.body && req.body.returnUrl
        });
        ok(res, { url: portal.url });
    } catch (e) {
        fail(res, 400, e.message, e.code || 'stripe_error');
    }
});

app.post('/checkout/cn', async (req, res) => {
    const a = authUser(req);
    if (!a) return fail(res, 401, 'unauthorized', 'unauthorized');
    if (!cnPay.ready() && !cnPay.allowSimulate()) {
        return fail(res, 403, 'China checkout is not available yet', 'checkout_not_ready');
    }
    const order = store.createOrder({
        userId: a.user.id,
        plan: 'pro_yearly',
        provider: cnPay.provider(),
        amountFen: cnPay.yearlyAmountFen(),
        currency: 'CNY'
    });
    try {
        const checkout = await cnPay.createOrder({
            orderId: order.id,
            userId: a.user.id,
            publicBaseUrl: PUBLIC_BASE
        });
        store.updateOrder(order.id, {
            provider: checkout.provider,
            checkoutUrl: checkout.checkoutUrl,
            codeUrl: checkout.codeUrl || ''
        });
        ok(res, {
            orderId: order.id,
            url: checkout.checkoutUrl,
            codeUrl: checkout.codeUrl || '',
            provider: checkout.provider,
            amountFen: checkout.amountFen,
            plan: 'pro_yearly'
        });
    } catch (e) {
        // 真实下单失败（网关/凭证问题）：回退到本地模拟页，避免白屏
        const fallback = await cnPay.createOrder({
            orderId: order.id,
            userId: a.user.id,
            publicBaseUrl: PUBLIC_BASE
        });
        store.updateOrder(order.id, { provider: fallback.provider, checkoutUrl: fallback.checkoutUrl });
        ok(res, {
            orderId: order.id,
            url: fallback.checkoutUrl,
            codeUrl: '',
            provider: fallback.provider,
            amountFen: fallback.amountFen,
            plan: 'pro_yearly',
            warning: e.message
        });
    }
});

// 前端轮询：拿到真实通道的二维码串（code_url / qr_code）
app.get('/checkout/cn/status', (req, res) => {
    const orderId = req.query && req.query.orderId;
    const order = orderId && store.findOrder(orderId);
    if (!order) return fail(res, 404, 'order not found', 'not_found');
    ok(res, {
        orderId: order.id,
        provider: order.provider,
        codeUrl: order.codeUrl || '',
        status: order.status,
        paid: order.status === 'paid'
    });
});

/** Local/dev: mark CN order paid (also used by simulate pay page). */
app.post('/checkout/cn/simulate-pay', (req, res) => {
    if (!cnPay.allowSimulate()) {
        return fail(res, 403, 'Simulate pay is disabled', 'checkout_not_ready');
    }
    const orderId = req.body && req.body.orderId;
    const order = store.findOrder(orderId);
    if (!order) return fail(res, 404, 'order not found', 'not_found');
    if (order.status === 'paid') {
        return ok(res, { orderId, alreadyPaid: true, entitlement: entitlementView(store.getEntitlement(order.userId)) });
    }
    const periodEnd = new Date(Date.now() + 365 * 86400000).toISOString();
    activatePaid(order.userId, 'pro_yearly', periodEnd, 1);
    store.updateOrder(orderId, { status: 'paid', paidAt: new Date().toISOString() });
    store.createSubscription({
        userId: order.userId,
        provider: order.provider || 'simulate',
        providerSubId: order.id,
        plan: 'pro_yearly',
        status: 'active',
        periodEnd
    });
    ok(res, { orderId, entitlement: entitlementView(store.getEntitlement(order.userId)) });
});

app.post('/webhooks/cn', express.raw({ type: () => true }), (req, res) => {
    const p = cnPay.provider();
    const raw = Buffer.isBuffer(req.body)
        ? req.body.toString('utf8')
        : (typeof req.body === 'string' ? req.body : JSON.stringify(req.body || {}));

    function grant(orderId, provider, providerSubId, payload) {
        const order = store.findOrder(orderId);
        if (!order) return fail(res, 404, 'order not found', 'not_found');
        if (order.status !== 'paid') {
            const periodEnd = new Date(Date.now() + 365 * 86400000).toISOString();
            activatePaid(order.userId, 'pro_yearly', periodEnd, 1);
            store.updateOrder(orderId, { status: 'paid', paidAt: new Date().toISOString(), notifyPayload: payload });
            store.createSubscription({
                userId: order.userId,
                provider: provider || order.provider || 'cn',
                providerSubId: String(providerSubId || orderId),
                plan: 'pro_yearly',
                status: 'active',
                periodEnd
            });
        }
        return ok(res, { orderId });
    }

    try {
        if (p === 'wechat') {
            const headers = {
                signature: req.headers['wechatpay-signature'],
                timestamp: req.headers['wechatpay-timestamp'],
                nonce: req.headers['wechatpay-nonce'],
                serial: req.headers['wechatpay-serial']
            };
            if (!cnPay.verifyWechatNotify(headers, raw, {
                platformPublicKey: paymentConfig.get('WECHAT_PLATFORM_PUBLIC_KEY')
            })) {
                return fail(res, 401, 'bad signature', 'bad_signature');
            }
            const evt = JSON.parse(raw);
            if (evt.event_type && evt.event_type !== 'TRANSACTION.SUCCESS') {
                return ok(res, { received: true, ignored: evt.event_type });
            }
            const decrypted = cnPay.decryptWechatResource(evt.resource, paymentConfig.get('WECHAT_API_V3_KEY'));
            if (decrypted.trade_state !== 'SUCCESS') {
                return ok(res, { received: true, trade_state: decrypted.trade_state });
            }
            return grant(decrypted.out_trade_no, 'wechat', decrypted.transaction_id, decrypted);
        }

        if (p === 'alipay') {
            const params = Object.fromEntries(new URLSearchParams(raw));
            if (!cnPay.verifyAlipay(params, params.sign, paymentConfig.get('ALIPAY_PUBLIC_KEY'))) {
                return fail(res, 401, 'bad signature', 'bad_signature');
            }
            const tradeStatus = params.trade_status;
            if (tradeStatus !== 'TRADE_SUCCESS' && tradeStatus !== 'TRADE_FINISHED') {
                return ok(res, { received: true, trade_status: tradeStatus });
            }
            return grant(params.out_trade_no, 'alipay', params.trade_no, params);
        }

        // simulate / 未配置：开发逃生口
        if (cnPay.allowSimulate()) return ok(res, { received: true, simulate: true });
        return fail(res, 403, 'cn not ready', 'checkout_not_ready');
    } catch (e) {
        return fail(res, 400, e.message, 'bad_request');
    }
});

/** Manual activation for support / offline wire transfer */
app.post('/admin/activate', (req, res) => {
    const key = req.headers['x-admin-key'] || (req.body && req.body.adminKey);
    if (!process.env.ADMIN_ACTIVATE_KEY || key !== process.env.ADMIN_ACTIVATE_KEY) {
        return fail(res, 403, 'forbidden', 'forbidden');
    }
    const login = (req.body && (req.body.login || req.body.email || req.body.phone)) || '';
    const days = parseInt((req.body && req.body.days) || '365', 10);
    const plan = (req.body && req.body.plan) || 'pro_yearly';
    const seats = parseInt((req.body && req.body.seats) || '1', 10);
    const user = store.findUserByLogin(login);
    if (!user) return fail(res, 404, 'user not found', 'not_found');
    const periodEnd = new Date(Date.now() + days * 86400000).toISOString();
    activatePaid(user.id, plan, periodEnd, seats);
    ok(res, { user: publicUser(user), entitlement: entitlementView(store.getEntitlement(user.id)) });
});

async function handleStripeEvent(event) {
    switch (event.type) {
        case 'checkout.session.completed': {
            const session = event.data.object;
            const userId = (session.metadata && session.metadata.userId) || session.client_reference_id;
            if (!userId) break;
            const plan = (session.metadata && session.metadata.plan) || 'pro_monthly';
            const days = plan === 'pro_yearly' ? 365 : 31;
            const periodEnd = new Date(Date.now() + days * 86400000).toISOString();
            activatePaid(userId, plan, periodEnd, 1);
            store.createSubscription({
                userId,
                provider: 'stripe',
                providerSubId: session.subscription || session.id,
                stripeCustomerId: session.customer || null,
                checkoutSessionId: session.id,
                plan,
                status: 'active',
                periodEnd
            });
            break;
        }
        case 'invoice.paid': {
            const invoice = event.data.object;
            const subId = invoice.subscription;
            if (!subId) break;
            let sub = store.findSubscriptionByProviderId('stripe', subId);
            const userId = sub && sub.userId;
            const periodEnd = invoice.lines && invoice.lines.data && invoice.lines.data[0] &&
                invoice.lines.data[0].period && invoice.lines.data[0].period.end
                ? new Date(invoice.lines.data[0].period.end * 1000).toISOString()
                : new Date(Date.now() + 31 * 86400000).toISOString();
            if (userId) {
                const plan = (sub && sub.plan) || 'pro_monthly';
                activatePaid(userId, plan, periodEnd, 1);
                if (sub) store.updateSubscription(sub.id, { status: 'active', periodEnd, stripeCustomerId: invoice.customer || sub.stripeCustomerId });
            }
            break;
        }
        case 'customer.subscription.deleted': {
            const subscription = event.data.object;
            const sub = store.findSubscriptionByProviderId('stripe', subscription.id);
            if (sub) {
                store.updateSubscription(sub.id, { status: 'canceled' });
                store.upsertEntitlement(sub.userId, { status: 'canceled' });
            }
            break;
        }
        default:
            break;
    }
}

app.get('/checkout/success', (_req, res) => {
    res.type('html').send('<!doctype html><meta charset=utf-8><title>Payment OK</title><p>Payment received. Return to SmartPilot and click Refresh subscription.</p>');
});
app.get('/checkout/cancel', (_req, res) => {
    res.type('html').send('<!doctype html><meta charset=utf-8><title>Cancelled</title><p>Checkout cancelled. You can close this window.</p>');
});

// ---- 支付配置管理 API（管理平台填配置即上线）----
// 受 ADMIN_ACTIVATE_KEY 保护（请求头 X-Admin-Key）。GET 返回脱敏配置 + 就绪状态；
// PUT 写入 config.json 并热重载到内存，无需重启 LicenseServer。
const ADMIN_KEY = process.env.ADMIN_ACTIVATE_KEY || 'change-me-admin-key';

function adminGuard(req, res, next) {
    const k = req.headers['x-admin-key'] || (req.query && req.query.adminKey);
    if (k !== ADMIN_KEY) return fail(res, 401, 'admin key required', 'unauthorized');
    next();
}

app.get('/api/admin/payment-config', adminGuard, (_req, res) => {
    ok(res, {
        config: paymentConfig.summary(),
        stripeReady: stripePay.ready(),
        cnReady: cnPay.ready()
    });
});

app.put('/api/admin/payment-config', adminGuard, (req, res) => {
    const partial = (req.body && req.body.config) || req.body || {};
    const cleaned = {};
    paymentConfig.KEYS.forEach(function (k) {
        if (partial[k] !== undefined && partial[k] !== null) cleaned[k] = String(partial[k]);
    });
    paymentConfig.update(cleaned);
    ok(res, {
        config: paymentConfig.summary(),
        stripeReady: stripePay.ready(),
        cnReady: cnPay.ready()
    });
});

// ---- 激活码吊销名单 ----
// 名单只含 SHA256(key) 前 16 位，不含任何激活码原文，因此可以公开分发：
// 客户端启动时拉一次，本地比对即可判断自己的激活码是否已被吊销。
app.get('/api/revoked', (_req, res) => {
    ok(res, revocation.list());
});

// 单个激活码的吊销状态查询（客户端心跳用，只传标识不传原文）
app.get('/api/license/verify', (req, res) => {
    const h = String((req.query && req.query.h) || '').trim();
    if (!isHash(h)) return fail(res, 400, 'bad hash', 'bad_hash');
    ok(res, { revoked: revocation.hasHash(h), version: revocation.list().version });
});

app.post('/api/admin/revoke', adminGuard, (req, res) => {
    const raw = String((req.body && (req.body.key || req.body.hash)) || '').trim();
    if (!raw) return fail(res, 400, 'key or hash required', 'bad_request');
    const r = revocation.add(raw);
    ok(res, {
        added: r.added,
        hash: r.hash,
        version: r.version,
        count: revocation.count()
    });
});

app.post('/api/admin/unrevoke', adminGuard, (req, res) => {
    const raw = String((req.body && (req.body.hash || req.body.key)) || '').trim();
    if (!raw) return fail(res, 400, 'key or hash required', 'bad_request');
    const r = revocation.remove(isHash(raw) ? raw : hashKey(raw));
    ok(res, {
        removed: r.removed,
        version: r.version,
        count: revocation.count()
    });
});

if (require.main === module) {
    const BIND_HOST = process.env.BIND_HOST || '';
    const onListen = () => {
        const where = BIND_HOST ? `${BIND_HOST}:${PORT}` : String(PORT);
        console.log(`[LicenseServer] listening on ${where}  data=${DATA_DIR}`);
        console.log(`[LicenseServer] trial=${TRIAL_DAYS}d  stripe=${!!paymentConfig.get(paymentConfig.SECRET_KEY)}  cn=${cnPay.provider()}`);
    };
    if (BIND_HOST) app.listen(PORT, BIND_HOST, onListen);
    else app.listen(PORT, onListen);
}

module.exports = { app, store, entitlementView, startTrial, activatePaid };
