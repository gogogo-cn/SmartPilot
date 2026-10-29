'use strict';
/**
 * 国内支付适配器（配置即上线框架）。
 *
 * 支持两个真实通道，均由 config 驱动，填好凭证即生效（cnPay.ready() 翻 true，
 * 插件购买按钮自动点亮）：
 *   - 微信支付 v3 Native（扫码支付）：下单拿 code_url → 前端渲染二维码。
 *   - 支付宝当面付 alipay.trade.precreate：拿 qr_code → 前端渲染二维码。
 *
 * 所有签名 / 验签 / 回调解密均用 Node 内置 crypto 实现，无第三方依赖。
 * 真实网络下单仅在 ready() 为真时执行；未配置或 provider=simulate 时回退到
 * 本地模拟页（开发用，allowSimulate 逃生口仍读实时 process.env）。
 *
 * ⚠️ 真实通道未在沙箱跑过（本机无商户凭证）。代码按官方文档实现，上线前请用
 *    微信/支付宝沙箱各跑一次下单 + 回调验签；门禁、签名、解密单元已测。
 */
const crypto = require('crypto');
const https = require('https');
const paymentConfig = require('./paymentConfig');

const WECHAT_GATEWAY = 'https://api.mch.weixin.qq.com';
const ALIPAY_GATEWAY_PROD = 'https://openapi.alipay.com/gateway.do';
const ALIPAY_GATEWAY_SANDBOX = 'https://openapi.alipaydev.com/gateway.do';

function provider() {
    return (paymentConfig.get('CN_PAY_PROVIDER') || 'simulate').toLowerCase();
}

/** 开发逃生口：仅读实时 process.env，不进托管配置缓存。生产必须不设置。 */
function allowSimulate() {
    return process.env.CN_PAY_ALLOW_SIMULATE === '1';
}

// 年付金额（分）。默认 19900 = 专业版年付 ¥199（2026-10-26 调价后口径，
// 与 OpsConsole/config.json 的 pro_yearly.amount=199 一致；原默认 29900 是旧价 ¥299，
// 曾出现"UI 卖 ¥199、国内支付实收 ¥299"的口径分裂）。可用 CN_YEARLY_AMOUNT_FEN 覆盖。
function yearlyAmountFen() {
    const n = parseInt(paymentConfig.get('CN_YEARLY_AMOUNT_FEN') || '19900', 10);
    return Number.isFinite(n) && n > 0 ? n : 19900;
}

function subject() {
    return paymentConfig.get('CN_PAY_SUBJECT') || 'SmartPilot 智航 年付授权';
}

// ---------------- 就绪判定 ----------------
function wechatConfigured() {
    return !!(
        paymentConfig.get('WECHAT_APP_ID') &&
        paymentConfig.get('WECHAT_MCH_ID') &&
        paymentConfig.get('WECHAT_API_V3_KEY') &&
        paymentConfig.get('WECHAT_SERIAL_NO') &&
        paymentConfig.get('WECHAT_PRIVATE_KEY') &&
        paymentConfig.get('CN_NOTIFY_URL')
    );
}

function alipayConfigured() {
    return !!(
        paymentConfig.get('ALIPAY_APP_ID') &&
        paymentConfig.get('ALIPAY_PRIVATE_KEY') &&
        paymentConfig.get('ALIPAY_PUBLIC_KEY') &&
        paymentConfig.get('CN_NOTIFY_URL')
    );
}

/**
 * 国内支付是否可上线：按通道校验必备凭证。
 * simulate 通道永远返回 false（不视为就绪，避免误开放）。
 */
function ready() {
    const p = provider();
    if (p === 'wechat') return wechatConfigured();
    if (p === 'alipay') return alipayConfigured();
    return false;
}

// ---------------- 通用签名工具 ----------------
function signRsaSha256(message, privateKeyPem) {
    const sign = crypto.createSign('RSA-SHA256');
    sign.update(message, 'utf8');
    return sign.sign(privateKeyPem, 'base64');
}

function verifyRsaSha256(message, signatureB64, publicKeyPem) {
    if (!publicKeyPem || !signatureB64) return false;
    const v = crypto.createVerify('RSA-SHA256');
    v.update(message, 'utf8');
    try {
        return v.verify(publicKeyPem, Buffer.from(signatureB64, 'base64'));
    } catch (_) {
        return false;
    }
}

// ---------------- 微信支付 v3 ----------------
/**
 * 构造微信 v3 请求签名头。
 * message = method\n+urlPath+\n+timestamp+\n+nonce+\n+body\n
 */
function buildWechatAuth(method, urlPath, bodyStr, opts) {
    const timestamp = Math.floor(Date.now() / 1000).toString();
    const nonce = crypto.randomBytes(16).toString('hex');
    const message = method + '\n' + urlPath + '\n' + timestamp + '\n' + nonce + '\n' + bodyStr + '\n';
    const signature = signRsaSha256(message, opts.privateKey);
    const header =
        'WECHATPAY2-SHA256-RSA2048 ' +
        'mchid="' + opts.mchId + '",' +
        'nonce_str="' + nonce + '",' +
        'signature="' + signature + '",' +
        'timestamp="' + timestamp + '",' +
        'serial_no="' + opts.serialNo + '"';
    return { header, timestamp, nonce, signature };
}

/** 验签微信回调：message = timestamp\n+nonce+\n+rawBody\n，用平台证书公钥。 */
function verifyWechatNotify(headers, rawBody, opts) {
    const publicKey = opts && opts.platformPublicKey;
    if (!publicKey) return false;
    const message =
        (headers.timestamp || '') + '\n' + (headers.nonce || '') + '\n' + rawBody + '\n';
    return verifyRsaSha256(message, headers.signature, publicKey);
}

/** 用 APIv3 密钥 AES-256-GCM 解密回调 resource，返回明文对象。 */
function decryptWechatResource(resource, apiV3Key) {
    if (!resource || !resource.ciphertext) throw new Error('missing ciphertext');
    const keyBuf = apiV3Key.length === 64 ? Buffer.from(apiV3Key, 'hex') : Buffer.from(apiV3Key, 'utf8');
    if (keyBuf.length !== 32) throw new Error('WECHAT_API_V3_KEY must be 32 bytes');
    const buf = Buffer.from(resource.ciphertext, 'base64');
    const authTag = buf.subarray(buf.length - 16);
    const data = buf.subarray(0, buf.length - 16);
    const decipher = crypto.createDecipheriv('aes-256-gcm', keyBuf, Buffer.from(resource.nonce || '', 'utf8'));
    decipher.setAuthTag(authTag);
    if (resource.associated_data) decipher.setAAD(Buffer.from(resource.associated_data, 'utf8'));
    const decrypted = Buffer.concat([decipher.update(data), decipher.final()]);
    return JSON.parse(decrypted.toString('utf8'));
}

// ---------------- 支付宝当面付 ----------------
function alipayGateway() {
    const g = paymentConfig.get('ALIPAY_GATEWAY');
    if (g) return g;
    return process.env.ALIPAY_SANDBOX === '1' ? ALIPAY_GATEWAY_SANDBOX : ALIPAY_GATEWAY_PROD;
}

/** 支付宝签名串：除 sign/sign_type 外，按 key 字典序拼 k=v&（空值跳过）。 */
function alipaySignContent(params) {
    return Object.keys(params)
        .filter((k) => k !== 'sign' && k !== 'sign_type' && params[k] !== undefined && params[k] !== '')
        .sort()
        .map((k) => k + '=' + params[k])
        .join('&');
}

function signAlipay(params, privateKeyPem) {
    return signRsaSha256(alipaySignContent(params), privateKeyPem);
}

function verifyAlipay(params, signature, alipayPublicKeyPem) {
    return verifyRsaSha256(alipaySignContent(params), signature, alipayPublicKeyPem);
}

// ---------------- 真实下单（gated） ----------------
function httpsRequest(opts, bodyStr) {
    return new Promise((resolve, reject) => {
        let u;
        try { u = new URL(opts.url); } catch (e) { return reject(e); }
        const data = bodyStr || '';
        const req = https.request(
            {
                hostname: u.hostname,
                port: u.port || 443,
                path: u.pathname + u.search,
                method: opts.method || 'POST',
                headers: Object.assign(
                    { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(data) },
                    opts.headers || {}
                ),
                timeout: 8000
            },
            (res) => {
                let b = '';
                res.on('data', (c) => { b += c; });
                res.on('end', () => {
                    let json = {};
                    try { json = JSON.parse(b); } catch (_) { json = { raw: b }; }
                    resolve({ status: res.statusCode, json, raw: b });
                });
            }
        );
        req.on('error', reject);
        req.setTimeout(8000, () => { req.destroy(new Error('request timeout')); });
        if (data) req.write(data);
        req.end();
    });
}

/**
 * 创建国内支付订单，返回 { provider, amountFen, orderId, codeUrl, checkoutUrl }。
 * 真实通道（wechat/alipay 且凭证齐全）会向支付网关发起请求拿二维码串；
 * 未配置或 simulate 时回退到本地模拟页 URL（不发起网络请求）。
 */
async function createOrder({ orderId, userId, amountFen, publicBaseUrl, description }) {
    const p = provider();
    const amount = amountFen || yearlyAmountFen();
    const desc = description || subject();
    const base = (publicBaseUrl || paymentConfig.get('PUBLIC_BASE_URL') || 'http://localhost:9530').replace(/\/$/, '');
    const notifyUrl = paymentConfig.get('CN_NOTIFY_URL') || base + '/webhooks/cn';
    const fallbackUrl =
        base + '/checkout/cn/pay.html?orderId=' + encodeURIComponent(orderId) +
        '&provider=simulate&amount=' + amount + '&userId=' + encodeURIComponent(userId || '');

    if (p === 'wechat' && wechatConfigured()) {
        const urlPath = '/v3/pay/transactions/native';
        const body = {
            mchid: paymentConfig.get('WECHAT_MCH_ID'),
            appid: paymentConfig.get('WECHAT_APP_ID'),
            description: desc,
            out_trade_no: orderId,
            notify_url: notifyUrl,
            amount: { total: amount, currency: 'CNY' }
        };
        const bodyStr = JSON.stringify(body);
        const auth = buildWechatAuth('POST', urlPath, bodyStr, {
            mchId: paymentConfig.get('WECHAT_MCH_ID'),
            serialNo: paymentConfig.get('WECHAT_SERIAL_NO'),
            privateKey: paymentConfig.get('WECHAT_PRIVATE_KEY')
        });
        const res = await httpsRequest({
            url: WECHAT_GATEWAY + urlPath,
            method: 'POST',
            headers: {
                Authorization: auth.header,
                Accept: 'application/json',
                'User-Agent': 'SmartPilot-License/1.0'
            }
        }, bodyStr);
        if (res.status === 200 && res.json && res.json.code_url) {
            return {
                provider: 'wechat',
                amountFen: amount,
                orderId,
                codeUrl: res.json.code_url,
                checkoutUrl: base + '/checkout/cn/pay.html?orderId=' + encodeURIComponent(orderId)
            };
        }
        throw new Error('WeChat order failed: ' + (res.json && (res.json.message || res.json.raw)) || res.status);
    }

    if (p === 'alipay' && alipayConfigured()) {
        const biz = {
            out_trade_no: orderId,
            total_amount: (amount / 100).toFixed(2),
            subject: desc
        };
        const params = {
            app_id: paymentConfig.get('ALIPAY_APP_ID'),
            method: 'alipay.trade.precreate',
            format: 'JSON',
            charset: 'utf-8',
            sign_type: 'RSA2',
            timestamp: new Date().toISOString().replace('T', ' ').slice(0, 19),
            version: '1.0',
            notify_url: notifyUrl,
            biz_content: JSON.stringify(biz)
        };
        params.sign = signAlipay(params, paymentConfig.get('ALIPAY_PRIVATE_KEY'));
        const query = new URLSearchParams(params).toString();
        const res = await httpsRequest({ url: alipayGateway() + '?' + query, method: 'GET' }, '');
        const resp = res.json && res.json.alipay_trade_precreate_response;
        if (resp && resp.code === '10000' && resp.qr_code) {
            return {
                provider: 'alipay',
                amountFen: amount,
                orderId,
                codeUrl: resp.qr_code,
                checkoutUrl: base + '/checkout/cn/pay.html?orderId=' + encodeURIComponent(orderId)
            };
        }
        throw new Error('Alipay order failed: ' + (resp && (resp.sub_msg || resp.code)) || res.raw || res.status);
    }

    // simulate / 未配置：本地模拟页（开发用）
    return {
        provider: 'simulate',
        amountFen: amount,
        orderId,
        codeUrl: '',
        checkoutUrl: fallbackUrl
    };
}

/**
 * 旧的兼容别名（保留以避免外部引用断裂）。新代码请用 createOrder。
 */
function createAnnualCheckout(opts) {
    return createOrder(opts);
}

module.exports = {
    provider,
    ready,
    allowSimulate,
    yearlyAmountFen,
    subject,
    wechatConfigured,
    alipayConfigured,
    // 微信 v3
    buildWechatAuth,
    verifyWechatNotify,
    decryptWechatResource,
    // 支付宝
    signAlipay,
    verifyAlipay,
    alipaySignContent,
    // 下单
    createOrder,
    createAnnualCheckout
};
