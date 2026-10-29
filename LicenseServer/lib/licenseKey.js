'use strict';
/**
 * licenseKey.js — 离线激活码算法（路线①：无授权服务器也能卖）。
 *
 * key 格式：  SPK1-<base64url(payload)>.<base64url(sig)>
 * payload：   v1|login|plan|seats|expUnixSeconds|nonce
 * sig：       HMAC-SHA256(payload, secret) 的 32 字节摘要，base64url
 * nonce 是 16 位随机 hex——没有它，同参数同秒批量生成的 key 会完全相同
 * （2026-10-20 实测踩过：一批 3 个 key 全等，等于所有客户共用一张）。
 *
 * 本文件是算法参考实现，与以下两端逐字节一致（改算法必须三处同步改）：
 *   - Bridges/LicenseKey.vb            （Excel 侧验签）
 *   - Hosts/Shared/HostLicense.vb      （Word/PPT 侧验签）
 *
 * 威胁模型说明：secret 内置于客户端，纯本地验签是「防君子不防小人」的轻度
 * 防护，用于挡住普通白嫖；远程吊销、换机绑定等强校验走 LicenseServer 在线
 * 链路（LicenseBridge / HostLicense 的 HTTP 方法）。生产环境应通过
 * LICENSE_KEY_SECRET 覆盖默认 secret，并同步修改两端常量后重新编译发布。
 */
const crypto = require('crypto');

const PREFIX = 'SPK1-';
// 注销码前缀（Deactivation Proof）：用户在本机销毁激活码后生成的退款凭证
const DEACT_PREFIX = 'SPD1-';
const VERSION = 'v1';
const DEFAULT_SECRET = 'SmartPilot-Offline-License-Key-v1-CHANGE-ME';
const FIELD_COUNT = 6;
const DEACT_FIELD_COUNT = 5;

function secret() {
    return process.env.LICENSE_KEY_SECRET || DEFAULT_SECRET;
}

function base64UrlEncode(buf) {
    return buf.toString('base64')
        .replace(/\+/g, '-')
        .replace(/\//g, '_')
        .replace(/=+$/, '');
}

function base64UrlDecode(text) {
    let s = String(text || '').replace(/-/g, '+').replace(/_/g, '/');
    switch (s.length % 4) {
        case 2: s += '=='; break;
        case 3: s += '='; break;
        default: break;
    }
    return Buffer.from(s, 'base64');
}

function sign(payloadText) {
    return base64UrlEncode(
        crypto.createHmac('sha256', secret()).update(payloadText, 'utf8').digest()
    );
}

function buildPayload({ login, plan, seats, expUnix, nonce }) {
    return [
        VERSION,
        login || '-',
        plan || 'pro_yearly',
        String(seats || 1),
        String(expUnix),
        nonce || crypto.randomBytes(8).toString('hex')
    ].join('|');
}

/**
 * 生成一个激活码。每次调用都带随机 nonce，同参数也绝不会重复。
 * @param {{login?:string, plan?:string, seats?:number, days?:number, expUnix?:number}} opts
 * @returns {string} SPK1-... 形式的激活码
 */
function createKey(opts) {
    opts = opts || {};
    const exp = opts.expUnix != null
        ? Math.floor(opts.expUnix)
        : Math.floor(Date.now() / 1000) + Math.floor(opts.days || 365) * 86400;
    const payload = buildPayload({
        login: opts.login,
        plan: opts.plan,
        seats: opts.seats,
        expUnix: exp
    });
    return PREFIX + base64UrlEncode(Buffer.from(payload, 'utf8')) + '.' + sign(payload);
}

function invalid(reason) {
    return { valid: false, reason, login: '', plan: '', seats: 0, periodEnd: null };
}

/**
 * 校验激活码。任何格式错误返回 { valid:false, reason }，不抛异常。
 * @returns {{valid:boolean, reason:string, login:string, plan:string, seats:number, periodEnd:string|null}}
 */
function verify(key) {
    const raw = String(key == null ? '' : key).trim();
    if (!raw.startsWith(PREFIX)) return invalid('malformed');
    const body = raw.slice(PREFIX.length);
    const dot = body.lastIndexOf('.');
    if (dot <= 0 || dot === body.length - 1) return invalid('malformed');

    let payloadText;
    try {
        payloadText = base64UrlDecode(body.slice(0, dot)).toString('utf8');
    } catch (_) {
        return invalid('malformed');
    }

    const sig = body.slice(dot + 1);
    const expected = sign(payloadText);
    if (sig.length !== expected.length) return invalid('bad_signature');
    if (!crypto.timingSafeEqual(Buffer.from(sig, 'utf8'), Buffer.from(expected, 'utf8'))) {
        return invalid('bad_signature');
    }

    const parts = payloadText.split('|');
    if (parts.length !== FIELD_COUNT || parts[0] !== VERSION) return invalid('bad_payload');
    const exp = Number(parts[4]);
    if (!Number.isFinite(exp) || exp <= 0) return invalid('bad_payload');
    if (exp <= Math.floor(Date.now() / 1000)) return invalid('expired');

    const seats = parseInt(parts[3], 10);
    return {
        valid: true,
        reason: 'ok',
        login: parts[1] === '-' ? '' : parts[1],
        plan: parts[2],
        seats: Number.isFinite(seats) && seats >= 1 ? seats : 1,
        periodEnd: new Date(exp * 1000).toISOString()
    };
}

/**
 * 激活码 → 不可逆标识（SHA256 前 16 位 hex）。
 * 与 lib/revocation.js 的 hashKey 必须逐字节一致（有守栏测试）。
 * 注销码里只带这个标识，绝不携带激活码原文——注销码是用户要发给客服的，
 * 一旦泄露也不能反推出可用激活码。
 */
function hashKey(key) {
    return crypto.createHash('sha256')
        .update(String(key == null ? '' : key), 'utf8')
        .digest('hex')
        .slice(0, 16);
}

/**
 * 生成「消 key 码」（退款凭证）。
 *
 * 业务闭环：用户在本机点「注销并申请退款」→ 本地激活被清除且该 key 在本机
 * 永久失效 → 生成此码 → 用户把它发给客服 → 运营台验签通过后才退款。
 * 这样「退款」与「不能再用」是同一个动作的两面，不需要强制联网。
 *
 * 格式：SPD1-<base64url(payload)>.<base64url(sig)>
 * payload：v1|keyHash|plan|destroyedUnix|nonce
 *
 * @param {{key:string, plan?:string, tsUnix?:number, nonce?:string}} opts
 * @returns {string}
 */
function createDeactivationProof(opts) {
    opts = opts || {};
    const ts = opts.tsUnix != null
        ? Math.floor(opts.tsUnix)
        : Math.floor(Date.now() / 1000);
    const payload = [
        VERSION,
        hashKey(opts.key),
        opts.plan || 'pro_yearly',
        String(ts),
        opts.nonce || crypto.randomBytes(8).toString('hex')
    ].join('|');
    return DEACT_PREFIX + base64UrlEncode(Buffer.from(payload, 'utf8')) + '.' + sign(payload);
}

/**
 * 校验「消 key 码」。任何格式错误返回 { valid:false, reason }，不抛异常。
 * @returns {{valid:boolean, reason:string, keyHash:string, plan:string, destroyedAt:string|null, nonce:string}}
 */
function verifyDeactivationProof(code) {
    const empty = { valid: false, reason: 'malformed', keyHash: '', plan: '', destroyedAt: null, nonce: '' };
    const raw = String(code == null ? '' : code).trim();
    if (!raw.startsWith(DEACT_PREFIX)) return empty;
    const body = raw.slice(DEACT_PREFIX.length);
    const dot = body.lastIndexOf('.');
    if (dot <= 0 || dot === body.length - 1) return empty;

    let payloadText;
    try {
        payloadText = base64UrlDecode(body.slice(0, dot)).toString('utf8');
    } catch (_) {
        return empty;
    }

    const sig = body.slice(dot + 1);
    const expected = sign(payloadText);
    if (sig.length !== expected.length) {
        return Object.assign({}, empty, { reason: 'bad_signature' });
    }
    if (!crypto.timingSafeEqual(Buffer.from(sig, 'utf8'), Buffer.from(expected, 'utf8'))) {
        return Object.assign({}, empty, { reason: 'bad_signature' });
    }

    const parts = payloadText.split('|');
    if (parts.length !== DEACT_FIELD_COUNT || parts[0] !== VERSION) {
        return Object.assign({}, empty, { reason: 'bad_payload' });
    }
    const ts = Number(parts[3]);
    if (!Number.isFinite(ts) || ts <= 0) {
        return Object.assign({}, empty, { reason: 'bad_payload' });
    }
    return {
        valid: true,
        reason: 'ok',
        keyHash: parts[1],
        plan: parts[2],
        destroyedAt: new Date(ts * 1000).toISOString(),
        nonce: parts[4]
    };
}

module.exports = {
    PREFIX,
    DEACT_PREFIX,
    VERSION,
    DEFAULT_SECRET,
    createKey,
    verify,
    buildPayload,
    sign,
    base64UrlEncode,
    base64UrlDecode,
    hashKey,
    createDeactivationProof,
    verifyDeactivationProof
};
