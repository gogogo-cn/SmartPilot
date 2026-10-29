'use strict';
/**
 * 支付配置单一来源（single source of truth）。
 *
 * 读取顺序：process.env（来自 .env / 系统环境） → 合并 config.json（平台可写，优先）。
 * 这样"管理平台填上配置即可上线"无需改代码、无需手改 .env：
 *   - OpsConsole 的「支付配置」页直接写 config.json；
 *   - LicenseServer 启动时加载，PUT 后台接口可热重载（reload），不必重启。
 *
 * 测试可用 LICENSE_CONFIG_PATH 覆盖 config.json 路径，避免污染真实文件。
 */

const fs = require('fs');
const path = require('path');

const ROOT = __dirname.replace(/[\\/]lib$/, '');
const CONFIG_PATH = process.env.LICENSE_CONFIG_PATH || path.join(ROOT, 'config.json');

// 平台可管理的全部支付相关键
const KEYS = [
    'STRIPE_SECRET_KEY',
    'STRIPE_WEBHOOK_SECRET',
    'STRIPE_PRICE_PRO_MONTHLY',
    'STRIPE_PRICE_PRO_YEARLY',
    'STRIPE_SUCCESS_URL',
    'STRIPE_CANCEL_URL',
    // ---- 国内支付（微信 Native / 支付宝当面付），配置即上线 ----
    'CN_PAY_PROVIDER',        // simulate | wechat | alipay
    'CN_ICP_BEIAN',           // 工信部备案号（回调域名需已备案）
    'CN_PAY_SUBJECT',         // 商品名称，如「SmartPilot 智航 年付授权」
    'CN_NOTIFY_URL',          // 支付结果异步通知地址（公网 HTTPS）
    'CN_RETURN_URL',          // 支付完成回跳地址
    'CN_YEARLY_AMOUNT_FEN',   // 年付金额（分）
    'WECHAT_APP_ID',          // 公众号/小程序 AppID
    'WECHAT_MCH_ID',          // 微信商户号
    'WECHAT_API_V3_KEY',      // APIv3 密钥（32 字节，用于回调报文解密）
    'WECHAT_SERIAL_NO',       // 商户 API 证书序列号
    'WECHAT_PRIVATE_KEY',     // 商户 API 私钥（PEM）
    'WECHAT_PLATFORM_PUBLIC_KEY', // 微信支付平台证书公钥（PEM，用于验签回调）
    'WECHAT_API_KEY',         // 旧版 v2 API 密钥（兼容保留，不用于 v3 主链路）
    'ALIPAY_APP_ID',          // 支付宝应用 AppID
    'ALIPAY_PRIVATE_KEY',     // 应用私钥（PEM）
    'ALIPAY_PUBLIC_KEY',      // 支付宝公钥（PEM，用于验签回调）
    'ALIPAY_GATEWAY',         // 网关：生产 https://openapi.alipay.com/gateway.do 或沙箱
    'PUBLIC_BASE_URL'
];

// 写文件时的必填校验：缺少这些键认为配置未就绪（与 stripePay.ready 口径一致）
const SECRET_KEY = 'STRIPE_SECRET_KEY';
const PRICE_MONTHLY = 'STRIPE_PRICE_PRO_MONTHLY';
const PRICE_YEARLY = 'STRIPE_PRICE_PRO_YEARLY';

let cache = {};

function load() {
    const fromEnv = {};
    KEYS.forEach(function (k) {
        if (process.env[k] !== undefined) fromEnv[k] = process.env[k];
    });

    let fromFile = {};
    try {
        if (fs.existsSync(CONFIG_PATH)) {
            const raw = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));
            KEYS.forEach(function (k) {
                if (raw[k] !== undefined && raw[k] !== null) fromFile[k] = String(raw[k]);
            });
        }
    } catch (_) {
        // config.json 损坏：忽略，回落到 env
    }

    // 文件优先于 env
    cache = Object.assign({}, fromEnv, fromFile);
    return cache;
}

function get(key) {
    return cache[key];
}

function getAll() {
    return Object.assign({}, cache);
}

function reload() {
    return load();
}

/** 写回 config.json（只保留已知键），并热重载到内存。 */
function update(partial) {
    const next = {};
    KEYS.forEach(function (k) {
        if (partial[k] !== undefined && partial[k] !== null) next[k] = String(partial[k]);
    });

    let raw = {};
    try {
        if (fs.existsSync(CONFIG_PATH)) raw = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));
    } catch (_) {
        raw = {};
    }
    Object.assign(raw, next);

    // 删除空字符串键，保持文件整洁
    Object.keys(raw).forEach(function (k) {
        if (raw[k] === '' || raw[k] === undefined) delete raw[k];
    });

    fs.writeFileSync(CONFIG_PATH, JSON.stringify(raw, null, 2) + '\n');
    return reload();
}

// 脱敏：密钥/私钥类字段只显示首尾，避免配置接口泄露
function mask(value) {
    if (!value) return '';
    if (value.length <= 10) return '****' + value.slice(-2);
    return value.slice(0, 4) + '…' + value.slice(-4);
}

function summary() {
    const out = {};
    KEYS.forEach(function (k) {
        const v = cache[k] || '';
        if (/SECRET|KEY|PRIVATE/.test(k)) out[k] = mask(v);
        else out[k] = v;
    });
    return out;
}

load();

module.exports = {
    KEYS: KEYS,
    CONFIG_PATH: CONFIG_PATH,
    get: get,
    getAll: getAll,
    reload: reload,
    update: update,
    summary: summary,
    SECRET_KEY: SECRET_KEY,
    PRICE_MONTHLY: PRICE_MONTHLY,
    PRICE_YEARLY: PRICE_YEARLY
};
