'use strict';
// 国内支付框架测试：门禁、微信 v3 签名/验签/解密、支付宝 RSA2 签名/验签、simulate 兜底。
// 不发起真实网络请求；只覆盖可在本机验证的纯逻辑 + 配置门禁。
const { test, before, after } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const crypto = require('crypto');

// 用独立临时配置，避免污染真实 config.json
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'lic-cnpay-'));
process.env.LICENSE_CONFIG_PATH = path.join(tmp, 'config.json');
const cnPay = require('../lib/cnPay');
const paymentConfig = require('../lib/paymentConfig');

function setCfg(obj) {
    paymentConfig.update(obj);
}

after(() => {
    try { fs.rmSync(tmp, { recursive: true, force: true }); } catch (_) {}
});

// 生成测试用 RSA 密钥对
function genRsa() {
    const kp = crypto.generateKeyPairSync('rsa', { modulusLength: 2048 });
    return {
        private: kp.privateKey.export({ type: 'pkcs8', format: 'pem' }),
        public: kp.publicKey.export({ type: 'spki', format: 'pem' })
    };
}

test('默认 simulate 通道 ready()=false', () => {
    setCfg({ CN_PAY_PROVIDER: '' });
    assert.strictEqual(cnPay.provider(), 'simulate');
    assert.strictEqual(cnPay.ready(), false);
});

test('wechat 凭证齐全 ready()=true，缺任一则 false', () => {
    setCfg({
        CN_PAY_PROVIDER: 'wechat',
        WECHAT_APP_ID: 'wx123',
        WECHAT_MCH_ID: 'mch123',
        WECHAT_API_V3_KEY: 'a'.repeat(32),
        WECHAT_SERIAL_NO: 'serial123',
        WECHAT_PRIVATE_KEY: '----PRIV----',
        CN_NOTIFY_URL: 'https://x.cn/webhooks/cn'
    });
    assert.strictEqual(cnPay.ready(), true);

    // 缺 notify url → false
    setCfg({ CN_NOTIFY_URL: '' });
    assert.strictEqual(cnPay.ready(), false);

    // 恢复
    setCfg({ CN_NOTIFY_URL: 'https://x.cn/webhooks/cn' });
    assert.strictEqual(cnPay.ready(), true);
});

test('alipay 凭证齐全 ready()=true，缺任一则 false', () => {
    setCfg({
        CN_PAY_PROVIDER: 'alipay',
        ALIPAY_APP_ID: 'appid123',
        ALIPAY_PRIVATE_KEY: '----PRIV----',
        ALIPAY_PUBLIC_KEY: '----PUB----',
        CN_NOTIFY_URL: 'https://x.cn/webhooks/cn'
    });
    assert.strictEqual(cnPay.ready(), true);

    setCfg({ ALIPAY_PUBLIC_KEY: '' });
    assert.strictEqual(cnPay.ready(), false);

    setCfg({ ALIPAY_PUBLIC_KEY: '----PUB----' });
    assert.strictEqual(cnPay.ready(), true);
});

test('微信 v3 请求签名头格式正确（含 serial）', () => {
    const kp = genRsa();
    const auth = cnPay.buildWechatAuth('POST', '/v3/pay/transactions/native', '{"a":1}', {
        mchId: 'mch123',
        serialNo: 'serial123',
        privateKey: kp.private
    });
    assert.ok(/^WECHATPAY2-SHA256-RSA2048 /.test(auth.header));
    assert.ok(auth.header.includes('serial_no="serial123"'));
});

test('微信 v3 回调验签（timestamp+nonce+body 串，官方格式）', () => {
    const kp = genRsa();
    const timestamp = '1700000000';
    const nonce = 'abc123';
    const rawBody = '{"id":"evt1"}';
    // 微信回调验签名串 = timestamp\nnonce\nbody\n（与请求签名不同，不含 method/url）
    const message = timestamp + '\n' + nonce + '\n' + rawBody + '\n';
    const signature = crypto.createSign('RSA-SHA256').update(message, 'utf8').sign(kp.private, 'base64');
    const headers = { signature, timestamp, nonce };
    assert.strictEqual(cnPay.verifyWechatNotify(headers, rawBody, { platformPublicKey: kp.public }), true);
    // 篡改报文 → 验签失败
    assert.strictEqual(cnPay.verifyWechatNotify(headers, '{"id":"evt2"}', { platformPublicKey: kp.public }), false);
});

test('微信回调解密 AES-256-GCM 往返', () => {
    const key = 'a'.repeat(32); // 32 字节 APIv3 密钥（ASCII）
    const iv = crypto.randomBytes(12).toString('utf8');
    const aad = 'transaction';
    const plain = JSON.stringify({ out_trade_no: 'O1', transaction_id: 'T1', trade_state: 'SUCCESS' });

    const cipher = crypto.createCipheriv('aes-256-gcm', Buffer.from(key, 'utf8'), Buffer.from(iv, 'utf8'));
    if (aad) cipher.setAAD(Buffer.from(aad, 'utf8'));
    const enc = Buffer.concat([cipher.update(plain, 'utf8'), cipher.final()]);
    const tag = cipher.getAuthTag();
    const resource = {
        ciphertext: Buffer.concat([enc, tag]).toString('base64'),
        nonce: iv,
        associated_data: aad
    };

    const decrypted = cnPay.decryptWechatResource(resource, key);
    assert.strictEqual(decrypted.out_trade_no, 'O1');
    assert.strictEqual(decrypted.transaction_id, 'T1');
    assert.strictEqual(decrypted.trade_state, 'SUCCESS');
});

test('支付宝 RSA2 签名可被支付宝公钥验签（往返）', () => {
    const kp = genRsa();
    const params = {
        app_id: 'appid123',
        method: 'alipay.trade.precreate',
        charset: 'utf-8',
        sign_type: 'RSA2',
        timestamp: '2026-10-20 12:00:00',
        version: '1.0',
        biz_content: JSON.stringify({ out_trade_no: 'O1', total_amount: '299.00' })
    };
    params.sign = cnPay.signAlipay(params, kp.private);
    assert.ok(params.sign);
    assert.strictEqual(cnPay.verifyAlipay(params, params.sign, kp.public), true);

    // 篡改金额 → 验签失败
    const tampered = Object.assign({}, params, { biz_content: JSON.stringify({ out_trade_no: 'O1', total_amount: '1.00' }) });
    assert.strictEqual(cnPay.verifyAlipay(tampered, params.sign, kp.public), false);
});

test('未配置凭证时 createOrder 走 simulate 兜底（不发起网络）', async () => {
    setCfg({ CN_PAY_PROVIDER: '', WECHAT_MCH_ID: '', ALIPAY_APP_ID: '' });
    const r = await cnPay.createOrder({ orderId: 'ORD-TEST-1', userId: 'u1' });
    assert.strictEqual(r.provider, 'simulate');
    assert.ok(/orderId=ORD-TEST-1/.test(r.checkoutUrl));
    assert.strictEqual(r.codeUrl, '');
});

test('备案号等国内配置键在 paymentConfig.KEYS 中', () => {
    const keys = paymentConfig.KEYS;
    ['CN_ICP_BEIAN', 'WECHAT_API_V3_KEY', 'WECHAT_SERIAL_NO', 'WECHAT_PRIVATE_KEY',
     'WECHAT_PLATFORM_PUBLIC_KEY', 'ALIPAY_GATEWAY', 'CN_PAY_SUBJECT'
    ].forEach((k) => assert.ok(keys.includes(k), 'missing key: ' + k));
});
