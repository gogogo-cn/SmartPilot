'use strict';
const { test, before, after } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');

// 用临时 config.json，避免污染真实文件
const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'lic-cfg-'));
const cfgPath = path.join(tmpDir, 'config.json');
process.env.LICENSE_CONFIG_PATH = cfgPath;

// 重新加载模块，使其指向临时文件
delete require.cache[require.resolve('../lib/paymentConfig')];
delete require.cache[require.resolve('../lib/stripePay')];
const paymentConfig = require('../lib/paymentConfig');
const stripePay = require('../lib/stripePay');

before(() => { process.env.STRIPE_SECRET_KEY = ''; });
after(() => { try { fs.unlinkSync(cfgPath); } catch (_) {} try { fs.rmdirSync(tmpDir); } catch (_) {} });

test('初始无配置：stripeReady=false', () => {
    paymentConfig.reload();
    assert.strictEqual(stripePay.ready(), false);
});

test('env 优先被 config.json 覆盖', () => {
    process.env.STRIPE_SECRET_KEY = 'sk_test_envonly';
    paymentConfig.reload();
    assert.strictEqual(paymentConfig.get('STRIPE_SECRET_KEY'), 'sk_test_envonly');
    // 写文件覆盖
    paymentConfig.update({ STRIPE_SECRET_KEY: 'sk_test_file' });
    assert.strictEqual(paymentConfig.get('STRIPE_SECRET_KEY'), 'sk_test_file');
});

test('put 完整配置后 stripeReady 变 true（无需重启）', () => {
    process.env.STRIPE_SECRET_KEY = '';
    paymentConfig.reload();
    paymentConfig.update({
        STRIPE_SECRET_KEY: 'sk_test_abc123',
        STRIPE_PRICE_PRO_MONTHLY: 'price_1MnthAb12',
        STRIPE_PRICE_PRO_YEARLY: 'price_1YrAb3400'
    });
    assert.strictEqual(stripePay.ready(), true);
    const sum = paymentConfig.summary();
    // 密钥脱敏
    assert.ok(sum.STRIPE_SECRET_KEY.startsWith('sk_t'));
    assert.ok(sum.STRIPE_SECRET_KEY.endsWith('123'));
    assert.strictEqual(sum.STRIPE_PRICE_PRO_MONTHLY, 'price_1MnthAb12');
});

test('非法 secret（无 sk_ 前缀）视为未就绪', () => {
    paymentConfig.update({ STRIPE_SECRET_KEY: 'garbage' });
    assert.strictEqual(stripePay.ready(), false);
});

test('占位 price（price_xxx）视为未就绪', () => {
    paymentConfig.update({
        STRIPE_SECRET_KEY: 'sk_test_abc123',
        STRIPE_PRICE_PRO_MONTHLY: 'price_xxx',
        STRIPE_PRICE_PRO_YEARLY: 'price_1YrAb3400'
    });
    assert.strictEqual(stripePay.ready(), false);
});

test('空字符串键会从文件清理', () => {
    paymentConfig.update({ STRIPE_SECRET_KEY: 'sk_test_abc123', STRIPE_PRICE_PRO_MONTHLY: 'price_m', STRIPE_PRICE_PRO_YEARLY: 'price_y' });
    paymentConfig.update({ STRIPE_WEBHOOK_SECRET: '' }); // 清空
    const raw = JSON.parse(fs.readFileSync(cfgPath, 'utf8'));
    assert.strictEqual(raw.STRIPE_WEBHOOK_SECRET, undefined);
    assert.strictEqual(raw.STRIPE_SECRET_KEY, 'sk_test_abc123');
});
