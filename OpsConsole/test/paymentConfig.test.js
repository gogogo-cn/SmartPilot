'use strict';
const { test, before, after } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const http = require('http');

const ROOT = path.join(__dirname, '..', '..');
const tmpLS = fs.mkdtempSync(path.join(os.tmpdir(), 'ops-ls-'));
const fakeConfig = path.join(tmpLS, 'config.json');
fs.writeFileSync(fakeConfig, JSON.stringify({ STRIPE_SECRET_KEY: 'sk_old_masked' }) + '\n');
fs.writeFileSync(path.join(tmpLS, '.env'), 'ADMIN_ACTIVATE_KEY=test-admin\nPORT=9530\n');

// 让 OpsConsole 指向这个假 LicenseServer 目录
process.env.LICENSE_SERVER_DIR = tmpLS;
delete require.cache[require.resolve('../server')];
delete require.cache[require.resolve('../lib/store')];
const { createServer } = require('../server');

let server, base;

function req(method, urlPath, body, headers) {
    return new Promise((resolve, reject) => {
        const u = new URL(urlPath, base);
        const data = body ? JSON.stringify(body) : null;
        const r = http.request({
            hostname: u.hostname, port: u.port, path: u.pathname, method,
            headers: Object.assign({ 'Content-Type': 'application/json' }, headers || {})
        }, (res) => {
            let buf = '';
            res.on('data', (c) => { buf += c; });
            res.on('end', () => { try { resolve({ status: res.statusCode, json: JSON.parse(buf || '{}') }); } catch (_) { resolve({ status: res.statusCode, json: {} }); } });
        });
        r.on('error', reject);
        if (data) r.write(data);
        r.end();
    });
}

before(async () => {
    server = await new Promise((resolve) => { const s = createServer({ store: { rows: () => [], search: () => [], customers: () => [], pricing: () => ({}) } }); s.listen(0, '127.0.0.1', () => resolve(s)); });
    base = 'http://127.0.0.1:' + server.address().port;
});
after(async () => {
    if (server) await new Promise((r) => server.close(r));
    try { fs.rmSync(tmpLS, { recursive: true, force: true }); } catch (_) {}
});

test('GET 支付配置：返回脱敏配置', async () => {
    const r = await req('GET', '/api/payment-config');
    assert.strictEqual(r.status, 200);
    assert.strictEqual(r.json.success, true);
    assert.ok(r.json.file.endsWith('config.json'));
    assert.strictEqual(r.json.config.STRIPE_SECRET_KEY, 'sk_o…sked');
});

test('PUT 支付配置：写入 LicenseServer/config.json（空 LicenseServer 时 written）', async () => {
    const payload = {
        STRIPE_SECRET_KEY: 'sk_test_newkey',
        STRIPE_PRICE_PRO_MONTHLY: 'price_m_9',
        STRIPE_PRICE_PRO_YEARLY: 'price_y_9'
    };
    const r = await req('PUT', '/api/payment-config', payload);
    assert.strictEqual(r.status, 200);
    assert.strictEqual(r.json.success, true);
    // 空 LicenseServer：reloaded=written（热重载调用失败，但文件已写）
    assert.strictEqual(r.json.reloaded, 'written');

    const raw = JSON.parse(fs.readFileSync(fakeConfig, 'utf8'));
    assert.strictEqual(raw.STRIPE_SECRET_KEY, 'sk_test_newkey');
    assert.strictEqual(raw.STRIPE_PRICE_PRO_MONTHLY, 'price_m_9');
});

test('PUT 不接受非白名单键', async () => {
    const before = fs.readFileSync(fakeConfig, 'utf8');
    await req('PUT', '/api/payment-config', { TOTALLY_RANDOM: 'x', STRIPE_SECRET_KEY: 'sk_test_keep' });
    const raw = JSON.parse(fs.readFileSync(fakeConfig, 'utf8'));
    assert.strictEqual(raw.TOTALLY_RANDOM, undefined);
    assert.strictEqual(raw.STRIPE_SECRET_KEY, 'sk_test_keep');
});
