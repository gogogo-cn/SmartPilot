'use strict';
const { test, before, after } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const http = require('http');

const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'lic-admin-'));
const cfgPath = path.join(tmpDir, 'config.json');
process.env.LICENSE_CONFIG_PATH = cfgPath;
process.env.DATA_DIR = fs.mkdtempSync(path.join(os.tmpdir(), 'lic-admin-data-'));
process.env.PORT = '0';
process.env.ADMIN_ACTIVATE_KEY = 'test-admin';

const { app } = require('../server');

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
    server = await new Promise((resolve) => { const s = app.listen(0, '127.0.0.1', () => resolve(s)); });
    base = 'http://127.0.0.1:' + server.address().port;
});
after(async () => {
    if (server) await new Promise((r) => server.close(r));
    try { fs.unlinkSync(cfgPath); } catch (_) {}
    try { fs.rmdirSync(tmpDir); } catch (_) {}
    try { fs.rmSync(process.env.DATA_DIR, { recursive: true, force: true }); } catch (_) {}
});

test('GET 无 admin key → 401', async () => {
    const r = await req('GET', '/api/admin/payment-config');
    assert.strictEqual(r.status, 401);
});

test('PUT 无 admin key → 401', async () => {
    const r = await req('PUT', '/api/admin/payment-config', { STRIPE_SECRET_KEY: 'sk_test_x' });
    assert.strictEqual(r.status, 401);
});

test('GET 有 key 返回脱敏配置 + stripeReady=false', async () => {
    const r = await req('GET', '/api/admin/payment-config', null, { 'X-Admin-Key': 'test-admin' });
    assert.strictEqual(r.status, 200);
    assert.strictEqual(r.json.success, true);
    assert.strictEqual(r.json.stripeReady, false);
    assert.ok('config' in r.json);
});

test('PUT 写完整配置 → 热重载 stripeReady=true，且写入 config.json', async () => {
    const payload = {
        STRIPE_SECRET_KEY: 'sk_test_abc123',
        STRIPE_PRICE_PRO_MONTHLY: 'price_1M9nthAb',
        STRIPE_PRICE_PRO_YEARLY: 'price_1Y9rAbCd'
    };
    const r = await req('PUT', '/api/admin/payment-config', payload, { 'X-Admin-Key': 'test-admin' });
    assert.strictEqual(r.status, 200);
    assert.strictEqual(r.json.stripeReady, true);
    assert.ok(r.json.config.STRIPE_SECRET_KEY.startsWith('sk_t'));

    const raw = JSON.parse(fs.readFileSync(cfgPath, 'utf8'));
    assert.strictEqual(raw.STRIPE_SECRET_KEY, 'sk_test_abc123');
});

test('PUT 后再 GET → 配置已持久化', async () => {
    const r = await req('GET', '/api/admin/payment-config', null, { 'X-Admin-Key': 'test-admin' });
    assert.strictEqual(r.json.config.STRIPE_PRICE_PRO_YEARLY, 'price_1Y9rAbCd');
});
