'use strict';
const { test, before, after } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const os = require('os');
const http = require('http');

const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'lic-api-'));
process.env.DATA_DIR = dataDir;
process.env.PORT = '0';
process.env.ADMIN_ACTIVATE_KEY = 'test-admin';
process.env.TRIAL_DAYS = '14';

const { app } = require('../server');

let server;
let base;

function req(method, urlPath, body, headers) {
    return new Promise((resolve, reject) => {
        const u = new URL(urlPath, base);
        const data = body ? JSON.stringify(body) : null;
        const r = http.request({
            hostname: u.hostname,
            port: u.port,
            path: u.pathname,
            method,
            headers: Object.assign({
                'Content-Type': 'application/json',
                'Content-Length': data ? Buffer.byteLength(data) : 0
            }, headers || {})
        }, (res) => {
            let buf = '';
            res.on('data', (c) => { buf += c; });
            res.on('end', () => {
                let json = {};
                try { json = JSON.parse(buf || '{}'); } catch (_) { json = { raw: buf }; }
                resolve({ status: res.statusCode, json });
            });
        });
        r.on('error', reject);
        if (data) r.write(data);
        r.end();
    });
}

before(async () => {
    server = await new Promise((resolve) => {
        const s = app.listen(0, '127.0.0.1', () => resolve(s));
    });
    const addr = server.address();
    base = `http://127.0.0.1:${addr.port}`;
});

after(async () => {
    await new Promise((resolve) => server.close(resolve));
});

test('register → entitlement active trial', async () => {
    const r = await req('POST', '/auth/register', {
        email: 'user@example.com',
        password: 'password12'
    });
    assert.strictEqual(r.status, 200);
    assert.ok(r.json.success);
    assert.ok(r.json.token);
    assert.strictEqual(r.json.entitlement.active, true);
    assert.strictEqual(r.json.entitlement.plan, 'trial');
});

test('login + me/entitlement', async () => {
    const login = await req('POST', '/auth/login', {
        login: 'user@example.com',
        password: 'password12'
    });
    assert.ok(login.json.token);
    const me = await req('GET', '/me/entitlement', null, {
        Authorization: 'Bearer ' + login.json.token
    });
    assert.strictEqual(me.json.success, true);
    assert.strictEqual(me.json.entitlement.active, true);
});

test('cn checkout stays closed unless simulate is explicitly allowed', async () => {
    const login = await req('POST', '/auth/login', {
        login: 'user@example.com',
        password: 'password12'
    });
    delete process.env.CN_PAY_ALLOW_SIMULATE;
    const blocked = await req('POST', '/checkout/cn', {}, {
        Authorization: 'Bearer ' + login.json.token
    });
    assert.strictEqual(blocked.status, 403);
    assert.strictEqual(blocked.json.code, 'checkout_not_ready');
    const health = await req('GET', '/health');
    assert.strictEqual(health.json.cnReady, false);
    assert.strictEqual(health.json.stripeReady, false);

    process.env.CN_PAY_ALLOW_SIMULATE = '1';
    const co = await req('POST', '/checkout/cn', {}, {
        Authorization: 'Bearer ' + login.json.token
    });
    assert.ok(co.json.url);
    assert.ok(co.json.orderId);
    const pay = await req('POST', '/checkout/cn/simulate-pay', { orderId: co.json.orderId });
    assert.strictEqual(pay.json.success, true);
    assert.strictEqual(pay.json.entitlement.plan, 'pro_yearly');
    assert.strictEqual(pay.json.entitlement.active, true);
    delete process.env.CN_PAY_ALLOW_SIMULATE;
});

test('admin activate', async () => {
    await req('POST', '/auth/register', {
        email: 'wire@example.com',
        password: 'password12'
    });
    const r = await req('POST', '/admin/activate', {
        login: 'wire@example.com',
        days: 30,
        plan: 'team',
        seats: 3
    }, { 'x-admin-key': 'test-admin' });
    assert.strictEqual(r.json.success, true);
    assert.strictEqual(r.json.entitlement.plan, 'team');
    assert.strictEqual(r.json.entitlement.seats, 3);
});
