'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { Store, hashPassword, verifyPassword, entitlementView } = require('../lib/store');

test('password hash verify', () => {
    const h = hashPassword('secret-pass');
    assert.ok(verifyPassword('secret-pass', h));
    assert.ok(!verifyPassword('wrong', h));
});

test('trial entitlement is active until periodEnd', () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'lic-'));
    const store = new Store(dir);
    const user = store.createUser({ email: 'a@b.com', passwordHash: hashPassword('password1') });
    const end = new Date(Date.now() + 7 * 86400000).toISOString();
    store.upsertEntitlement(user.id, { plan: 'trial', seats: 1, status: 'active', periodEnd: end });
    const v = entitlementView(store.getEntitlement(user.id));
    assert.strictEqual(v.active, true);
    assert.strictEqual(v.plan, 'trial');
});

test('expired entitlement is not active', () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'lic-'));
    const store = new Store(dir);
    const user = store.createUser({ email: 'c@d.com', passwordHash: hashPassword('password1') });
    const end = new Date(Date.now() - 86400000).toISOString();
    store.upsertEntitlement(user.id, { plan: 'pro_yearly', seats: 1, status: 'active', periodEnd: end });
    const v = entitlementView(store.getEntitlement(user.id));
    assert.strictEqual(v.active, false);
    assert.strictEqual(v.reason, 'expired');
});
