'use strict';
/**
 * 消 key 退款闭环测试（运营台侧）。
 *
 * 核心要守住的三件事：
 *   1. 伪造/篡改的注销码一律拒绝 —— 不能让人不销 key 就骗到退款
 *   2. 确认退款必须同时写吊销名单 —— 否则他在别处备份的那份还能用
 *   3. 同一张激活码不能退两次
 */
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { OpsStore } = require('../lib/store');
const { verifyProof, refundByProof, hashKey } = require('../lib/deactivation');
const { Revocation } = require('../../LicenseServer/lib/revocation.js');
const lk = require('../../LicenseServer/lib/licenseKey.js');

function mk() {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ops-deact-'));
    const store = new OpsStore({
        ledgerPath: path.join(dir, 'issued-keys.csv'),
        statePath: path.join(dir, 'ops-state.json')
    });
    return { store, dir };
}

/** 造一张已发货的激活码，返回 { key, code, store, dir } */
function issued(opts) {
    const ctx = mk();
    const rows = ctx.store.addKeys(Object.assign({
        login: 'a@b.com', plan: 'pro_yearly', days: 365, count: 1
    }, opts || {}));
    const key = rows[0].key;
    return Object.assign(ctx, {
        key: key,
        code: lk.createDeactivationProof({ key: key, plan: 'pro_yearly' })
    });
}

test('有效注销码 → 验签通过并定位到台账客户', () => {
    const c = issued();
    const v = verifyProof(c.code, c.store.rows());
    assert.strictEqual(v.ok, true, '应验签通过');
    assert.strictEqual(v.row.key, c.key, '应定位到同一张激活码');
    assert.strictEqual(v.row.login, 'a@b.com');
    assert.strictEqual(v.alreadyRefunded, false);
    assert.ok(v.destroyedAt, '应带本机销毁时间');
});

test('篡改注销码 → 拒绝（不销 key 骗不到退款）', () => {
    const c = issued();
    const body = c.code.slice(lk.DEACT_PREFIX.length);
    const dot = body.lastIndexOf('.');
    const payloadText = lk.base64UrlDecode(body.slice(0, dot)).toString('utf8');
    const forged = lk.DEACT_PREFIX
        + lk.base64UrlEncode(Buffer.from(payloadText.replace('pro_yearly', 'pro_monthly'), 'utf8'))
        + '.' + body.slice(dot + 1);

    const v = verifyProof(forged, c.store.rows());
    assert.strictEqual(v.ok, false);
    assert.ok(/篡改|无效/.test(v.error), '应提示不要退款，实际：' + v.error);
});

test('乱码 / 空值 → 拒绝', () => {
    const c = issued();
    ['', 'abc', 'SPK1-xxx.yyy', null].forEach((bad) => {
        const v = verifyProof(bad, c.store.rows());
        assert.strictEqual(v.ok, false, String(bad) + ' 应被拒绝');
    });
});

test('非本系统签发的激活码 → 台账查不到 → 拒绝', () => {
    const c = issued();
    // 伪造者用自己的 secret 签一个合法格式的注销码，但 key 不在我们台账里
    const foreignKey = 'SPK1-foreign-key-not-in-ledger';
    const foreign = lk.createDeactivationProof({ key: foreignKey, plan: 'pro_yearly' });
    const v = verifyProof(foreign, c.store.rows());
    assert.strictEqual(v.ok, false);
    assert.ok(/查不到/.test(v.error), '应提示台账查不到，实际：' + v.error);
});

test('确认退款 → 标记 refunded 且写入吊销名单', () => {
    const c = issued();
    const dataDir = path.join(c.dir, 'lsdata');
    const r = refundByProof({ code: c.code, store: c.store, dataDir: dataDir });

    assert.strictEqual(r.ok, true);
    assert.strictEqual(r.state.status, 'refunded', '运营状态应为已退款');

    const rev = new Revocation(dataDir);
    assert.strictEqual(rev.hasKey(c.key), true, '该激活码必须进吊销名单（换机备份也要作废）');
    assert.strictEqual(rev.count(), 1);
    // 名单里只能是标识，绝不能是激活码原文
    const list = rev.list();
    assert.ok(!list.revoked.some(h => h.indexOf(c.key) >= 0), '名单不得含激活码原文');
    assert.strictEqual(list.revoked[0], hashKey(c.key));
});

test('重复退款 → 拒绝（同一张激活码不能退两次）', () => {
    const c = issued();
    const dataDir = path.join(c.dir, 'lsdata');
    assert.strictEqual(refundByProof({ code: c.code, store: c.store, dataDir }).ok, true);

    const second = refundByProof({ code: c.code, store: c.store, dataDir });
    assert.strictEqual(second.ok, false);
    assert.ok(/重复/.test(second.error), '应提示重复退款，实际：' + second.error);
});

test('仅验证不退款时，不得改动任何状态', () => {
    const c = issued();
    const v = verifyProof(c.code, c.store.rows());
    assert.strictEqual(v.ok, true);
    const after = c.store.rows().find(r => r.key === c.key);
    assert.notStrictEqual(after.status, 'refunded', 'verify 是只读的，不能改状态');
});
