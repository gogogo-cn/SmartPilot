'use strict';
/**
 * license-key.test.js — 离线激活码算法（JS 参考实现）回归
 * 运行：cd LicenseServer && node --test test/*.test.js
 */
const { test } = require('node:test');
const assert = require('node:assert');
const path = require('path');
const fs = require('fs');
const { execFileSync } = require('child_process');

const lk = require(path.join(__dirname, '..', 'lib', 'licenseKey.js'));

const KEY_RE = /^SPK1-[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$/;

test('生成的 key 格式合法且能通过验签', () => {
    const key = lk.createKey({ login: 'user@example.com', plan: 'pro_yearly', seats: 1, days: 365 });
    assert.match(key, KEY_RE);
    const v = lk.verify(key);
    assert.strictEqual(v.valid, true);
    assert.strictEqual(v.reason, 'ok');
    assert.strictEqual(v.login, 'user@example.com');
    assert.strictEqual(v.plan, 'pro_yearly');
    assert.strictEqual(v.seats, 1);
    assert.ok(v.periodEnd && new Date(v.periodEnd).getTime() > Date.now());
});

test('payload 是 v1|login|plan|seats|expUnix|nonce 六段，nonce 为 16 位 hex', () => {
    const exp = Math.floor(Date.now() / 1000) + 86400;
    const payload = lk.buildPayload({ login: 'a@b.c', plan: 'team', seats: 5, expUnix: exp });
    const parts = payload.split('|');
    assert.strictEqual(parts.length, 6);
    assert.strictEqual(parts[0], 'v1');
    assert.strictEqual(parts[1], 'a@b.c');
    assert.strictEqual(parts[2], 'team');
    assert.strictEqual(parts[3], '5');
    assert.strictEqual(parts[4], String(exp));
    assert.match(parts[5], /^[0-9a-f]{16}$/, 'nonce 应为 16 位随机 hex');
});

test('显式传入 nonce 时 payload 确定可复现', () => {
    const exp = 1798761600;
    const payload = lk.buildPayload({ login: 'a@b.c', plan: 'team', seats: 5, expUnix: exp, nonce: 'abc123' });
    assert.strictEqual(payload, 'v1|a@b.c|team|5|1798761600|abc123');
});

test('回归：同参数批量生成的 key 必须互不相同（nonce）', () => {
    // 2026-10-20 踩过的坑：payload 无随机成分时，--batch 3 同参数同秒
    // 生成的 key 完全一样，等于所有客户共用一张、无法区分归属。
    const seen = new Set();
    for (let i = 0; i < 50; i++) {
        const key = lk.createKey({ login: 'batch@example.com', plan: 'pro_yearly', seats: 1, days: 365 });
        assert.strictEqual(lk.verify(key).valid, true, '第 ' + i + ' 个 key 应有效');
        seen.add(key);
    }
    assert.strictEqual(seen.size, 50, '50 个同参数 key 应全部不同');
});

test('缺 nonce 的旧五段 payload -> bad_payload（三端口径一致）', () => {
    const exp = Math.floor(Date.now() / 1000) + 86400;
    const legacy = 'v1|legacy@example.com|pro_yearly|1|' + exp;
    const legacyKey = lk.PREFIX + lk.base64UrlEncode(Buffer.from(legacy, 'utf8')) + '.' + lk.sign(legacy);
    const v = lk.verify(legacyKey);
    assert.strictEqual(v.valid, false);
    assert.strictEqual(v.reason, 'bad_payload', '五段旧格式必须被拒，防止 VB 端与 JS 端口径漂移');
});

test('无 login 时用 - 占位，verify 回空串', () => {
    const key = lk.createKey({ plan: 'pro_monthly', days: 31 });
    const v = lk.verify(key);
    assert.strictEqual(v.valid, true);
    assert.strictEqual(v.login, '');
});

test('篡改 payload 任意字符 -> bad_signature', () => {
    const key = lk.createKey({ login: 'u@x.com', plan: 'pro_yearly', seats: 1, days: 365 });
    const dot = key.lastIndexOf('.');
    const payloadB64 = key.slice(lk.PREFIX.length, dot);
    // 改 payload 第一个 base64url 字符（保持字符集合法）
    const flipped = (payloadB64[0] === 'A' ? 'B' : 'A') + payloadB64.slice(1);
    const v = lk.verify(lk.PREFIX + flipped + '.' + key.slice(dot + 1));
    assert.strictEqual(v.valid, false);
    assert.strictEqual(v.reason, 'bad_signature');
});

test('篡改签名 -> bad_signature', () => {
    const key = lk.createKey({ plan: 'pro_yearly', days: 365 });
    const dot = key.lastIndexOf('.');
    const sig = key.slice(dot + 1);
    const flippedSig = (sig[0] === 'A' ? 'B' : 'A') + sig.slice(1);
    const v = lk.verify(key.slice(0, dot + 1) + flippedSig);
    assert.strictEqual(v.valid, false);
    assert.strictEqual(v.reason, 'bad_signature');
});

test('过期 key -> expired', () => {
    const key = lk.createKey({ plan: 'pro_yearly', expUnix: Math.floor(Date.now() / 1000) - 10 });
    const v = lk.verify(key);
    assert.strictEqual(v.valid, false);
    assert.strictEqual(v.reason, 'expired');
});

test('非 SPK1 前缀 / 空 / 垃圾输入 -> malformed 且不抛异常', () => {
    for (const bad of ['', '   ', 'XXXX-abc.def', 'SPK1-', 'SPK1-abc', 'SPK1-abc.', 'SPK1-.abc']) {
        const v = lk.verify(bad);
        assert.strictEqual(v.valid, false, '应当拒绝: ' + JSON.stringify(bad));
    }
});

test('LICENSE_KEY_SECRET 覆盖后旧 key 失效、新 key 有效', () => {
    const old = process.env.LICENSE_KEY_SECRET;
    try {
        delete process.env.LICENSE_KEY_SECRET;
        const defaultKey = lk.createKey({ plan: 'pro_yearly', days: 365 });
        process.env.LICENSE_KEY_SECRET = 'another-secret-entirely';
        assert.strictEqual(lk.verify(defaultKey).reason, 'bad_signature');
        const rotated = lk.createKey({ plan: 'pro_yearly', days: 365 });
        assert.strictEqual(lk.verify(rotated).valid, true);
    } finally {
        if (old === undefined) delete process.env.LICENSE_KEY_SECRET;
        else process.env.LICENSE_KEY_SECRET = old;
    }
});

test('base64url 编解码往返（含 + / 字符的载荷）', () => {
    const raw = Buffer.from('v1|a+b/c=d|pro_yearly|1|1798761600', 'utf8');
    const enc = lk.base64UrlEncode(raw);
    assert.ok(!/[+/=]/.test(enc), 'base64url 不应出现 + / =');
    assert.deepStrictEqual(lk.base64UrlDecode(enc), raw);
});

test('VB 两端常量与 JS 默认 secret 一致（防三处漂移）', () => {
    const root = path.join(__dirname, '..', '..');
    const jsHas = fs.readFileSync(path.join(__dirname, '..', 'lib', 'licenseKey.js'), 'utf8')
        .includes("DEFAULT_SECRET = '" + lk.DEFAULT_SECRET + "'");
    assert.ok(jsHas, 'lib/licenseKey.js 应暴露 DEFAULT_SECRET 常量');
    const excelVb = fs.readFileSync(path.join(root, 'Bridges', 'LicenseKey.vb'), 'utf8');
    const hostVb = fs.readFileSync(path.join(root, 'Hosts', 'Shared', 'HostLicense.vb'), 'utf8');
    assert.ok(excelVb.includes('Private Const SECRET As String = "' + lk.DEFAULT_SECRET + '"'),
        'Bridges/LicenseKey.vb 的 SECRET 必须与 JS 默认值一致');
    assert.ok(hostVb.includes('Private Const KEY_SECRET As String = "' + lk.DEFAULT_SECRET + '"'),
        'Hosts/Shared/HostLicense.vb 的 KEY_SECRET 必须与 JS 默认值一致');
});

test('CLI 生成器：stdout 出一个可验签的 key', () => {
    const out = execFileSync(process.execPath, [
        path.join(__dirname, '..', 'tools', 'gen-license-key.js'),
        '--login', 'cli@example.com', '--plan', 'pro_yearly', '--days', '365'
    ], { encoding: 'utf8' });
    const key = out.split(/\r?\n/)[0].trim();
    assert.match(key, KEY_RE);
    assert.strictEqual(lk.verify(key).valid, true);
});

test('CLI --verify 对有效 key 返回 0，对无效 key 返回 1', () => {
    const good = lk.createKey({ plan: 'pro_yearly', days: 365 });
    const bad = 'SPK1-AAAA.BBBB';
    const run = (k) => {
        try {
            execFileSync(process.execPath, [
                path.join(__dirname, '..', 'tools', 'gen-license-key.js'), '--verify', k
            ], { encoding: 'utf8', stdio: 'pipe' });
            return 0;
        } catch (e) {
            return e.status;
        }
    };
    assert.strictEqual(run(good), 0);
    assert.strictEqual(run(bad), 1);
});

test('CLI --out 同时登记发货台账，且台账里的 key 真实可用', () => {
    const ledger = path.join(__dirname, '..', 'data', 'issued-keys.csv');
    const before = fs.existsSync(ledger) ? fs.readFileSync(ledger, 'utf8').length : 0;
    const tmpCsv = path.join(__dirname, 'tmp-gen-test.csv');
    execFileSync(process.execPath, [
        path.join(__dirname, '..', 'tools', 'gen-license-key.js'),
        '--login', 'ledger@example.com', '--plan', 'pro_yearly', '--days', '365',
        '--out', tmpCsv, '--note', 'unit-test-order'
    ], { encoding: 'utf8', stdio: 'pipe' });
    try {
        const content = fs.readFileSync(ledger, 'utf8');
        assert.ok(content.length > before, '台账应被追加');
        const lastLine = content.trimEnd().split(/\r?\n/).pop();
        const cols = lastLine.split(',');
        assert.strictEqual(cols.length, 7, '台账应为 7 列: issuedAt,key,login,plan,seats,expiresAt,note');
        assert.strictEqual(cols[2], 'ledger@example.com');
        assert.strictEqual(cols[6], 'unit-test-order');
        const v = lk.verify(cols[1]);
        assert.strictEqual(v.valid, true, '台账里登记的 key 必须能通过验签');
        assert.strictEqual(v.login, 'ledger@example.com');
    } finally {
        fs.unlinkSync(tmpCsv);
    }
});
