'use strict';
/**
 * 消 key 码（退款凭证）测试。
 *
 * 业务闭环：用户本机销毁激活码 → 生成注销码 → 发给客服 → 运营台验签后退款。
 * 因此「注销码不能反推出激活码」是硬要求——它是要经微信/邮件传给客服的。
 */
const test = require('node:test');
const assert = require('node:assert');
const lk = require('../lib/licenseKey');
const revocation = require('../lib/revocation');

const KEY = lk.createKey({ login: 'a@b.com', plan: 'pro_yearly', days: 365 });

test('生成注销码 → 验签通过并解析出字段', () => {
    const code = lk.createDeactivationProof({ key: KEY, plan: 'pro_yearly' });
    assert.ok(code.startsWith(lk.DEACT_PREFIX), '应以 SPD1- 开头');

    const r = lk.verifyDeactivationProof(code);
    assert.strictEqual(r.valid, true);
    assert.strictEqual(r.reason, 'ok');
    assert.strictEqual(r.keyHash, lk.hashKey(KEY));
    assert.strictEqual(r.plan, 'pro_yearly');
    assert.ok(r.destroyedAt, '应带销毁时间');
});

test('注销码不得包含激活码原文（它是要发给客服的）', () => {
    const code = lk.createDeactivationProof({ key: KEY, plan: 'pro_yearly' });
    // payload 是 base64url，但原文若被塞进去，解出来一定会暴露
    const body = code.slice(lk.DEACT_PREFIX.length);
    const payloadText = lk.base64UrlDecode(body.slice(0, body.lastIndexOf('.'))).toString('utf8');
    assert.ok(!payloadText.includes(KEY), 'payload 不得含激活码原文');
    assert.ok(!code.includes(KEY), '整个注销码不得含激活码原文');
    // 抽查：激活码里最独特的那段（签名）不能出现在注销码里
    const keySig = KEY.slice(KEY.lastIndexOf('.') + 1);
    assert.ok(keySig.length > 10 && !code.includes(keySig), '不得含激活码签名段');
});

test('篡改 payload → 验签失败', () => {
    const code = lk.createDeactivationProof({ key: KEY, plan: 'pro_yearly' });
    const body = code.slice(lk.DEACT_PREFIX.length);
    const dot = body.lastIndexOf('.');
    const sig = body.slice(dot + 1);
    const payloadText = lk.base64UrlDecode(body.slice(0, dot)).toString('utf8');
    const tampered = payloadText.replace('pro_yearly', 'pro_monthly');
    const forged = lk.DEACT_PREFIX + lk.base64UrlEncode(Buffer.from(tampered, 'utf8')) + '.' + sig;

    const r = lk.verifyDeactivationProof(forged);
    assert.strictEqual(r.valid, false);
    assert.strictEqual(r.reason, 'bad_signature');
});

test('篡改签名 → 验签失败', () => {
    const code = lk.createDeactivationProof({ key: KEY, plan: 'pro_yearly' });
    const forged = code.slice(0, -4) + (code.slice(-4) === 'AAAA' ? 'BBBB' : 'AAAA');
    const r = lk.verifyDeactivationProof(forged);
    assert.strictEqual(r.valid, false);
    assert.strictEqual(r.reason, 'bad_signature');
});

test('激活码不能当注销码用（前缀隔离）', () => {
    const r = lk.verifyDeactivationProof(KEY);
    assert.strictEqual(r.valid, false);
    assert.strictEqual(r.reason, 'malformed');

    const code = lk.createDeactivationProof({ key: KEY });
    const r2 = lk.verify(code);
    assert.strictEqual(r2.valid, false, '注销码也不能当激活码用');
});

test('乱码/空值 → malformed，不抛异常', () => {
    ['', '   ', 'abc', 'SPD1-', 'SPD1-xxx', 'SPK1-abc.def', null, undefined].forEach((bad) => {
        const r = lk.verifyDeactivationProof(bad);
        assert.strictEqual(r.valid, false, String(bad) + ' 应无效');
        assert.strictEqual(r.reason, 'malformed');
    });
});

test('hashKey 与 revocation.hashKey 逐字节一致（防两处实现漂移）', () => {
    const keys = [KEY, lk.createKey({ login: 'x@y.com', days: 30 }), '', 'SPK1-whatever'];
    keys.forEach((k) => {
        assert.strictEqual(lk.hashKey(k), revocation.hashKey(k), '同一 key 两处哈希必须相同');
    });
    // 且必须符合同一套吊销标识规则（16 位小写 hex）
    assert.ok(revocation.isHash(lk.hashKey(KEY)), '应能被 revocation 识别为合法吊销标识');
});

test('同一激活码两次销毁：码不同（nonce）但标识相同（可查同一条台账）', () => {
    const a = lk.createDeactivationProof({ key: KEY });
    const b = lk.createDeactivationProof({ key: KEY });
    assert.notStrictEqual(a, b, 'nonce 应使两次生成的注销码不同');
    assert.strictEqual(
        lk.verifyDeactivationProof(a).keyHash,
        lk.verifyDeactivationProof(b).keyHash,
        '同一 key 的标识必须稳定，否则运营台无法对上台账'
    );
});

test('不同激活码 → 不同标识', () => {
    const other = lk.createKey({ login: 'c@d.com', days: 365 });
    assert.notStrictEqual(lk.hashKey(KEY), lk.hashKey(other));
});
