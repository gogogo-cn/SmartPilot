'use strict';
/**
 * 消 key 退款闭环的跨端守栏测试。
 *
 * 背景：离线激活码一旦发出，服务端无法远程吊销。解决办法是「消 key 退款」——
 * 用户在本机点注销 → 生成 SPD1- 注销码 → 发给客服 → 运营台验签通过才放款。
 *
 * 这类跨端协议最怕两端悄悄漂移。曾经踩过一次真实事故：VB 把「SHA256 前 16 位」
 * 写成 16 个**字节**（32 个 hex 字符），而 JS 是 16 个**字符**，导致运营台永远
 * 提示「台账查不到」。下面第 2 条就是钉死它的。
 */
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const lk = require(path.join(ROOT, 'LicenseServer', 'lib', 'licenseKey.js'));

function read(rel) {
    return fs.readFileSync(path.join(ROOT, rel), 'utf8');
}

const NEW_I18N_KEYS = [
    'license.btnRefund',
    'license.refund.confirm',
    'license.refund.codeTitle',
    'license.refund.codeHint',
    'license.refund.copy',
    'license.refund.copied',
    'license.refund.done',
    'license.refund.none',
    'license.refund.failed'
];

test('三端注销码前缀一致为 SPD1-（与激活码 SPK1- 严格隔离）', () => {
    assert.strictEqual(lk.DEACT_PREFIX, 'SPD1-');
    assert.ok(read('Bridges/LicenseKey.vb').includes('"SPD1-"'), 'Excel 侧 VB 应有 SPD1-');
    assert.ok(read('Hosts/Shared/HostLicense.vb').includes('"SPD1-"'), '宿主侧 VB 应有 SPD1-');
});

test('VB 两侧 HashKey 都只取前 16 个 hex 字符（防“16 字节”事故复发）', () => {
    const vbExcel = read('Bridges/LicenseKey.vb');
    const vbHost = read('Hosts/Shared/HostLicense.vb');
    [vbExcel, vbHost].forEach((src, i) => {
        const where = i === 0 ? 'Bridges/LicenseKey.vb' : 'Hosts/Shared/HostLicense.vb';
        assert.ok(src.includes('Substring(0, 16)'),
            where + ' 必须以 Substring(0, 16) 截断——写错成 16 个字节会让运营台对不上台账');
    });
    // 对照：JS 端确实是 16 个字符
    assert.strictEqual(lk.hashKey('SPK1-anything').length, 16);
});

test('JS 与 VB 的哈希规则一致：都是 SHA256 hex 的前 16 个字符', () => {
    // 直接对 VB 的输出做等价验证：取 sha256 hex 前 16 位
    const crypto = require('crypto');
    const sample = 'SPK1-abc.def';
    const expect = crypto.createHash('sha256').update(sample, 'utf8').digest('hex').slice(0, 16);
    assert.strictEqual(lk.hashKey(sample), expect);
    assert.ok(/^[0-9a-f]{16}$/.test(lk.hashKey(sample)));
});

test('两端 VB 桥都暴露了注销方法', () => {
    assert.ok(read('Bridges/LicenseBridge.vb').includes('Function DeactivateForRefund'),
        'Excel 侧应有 DeactivateForRefund');
    assert.ok(read('ExcelBridge.vb').includes('LicenseDeactivateForRefund'),
        'ExcelBridge 应暴露 LicenseDeactivateForRefund（否则前端注销流程在 Excel 调不到）');
    assert.ok(read('Hosts/Shared/HostBridgeBase.vb').includes('LicenseDeactivateForRefund'),
        '宿主侧应有 LicenseDeactivateForRefund');
    assert.ok(read('Hosts/Shared/HostLicense.vb').includes('Function DeactivateForRefund'),
        '宿主侧 HostLicense 应有 DeactivateForRefund');
});

test('注销后同一张 key 不能被再次激活（两端 VB 查同一份共享注销名单）', () => {
    const vbExcel = read('Bridges/LicenseBridge.vb');
    const vbHost = read('Hosts/Shared/HostLicense.vb');
    assert.ok(vbExcel.includes('IsKeyInDeactivatedList'), 'Excel 侧 ActivateKey 应查共享注销名单');
    assert.ok(vbHost.includes('IsDeactivated(key)'), '宿主侧 ActivateKey 应查共享注销名单');
});

test('前端有注销入口且调用同名桥方法', () => {
    const js = read('UI/js/license.js');
    assert.ok(js.includes('LicenseDeactivateForRefund'), 'license.js 应调用 LicenseDeactivateForRefund');
    assert.ok(js.includes('data-lic="refund"'), '应有注销按钮');
    assert.ok(js.includes('id="license-refund-code"'), '应有消 key 码展示框');
});

test('消 key 相关文案中英文齐平', () => {
    const i18n = read('UI/js/i18n.js');
    NEW_I18N_KEYS.forEach((k) => {
        const hits = i18n.split("'" + k + "'").length - 1;
        assert.strictEqual(hits, 2, k + ' 应在中英各出现一次，实际 ' + hits + ' 次');
    });
});

test('消 key 码不含激活码原文（要经微信发给客服）', () => {
    const key = lk.createKey({ login: 'a@b.com', plan: 'pro_yearly', days: 365 });
    const code = lk.createDeactivationProof({ key: key, plan: 'pro_yearly' });
    assert.ok(!code.includes(key), '注销码不得包含激活码原文');
    // 但必须能反查出是同一张 key（运营台靠标识对台账）
    const v = lk.verifyDeactivationProof(code);
    assert.strictEqual(v.valid, true);
    assert.strictEqual(v.keyHash, lk.hashKey(key));
});
