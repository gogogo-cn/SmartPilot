'use strict';
/**
 * 服务端 / 客户端文件边界守栏。
 *
 * 防止两件事：
 *  1) 上传包夹带密钥或用户数据（.env / config.json / data / logs / node_modules / test）；
 *  2) 支付密钥（Stripe / 微信 / 支付宝私钥）被写进客户端源码（UI/js、Bridges、Hosts/Shared）。
 *
 * 注：离线激活码用的 HMAC 密钥**按设计**在客户端（Bridges/LicenseKey.vb、
 * Hosts/Shared/HostLicense.vb），离线验签必须用同一把对称密钥，无法避免；
 * 本测试只拦"支付私钥"，不拦它。
 */
const { test, before, after } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const pkg = require(path.join(ROOT, 'Tools', 'package-server.js'));

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'srv-pkg-'));
const outDir = path.join(tmp, 'server');

after(() => {
    try { fs.rmSync(tmp, { recursive: true, force: true }); } catch (_) {}
});

let result;

before(() => {
    // 临时目录在仓库外，需显式放行
    result = pkg.buildServerPackage({ outDir, allowOutside: true });
});

test('上传包包含服务端运行必需文件', () => {
    ['server.js', 'package.json', 'package-lock.json', '.env.example',
     'lib/cnPay.js', 'lib/paymentConfig.js', 'lib/store.js', 'lib/licenseKey.js',
     'public/cn/pay.html', 'tools/gen-license-key.js'
    ].forEach((rel) => {
        assert.ok(fs.existsSync(path.join(outDir, rel)), '缺少: ' + rel);
    });
});

test('上传包绝不包含密钥 / 用户数据目录', () => {
    ['.env', 'config.json', 'data', 'data/license-db.json', 'data/issued-keys.csv',
     'logs', 'node_modules', 'test'
    ].forEach((rel) => {
        assert.strictEqual(fs.existsSync(path.join(outDir, rel)), false, '不该包含: ' + rel);
    });
});

test('上传包通过真实密钥扫描', () => {
    assert.deepStrictEqual(result.secretFindings, []);
});

test('密钥扫描器确实能抓到密钥（阳性对照，防止扫描器形同虚设）', () => {
    const bad = path.join(tmp, 'bad');
    fs.mkdirSync(bad, { recursive: true });
    fs.writeFileSync(path.join(bad, 'x.txt'), 'STRIPE_SECRET_KEY=sk_live_51Nabcdefghijklmnop');
    fs.writeFileSync(path.join(bad, 'y.txt'), '-----BEGIN RSA PRIVATE KEY-----\nabc\n-----END RSA PRIVATE KEY-----');
    const found = pkg.scanForSecrets(bad);
    assert.ok(found.length >= 2, '扫描器应至少命中 2 处，实际: ' + found.length);
});

test('占位值不会被误判为密钥（.env.example 可安全进包）', () => {
    const ex = fs.readFileSync(path.join(ROOT, 'LicenseServer', '.env.example'), 'utf8');
    assert.ok(/sk_test_xxx/.test(ex), '.env.example 应含占位值');
    const findings = pkg.scanForSecrets(outDir);
    const envFindings = findings.filter((f) => f.file.includes('.env.example'));
    assert.deepStrictEqual(envFindings, [], '.env.example 的占位值不应被判为密钥');
});

test('客户端源码不含支付私钥（Stripe / 微信 / 支付宝）', () => {
    const clientDirs = ['UI/js', 'Bridges', 'Hosts/Shared'];
    const patterns = pkg.SECRET_PATTERNS;
    const hits = [];
    clientDirs.forEach((rel) => {
        const dir = path.join(ROOT, rel);
        if (!fs.existsSync(dir)) return;
        const walk = (d) => {
            for (const entry of fs.readdirSync(d)) {
                const p = path.join(d, entry);
                const st = fs.statSync(p);
                if (st.isDirectory()) { walk(p); continue; }
                if (!/\.(js|vb|html|json)$/i.test(entry)) continue;
                if (st.size > 2 * 1024 * 1024) continue;
                let text = '';
                try { text = fs.readFileSync(p, 'utf8'); } catch (_) { continue; }
                patterns.forEach((pat) => {
                    if (pat.re.test(text)) hits.push(path.relative(ROOT, p) + ' → ' + pat.name);
                });
            }
        };
        walk(dir);
    });
    assert.deepStrictEqual(hits, [], '客户端源码不应含支付私钥: ' + hits.join(', '));
});

test('--include-data 会带上 data/，但仍排除 .bak 备份', () => {
    const out2 = path.join(tmp, 'server-with-data');
    const res2 = pkg.buildServerPackage({ outDir: out2, includeData: true, allowOutside: true });
    assert.ok(fs.existsSync(path.join(out2, 'data')), 'includeData 应包含 data/');
    const ledger = fs.existsSync(path.join(out2, 'data', 'issued-keys.csv'));
    assert.ok(ledger, '台账应被带上（用于迁移）');
    const bak = fs.readdirSync(path.join(out2, 'data')).some((f) => f.includes('.bak'));
    assert.strictEqual(bak, false, '.bak 备份不应进包');
});
