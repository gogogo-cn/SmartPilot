'use strict';
/**
 * ============================================================
 * Tests/no-server-required.test.js — 「零服务器」守栏
 * ============================================================
 * 业务前提：**卖软件 + 激活 + 退款整条链路不需要服务器**。
 *   发码    → 本地 CLI（LicenseServer/lib/licenseKey.js，纯 crypto）
 *   验签    → 客户端 VB 本地 HMAC
 *   注销    → 客户端 VB 本地生成 SPD1- 消 key 码
 *   验注销码 → 运营台本地 Node（当**库**用，不起服务）
 *   退款    → 运营台写本地 CSV / JSON
 *
 * 服务器（LicenseServer 作为进程运行）只是**可选增强**：在线订阅、Stripe/国内支付、
 * 吊销名单下发。它挂了、甚至根本不部署，上面这条闭环照常工作。
 *
 * 本套件就是要把这个前提钉死：一旦有人在链路上加联网调用（哪怕只是"顺便校验一下"），
 * 测试立刻红。
 */
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const UI_JS = path.join(ROOT, 'UI', 'js');

/** 离线闭环涉及的模块：这些文件一旦引入网络库就破坏了「无服务器可用」的前提 */
const OFFLINE_LOOP_MODULES = [
    'LicenseServer/lib/licenseKey.js',
    'LicenseServer/lib/revocation.js',
    'OpsConsole/lib/deactivation.js',
    'OpsConsole/lib/ledger.js',
    'OpsConsole/lib/store.js'
];

/** 允许的内置模块（离线计算用得到） */
const ALLOWED_BUILTINS = new Set(['fs', 'path', 'crypto', 'os', 'url', 'util', 'assert']);

/** 一旦出现即视为破坏离线前提 */
const FORBIDDEN_MODULES = [
    'http', 'https', 'net', 'dns', 'http2', 'tls', 'dgram',
    'express', 'axios', 'node-fetch', 'got', 'superagent', 'request', 'ws', 'socket.io'
];

function read(rel) {
    return fs.readFileSync(path.join(ROOT, rel), 'utf8');
}

/** 抽出源码里所有 require('X') 的模块名 */
function requireTargets(src) {
    const out = [];
    const re = /require\(\s*['"]([^'"]+)['"]\s*\)/g;
    let m;
    while ((m = re.exec(src)) !== null) out.push(m[1]);
    return out;
}

// ---------------------------------------------------------------------------
// 1. 静态：离线闭环模块不依赖任何网络库
// ---------------------------------------------------------------------------
test('离线闭环模块不 require 任何网络库', () => {
    for (const rel of OFFLINE_LOOP_MODULES) {
        const src = read(rel);
        for (const dep of requireTargets(src)) {
            if (dep.startsWith('.') || dep.startsWith('/')) continue; // 本地相对引用，OK
            assert.ok(
                ALLOWED_BUILTINS.has(dep),
                rel + ' 引入了非离线依赖 "' + dep + '" —— 这会破坏「无服务器可用」的前提'
            );
            assert.ok(
                FORBIDDEN_MODULES.indexOf(dep) < 0,
                rel + ' 引入了网络依赖 "' + dep + '" —— 离线闭环不得联网'
            );
        }
    }
});

// ---------------------------------------------------------------------------
// 2. 静态：离线闭环不使用 fetch
// ---------------------------------------------------------------------------
test('离线闭环模块不使用 fetch', () => {
    for (const rel of OFFLINE_LOOP_MODULES) {
        const src = read(rel);
        assert.ok(!/\bfetch\s*\(/.test(src), rel + ' 里出现了 fetch( —— 离线闭环不得联网');
    }
});

// ---------------------------------------------------------------------------
// 3. 运行时：把网络全部毒化，整条闭环仍必须跑通
//    这是最硬的一条证据——不是"看代码觉得没联网"，而是真断网也过。
// ---------------------------------------------------------------------------
test('毒化全部网络出口后，发码→销码→退款闭环仍完整跑通', () => {
    const http = require('http');
    const https = require('https');
    const net = require('net');
    const dns = require('dns');

    const saved = {
        httpRequest: http.request,
        httpGet: http.get,
        httpsRequest: https.request,
        httpsGet: https.get,
        netConnect: net.connect,
        dnsLookup: dns.lookup,
        fetch: globalThis.fetch
    };
    const boom = function () { throw new Error('NETWORK_FORBIDDEN: 离线闭环不得发起任何网络请求'); };

    http.request = boom;
    http.get = boom;
    https.request = boom;
    https.get = boom;
    net.connect = boom;
    dns.lookup = boom;
    globalThis.fetch = boom;

    try {
        const { OpsStore } = require(path.join(ROOT, 'OpsConsole', 'lib', 'store'));
        const { verifyProof, refundByProof, hashKey } = require(path.join(ROOT, 'OpsConsole', 'lib', 'deactivation'));
        const { Revocation } = require(path.join(ROOT, 'LicenseServer', 'lib', 'revocation.js'));
        const lk = require(path.join(ROOT, 'LicenseServer', 'lib', 'licenseKey.js'));

        const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'noserver-'));
        const store = new OpsStore({
            ledgerPath: path.join(dir, 'issued-keys.csv'),
            statePath: path.join(dir, 'ops-state.json')
        });

        // ① 发码（等价于 gen-key CLI）
        const rows = store.addKeys({ login: 'a@b.com', plan: 'pro_yearly', days: 365, count: 1 });
        const key = rows[0].key;
        assert.ok(key.indexOf(lk.PREFIX) === 0, '应生成 SPK1- 激活码');

        // ② 用户在本机销 key，拿到消 key 码
        const code = lk.createDeactivationProof({ key: key, plan: 'pro_yearly' });
        assert.ok(code.indexOf(lk.DEACT_PREFIX) === 0, '应生成 SPD1- 消 key 码');

        // ③ 运营台验签 + 定位台账
        const v = verifyProof(code, store.rows());
        assert.strictEqual(v.ok, true, '断网状态下验签必须仍然通过');
        assert.strictEqual(v.row.key, key);

        // ④ 确认退款（同时写吊销名单）
        const r = refundByProof({ code: code, store: store, dataDir: dir });
        assert.strictEqual(r.ok, true, '断网状态下退款必须仍然成功');
        assert.strictEqual(r.state.status, 'refunded');
        assert.ok(r.revoked && !r.revoked.error, '吊销名单应写入成功：' + JSON.stringify(r.revoked));

        const rev = new Revocation(dir);
        assert.strictEqual(rev.hasKey(key), true, '吊销名单应含该 key 的不可逆标识');
        assert.strictEqual(rev.hasHash(hashKey(key)), true);
        assert.ok(fs.existsSync(path.join(dir, 'revoked.json')), 'revoked.json 应落盘');
    } finally {
        http.request = saved.httpRequest;
        http.get = saved.httpGet;
        https.request = saved.httpsRequest;
        https.get = saved.httpsGet;
        net.connect = saved.netConnect;
        dns.lookup = saved.dnsLookup;
        globalThis.fetch = saved.fetch;
    }
});

// ---------------------------------------------------------------------------
// 4. Excel 侧：LicenseKey.vb 纯本地运算，绝不联网
// ---------------------------------------------------------------------------
test('Excel 侧 LicenseKey.vb 无任何 HTTP 客户端（纯本地 HMAC）', () => {
    const src = read('Bridges/LicenseKey.vb');
    for (const bad of ['HttpClient', 'WebRequest', 'HttpWebRequest', 'WebClient', 'System.Net']) {
        assert.ok(src.indexOf(bad) < 0, 'LicenseKey.vb 出现了 ' + bad + ' —— 验签/发码必须是纯本地运算');
    }
});

// ---------------------------------------------------------------------------
// 5. 宿主侧：验签与生成注销码这两个函数体内不得联网
//    （HostLicense.vb 整体有 HttpClient，那是**在线登录**用的，不是验签用的，
//     所以断言必须收到具体函数上，不能一刀切扫整个文件。）
// ---------------------------------------------------------------------------
test('宿主侧 VerifyKey / CreateDeactivationProof 函数体不联网', () => {
    const src = read('Hosts/Shared/HostLicense.vb');

    function bodyOf(sig) {
        const i = src.indexOf(sig);
        assert.ok(i >= 0, '应能找到函数：' + sig);
        const j = src.indexOf('End Function', i);
        assert.ok(j > i, '函数应有 End Function：' + sig);
        return src.slice(i, j);
    }

    for (const sig of ['Private Function VerifyKey(', 'Private Function CreateDeactivationProof(']) {
        const body = bodyOf(sig);
        assert.ok(!/\bHttp\b/.test(body), sig + ' 函数体内出现了 Http —— 该运算必须离线完成');
        assert.ok(body.indexOf('HttpClient') < 0, sig + ' 不得使用 HttpClient');
    }
});

// ---------------------------------------------------------------------------
// 6. 前端门禁：离线激活命中时，先于任何联网探测返回
// ---------------------------------------------------------------------------
test('前端门禁在离线激活命中时直接放行，不先探测服务器', () => {
    const licenseJs = read(path.join('UI', 'js', 'license.js'));
    const start = licenseJs.indexOf('async function ensureSubscriptionActive');
    assert.ok(start >= 0, '应能定位 ensureSubscriptionActive');
    const body = licenseJs.slice(start, start + 2600);

    const localIdx = body.indexOf('LicenseGetLocalActivation');
    const pingIdx = body.indexOf('pingLicenseServer');
    assert.ok(localIdx >= 0, '门禁应先查本地离线激活');

    // 注意：第一个 return true 是 softBypassEnabled()（开发绕过，本身也离线），
    // 它排在离线激活检查之前是合理的。要守的是「本地激活命中后的那个放行分支」。
    const returnAfterLocal = body.indexOf('return true', localIdx);
    assert.ok(returnAfterLocal > localIdx, '本地激活检查之后必须紧跟一个放行分支');
    assert.ok(
        returnAfterLocal < pingIdx || pingIdx < 0,
        '离线激活的放行分支必须出现在任何服务器探测之前 —— 否则无服务器时门禁会卡住'
    );
});

// ---------------------------------------------------------------------------
// 7. 运营台退款接口不得回调授权服务（避免"没部署服务器就不能退款"）
// ---------------------------------------------------------------------------
test('运营台退款接口纯本地，不回调授权服务', () => {
    const src = read(path.join('OpsConsole', 'server.js'));
    const start = src.indexOf("route === '/api/deactivation/verify'");
    const end = src.indexOf("route === '/api/payment-config'") >= 0
        ? src.indexOf("route === '/api/payment-config'")
        : src.length;
    assert.ok(start >= 0, '应能定位退款相关路由');

    // 往前挪一点，把上一行注释也包进来
    const block = src.slice(Math.max(0, start - 200), end);
    for (const bad of ['httpJson', 'fetch(']) {
        assert.ok(block.indexOf(bad) < 0, '退款接口不得调用 ' + bad + ' —— 无服务器时也必须能退款');
    }
});
