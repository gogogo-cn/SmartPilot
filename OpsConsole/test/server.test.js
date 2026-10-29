'use strict';
/**
 * server.test.js — 运营台 API 端到端（临时台账 + 临时状态层 + 真实 public/index.html）
 */
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { createServer } = require('../server');
const { OpsStore } = require('../lib/store');
const lk = require('../../LicenseServer/lib/licenseKey.js');

const INDEX_HTML = path.join(__dirname, '..', 'public', 'index.html');

function withServer(fn) {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ops-server-'));
    const store = new OpsStore({
        ledgerPath: path.join(dir, 'issued-keys.csv'),
        statePath: path.join(dir, 'ops-state.json')
    });
    const server = createServer({ store, indexHtml: INDEX_HTML });
    return new Promise((resolve, reject) => {
        server.listen(0, '127.0.0.1', async () => {
            const base = 'http://127.0.0.1:' + server.address().port;
            try {
                await fn(base);
            } catch (e) {
                reject(e);
            } finally {
                server.close(() => resolve());
            }
        });
    });
}

function api(base, route, opts) {
    return fetch(base + route, opts).then(r => r.json());
}

test('GET / 返回运营台页面', async () => {
    await withServer(async (base) => {
        const r = await fetch(base + '/');
        assert.strictEqual(r.status, 200);
        const html = await r.text();
        assert.ok(html.includes('SmartPilot 运营台'));
        assert.ok(html.includes('id="gen-form"'));
    });
});

test('完整链路：生成 → 台账可见 → 标记已发货 → 统计更新 → 客户文案', async () => {
    await withServer(async (base) => {
        // 1) 生成 2 个 key
        const gen = await api(base, '/api/keys', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ login: 'e2e@example.com', plan: 'pro_yearly', days: 365, note: 'e2e订单', count: 2 })
        });
        assert.strictEqual(gen.success, true);
        assert.strictEqual(gen.rows.length, 2);
        for (const r of gen.rows) assert.strictEqual(lk.verify(r.key).valid, true, 'API 生成的 key 必须能验签');

        // 2) 台账里能查到（含 login 后缀）
        const list = await api(base, '/api/keys');
        assert.strictEqual(list.rows.length, 2);
        assert.deepStrictEqual(list.rows.map(r => r.login).sort(),
            ['e2e+1@example.com', 'e2e+2@example.com']);

        // 3) 标记一个已发货
        const key = gen.rows[0].key;
        const st = await api(base, '/api/keys/status', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ key, status: 'sent' })
        });
        assert.strictEqual(st.success, true);
        assert.strictEqual(st.state.status, 'sent');
        const after = await api(base, '/api/keys?status=sent');
        assert.strictEqual(after.rows.length, 1);
        assert.strictEqual(after.rows[0].key, key);

        // 4) 统计：pending 1 / sent 1
        const ov = await api(base, '/api/overview');
        assert.strictEqual(ov.stats.pendingCount, 1);
        assert.strictEqual(ov.stats.sentCount, 1);
        assert.strictEqual(ov.stats.totalSales, 398);

        // 5) 客户文案
        const text = await api(base, '/api/customer-text?key=' + encodeURIComponent(key));
        assert.strictEqual(text.success, true);
        assert.ok(text.text.includes(key));
    });
});

test('POST /api/keys 缺 login → 400；标记不存在 key → 404；坏 JSON → 500 不崩', async () => {
    await withServer(async (base) => {
        const bad = await api(base, '/api/keys', {
            method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ plan: 'team' })
        });
        assert.strictEqual(bad.success, false);
        assert.ok(bad.error.includes('login'));

        const nf = await api(base, '/api/keys/status', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ key: 'SPK1-nope.nope', status: 'sent' })
        });
        assert.strictEqual(nf.success, false);
        assert.ok(nf.error.includes('不存在'));

        // 坏 JSON：fetch 层抛错或返回 500，两种情况都不应让服务挂掉
        let alive = false;
        try {
            const r = await fetch(base + '/api/keys', { method: 'POST', body: '{oops' });
            alive = r.status === 500;
        } catch (_) { /* 连接被拒也算失败，但这里应拿到 500 */ }
        const still = await api(base, '/api/overview');
        assert.ok(alive && still.success, '坏 JSON 后服务必须仍存活');
    });
});

test('GET /api/export.csv 带 BOM 且行数与台账一致', async () => {
    await withServer(async (base) => {
        // 团队版 5 人起订（2026-10-26）：count=5 踩线发码，导出 5 行
        await api(base, '/api/keys', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ login: 'exp@example.com', plan: 'team', days: 365, count: 5 })
        });
        const r = await fetch(base + '/api/export.csv');
        assert.strictEqual(r.status, 200);
        const buf = Buffer.from(await r.arrayBuffer());
        assert.strictEqual(buf[0], 0xEF, 'CSV 应以 UTF-8 BOM 开头（Excel 兼容）');
        assert.strictEqual(buf[1], 0xBB);
        assert.strictEqual(buf[2], 0xBF);
        const text = buf.toString('utf8');
        const lines = text.replace(/^\uFEFF/, '').split('\r\n').filter(Boolean);
        assert.strictEqual(lines.length, 6, '表头 + 5 行');
        assert.ok(lines[0].startsWith('issuedAt,key,login,plan,seats,expiresAt,status,amount'));
    });
});

test('POST /api/keys：团队版不满起订量 → 400（不写进台账）', async () => {
    await withServer(async (base) => {
        const bad = await api(base, '/api/keys', {
            method: 'POST', headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ login: 'small@example.com', plan: 'team', seats: 3, days: 365 })
        });
        assert.strictEqual(bad.success, false);
        assert.strictEqual(bad.error.indexOf('5 人起订') >= 0, true, '错误信息要写起订量：' + bad.error);
        // 被拒的单子不能落台账
        const csv = await fetch(base + '/api/export.csv');
        const text = Buffer.from(await csv.arrayBuffer()).toString('utf8');
        assert.ok(text.indexOf('small@example.com') < 0, '起订不足的单子不应写进台账');
    });
});

test('未知路由 404', async () => {
    await withServer(async (base) => {
        const j = await api(base, '/api/nope');
        assert.strictEqual(j.success, false);
    });
});

test('CORS：file://（Origin: null）与本机源放行，其他源不给头', async () => {
    await withServer(async (base) => {
        // file:// 双击打开时 Origin 为 null——不放行就连不上 API
        const r1 = await fetch(base + '/api/overview', { headers: { Origin: 'null' } });
        assert.strictEqual(r1.status, 200);
        assert.strictEqual(r1.headers.get('access-control-allow-origin'), 'null');
        // 本机 http 源同样放行
        const r2 = await fetch(base + '/api/overview', { headers: { Origin: 'http://127.0.0.1:9600' } });
        assert.strictEqual(r2.headers.get('access-control-allow-origin'), 'http://127.0.0.1:9600');
        // 其他源一律不给 CORS 头（本机运营台不能被任意网页读）
        const r3 = await fetch(base + '/api/overview', { headers: { Origin: 'https://evil.example.com' } });
        assert.strictEqual(r3.status, 200);
        assert.strictEqual(r3.headers.get('access-control-allow-origin'), null);
        // 预检：PUT 带 X-Admin-Key / POST 带 JSON 都会触发 OPTIONS
        const pre = await fetch(base + '/api/keys/status', {
            method: 'OPTIONS',
            headers: {
                Origin: 'null',
                'Access-Control-Request-Method': 'POST',
                'Access-Control-Request-Headers': 'content-type,x-admin-key'
            }
        });
        assert.strictEqual(pre.status, 204);
        assert.strictEqual(pre.headers.get('access-control-allow-origin'), 'null');
        assert.ok(pre.headers.get('access-control-allow-headers').includes('X-Admin-Key'),
            '预检必须放行 X-Admin-Key，否则支付配置写不进去');
    });
});

test('守栏：file:// 双击打开可用（api 绝对基址 + 探测式帮助条 + 一键启动器）', () => {
    // 2026-10-21 用户反馈："怎么启动，不能给个直接点的按钮吗"。
    // 三件套：api() 在 file:// 下自动指向本机服务；CORS 放行 null 源（上面已功能验证）；
    // 服务没起时帮助条给出双击启动器 / 复制命令，而不是一句"所有功能都不可用"。
    const html = fs.readFileSync(INDEX_HTML, 'utf8');
    assert.ok(/location\.protocol === 'file:' \? \('http:\/\/127\.0\.0\.1:' \+ OPS_PORT\)/.test(html),
        'api() 在 file:// 下必须用绝对基址（/api/... 会解析成 file:///api/...）');
    assert.ok(/\[\?&\]ops=\(\\d\{2,5\}\)/.test(html), '应支持 ?ops= 端口覆盖（对应 OPS_PORT）');
    // 帮助条必须是探测式的：服务在跑就安静，不在才出现
    assert.ok(/api\('\/api\/overview'\)\.then\(function \(\) \{[\s\S]*?\}\)\.catch\(function \(\) \{/.test(html),
        '帮助条应由探测结果触发，而非无条件红条');
    assert.ok(html.indexOf('fb-copy') >= 0 && html.indexOf('copyText(') >= 0, '帮助条应有复制启动命令按钮');
    assert.ok(html.indexOf('fb-open') >= 0, '帮助条应有直接打开运营台的按钮');
    assert.ok(html.indexOf('start-ops-console.bat') >= 0, '帮助条应指明一键启动器文件名');

    // 一键启动器本体
    const bat = path.join(__dirname, '..', 'start-ops-console.bat');
    assert.ok(fs.existsSync(bat), '应有 OpsConsole/start-ops-console.bat 一键启动器');
    const batSrc = fs.readFileSync(bat, 'utf8');
    assert.ok(/server\.js/.test(batSrc), '启动器必须启动 server.js');
    // 端口：默认 9600，但全部走 OPS_PORT 变量（否则 bat 里"换端口"提示是空话：
    // 服务端会听 9610，启动器却探测/打开 9600，必然超时）
    assert.ok(/if not defined OPS_PORT set "OPS_PORT=9600"/.test(batSrc), '启动器端口应默认 9600');
    assert.ok(/%OPS_PORT%/.test(batSrc), '启动器探测/打开地址应走 OPS_PORT 变量');
    // 探测实现：必须用 node 原生 http 模块，禁用任何 HTTP 客户端命令行工具
    // 2026-10-21 实测：设了 http_proxy/HTTP_PROXY 的机器（含公司网络）上，
    // curl 对死端口也返回 0——代理替它应答。启动器因此误判"已在运行"、
    // 跳过启动、浏览器打开一个并不存在的页面。node 的 http 模块不吃代理变量。
    const batCode = batSrc.split(/\r?\n/).filter(l => !/^\s*REM\b/i.test(l)).join('\n');
    assert.ok(!/\b(curl|wget)\b/.test(batCode), '启动器探测禁用 curl/wget：代理环境下会把死端口报成已启动');
    assert.ok(/require\('http'\)\.get\('http:\/\/127\.0\.0\.1:%OPS_PORT%\/api\/overview'/.test(batCode),
        '探测应直接用 node http 模块请求 /api/overview');
    // 闪退根因守栏：cmd 按系统 ANSI 代码页解析 .bat，任何非 ASCII 行都会被
    // 字节错位拆成乱码"命令"，窗口一闪即退。全 ASCII 是唯一可靠形态
    // （chcp 65001 只管输出显示，管不了文件解析）。
    assert.ok(!/[^\x00-\x7F]/.test(batSrc), '启动器必须全 ASCII：非 ASCII 行会被 cmd 拆成乱码命令导致闪退');
});
