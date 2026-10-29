/**
 * ============================================================
 * Tests/mobile-bridge-auth.test.js — 本地遥控服务（9528）鉴权与 CORS 护栏
 * ============================================================
 * 背景（2026-10-09 审查发现，P0 安全问题）：
 *   Bridges/MobileControlBridge.vb 自建 TcpListener 监听 9528 并对所有响应
 *   无条件输出 `Access-Control-Allow-Origin: *`；而 `/api/status` 属于公开路径
 *   （IsPublicPath 放行、无需 token）且响应里直接带 `sessionToken`。
 *   ⇒ 任意第三方网页只要 `fetch('http://127.0.0.1:9528/api/status')` 就能读到
 *     token，随后凭 token 调用 /api/upload 等接口操作本机的 Excel/文档。
 *
 * 修复：
 *   ① 响应头不再无条件 "*"，改为按 Origin 白名单回显（ResolveCorsOrigin）；
 *   ② /api/status 的 sessionToken 仅下发给「无 Origin 的同源请求」或可信来源
 *      （localhost / 127.0.0.1 / ::1 / 宿主虚拟主机名 / 服务自身局域网 IP）。
 *
 * 本套件用源码级断言固化这两条约束（VB 无法在 node 中执行，只能静态校验）。
 * 运行：node Tests/mobile-bridge-auth.test.js（随 run-tests.js 自动发现）
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const SRC = path.join(ROOT, 'Bridges', 'MobileControlBridge.vb');

function src() {
    assert.ok(fs.existsSync(SRC), 'Bridges/MobileControlBridge.vb 应存在');
    return fs.readFileSync(SRC, 'utf8');
}

/** 去掉 VB 注释行（''' 开头），只保留代码，避免注释里的历史描述误命中 */
function code() {
    return src().split(/\r?\n/).filter(l => !/^\s*'''/.test(l)).join('\n');
}

/** 抽取 GetApiStatus 函数体 */
function apiStatusBody() {
    const s = code();
    const start = s.indexOf('Private Function GetApiStatus(');
    assert.ok(start > -1, '应有 GetApiStatus');
    const end = s.indexOf('End Function', start);
    return s.slice(start, end);
}

test('安全: 不得再无条件输出 Access-Control-Allow-Origin: *', () => {
    const s = code();
    assert.ok(!/Access-Control-Allow-Origin:\s*\*/.test(s),
        '不应再有字面量 "Access-Control-Allow-Origin: *"（会使任意网页可读本地服务响应体）');
    console.log('  [ok] 无通配 CORS 响应头字面量');
});

test('安全: CORS 来源按白名单回显（ResolveCorsOrigin + IsTrustedOrigin）', () => {
    const s = src();
    assert.ok(/Private Function ResolveCorsOrigin\(/.test(s), '应有 ResolveCorsOrigin 计算可信来源');
    assert.ok(/Private Function IsTrustedOrigin\(/.test(s), '应有 IsTrustedOrigin 白名单判定');
    // 白名单须包含本机与宿主虚拟主机名
    for (const host of ['localhost', '127.0.0.1', 'app.excelai.local']) {
        assert.ok(s.includes(`"${host}"`) || s.includes(host), `可信来源白名单应包含 ${host}`);
    }
    console.log('  [ok] 来源白名单判定就位');
});

test('安全: /api/status 的 sessionToken 受可信来源约束，不再裸下发', () => {
    const s = code();
    assert.ok(/Private Function CanDiscloseSessionToken\(/.test(s), '应有 CanDiscloseSessionToken 门槛');
    const body = apiStatusBody();
    // GetApiStatus 必须接收 request 并按门槛决定是否下发 token
    assert.ok(/Private Function GetApiStatus\(request As HttpRequestInfo\)/.test(body),
        'GetApiStatus 必须接收 request，以便判断来源是否可信');
    assert.ok(/CanDiscloseSessionToken\(request\)/.test(body),
        'GetApiStatus 必须经 CanDiscloseSessionToken(request) 决定是否下发 token');
    // 不得再直接把 _sessionToken 写进 /api/status 响应（历史上正是这里泄露）
    assert.ok(!/\.sessionToken = _sessionToken\b/.test(body),
        '不得再把 _sessionToken 无条件写进 /api/status 响应');
    console.log('  [ok] /api/status 的 token 下发已受来源约束');
});

test('安全: 响应写入函数支持按来源输出 ACAO（不可信来源不输出该头）', () => {
    const s = code();
    for (const fn of ['SendHttpResponse', 'SendHttpResponseBytes']) {
        const re = new RegExp(`Private Sub ${fn}\\([\\s\\S]*?Optional corsOrigin As String`);
        assert.ok(re.test(s), `${fn} 应支持 Optional corsOrigin 参数`);
    }
    assert.ok(/If Not String\.IsNullOrEmpty\(corsOrigin\) Then sb\.AppendLine/.test(s),
        'corsOrigin 为空时应跳过 ACAO 头输出（浏览器因此阻断响应体）');
    console.log('  [ok] 响应函数按来源输出 ACAO');
});

test('安全: 鉴权仍保留（公开路径之外的 /api/* 需要 session token）', () => {
    const s = src();
    assert.ok(/If Not IsPublicPath\(path, method\) Then/.test(s), '仍应对非公开路径做鉴权');
    assert.ok(/Private Function ValidateRequestToken\(/.test(s), '仍应有 token 校验');
    console.log('  [ok] 鉴权链路保留');
});
