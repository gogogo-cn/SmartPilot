/**
 * BYOK 席位订阅：客户端与法务文案护栏
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');

test('LicenseServer 存在且不计量 Token', () => {
    const server = fs.readFileSync(path.join(ROOT, 'LicenseServer', 'server.js'), 'utf8');
    assert.ok(server.includes('/me/entitlement'));
    assert.ok(server.includes('/checkout/stripe'));
    assert.ok(server.includes('/checkout/cn'));
    assert.ok(!/consume_quota|usage_ledger|token.?meter/i.test(server));
});

test('LicenseBridge 与 ExcelBridge 暴露席位 API', () => {
    const bridge = fs.readFileSync(path.join(ROOT, 'Bridges', 'LicenseBridge.vb'), 'utf8');
    const excel = fs.readFileSync(path.join(ROOT, 'ExcelBridge.vb'), 'utf8');
    assert.ok(bridge.includes('IsSubscriptionActive'));
    assert.ok(bridge.includes('StartCheckout'));
    assert.ok(bridge.includes('PingHealth'));
    assert.ok(excel.includes('LicenseGetEntitlement'));
    assert.ok(excel.includes('LicenseStartCheckout'));
    assert.ok(excel.includes('LicensePingHealth'));
    assert.ok(!excel.includes('LicenseStartLocalServer'), '插件不能提供启动授权服务的入口');
    assert.ok(!bridge.includes('StartLocalTestServer'));
});

test('插件不内置授权服务启动入口', () => {
    const proj = fs.readFileSync(path.join(ROOT, 'ExcelAIVSTOAddIn.vbproj'), 'utf8');
    const startup = fs.readFileSync(path.join(ROOT, 'ThisAddIn.Startup.vb'), 'utf8');
    const shutdown = fs.readFileSync(path.join(ROOT, 'ThisAddIn.Shutdown.vb'), 'utf8');
    const licenseJs = fs.readFileSync(path.join(ROOT, 'UI', 'js', 'license.js'), 'utf8');
    assert.ok(!fs.existsSync(path.join(ROOT, 'Services', 'LicenseServerLauncher.vb')));
    assert.ok(!/LicenseServerLauncher/.test(proj + startup + shutdown));
    assert.ok(!/start-server|LicenseStartLocalServer|启动本地测试服务/.test(licenseJs));
    assert.ok(!/请先运行 LicenseServer/.test(licenseJs));
    const starter = fs.readFileSync(path.join(ROOT, 'LicenseServer', 'start-test.ps1'), 'utf8');
    assert.match(starter, /127\.0\.0\.1/);
    assert.match(starter, /9530/);
    assert.ok(!/ExcelBridge|LicenseStartLocalServer/.test(starter));
});

test('账户入口在主界面一级，并覆盖 OCR', () => {
    const index = fs.readFileSync(path.join(ROOT, 'UI', 'index.html'), 'utf8');
    const settings = fs.readFileSync(path.join(ROOT, 'UI', 'ai-settings-panel.html'), 'utf8');
    const ocrSettings = fs.readFileSync(path.join(ROOT, 'UI', 'ocr-settings-panel.html'), 'utf8');
    const ocrPanel = fs.readFileSync(path.join(ROOT, 'UI', 'ocr-panel.html'), 'utf8');
    const entry = fs.readFileSync(path.join(ROOT, 'UI', 'js', 'chat-v2', 'chat-entry.js'), 'utf8');
    const licenseJs = fs.readFileSync(path.join(ROOT, 'UI', 'js', 'license.js'), 'utf8');
    assert.ok(index.includes('openAccountPanel()') || index.includes('openAccountPanel'));
    assert.ok(index.includes('side-account-btn'));
    assert.ok(!index.includes('license-status-badge'), '不应再有第二个账户徽章');
    assert.ok(licenseJs.includes('openAccountPanel'));
    assert.ok(licenseJs.includes('LICENSE_SOFT_BYPASS'));
    assert.ok(licenseJs.includes('softBypassEnabled'));
    assert.ok(!licenseJs.includes('license server not running'), '已移除「授权服务未启动则放行」');
    assert.ok(!settings.includes('openAccountPanel') && !settings.includes('license-card'), 'AI 设置不再放登录入口');
    assert.ok(!ocrSettings.includes('license-card') && !ocrSettings.includes('license.js'), 'OCR 设置不再放登录表单');
    assert.ok(licenseJs.includes("getElementById('side-account-btn')"), '独立窗格不自动弹出第二套登录');
    assert.ok(ocrPanel.includes('ensureSubscriptionActive') && ocrPanel.includes('license.js'));
    assert.ok(entry.includes('ensureSubscriptionActive'));
    const sendFn = entry.slice(entry.indexOf('async function sendMessage'), entry.indexOf('function cancelCurrentRequest'));
    const uiAt = sendFn.indexOf('setProcessingUI(true)');
    const gateAt = sendFn.indexOf('ensureSubscriptionActive');
    assert.ok(uiAt >= 0 && gateAt > uiAt, '发送按钮应在席位校验之前切到停止态，避免授权 HTTP 期间点击像没反应');
    assert.ok(licenseJs.includes('GATE_TTL_MS'), '席位校验应复用刚拉过的权益，不要每次发送都打授权服务');
    assert.ok(licenseJs.includes('res.loggedIn === false'), '本机识别只在明确未登录时放行');
    const bridge = fs.readFileSync(path.join(ROOT, 'Bridges', 'LicenseBridge.vb'), 'utf8');
    const hostLic = fs.readFileSync(path.join(ROOT, 'Hosts', 'Shared', 'HostLicense.vb'), 'utf8');
    assert.ok(bridge.includes('TryCachedSubscriptionJson') && hostLic.includes('TryCachedSubscriptionJson'),
        '已有权益缓存时 IsSubscriptionActive 应立刻返回，授权 HTTP 放到后台');
    assert.ok(licenseJs.includes("code !== 'bridge_failed'"));
    const hostBridge = fs.readFileSync(path.join(ROOT, 'Hosts', 'Shared', 'HostBridgeBase.vb'), 'utf8');
    assert.ok(hostBridge.includes('LicenseIsSubscriptionActive'));
    assert.ok(hostBridge.includes('当前宿主没有手机遥控服务'));
});

test('法务写明订阅不含模型费', () => {
    const eula = fs.readFileSync(path.join(ROOT, 'Legal', 'EULA.md'), 'utf8');
    const privacy = fs.readFileSync(path.join(ROOT, 'Legal', 'PRIVACY.md'), 'utf8');
    assert.ok(eula.includes('不包含') && eula.includes('模型'));
    assert.ok(privacy.includes('软件席位') || privacy.includes('授权服务'));
    assert.ok(privacy.includes('API Key'));
});
