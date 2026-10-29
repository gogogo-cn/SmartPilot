/**
 * license-offline-activation.test.js — 路线①离线激活码护栏
 *
 * 覆盖：
 *   1) 算法冒烟（与 LicenseServer/test/license-key.test.js 同源，主套件独立可跑）；
 *   2) 三端常量一致（JS / Bridges/LicenseKey.vb / Hosts/Shared/HostLicense.vb）；
 *   3) Excel 桥与宿主桥暴露 Activate / GetLocal / ClearLocal 三个方法；
 *   4) 前端门禁「离线激活优先」且激活 UI 齐全；
 *   5) i18n 中英文案齐平；
 *   6) 改动文件语法可解析（node --check）。
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const UI_JS = path.join(ROOT, 'UI', 'js');
const lk = require(path.join(ROOT, 'LicenseServer', 'lib', 'licenseKey.js'));

const NEW_KEYS = [
    'license.key', 'license.key.hint', 'license.key.placeholder',
    'license.btnActivate', 'license.btnClearKey', 'license.local.active',
    'license.local.none', 'license.key.ok', 'license.key.cleared',
    'license.key.empty', 'license.ok.auth', 'license.ok.logout',
    // 购买 / 续费面板（路线①人工收款）
    'license.buy.title', 'license.buy.titleRenew', 'license.buy.price',
    'license.buy.planMonthly', 'license.buy.planYearly', 'license.buy.planTeam',
    'license.buy.planHint',
    'license.buy.step1', 'license.buy.step2', 'license.buy.step3',
    'license.buy.qr', 'license.buy.qrMissing', 'license.buy.contactLabel',
    'license.buy.copyContact', 'license.buy.copied', 'license.buy.copyFailed',
    'license.buy.contactEmpty',
    // 已激活用户的账户面板（2026-10-21：消除"未登录 + 离线激活"并存的迷惑 UI）
    'license.status.activeOffline', 'license.server.offline',
    // 已激活时激活码区切续费话术（2026-10-21：消除权益信息重复显示）
    'license.key.hintRenew',
    // 已激活时收起激活入口，改揭示按钮（2026-10-22）
    'license.btnRenewKey'
];

test('激活码算法冒烟：有效 / 篡改 / 过期', () => {
    const key = lk.createKey({ login: 'smoke@example.com', plan: 'pro_yearly', seats: 2, days: 365 });
    assert.match(key, /^SPK1-[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$/);
    const v = lk.verify(key);
    assert.strictEqual(v.valid, true);
    assert.strictEqual(v.seats, 2);

    const dot = key.lastIndexOf('.');
    const tampered = lk.PREFIX + 'B' + key.slice(lk.PREFIX.length + 1, dot) + '.' + key.slice(dot + 1);
    assert.strictEqual(lk.verify(tampered).reason, 'bad_signature');
    const expired = lk.createKey({ plan: 'pro_yearly', expUnix: Math.floor(Date.now() / 1000) - 5 });
    assert.strictEqual(lk.verify(expired).reason, 'expired');
    assert.strictEqual(lk.verify('not-a-key').reason, 'malformed');
});

test('三端 secret 常量一致（JS / LicenseKey.vb / HostLicense.vb）', () => {
    const keyVb = fs.readFileSync(path.join(ROOT, 'Bridges', 'LicenseKey.vb'), 'utf8');
    const hostVb = fs.readFileSync(path.join(ROOT, 'Hosts', 'Shared', 'HostLicense.vb'), 'utf8');
    const jsLib = fs.readFileSync(path.join(ROOT, 'LicenseServer', 'lib', 'licenseKey.js'), 'utf8');
    assert.ok(keyVb.includes('Private Const SECRET As String = "' + lk.DEFAULT_SECRET + '"'));
    assert.ok(hostVb.includes('Private Const KEY_SECRET As String = "' + lk.DEFAULT_SECRET + '"'));
    assert.ok(keyVb.includes('HMACSHA256') && hostVb.includes('HMACSHA256'), '两端都用 HMAC-SHA256');
    assert.ok(keyVb.includes('SPK1-') && hostVb.includes('SPK1-'), '两端都用 SPK1- 前缀');
    // payload 字段数（v1|login|plan|seats|expUnix|nonce = 6）三端必须一致，
    // 否则一端会把新格式判成 bad_payload（2026-10-20 nonce 修复后新增守栏）。
    assert.ok(jsLib.includes('FIELD_COUNT = 6'), 'JS lib 应为 6 段 payload');
    assert.ok(keyVb.includes('Private Const FIELD_COUNT As Integer = 6'), 'LicenseKey.vb 应为 6 段');
    assert.ok(hostVb.includes('Private Const KEY_FIELD_COUNT As Integer = 6'), 'HostLicense.vb 应为 6 段');
});

test('CLI 批量生成的 key 互不相同（nonce 防重复）', () => {
    const cli = path.join(ROOT, 'LicenseServer', 'tools', 'gen-license-key.js');
    const out = spawnSync(process.execPath, [
        cli, '--login', 'batch-guard@example.com', '--plan', 'pro_yearly',
        '--days', '365', '--batch', '5'
    ], { encoding: 'utf8' });
    assert.strictEqual(out.status, 0, out.stderr);
    const keys = out.stdout.split(/\r?\n/).map((s) => s.trim()).filter(Boolean);
    assert.strictEqual(keys.length, 5, '应输出 5 个 key');
    assert.strictEqual(new Set(keys).size, 5, '同参数批量生成的 5 个 key 必须互不相同');
    for (const k of keys) assert.strictEqual(lk.verify(k).valid, true);
});

test('Excel 桥暴露离线激活五方法，且激活存储统一到共享会话', () => {
    const bridge = fs.readFileSync(path.join(ROOT, 'Bridges', 'LicenseBridge.vb'), 'utf8');
    const excel = fs.readFileSync(path.join(ROOT, 'ExcelBridge.vb'), 'utf8');
    const proj = fs.readFileSync(path.join(ROOT, 'ExcelAIVSTOAddIn.vbproj'), 'utf8');
    for (const m of ['ActivateKey', 'GetLocalActivation', 'ClearLocalActivation', 'DeactivateForRefund', 'IsKeyDeactivated']) {
        assert.ok(bridge.includes('Public Function ' + m), 'LicenseBridge 应有 ' + m);
        assert.ok(excel.includes('Public Function License' + m), 'ExcelBridge 应暴露 License' + m);
    }
    assert.ok(proj.includes('<Compile Include="Bridges\\LicenseKey.vb" />'), '新文件必须登记进 vbproj');
    // 统一门禁：激活记录只能落在共享会话，Excel 不允许自建第二份存储
    assert.ok(bridge.includes('HostLicense.RememberActivation'), '激活记录必须写进共享会话');
    assert.ok(bridge.includes('HostLicense.ActivationIsAuthoritative'), '必须识别共享会话的权威位');
    assert.ok(!/cfg\.SetSetting\(ACTIVATION_SETTING,\s*record/.test(bridge),
        '不允许再把激活记录写回 ConfigurationManager——那会重新劈出双轨');
    // LICENSE_ACTIVATION 只应作为旧版迁移源出现
    assert.ok(bridge.includes('EnsureSharedActivation'), '应有一次性迁移入口');
});

test('宿主桥（Word/PPT）暴露离线激活五方法且记录落共享盘', () => {
    const hostVb = fs.readFileSync(path.join(ROOT, 'Hosts', 'Shared', 'HostLicense.vb'), 'utf8');
    const hostBase = fs.readFileSync(path.join(ROOT, 'Hosts', 'Shared', 'HostBridgeBase.vb'), 'utf8');
    for (const m of ['ActivateKey', 'GetLocalActivation', 'ClearLocalActivation', 'DeactivateForRefund', 'IsKeyDeactivated']) {
        assert.ok(hostVb.includes('Public Function ' + m), 'HostLicense 应有 ' + m);
        assert.ok(hostBase.includes('HostLicense.' + m), 'HostBridgeBase 应转发 ' + m);
    }
    assert.ok(hostVb.includes('activationProtected'), '激活记录须 DPAPI 保护后落盘');
    assert.ok(hostVb.includes('VerifyKey'), '宿主侧应有独立验签实现（不能引用 Excel 程序集）');
});

test('前端：离线激活优先于授权服务，且激活 UI 齐全', () => {
    const licenseJs = fs.readFileSync(path.join(UI_JS, 'license.js'), 'utf8');
    const gate = licenseJs.slice(licenseJs.indexOf('async function ensureSubscriptionActive'));
    const localIdx = gate.indexOf("call('LicenseGetLocalActivation')");
    const serverIdx = gate.indexOf('gateCache &&');
    assert.ok(localIdx >= 0, '门禁应先查本地离线激活');
    assert.ok(serverIdx >= 0 && localIdx < serverIdx, '离线激活判断必须先于授权服务缓存');
    assert.ok(licenseJs.includes("data-lic=\"activate\"") || licenseJs.includes("data-lic='activate'"), '应有激活按钮');
    assert.ok(licenseJs.includes('license-key-input'), '应有激活码输入框');
    assert.ok(licenseJs.includes('activateKey: activateKey'), '应导出 activateKey');
    assert.ok(licenseJs.includes('clearLocalActivation: clearLocalActivation'), '应导出 clearLocalActivation');
    assert.ok(licenseJs.includes('LicenseActivateKey'), '应调用桥的激活方法');
    assert.ok(!/start-server|LicenseStartLocalServer/.test(licenseJs), '插件不提供启动授权服务入口');
});

test('购买 / 续费面板：配置、UI 区块、复制按钮、标题切换齐全', () => {
    const licenseJs = fs.readFileSync(path.join(UI_JS, 'license.js'), 'utf8');
    // 配置项（开卖前改这里）
    assert.ok(/var PURCHASE = \{[\s\S]*?qrImage:[\s\S]*?contact:/.test(licenseJs),
        '应暴露 PURCHASE 配置（qrImage + contact）');
    // 购买区块挂进账户抽屉卡片
    assert.ok(licenseJs.includes('id="license-buy"'), '应有购买区块');
    assert.ok(licenseJs.includes('id="license-buy-qr"'), '应有收款码图片位');
    assert.ok(licenseJs.includes('id="license-buy-contact"'), '应有联系方式位');
    assert.ok(licenseJs.includes("data-lic=\"copy-contact\""), '应有复制联系方式按钮');
    assert.ok(licenseJs.includes('function copyPurchaseContact'), '应实现复制逻辑');
    assert.ok(licenseJs.includes('legacyCopy'), '应兼容 navigator.clipboard 不可用的场景');
    // 收款码加载失败要有文字兜底，不能露破图
    assert.ok(/addEventListener\('error'[\s\S]*?license-buy-qr-missing/.test(licenseJs),
        '收款码加载失败应显示 qrMissing 兜底文案');
    // 已激活用户标题切「续费」
    assert.ok(/license-buy-title[\s\S]*?titleRenew/.test(licenseJs) || licenseJs.includes("t(active ? 'license.buy.titleRenew'"),
        '已激活时标题应切到 license.buy.titleRenew');
    // 联系方式进 HTML 前必须转义（防配置里手滑写出的字符破坏结构）
    assert.ok(/esc\(PURCHASE\.contact\)/.test(licenseJs), 'contact 必须经 esc() 转义后插入');
    // 配置里指向的二维码图片必须真实存在（防改了路径忘放图，线上只露兜底文案）
    const qrPath = /qrImage:\s*'([^']*)'/.exec(licenseJs);
    assert.ok(qrPath && qrPath[1], 'PURCHASE.qrImage 应配置图片路径');
    if (qrPath && qrPath[1]) {
        assert.ok(fs.existsSync(path.join(ROOT, 'UI', qrPath[1])),
            '二维码图片缺失: UI/' + qrPath[1] + '（放置说明见 UI/images/README.txt）');
    }
});

test('i18n 中英文案齐平（新键都必须有且占位符一致）', () => {
    const i18n = fs.readFileSync(path.join(UI_JS, 'i18n.js'), 'utf8');
    for (const k of NEW_KEYS) {
        const zh = i18n.match(new RegExp("'" + k.replace('.', '\\.') + "':\\s*'([^']*)'"));
        assert.ok(zh, 'zh 词条缺失: ' + k);
        const en = i18n.match(new RegExp("'" + k.replace('.', '\\.') + "':\\s*'([^']*)'", 'g'));
        assert.ok(en && en.length >= 2, 'en 词条缺失: ' + k);
    }
    const placeholders = (s) => (s.match(/\{\w+\}/g) || []).sort().join(',');
    const zhActive = /'license\.local\.active':\s*'([^']*)'/.exec(i18n);
    const enMatches = [...i18n.matchAll(/'license\.local\.active':\s*'([^']*)'/g)];
    assert.strictEqual(placeholders(zhActive[1]), placeholders(enMatches[1][1]),
        'license.local.active 占位符中英必须一致');
});

test('改动文件语法可解析', () => {
    for (const f of ['license.js', 'i18n.js']) {
        const r = spawnSync(process.execPath, ['--check', path.join(UI_JS, f)], { encoding: 'utf8' });
        assert.strictEqual(r.status, 0, f + ' 语法错误:\n' + r.stderr);
    }
    for (const f of [
        path.join(ROOT, 'LicenseServer', 'lib', 'licenseKey.js'),
        path.join(ROOT, 'LicenseServer', 'tools', 'gen-license-key.js')
    ]) {
        const r = spawnSync(process.execPath, ['--check', f], { encoding: 'utf8' });
        assert.strictEqual(r.status, 0, path.basename(f) + ' 语法错误:\n' + r.stderr);
    }
});
