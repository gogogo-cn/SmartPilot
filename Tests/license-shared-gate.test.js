'use strict';
/**
 * license-shared-gate.test.js — 统一激活门禁守栏
 *
 * 架构不变量（2026-10-21 定型）：
 *   离线激活码（路线①）的唯一存储是 %APPDATA%\SmartPilotOffice\license_session.json。
 *   Excel / Word / PPT / 其它端共用同一份激活、同一份注销名单：
 *     - 任何一端激活 → 全端激活；
 *     - 任何一端注销退款 → 全端失效，且该 key 不能再在任何端激活。
 *   旧版 Excel 曾把激活记在 ConfigurationManager（双轨存储），现降级为「迁移源」。
 *
 * 这条守栏存在的原因：双轨存储不会报错，只会让两件事悄悄发生——
 *   「在 A 端激活、B 端照样要激活」（用户骂娘）；
 *   「在 A 端退款、B 端照常能用」（这是钱的问题）。
 */
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');

function read(rel) {
    return fs.readFileSync(path.join(ROOT, rel), 'utf8');
}

/** 取某个 VB 函数体（从签名到最近的 End Function），用于逐方法断言。 */
function fnBody(src, signature) {
    const i = src.indexOf(signature);
    assert.ok(i >= 0, '找不到函数: ' + signature);
    const end = src.indexOf('End Function', i);
    assert.ok(end > i, '函数体没有闭合: ' + signature);
    return src.slice(i, end);
}

test('Excel 侧五个离线方法全部经共享会话，不自建第二份存储', () => {
    const bridge = read('Bridges/LicenseBridge.vb');
    const sigs = [
        'Public Function ActivateKey(bodyJson As String) As String',
        'Public Function GetLocalActivation() As String',
        'Public Function ClearLocalActivation() As String',
        'Public Function DeactivateForRefund() As String',
        'Public Function IsKeyDeactivated(bodyJson As String) As String'
    ];
    for (const sig of sigs) {
        const body = fnBody(bridge, sig);
        assert.ok(body.includes('HostLicense.'), sig + ' 必须委托 HostLicense 共享会话');
        assert.ok(body.includes('EnsureSharedActivation'), sig + ' 必须先确保共享会话已接管（含一次性迁移）');
    }
    assert.ok(!/cfg\.SetSetting\(ACTIVATION_SETTING,\s*record/.test(bridge),
        '不允许把激活记录写回 ConfigurationManager——那会重新劈出双轨');
});

test('ExcelBridge 暴露全部五个 License* 离线方法（前端同一套调用）', () => {
    const excel = read('ExcelBridge.vb');
    for (const m of ['LicenseActivateKey', 'LicenseGetLocalActivation', 'LicenseClearLocalActivation',
                     'LicenseDeactivateForRefund', 'LicenseIsKeyDeactivated']) {
        assert.ok(excel.includes('Public Function ' + m), 'ExcelBridge 缺 ' + m);
    }
});

test('宿主桥同样暴露五个方法并转发 HostLicense', () => {
    const base = read('Hosts/Shared/HostBridgeBase.vb');
    const short = ['ActivateKey', 'GetLocalActivation', 'ClearLocalActivation',
                   'DeactivateForRefund', 'IsKeyDeactivated'];
    for (const m of short) {
        assert.ok(base.includes('Function License' + m), 'HostBridgeBase 缺 License' + m);
        assert.ok(base.includes('HostLicense.' + m), 'HostBridgeBase 应转发 ' + m);
    }
});

test('共享存储层提供迁移与跨端注销名单入口', () => {
    const host = read('Hosts/Shared/HostLicense.vb');
    assert.ok(host.includes('Public Sub RememberActivation'), 'RememberActivation 应公开（Excel 迁移写入用）');
    assert.ok(host.includes('Public Function ActivationIsAuthoritative'), '应有权威位判断');
    assert.ok(host.includes('Public Sub MergeDeactivatedList'), '应有注销名单合并（旧 Excel 名单并进来）');
    assert.ok(host.includes('Public Function IsKeyInDeactivatedList'), '应有跨端注销查询');
});

test('共享会话路径只有一份定义（HostLicense.SessionPath），Excel 侧不重复造路径', () => {
    const bridge = read('Bridges/LicenseBridge.vb');
    // 注释里可以提文件名（文档），但不允许出现路径字面量定义或第二份 SessionPath
    assert.ok(!bridge.includes('"license_session.json"'), 'LicenseBridge 不得自建共享文件路径字面量');
    assert.ok(!bridge.includes('SessionPath'), 'LicenseBridge 不得定义自己的 SessionPath');
    const host = read('Hosts/Shared/HostLicense.vb');
    assert.ok(host.includes('SmartPilotOffice'), 'HostLicense 应定义共享会话路径');
});

test('共享会话读写带重试（三端 Office 同时跑，文件锁冲突是常态）', () => {
    const host = read('Hosts/Shared/HostLicense.vb');
    assert.ok(/For attempt As Integer = 1 To SHARED_FILE_RETRY_TIMES/.test(host), '读写都应有重试循环');
    assert.ok(host.includes('Catch ex As JsonException'), '读到对方写一半的内容要重试，不能降级成空会话');
    assert.ok(host.includes('Catch ex As IOException'), 'sharing violation 要重试');
});

test('旧 Excel 私有存储降级为迁移源，且迁移后清理', () => {
    const bridge = read('Bridges/LicenseBridge.vb');
    assert.ok(bridge.includes('EnsureSharedActivation'), '应有迁移入口');
    assert.ok(bridge.includes('MergeDeactivatedList(legacy)'), '旧注销名单必须并进共享名单');
    assert.ok(bridge.includes('ForgetLegacyActivation'), '清除/注销后应清理旧存储，防止幽灵数据');
    assert.ok(bridge.includes('GetSetting(ACTIVATION_SETTING'), '迁移应读取旧 LICENSE_ACTIVATION');
});

test('前端门禁调用面三端一致（同一份 license.js，方法名两端都已暴露）', () => {
    const js = read('UI/js/license.js');
    for (const m of ['LicenseGetLocalActivation', 'LicenseActivateKey',
                     'LicenseClearLocalActivation', 'LicenseDeactivateForRefund']) {
        assert.ok(js.includes(m), 'license.js 应调用 ' + m);
    }
});

test('已激活用户的账户面板不再显示"未登录"式迷惑 UI', () => {
    const js = read('UI/js/license.js');
    // 离线激活生效时，主状态必须切到激活码权益文案（不能停留在"未登录"）
    assert.ok(/if \(statusEl\) statusEl\.textContent = formatEntitlement\(local\.entitlement\);/.test(js),
        '离线激活生效时主状态应显示激活码权益');
    assert.ok(js.includes("ent.source === 'offline'"), 'formatEntitlement 应识别离线来源');
    // 登录表单、试用说明对已激活用户隐藏
    assert.ok(/if \(authBox\) authBox\.style\.display = 'none';/.test(js),
        '离线激活生效时应隐藏登录表单');
    assert.ok(/trialHint\.style\.display = 'none';/.test(js),
        '已激活用户不应再看到"注册试用 14 天"说明');
    // 授权服务连不上时，对已激活用户用中性文案而非红色警告
    assert.ok(js.includes('offlineActivationActive'), '应跟踪离线激活状态');
    assert.ok(/offlineActivationActive\)\s*\{\s*\n\s*el\.textContent = t\('license\.server\.offline'\)/.test(js),
        '授权服务未连接时已激活用户应看到中性文案');
    assert.ok(js.includes("license.server.offline"), 'license.js 内置字典应备份 server.offline 文案');
    // 已激活时激活码区不得再重复一遍权益（主状态行已经说了）
    assert.ok(/el\.style\.display = 'none';[\s\S]*?已激活/.test(js) || /if \(active\) \{[\s\S]*?el\.style\.display = 'none';/.test(js),
        '离线激活生效时 license-local-status 必须隐藏，不能和主状态重复');
    assert.ok(js.includes('license-key-hint-renew'), '激活码区应有续费话术提示位');
    assert.ok(/hint\) hint\.style\.display = active \? 'none' : '';/.test(js),
        '未激活提示与续费提示应按激活状态互斥切换');
    // 已激活用户不该看到激活入口（输入框 + 激活按钮），只能经揭示按钮展开续费用
    assert.ok(/if \(inputRow\) inputRow\.style\.display = active \? 'none' : '';/.test(js),
        '已激活时必须收起激活码输入行');
    assert.ok(/if \(activateBtn\) activateBtn\.style\.display = active \? 'none' : '';/.test(js),
        '已激活时必须收起「激活」按钮');
    assert.ok(js.includes('license-renew-btn'), '应提供「输入新激活码」揭示按钮');
    assert.ok(js.includes('function revealKeyInput'), '应实现揭示逻辑');
    assert.ok(js.includes("act === 'renew-key'"), '揭示按钮应有点击绑定');
    const i18n = read('UI/js/i18n.js');
    for (const k of ['license.status.activeOffline', 'license.server.offline']) {
        const hits = [...i18n.matchAll(new RegExp("'" + k.replace('.', '\\.') + "':\\s*'([^']*)'", 'g'))];
        assert.ok(hits.length >= 2, 'i18n 中英文都应配置 ' + k);
    }
});
