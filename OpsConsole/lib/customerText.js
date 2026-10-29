'use strict';
/**
 * customerText.js — 付款后发给客户的激活文案模板（运营台一键复制用）。
 * 纯函数，无 IO；改文案只动这里，测试在 OpsConsole/test/customerText.test.js。
 */

/**
 * @param {{key:string, plan?:string, expiresAt:string, contact?:string}} o
 * @returns {string} 多行文案（\n 分隔，直接粘微信）
 */
function customerMessage(o) {
    o = o || {};
    const end = String(o.expiresAt || '').slice(0, 10);
    const lines = [
        '【SmartPilot 智航 激活码】',
        '您的激活码：' + String(o.key || ''),
        '使用方法：打开 Excel → SmartPilot 面板 → 右上角「账户」→ 粘贴激活码 → 点击「激活」。',
        end ? ('有效期至 ' + end + '，激活后无需联网验证。') : '激活后无需联网验证。'
    ];
    if (o.contact) lines.push('问题反馈：' + o.contact);
    return lines.join('\n');
}

module.exports = { customerMessage };
