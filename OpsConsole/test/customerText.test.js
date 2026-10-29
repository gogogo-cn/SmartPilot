'use strict';
/**
 * customerText.test.js — 客户文案模板
 */
const { test } = require('node:test');
const assert = require('node:assert');
const { customerMessage } = require('../lib/customerText');

test('客户文案包含 key、到期日、激活步骤', () => {
    const text = customerMessage({
        key: 'SPK1-abcdef.ghijkl',
        plan: 'pro_yearly',
        expiresAt: '2027-10-20T00:00:00.000Z',
        contact: '微信：hellokitty360730'
    });
    assert.ok(text.includes('SPK1-abcdef.ghijkl'));
    assert.ok(text.includes('2027-10-20'));
    assert.ok(text.includes('账户'));
    assert.ok(text.includes('激活'));
    assert.ok(text.includes('微信：hellokitty360730'), '有 contact 时应带上');
    assert.ok(!text.includes('undefined'));
    assert.ok(!text.includes('null'));
});

test('无 contact / 无到期日时不出现空行或占位符', () => {
    const text = customerMessage({ key: 'SPK1-a.b', expiresAt: '' });
    assert.ok(text.includes('SPK1-a.b'));
    assert.ok(!text.includes('问题反馈'));
    assert.ok(!text.includes('有效期至 '));
    assert.ok(!text.includes('undefined'));
});
