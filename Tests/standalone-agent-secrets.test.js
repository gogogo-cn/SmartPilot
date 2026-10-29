/**
 * 独立服务 9529：CORS 不得对任意网页放开，/model-config 不得回传明文 API 密钥。
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const SRC = path.join(__dirname, '..', 'StandaloneAgent', 'standalone_agent.py');

function code() {
    const raw = fs.readFileSync(SRC, 'utf8');
    return raw.split(/\r?\n/).filter(function (line) {
        return !/^\s*#/.test(line);
    }).join('\n');
}

test('独立服务不得写死 Access-Control-Allow-Origin: *', function () {
    const s = code();
    assert.ok(!/Access-Control-Allow-Origin["']?\s*,\s*["']\*["']/.test(s),
        '不可再把 ACAO 写成字面量 *');
    assert.ok(s.includes('def _cors_allow_origin'), '应有来源白名单函数');
    assert.ok(s.includes('app.excelai.local'), '白名单应包含宿主虚拟主机');
});

test('/model-config 与保存响应使用脱敏配置', function () {
    const s = code();
    assert.ok(s.includes('def _public_llm_config'), '应有脱敏函数');
    const idx = s.indexOf('path == "/model-config"');
    assert.ok(idx > -1, '应有 /model-config');
    const slice = s.slice(idx, idx + 400);
    assert.ok(slice.includes('_public_llm_config'), '/model-config 应返回脱敏配置');
    assert.ok(!/"data":\s*cfg\b/.test(slice), '/model-config 不得回传原始 cfg');
    assert.ok(!/"data":\s*payload\b/.test(s), '保存响应不得回传含密钥的 payload');
});

test('法务三件套存在', function () {
    for (const name of ['EULA.md', 'PRIVACY.md', 'THIRD_PARTY_NOTICES.md']) {
        const p = path.join(__dirname, '..', 'Legal', name);
        assert.ok(fs.existsSync(p), '缺少 Legal/' + name);
        const text = fs.readFileSync(p, 'utf8');
        assert.ok(text.length > 200, name + ' 过短');
    }
    const privacy = fs.readFileSync(path.join(__dirname, '..', 'Legal', 'PRIVACY.md'), 'utf8');
    assert.ok(privacy.includes('API 地址'), '隐私说明应写明数据发往用户配置的 API 地址');
});
