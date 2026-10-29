/**
 * ============================================================
 * Tests/skill-routing-domain-guard.test.js — 技能路由判定不得吃掉确定性硬短路
 * ============================================================
 * 背景（实测 2026-10-06 01:10，Word 宿主）：
 *   host_bridge.log 记录 `McpDispatcher 被跳过 … skillDomain=true`，
 *   于是"将当前文档导出为 PDF"根本没进 dispatcher → 硬短路全失效 → 落到 agent-loop，
 *   模型零工具调用却声称"已导出"。
 * 根因：skill-domain.js 的 isSkillRoutingMessage 遍历**全部** ROUTING_POLICIES，
 *   而 SmartTableBeautifier / SmartFormulaAssistant / DashboardGenerator 只声明了
 *   block / boostMinScore、没有 intent → 走 _testPolicyIntent 的通用门
 *   （hasGenericVerb || hasStrongDomainIntent）。GENERIC_CREATE_VERBS 含"导出/生成/保存"
 *   这类普通动词 → 任意正常命令都被判成"技能路由消息"。
 * 修复：只认显式声明了 intent 的策略（SkillCreator / GenerateOCRParseScript）。
 * 运行：node Tests/skill-routing-domain-guard.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

const { App } = loadChatV2(['config/skill-domain.js']);
const sd = App.ChatV2.skillDomain;

test('导出/另存类请求不得被判为技能路由消息（否则 dispatcher 被跳过）', () => {
    for (const msg of ['将当前文档导出为 PDF。', '帮我把这份文档另存为 PDF', '把演示文稿导出为图片']) {
        assert.strictEqual(sd.isSkillRoutingMessage(msg), false,
            `"${msg}" 必须放行给 McpDispatcher，不能被 skill 路由守卫截走`);
    }
});

test('真正声明了 intent 的策略仍生效（SkillCreator / 解析脚本生成器）', () => {
    assert.strictEqual(sd.isSkillRoutingMessage('帮我创建一个 skill'), true);
    assert.strictEqual(sd.isSkillRoutingMessage('生成一个解析器'), true);
});

test('确定性域关键词仍照旧命中（美化/格式化走 skill-pipeline）', () => {
    assert.strictEqual(sd.isDomainMessage('美化这个表格'), true);
    assert.strictEqual(sd.isDomainMessage('把表头加粗'), true);
    assert.strictEqual(sd.isSkillRoutingMessage('美化这个表格'), true);
});

test('域关键词必须来自 DESCRIPTORS，普通动词不算域命中', () => {
    assert.deepStrictEqual(sd.matchDomains('将当前文档导出为 PDF。'), []);
});
