/**
 * ============================================================
 * Tests/skill-spec-e2e.test.js — Skill Spec 端到端集成测试
 * ============================================================
 * 验证：Skill 声明 processingMode=SkillSpec 时，skill-pipeline 走
 *       "LLM 规划 spec → 引擎渲染 VBA → excel_execute_vba 执行"
 * 场景：用户说"美化表格，要科技感的绿色风格，销售额加数据条，冻结表头"
 * 运行：node Tests/skill-spec-e2e.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2, DEFAULT_ORDER, stubDangerConfirm } = require('./helpers/harness');

// 与真实 index.html 一致的加载顺序，但截断到 chat-orchestrator（不加载 chat-entry/index，避免入口副作用）
const ORDER = DEFAULT_ORDER.slice(0, DEFAULT_ORDER.indexOf('ui/chat-renderer.js'));

test('skill-spec e2e: 美化表格带样式要求 → LLM 规划 → 确定性 VBA 执行', async () => {
    const App = loadChatV2(ORDER).App;
    // excel_execute_vba 是高危工具：所有思考模式（含下面的 expert）都会被【安全守卫】
    // 拦下并弹确认框。本用例验证的是"spec → 确定性 VBA 渲染"，故模拟用户点「仍要继续执行」。
    stubDangerConfirm(true);

    // 1. 注册 SkillSpec 模式的 skill（模拟 SmartTableBeautifier 声明）
    App.cachedSkills = [{
        name: 'SmartTableBeautifier', displayName: '智能表格美化',
        keywords: ['美化', '美化表格', '一键美化'],
        inputContext: ['selection'], pipeline: 'skill',
        processingMode: 'SkillSpec'
    }];
    const tr = App.ChatV2.toolRegistry;
    tr.register({
        name: 'skill.SmartTableBeautifier', displayName: '智能表格美化', description: '美化',
        inputSchema: { type: 'object', properties: {} }, source: 'skill', skillInfo: App.cachedSkills[0],
        async execute() { return { ok: true, result: '传统模式' }; }
    });

    // 2. 捕获 excel_execute_vba 收到的代码
    let capturedVba = null;
    tr.register({
        name: 'excel_execute_vba', displayName: '执行 VBA', description: '执行 VBA 代码',
        inputSchema: { type: 'object', properties: { code: { type: 'string' } } }, source: 'mcp',
        async execute(args) { capturedVba = args?.code || ''; return { ok: true, result: { output: '执行成功' } }; }
    });

    // 3. stub LLM：规划 spec（科技绿 + 数据条 + 冻结表头）
    let llmCalls = 0;
    App.ChatV2.llmClient = {
        stream({ messages }) {
            return (async function* () {
                llmCalls++;
                const isPlanning = messages.some(m => m.role === 'system' && String(m.content).includes('表格美化规划师'));
                if (isPlanning) {
                    const spec = {
                        theme: 'tech_green', headerStyle: 'bold_colored', alignment: 'center',
                        autofit: true, freezeHeader: true, addBorders: true,
                        bandedRows: false, conditionalFormat: 'dataBar', numberFormat: null, title: '销售报表'
                    };
                    yield { type: 'token', delta: JSON.stringify(spec) };
                } else {
                    yield { type: 'token', delta: '{}' };
                }
                yield { type: 'done', finishReason: 'stop' };
            })();
        }
    };

    // 4. 模拟 turn（expert 模式放行 VBA 守卫）
    const sessionStore = App.ChatV2.sessionStore;
    const turn = sessionStore.startTurn({
        message: '美化表格，要科技感绿色风格，销售额加数据条，冻结表头',
        kind: 'text',
        context: { selection: { address: 'A1:F30', sheetName: 'Sheet1' }, _matchedSkill: App.cachedSkills[0] }
    });
    turn.mode = 'expert';

    // 5. 运行 skill-pipeline
    const pipeline = App.ChatV2.pipelines.skill;
    const gen = pipeline.execute(turn, new AbortController().signal);
    const events = [];
    for await (const evt of gen) {
        events.push(evt.type);
        if (evt.type === 'status') console.log('  [status]', evt.text);
    }

    // 6. 断言
    assert.ok(capturedVba, '应通过 excel_execute_vba 执行渲染的 VBA');
    console.log('\n=== 执行的 VBA（关键片段）===');
    console.log(capturedVba.split('\n').filter(l => /RGB|Freeze|Databar|DataBar|Sub |End Sub|MOD\(ROW/.test(l)).join('\n'));

    // 主题：科技绿表头
    assert.ok(capturedVba.includes('RGB(0, 128, 128)'), '应应用科技绿主题');
    // 冻结表头
    assert.ok(capturedVba.includes('FreezePanes'), '应冻结表头');
    // 数据条
    assert.ok(/AddDatabar/i.test(capturedVba), '应加数据条');
    // 标题行
    assert.ok(capturedVba.includes('销售报表'), '应插入标题行');
    // LLM 规划被调用
    assert.ok(llmCalls >= 1, '应调用 LLM 规划');
    // pipeline 完成
    assert.ok(events.includes('pipeline_done'), '应有 pipeline_done');
    console.log('\n✅ 端到端：LLM 规划 spec → 确定性 VBA 执行成功');
    console.log('✅ 事件流:', events.join(', '));
});

if (require.main === module) { setTimeout(() => process.exit(0), 5000).unref(); } // 仅独立运行时强制退出；聚合运行（run-tests.js require）不注册，避免进程提前退出
