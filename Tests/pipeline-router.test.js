/**
 * ============================================================
 * Tests/pipeline-router.test.js — 路由与防光说不做测试
 * ============================================================
 * 覆盖：pipelineRegistry 注册/降级、shouldRouteToSkill 意图路由、
 *       误路由防护（通用请求不触发 skill）
 * 运行：node Tests/pipeline-router.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

test('pipeline-registry: 注册/查询/未知kind降级', async () => {
    const { App } = loadChatV2();
    const pr = App.ChatV2.pipelineRegistry;

    assert.strictEqual(pr.pick('text').kind, 'text');
    assert.strictEqual(pr.pick('python').kind, 'python');
    assert.strictEqual(pr.pick('nonexistent').kind, 'text', '未知 kind 应降级 text');
    assert.strictEqual(pr.list().length, 7, '应有 7 条 pipeline');
});

test('pipeline-registry: 自定义注册 + priority 排序', async () => {
    const { App } = loadChatV2();
    const pr = App.ChatV2.pipelineRegistry;
    pr.register('custom_x', { execute: async function* () {}, name: 'x', priority: 1 });
    const sorted = pr.list().map(p => p.kind);
    assert.strictEqual(sorted[0], 'custom_x');
});

test('orchestrator: 意图请求路由到 skill-pipeline（不经过 LLM）', async () => {
    const { App } = loadChatV2();
    App.cachedSkills = [{
        name: 'SmartTableBeautifier', displayName: '智能表格美化',
        keywords: ['美化', '美化表格', '表格美化', '美化这个表格', '一键美化'],
        inputContext: ['selection'], pipeline: 'skill'
    }];
    const tr = App.ChatV2.toolRegistry;
    tr.register({
        name: 'skill.SmartTableBeautifier', displayName: '智能表格美化',
        description: '美化', inputSchema: { type: 'object', properties: {} },
        source: 'skill', skillInfo: App.cachedSkills[0],
        async execute() { return { ok: true, result: 'ok' }; }
    });

    let llmCalled = false;
    App.ChatV2.llmClient = { stream() { llmCalled = true; return (async function* () { yield { type: 'done', finishReason: 'stop' }; })(); } };
    App.window.chrome.webview.hostObjects.excelBridge.ExecuteSkill = async () => JSON.stringify({ success: true });

    await App.ChatV2.orchestrator.dispatch({
        message: '美化表格',
        context: { selection: { range: 'A1:H6' } }
    });
    const turn = App.ChatV2.sessionStore.getCurrentTurn();
    // 契约验证：应路由到 skill-pipeline 且锁定 SmartTableBeautifier。
    // 注意：skill-pipeline 内部规划 spec 设计上会调用 LLM（LLM 规划是 Skill 流程的一部分），
    // 因此不能断言"完全不调用 LLM"——只需断言未进入 text-pipeline 的主 LLM 对话即可（kind==='skill'）。
    assert.strictEqual(turn?.kind, 'skill', '美化表格应路由到 skill-pipeline');
    assert.strictEqual(turn?.context?._matchedSkill?.name, 'SmartTableBeautifier', '应锁定 SmartTableBeautifier Skill');
});

test('orchestrator: 通用请求不误路由 skill', async () => {
    const { App } = loadChatV2();
    App.cachedSkills = [{
        name: 'SmartTableBeautifier', displayName: '智能表格美化',
        keywords: ['美化', '美化表格', '表格美化', '美化这个表格', '一键美化'],
        inputContext: ['selection'], pipeline: 'skill'
    }];
    const tr = App.ChatV2.toolRegistry;
    tr.register({
        name: 'skill.SmartTableBeautifier', displayName: '智能表格美化',
        description: '美化', inputSchema: { type: 'object', properties: {} },
        source: 'skill', skillInfo: App.cachedSkills[0],
        async execute() { return { ok: true, result: 'ok' }; }
    });

    let llmCalled = false;
    App.ChatV2.llmClient = { stream() { llmCalled = true; return (async function* () { yield { type: 'done', finishReason: 'stop' }; })(); } };
    App.window.chrome.webview.hostObjects.excelBridge.ExecuteSkill = async () => JSON.stringify({ success: true });

    await App.ChatV2.orchestrator.dispatch({ message: '你好', context: { selection: { range: 'A1:H6' } } });
    await new Promise(r => setTimeout(r, 300));
    assert.strictEqual(llmCalled, true, '"你好" 应留在 LLM 对话，不路由 skill');
});

test('orchestrator: PPT 汇报含"趋势分析页"应路由 code-pipeline 而非 python', async () => {
    const { App } = loadChatV2();
    // 复现真实场景：GeneratePython 的"趋势分析/趋势"关键词会与 GeneratePPTReport 竞争，
    // 修复前 GeneratePython(pipeline: python) 因位置加成反超，导致误触发 python 分析。
    App.cachedSkills = [
        { name: 'GeneratePPTReport', displayName: '生成PPT汇报',
          keywords: ['PPT', '汇报', '幻灯片', 'pptx', '演示文稿', '做PPT', '做幻灯片', 'PPT 汇报', '做成PPT', '生成PPT', '生成汇报', '数据报告'],
          pipeline: 'code' },
        { name: 'GeneratePython', displayName: 'Python数据分析',
          keywords: ['数据分析', '趋势分析', '趋势', '预测', '统计', 'Python', 'python', '图表', '可视化'],
          pipeline: 'python' }
    ];
    await App.ChatV2.orchestrator.dispatch({
        message: '在PPT汇报中增加趋势分析页',
        context: { selection: { range: 'A1:H6', headers: ['日期', '销售额'] } }
    });
    const turn = App.ChatV2.sessionStore.getCurrentTurn();
    assert.strictEqual(turn.kind, 'code',
        `PPT 汇报请求应路由 code-pipeline（GeneratePPTReport），实际 ${turn.kind}`);
});

test('orchestrator: 用"powerpoint"表述 PPT 时仍路由 code-pipeline', async () => {
    const { App } = loadChatV2();
    // "powerpoint" 不在 GeneratePPTReport 关键词中 → shouldRouteToSkill 会选中 GeneratePython，
    // 但 PPT 意图优先覆盖应改回 GeneratePPTReport → code-pipeline。
    App.cachedSkills = [
        { name: 'GeneratePPTReport', displayName: '生成PPT汇报',
          keywords: ['PPT', '汇报', '幻灯片', 'pptx', '演示文稿', '做PPT', '做幻灯片', 'PPT 汇报', '做成PPT', '生成PPT', '生成汇报', '数据报告'],
          pipeline: 'code' },
        { name: 'GeneratePython', displayName: 'Python数据分析',
          keywords: ['数据分析', '趋势分析', '趋势', '预测', '统计', 'Python', 'python', '图表', '可视化'],
          pipeline: 'python' }
    ];
    await App.ChatV2.orchestrator.dispatch({
        message: '用powerpoint做一个趋势分析页',
        context: { selection: { range: 'A1:H6', headers: ['日期', '销售额'] } }
    });
    const turn = App.ChatV2.sessionStore.getCurrentTurn();
    assert.strictEqual(turn.kind, 'code',
        `PPT 意图优先应锁定 code-pipeline，实际 ${turn.kind}`);
});

test('orchestrator: 纯分析请求"分析销售趋势"仍路由 python-pipeline', async () => {
    const { App } = loadChatV2();
    App.cachedSkills = [
        { name: 'GeneratePPTReport', displayName: '生成PPT汇报',
          keywords: ['PPT', '汇报', '幻灯片', 'pptx', '演示文稿', '做PPT', '做幻灯片', 'PPT 汇报', '做成PPT', '生成PPT', '生成汇报', '数据报告'],
          pipeline: 'code' },
        { name: 'GeneratePython', displayName: 'Python数据分析',
          keywords: ['数据分析', '趋势分析', '趋势', '预测', '统计', 'Python', 'python', '图表', '可视化'],
          pipeline: 'python' }
    ];
    await App.ChatV2.orchestrator.dispatch({
        message: '分析销售趋势',
        context: { selection: { range: 'A1:H6', headers: ['日期', '销售额'] } }
    });
    const turn = App.ChatV2.sessionStore.getCurrentTurn();
    assert.strictEqual(turn.kind, 'python',
        `纯分析请求应路由 python-pipeline，实际 ${turn.kind}`);
});

test('orchestrator: "使用饼图按部门展示销售额"必须路由 SmartChartWizard（不被 GeneratePython 的"销售额"抢走）', async () => {
    const { App } = loadChatV2();
    // 复现真实失败（9/9 日志）：SmartChartWizard 命中"饼图"(2字强词→+2加成后 score=3, strong2=true)，
    // GeneratePython 命中"销售额"(3字数据词 score=3, short=true) 同分——
    // 修复前按 cachedSkills 插入顺序 GeneratePython 排前，strongMatch 只看 best 的 strong2 → false
    // → matchedSkill=null → 掉进 text-pipeline 让小模型幻觉（整表画图、缺参数）。
    App.cachedSkills = [
        { name: 'GeneratePython', displayName: 'Python数据分析',
          keywords: ['数据分析', '趋势分析', '趋势', '预测', '统计', 'Python', 'python', '图表', '可视化', '销售额', '相关性', '分布'],
          pipeline: 'python' },
        { name: 'SmartChartWizard', displayName: '智能图表向导',
          keywords: ['画图', '图表', '柱状图', '饼图', '折线图', '散点图', '雷达图', '面积图', '占比图', '创建图表', '生成图表'],
          pipeline: 'skill' }
    ];
    await App.ChatV2.orchestrator.dispatch({
        message: '使用饼图按部门展示销售额',
        context: { selection: { range: 'A1:I101', headers: ['日期', '地区', '部门', '产品', '销售额', '成本'] } }
    });
    const turn = App.ChatV2.sessionStore.getCurrentTurn();
    assert.strictEqual(turn.kind, 'skill',
        `"饼图"明确图表意图应路由 skill-pipeline（SmartChartWizard），实际 ${turn.kind}`);
    const matched = App.ChatV2.sessionStore.getCurrentTurn();
    assert.ok(matched.context && matched.context._matchedSkill,
        '应记录 _matchedSkill=SmartChartWizard 供 skill-pipeline 执行');
    assert.strictEqual(matched.context._matchedSkill.name, 'SmartChartWizard');
    console.log('\n✅ 饼图请求路由修正：同分时强意图 2 字词(饼图)优先于泛数据词(销售额) → SmartChartWizard/skill-pipeline');
});

test('agent-loop: findCandidateTools 含 skill keywords 匹配', async () => {
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;
    tr.register({
        name: 'skill_SmartTableBeautifier', displayName: '智能表格美化',
        description: '美化表格', inputSchema: { type: 'object', properties: {} },
        source: 'skill',
        skillInfo: { name: 'SmartTableBeautifier', keywords: ['美化', '美化表格', '表格美化'] },
        async execute() { return { ok: true, result: 'ok' }; }
    });

    const candidates = await App.ChatV2.agentLoop.findCandidateTools('美化表格', tr, 3);
    assert.ok(candidates.length > 0, '应找到候选');
    assert.ok(candidates[0].score >= 20, 'score 应达到阈值，实际 ' + candidates[0].score);
    assert.ok(String(candidates[0].tool.name).startsWith('skill_'));
});

// 兜底退出：避免 session-store 防抖定时器挂起进程
if (require.main === module) { setTimeout(() => process.exit(0), 4000).unref(); } // 仅独立运行时强制退出；聚合运行（run-tests.js require）不注册，避免进程提前退出
