/**
 * 全量 skill 路由验证（真实 Skill/*.md 定义）
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const { loadChatV2 } = require('./helpers/harness');

/** 解析 Skill/*.md 的 keywords */
function parseSkillKeywords(mdPath) {
    const lines = fs.readFileSync(mdPath, 'utf8').split(/\r?\n/);
    const kwStart = lines.findIndex(l => /^keywords:/.test(l));
    if (kwStart < 0) return [];
    const kws = [];
    for (let i = kwStart + 1; i < lines.length; i++) {
        const m = lines[i].match(/^\s*-\s*(.+)$/);
        if (!m) break;
        if (m[1].trim().startsWith('name:')) break;
        if (m[1].includes('反例')) continue;
        kws.push(m[1].trim());
    }
    return kws;
}

/** 解析 Skill/*.md 的 pipeline / autoExecute（与后端 SkillManager.Loading 对齐） */
function parseSkillMeta(mdPath) {
    const content = fs.readFileSync(mdPath, 'utf8');
    const fm = content.match(/^---\r?\n([\s\S]*?)\r?\n---/);
    if (!fm) return {};
    const yaml = fm[1];
    const plMatch = yaml.match(/^pipeline:\s*(.+)$/m);
    const aeMatch = yaml.match(/^autoExecute:\s*(.+)$/m);
    return {
        pipeline: plMatch ? plMatch[1].trim() : '',
        autoExecute: aeMatch ? aeMatch[1].trim().toLowerCase() === 'true' : false
    };
}

function loadAllSkills() {
    const dir = path.join(__dirname, '..', 'Skills');
    const skills = [];
    for (const f of fs.readdirSync(dir).filter(n => n.endsWith('.md') && n !== 'ThinkingFramework.md')) {
        skills.push({ name: f.replace('.md', ''), keywords: parseSkillKeywords(path.join(dir, f)), ...parseSkillMeta(path.join(dir, f)) });
    }
    return skills;
}

test('skill 路由: 意图请求命中对应 skill（经 orchestrator 后端执行）', async () => {
    const { App } = loadChatV2();
    App.cachedSkills = loadAllSkills();

    const tr = App.ChatV2.toolRegistry;
    for (const s of App.cachedSkills) {
        tr.register({
            name: 'skill.' + s.name, displayName: s.name, description: s.name,
            inputSchema: { type: 'object', properties: {} }, source: 'skill', skillInfo: s,
            async execute() { return { ok: true, result: 'ok' }; }
        });
    }
    // 阻止真正调 LLM（路由失败会走到这里）
    App.ChatV2.llmClient = { stream() { throw new Error('不应调用 LLM'); } };
    window.chrome.webview.hostObjects.excelBridge.ExecuteSkill = async () => JSON.stringify({ success: true });

    const orch = App.ChatV2.orchestrator;
    // 注意：**每次 dispatch 必须传入全新的 context 对象**。
    // orchestrator 会把路由结论写回 context（context._matchedSkill = matchedSkill），
    // 若复用同一个对象，下一轮 dispatch 进来时 context 里已带着上一轮的技能，
    // 后续继承/守卫逻辑会直接沿用 → 表现为"每条消息都路由到上一条的技能"
    // （实测：总结这段内容 → GeneratePPTReport，正是上一例"生成PPT汇报"的技能）。
    // 这会让本用例退化成测"继承"而非测"意图路由"，必须隔离。
    const makeCtx = () => ({ selection: { range: 'A1:H6' } });

    // [消息, 期望 skill 名] —— 走 skill-pipeline 的场景
    const skillCases = [
        ['美化表格', 'SmartTableBeautifier'],
        ['清洗数据', 'SmartDataCleaner'],
        ['对比这两个表', 'DataDiffAnalyzer'],
        ['合并多个表', 'DataMerger'],
        ['合并所有Sheet', 'DataMerger'],
        ['按部门拆分成不同Sheet', 'DataSplitter'],
        ['每1000行拆一个文件', 'DataSplitter'],
        ['跨表匹配一下这两张表', 'TableJoiner'],
        ['生成PPT汇报', 'GeneratePPTReport'],
        ['总结这段内容', 'SmartSummary'],
        ['画个柱状图', 'SmartChartWizard'],
        // 【回归】"创建解析器/解析脚本"必须落到 GenerateOCRParseScript：
        // 曾因 ①该 Skill 无 ROUTING_POLICIES 被意图守卫静默丢弃 ②路由键里裸的 `脚本`
        // 把它判成代码生成意图 → 被 skill_SkillCreator / GenerateVBA 抢走执行。
        ['为火车票创建一个解析器', 'GenerateOCRParseScript'],
        ['帮我做一个发票解析脚本', 'GenerateOCRParseScript']
    ];
    let pass = 0;
    for (const [msg, expect] of skillCases) {
        try {
            await orch.dispatch({ message: msg, context: makeCtx() });
            // 断言路由锁定结果（orchestrator 路由阶段同步写入 turn.context._matchedSkill）。
            // 不用共享全局 window._lastSkillExecuted 轮询：聚合运行器（run-tests.js）下
            // 前序套件的孤儿异步链（dispatch 提前 resolve 后仍在跑的 ExecuteSkill）会跨测试
            // 覆写该全局，导致 "美化表格 → SmartChartWizard" 这类假失败（实测复现）。
            const turn = App.ChatV2.sessionStore.getCurrentTurn();
            const got = (turn && turn.context && turn.context._matchedSkill && (turn.context._matchedSkill.name || turn.context._matchedSkill.Name)) || null;
            const ok = got === expect;
            if (ok) pass++;
            console.log(`  ${ok ? '✅' : '❌'} ${msg} → ${got || '(未路由)'} ${ok ? '' : '(期望 ' + expect + ')'}`);
        } catch (e) {
            console.log(`  ❌ ${msg} → 异常: ${e.message}`);
        }
    }
    assert.strictEqual(pass, skillCases.length, `${pass}/${skillCases.length} skill 路由命中`);
});

test('skill 路由: 通用请求不误路由', async () => {
    const { App } = loadChatV2();
    App.cachedSkills = loadAllSkills();
    const tr = App.ChatV2.toolRegistry;
    for (const s of App.cachedSkills) {
        tr.register({
            name: 'skill.' + s.name, displayName: s.name, description: s.name,
            inputSchema: { type: 'object', properties: {} }, source: 'skill', skillInfo: s,
            async execute() { return { ok: true, result: 'ok' }; }
        });
    }
    let llmCalled = false;
    App.ChatV2.llmClient = { stream() { llmCalled = true; return (async function* () { yield { type: 'done', finishReason: 'stop' }; })(); } };
    window.chrome.webview.hostObjects.excelBridge.ExecuteSkill = async () => JSON.stringify({ success: true });

    await App.ChatV2.orchestrator.dispatch({ message: '你好', context: { selection: { range: 'A1' } } });
    await new Promise(r => setTimeout(r, 300));
    assert.strictEqual(llmCalled, true, '通用请求应留在 LLM');

    // 【回归守护 · 2026-10-09 意图守卫放宽后】
    // 放宽"强关键词豁免"（长词 → 中词 / 强 2 字词）与"句首查询词 = 技能动作词"放行之后，
    // 必须保证**疑问句仍不路由到 skill 执行**：
    //   "对比和合并有什么区别" 若被放行，会被当成"对比"指令真去跑 DataDiffAnalyzer；
    //   "清洗数据是什么意思" 放宽前就已误路由（清洗数据命中域关键词 + canClean），一并守护。
    // 靠"句首锚定的查询拦截"拦不住这类疑问在句中的问法，故由疑问语气兜底。
    const queryCases = [
        '什么是合并所有Sheet',
        '对比和合并有什么区别',
        '总结是什么意思',
        '跨表匹配怎么用',
        '清洗数据是什么意思'
    ];
    for (const msg of queryCases) {
        llmCalled = false;
        // 只等路由结论落定（dispatch 内同步写入 turn.context），
        // 不做 300ms 等待——本文件独立运行时有 5000ms 强制退出定时器，5 例会把它撑爆导致汇总丢失。
        await App.ChatV2.orchestrator.dispatch({ message: msg, context: { selection: { range: 'A1:H6' } } });
        await new Promise(r => setTimeout(r, 20));
        assert.strictEqual(llmCalled, true, `疑问句不应路由到 skill 执行，应交回 LLM 作答: ${msg}`);
    }
});

test('skill 路由: 写代码请求强制 code-pipeline，禁止掉进执行管道', async () => {
    const { App } = loadChatV2();
    App.cachedSkills = loadAllSkills();
    const tr = App.ChatV2.toolRegistry;
    for (const s of App.cachedSkills) {
        tr.register({
            name: 'skill.' + s.name, displayName: s.name, description: s.name,
            inputSchema: { type: 'object', properties: {} }, source: 'skill', skillInfo: s,
            async execute() { return { ok: true, result: 'ok' }; }
        });
    }
    // 阻止真调 LLM / 真执行
    App.ChatV2.llmClient = {
        stream() {
            return (async function* () {
                yield { type: 'reasoning', delta: '' };
                yield { type: 'done', finishReason: 'stop' };
            })();
        }
    };
    window.chrome.webview.hostObjects.excelBridge.ExecuteSkill = async () => JSON.stringify({ success: true });
    window.chrome.webview.hostObjects.excelBridge.ExecuteVBA = async () => JSON.stringify({ success: true, message: 'ok' });

    const orch = App.ChatV2.orchestrator;
    const ctx = { selection: { range: 'A1:H6', sheetName: 'Sheet1', workbookName: 'Book1', headers: ['地区', '销量'] } };

    // [消息, 期望 skill, 期望 kind]
    // 写代码类请求必须落 code-pipeline（只生成展示），即使关键词未命中 GenerateVBA 也要兜底强制；
    // 纯执行请求（无写代码意图）仍走 DataSplitter 执行管道。
    const cases = [
        ['帮我写一段 VBA 代码：拆分工作簿', 'GenerateVBA', 'code'],
        ['帮我写一个按工作表拆分工作簿的 VBA 宏', 'GenerateVBA', 'code'],
        ['帮我写一个拆分工作簿的宏', 'GenerateVBA', 'code'],
        ['帮我拆分工作簿', 'DataSplitter', 'skill']
    ];
    let pass = 0;
    for (const [msg, expectSkill, expectKind] of cases) {
        try {
            await orch.dispatch({ message: msg, context: { ...ctx } });
        } catch (e) {
            // pipeline 异步执行异常可接受（路由断言在 turn.context，dispatch 同步完成路由）
        }
        const turn = App.ChatV2.sessionStore.getCurrentTurn();
        const gotSkill = (turn && turn.context && turn.context._matchedSkill && (turn.context._matchedSkill.name || turn.context._matchedSkill.Name)) || null;
        const gotKind = turn ? turn.kind : null;
        const ok = gotSkill === expectSkill && gotKind === expectKind;
        if (ok) pass++;
        console.log(`  ${ok ? '✅' : '❌'} ${msg} → skill=${gotSkill} kind=${gotKind} ${ok ? '' : '(期望 ' + expectSkill + '/' + expectKind + ')'}`);
    }
    assert.strictEqual(pass, cases.length, `${pass}/${cases.length} 写代码路由命中`);
});

test('快捷入口直连: context._forcedSkillName 钉死技能，不被意图路由改判', async () => {
    const { App } = loadChatV2();
    App.cachedSkills = loadAllSkills();
    const tr = App.ChatV2.toolRegistry;
    for (const s of App.cachedSkills) {
        tr.register({
            name: 'skill.' + s.name, displayName: s.name, description: s.name,
            inputSchema: { type: 'object', properties: {} }, source: 'skill', skillInfo: s,
            async execute() { return { ok: true, result: 'ok' }; }
        });
    }
    App.ChatV2.llmClient = { stream() { throw new Error('不应调用 LLM'); } };
    window.chrome.webview.hostObjects.excelBridge.ExecuteSkill = async () => JSON.stringify({ success: true });

    const orch = App.ChatV2.orchestrator;
    // [消息, 绑定的技能, 期望 skill, 期望 kind]
    // 背景（实测 2026-10-02 23:37）：快捷入口原先只把提示词填进输入框、丢掉 skillName，
    // 于是"为火车票创建一个解析器"（GenerateOCRParseScript 快捷入口的提示词）被
    // skill_SkillCreator 抢走执行。绑定后必须无条件用绑定的那个技能。
    const cases = [
        // 绑定的技能与关键词路由结论相反 → 必须听绑定的
        ['创建一个帮我整理表格的skill', 'GenerateOCRParseScript', 'GenerateOCRParseScript', 'skill'],
        ['为火车票创建一个解析器', 'SkillCreator', 'SkillCreator', 'skill'],
        // 消息本身会触发"代码生成意图"（含"脚本"）→ 绑定技能优先于 code-pipeline
        ['帮我写一个解析脚本', 'GenerateOCRParseScript', 'GenerateOCRParseScript', 'skill'],
        // 绑定了一个不存在的技能 → 安全回退通用路由
        ['为火车票创建一个解析器', 'NoSuchSkillXYZ', 'GenerateOCRParseScript', 'skill']
    ];
    let pass = 0;
    for (const [msg, pinned, expectSkill, expectKind] of cases) {
        try {
            await orch.dispatch({
                message: msg,
                context: { selection: { range: 'A1:H6' }, _forcedSkillName: pinned, _forcedSkillSource: 'quickShortcut' }
            });
        } catch (e) {
            // pipeline 异步执行异常可接受（路由断言在 turn.context）
        }
        const turn = App.ChatV2.sessionStore.getCurrentTurn();
        const gotSkill = (turn && turn.context && turn.context._matchedSkill && (turn.context._matchedSkill.name || turn.context._matchedSkill.Name)) || null;
        const gotKind = turn ? turn.kind : null;
        const ok = gotSkill === expectSkill && gotKind === expectKind;
        if (ok) pass++;
        console.log(`  ${ok ? '✅' : '❌'} 绑定[${pinned}] "${msg}" → skill=${gotSkill} kind=${gotKind} ${ok ? '' : '(期望 ' + expectSkill + '/' + expectKind + ')'}`);
    }
    assert.strictEqual(pass, cases.length, `${pass}/${cases.length} 快捷入口直连命中`);
});

// 兜底退出
if (require.main === module) { setTimeout(() => process.exit(0), 5000).unref(); } // 仅独立运行时强制退出；聚合运行（run-tests.js require）不注册，避免进程提前退出
