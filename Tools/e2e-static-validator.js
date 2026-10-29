/**
 * ============================================================
 * e2e-static-validator.js — 端到端静态校验器
 * ============================================================
 * 功能：在不启动 Excel / 浏览器的情况下，验证 chat-v2 重构后的
 *       关键代码约束、旧架构清理、构建产物同步、版本戳完整性。
 *
 * 运行方式：
 *   cd Tools
 *   node e2e-static-validator.js
 * ============================================================
 */

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');

function read(rel) {
    return fs.readFileSync(path.join(ROOT, rel), 'utf8');
}

function normalizeLineEndings(text) {
    return text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
}

function getVersionStamps(html) {
    const set = new Set();
    const regex = /\?v=([^"'\s>]+)/g;
    let m;
    while ((m = regex.exec(html)) !== null) {
        set.add(m[1]);
    }
    return Array.from(set);
}

const results = {
    passed: 0,
    failed: 0,
    errors: []
};

function check(name, testFn, details = '') {
    try {
        const ok = testFn();
        if (ok) {
            console.log(`✅ ${name}`);
            results.passed++;
        } else {
            console.error(`❌ ${name}`);
            if (details) console.error(`   ${details}`);
            results.failed++;
            results.errors.push(name);
        }
    } catch (e) {
        console.error(`💥 ${name} 执行异常: ${e.message}`);
        results.failed++;
        results.errors.push(name);
    }
}

console.log('='.repeat(70));
console.log('🔍 ExcelAIVSTOAddIn 端到端静态校验');
console.log('='.repeat(70));

// ============================================================
// A 阶段：旧架构清理确认（已删除文件必须保持不存在）
// ============================================================
console.log('\n📦 A 阶段：旧架构清理确认');

const deletedLegacyFiles = [
    'UI/js/app-ai-send.js',
    'UI/js/app-ai-calls.js',
    'UI/js/app-ai-skill.js',
    'UI/js/app-ai-task.js',
    'UI/js/task-agent/task-agent.js',
    'UI/js/chat-v2/services/subagent-registry.js'
];

for (const rel of deletedLegacyFiles) {
    const binRel = rel.replace('UI/', 'bin/Debug/UI/');
    check(`旧架构文件已删除: ${path.basename(rel)}`, () => {
        return !fs.existsSync(path.join(ROOT, rel)) && !fs.existsSync(path.join(ROOT, binRel));
    }, `${rel} 或 ${binRel} 仍存在，请清理`);
}

// ============================================================
// B 阶段：chat-v2 架构关键约束
// ============================================================
console.log('\n🧭 B 阶段：chat-v2 架构关键约束');

check('chat-orchestrator.js 提供路由编排', () => {
    const c = read('UI/js/chat-v2/chat-orchestrator.js');
    return c.includes('function pickPipeline') &&
           c.includes('function inferKind') &&
           c.includes('function shouldRouteToSkill') &&
           c.includes('AbortController');
}, '路由编排关键函数缺失');

check('agent-loop.js 提供多轮循环与护栏', () => {
    const c = read('UI/js/chat-v2/agent-loop.js');
    return c.includes('MAX_TOOL_ROUNDS = 5') &&
           c.includes('function parseTextToolCalls') &&
           c.includes('FAKE_EXECUTION_PATTERN');
}, '多轮循环或护栏缺失');

check('tool-registry.js 提供工具注册表', () => {
    const c = read('UI/js/chat-v2/services/tool-registry.js');
    return c.includes('listBySource(source)') && c.includes('registerAll(tools)');
}, '工具注册表关键方法缺失');

check('pipeline-registry.js 提供管线注册化', () => {
    const c = read('UI/js/chat-v2/services/pipeline-registry.js');
    return c.includes('register(kind, definition)') && c.includes('pick(kind)');
}, '管线注册化关键方法缺失');

check('skill-spec-engine.js 提供 SkillSpec 规划执行', () => {
    const c = read('UI/js/chat-v2/services/skill-spec-engine.js');
    return c.includes('function planAndRender');
}, 'SkillSpec 引擎缺失');

check('chat-entry.js 提供消息入口', () => {
    const c = read('UI/js/chat-v2/chat-entry.js');
    return c.includes('function sendMessage') || c.includes('function gatherContext');
}, '消息入口缺失');

check('chat-renderer.js 提供消息渲染', () => {
    const c = read('UI/js/chat-v2/ui/chat-renderer.js');
    return c.includes('App.ChatV2.renderer = {') && c.includes('renderUserBubble');
}, '消息渲染器缺失');

// ============================================================
// C 阶段：构建产物同步
// ============================================================
console.log('\n🔄 构建产物同步');

const syncFiles = [
    'UI/index.html',
    'UI/css/styles.css',
    'UI/js/chat-v2/index.js',
    'UI/js/chat-v2/chat-orchestrator.js',
    'UI/js/chat-v2/agent-loop.js',
    'UI/js/chat-v2/services/tool-registry.js',
    'UI/js/chat-v2/services/skill-spec-engine.js',
    'UI/js/chat-v2/services/pipeline-registry.js'
];

for (const file of syncFiles) {
    const binFile = file.replace('UI/', 'bin/Debug/UI/');
    check(`bin/Debug 与 source 同步: ${path.basename(file)}`, () => {
        return normalizeLineEndings(read(file)) === normalizeLineEndings(read(binFile));
    }, `${file} 与 ${binFile} 内容不一致，请重新同步`);
}

// ============================================================
// D 阶段：版本戳完整性（每个外链资源都必须带 ?v= 版本戳）
// ============================================================
console.log('\n🏷️ D 阶段：版本戳完整性');

function versionCheck(rel) {
    const html = read(rel);
    const stamps = getVersionStamps(html);
    const externalLinks = (html.match(/<script[^>]+src="[^"]+"[^>]*>/g) || [])
        .concat(html.match(/<link[^>]+stylesheet[^>]+>/g) || []);
    const withoutVersion = externalLinks.filter(l => !/\?v=/.test(l));
    return { stamps, withoutVersion, total: externalLinks.length };
}

check('UI/index.html 所有外链资源带版本戳', () => {
    const v = versionCheck('UI/index.html');
    return v.withoutVersion.length === 0;
}, `存在无版本戳资源: ${versionCheck('UI/index.html').withoutVersion.join(', ')}`);

check('bin/Debug/UI/index.html 所有外链资源带版本戳', () => {
    const v = versionCheck('bin/Debug/UI/index.html');
    return v.withoutVersion.length === 0;
}, `存在无版本戳资源: ${versionCheck('bin/Debug/UI/index.html').withoutVersion.join(', ')}`);

check('source 与 bin/Debug index.html 版本戳集合一致', () => {
    const srcStamps = getVersionStamps(read('UI/index.html'));
    const binStamps = getVersionStamps(read('bin/Debug/UI/index.html'));
    return srcStamps.length === binStamps.length && srcStamps.every(s => binStamps.includes(s));
}, `source=${getVersionStamps(read('UI/index.html'))}, bin=${getVersionStamps(read('bin/Debug/UI/index.html'))}`);

// ============================================================
// 总结
// ============================================================
console.log('\n' + '='.repeat(70));
console.log(`📊 静态校验完成：通过 ${results.passed} / 失败 ${results.failed}`);
if (results.failed > 0) {
    console.log('\n❌ 失败项：');
    results.errors.forEach(e => console.log(`   • ${e}`));
    process.exit(1);
} else {
    console.log('\n🎉 所有静态约束校验通过！');
}
