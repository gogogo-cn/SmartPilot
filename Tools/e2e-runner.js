/**
 * ============================================================
 * e2e-runner.js — 端到端测试矩阵运行器
 * ============================================================
 * 功能：
 *   1. 运行 L1 静态校验（e2e-static-validator.js）
 *   2. 运行 L2 VBA 安全验证器单元测试
 *   3. 运行额外的 Node.js 可执行规则测试（路由/工具过滤）
 *   4. 生成 JSON + Markdown 测试报告
 *
 * 运行方式：
 *   cd Tools
 *   node e2e-runner.js
 * ============================================================
 */

const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const ROOT = path.resolve(__dirname, '..');
const TOOLS_DIR = __dirname;
const REPORT_DIR = path.join(ROOT, 'Tools', 'reports');

if (!fs.existsSync(REPORT_DIR)) {
    fs.mkdirSync(REPORT_DIR, { recursive: true });
}

const report = {
    meta: {
        project: 'ExcelAIVSTOAddIn',
        phase: 'D2',
        timestamp: new Date().toISOString(),
        runner: 'e2e-runner.js'
    },
    suites: [],
    summary: { passed: 0, failed: 0, total: 0 }
};

function logSection(title) {
    console.log('\n' + '='.repeat(70));
    console.log(title);
    console.log('='.repeat(70));
}

function addSuite(name, status, details = []) {
    const suite = { name, status, details, timestamp: new Date().toISOString() };
    report.suites.push(suite);
    if (status === 'PASSED') {
        report.summary.passed++;
    } else {
        report.summary.failed++;
    }
    report.summary.total++;
    return suite;
}

function runExternal(command, args, cwd) {
    const result = spawnSync(command, args, {
        cwd: cwd || TOOLS_DIR,
        encoding: 'utf8',
        timeout: 120000,
        env: process.env
    });
    return {
        exitCode: result.status,
        stdout: result.stdout || '',
        stderr: result.stderr || '',
        signal: result.signal
    };
}

// ============================================================
// L1 静态校验
// ============================================================
logSection('🔍 L1 静态约束校验');
const staticResult = runExternal('node', ['e2e-static-validator.js'], TOOLS_DIR);
console.log(staticResult.stdout);
if (staticResult.stderr) console.error(staticResult.stderr);

const staticPassed = staticResult.exitCode === 0;
addSuite(
    'L1 静态约束校验',
    staticPassed ? 'PASSED' : 'FAILED',
    staticPassed
        ? ['18 项静态检查全部通过', '构建产物同步、版本戳一致']
        : ['静态校验失败，详见上方输出', staticResult.stderr].filter(Boolean)
);

// ============================================================
// L2 VBA 安全验证器单元测试
// ============================================================
logSection('🛡️ L2 VBA 安全验证器测试');
const vbaResult = runExternal('node', ['test-validator.js'], TOOLS_DIR);
console.log(vbaResult.stdout);
if (vbaResult.stderr) console.error(vbaResult.stderr);

const vbaPassed = vbaResult.exitCode === 0 && vbaResult.stdout.includes('验证通过');
addSuite(
    'L2 VBA 安全验证器',
    vbaPassed ? 'PASSED' : 'FAILED',
    vbaPassed
        ? ['危险模式识别、自动修复、执行计划生成均正常']
        : ['VBA 验证器测试失败', vbaResult.stderr].filter(Boolean)
);

// ============================================================
// L2 路由规则 & 护栏规则测试（chat-v2 当前架构）
// ============================================================
logSection('🧭 L2 路由规则 & 护栏规则测试');

const routerTests = [];

function testRouterCase(name, fn) {
    try {
        const ok = fn();
        routerTests.push({ name, status: ok ? 'PASSED' : 'FAILED' });
        console.log(`${ok ? '✅' : '❌'} ${name}`);
        return ok;
    } catch (e) {
        routerTests.push({ name, status: 'FAILED', error: e.message });
        console.error(`💥 ${name}: ${e.message}`);
        return false;
    }
}

// 读取 chat-v2 当前架构源码进行规则级验证
// 旧架构 smart-router/app-ai-calls/app-ai-send 已删除，改测 chat-v2 路由与护栏
const orchestratorSrc = fs.readFileSync(path.join(ROOT, 'UI/js/chat-v2/chat-orchestrator.js'), 'utf8');
const agentLoopSrc = fs.readFileSync(path.join(ROOT, 'UI/js/chat-v2/agent-loop.js'), 'utf8');
const toolRegistrySrc = fs.readFileSync(path.join(ROOT, 'UI/js/chat-v2/services/tool-registry.js'), 'utf8');

const allRouterPassed =
    testRouterCase('orchestrator 提供 pickPipeline 管线选择', () =>
        orchestratorSrc.includes('function pickPipeline') && orchestratorSrc.includes('pipelines.text')
    ) &&
    testRouterCase('orchestrator 提供 inferKind 附件类型推断', () =>
        orchestratorSrc.includes('function inferKind')
    ) &&
    testRouterCase('orchestrator 提供 shouldRouteToSkill 通用 Skill 意图路由', () =>
        orchestratorSrc.includes('function shouldRouteToSkill')
    ) &&
    testRouterCase('orchestrator 支持 Skill 声明式 pipeline 路由（非硬编码）', () =>
        orchestratorSrc.includes('pipeline: "skill"') &&
        orchestratorSrc.includes('pipeline: "html"') &&
        orchestratorSrc.includes('pipeline: "python"') &&
        orchestratorSrc.includes('pipeline: "text"')
    ) &&
    testRouterCase('orchestrator 检测编程语言指令（python/vba）', () =>
        orchestratorSrc.includes('function detectLanguageDirective') &&
        orchestratorSrc.includes("'python'") &&
        orchestratorSrc.includes("'vba'")
    ) &&
    testRouterCase('orchestrator 支持 AbortController 取消能力', () =>
        orchestratorSrc.includes('AbortController')
    ) &&
    testRouterCase('agent-loop 限制工具循环上限（MAX_TOOL_ROUNDS=5）', () =>
        agentLoopSrc.includes('MAX_TOOL_ROUNDS = 5')
    ) &&
    testRouterCase('agent-loop 工具结果截断（4000字符）', () =>
        agentLoopSrc.includes('MAX_TOOL_RESULT_CHARS = 4000')
    ) &&
    testRouterCase('agent-loop 拦截伪造工具名（parseTextToolCalls）', () =>
        agentLoopSrc.includes('function parseTextToolCalls')
    ) &&
    testRouterCase('agent-loop 拦截"光说不做"假执行（FAKE_EXECUTION_PATTERN）', () =>
        agentLoopSrc.includes('FAKE_EXECUTION_PATTERN')
    ) &&
    testRouterCase('agent-loop 拦截 OCR 工具误用（非图片场景）', () =>
        agentLoopSrc.includes('isOcrTool') && agentLoopSrc.includes('excel_read_range')
    ) &&
    testRouterCase('agent-loop 工具执行后强制总结（不再提供 tools）', () =>
        agentLoopSrc.includes('toolsExecuted') && agentLoopSrc.includes('shouldProvideTools')
    ) &&
    testRouterCase('tool-registry 支持按 source 过滤（listBySource）', () =>
        toolRegistrySrc.includes('listBySource(source)')
    ) &&
    testRouterCase('tool-registry 支持注册表增删（register/unregister）', () =>
        toolRegistrySrc.includes('registerAll(tools)') && toolRegistrySrc.includes('unregister(name)')
    );

addSuite(
    'L2 路由规则 & 护栏规则',
    allRouterPassed ? 'PASSED' : 'FAILED',
    routerTests.map(t => `${t.status === 'PASSED' ? '✅' : '❌'} ${t.name}${t.error ? ': ' + t.error : ''}`)
);

// ============================================================
// L2 构建产物同步检查
// ============================================================
logSection('🔄 L2 构建产物同步抽查');
// 旧架构文件已删除，改抽查 chat-v2 当前关键文件（源 vs bin/Debug）
const syncFiles = [
    'UI/index.html',
    'UI/css/styles.css',
    'UI/js/chat-v2/index.js',
    'UI/js/chat-v2/chat-orchestrator.js',
    'UI/js/chat-v2/agent-loop.js'
];

const syncTests = [];
let allSyncPassed = true;
for (const file of syncFiles) {
    const binFile = file.replace('UI/', 'bin/Debug/UI/');
    const srcPath = path.join(ROOT, file);
    const binPath = path.join(ROOT, binFile);
    let passed = false;
    let detail = '';
    try {
        const src = fs.readFileSync(srcPath, 'utf8').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
        const bin = fs.readFileSync(binPath, 'utf8').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
        passed = src === bin;
        detail = passed ? '内容一致' : '内容不一致';
    } catch (e) {
        passed = false;
        detail = `读取失败: ${e.message}`;
    }
    syncTests.push({ file, status: passed ? 'PASSED' : 'FAILED', detail });
    console.log(`${passed ? '✅' : '❌'} ${path.basename(file)} — ${detail}`);
    if (!passed) allSyncPassed = false;
}

addSuite(
    'L2 构建产物同步抽查',
    allSyncPassed ? 'PASSED' : 'FAILED',
    syncTests.map(t => `${t.status === 'PASSED' ? '✅' : '❌'} ${t.file} — ${t.detail}`)
);

// ============================================================
// 生成报告
// ============================================================
logSection('📊 测试报告生成');

const jsonReportPath = path.join(REPORT_DIR, `e2e-report-${Date.now()}.json`);
fs.writeFileSync(jsonReportPath, JSON.stringify(report, null, 2), 'utf8');
console.log(`✅ JSON 报告: ${jsonReportPath}`);

const mdLines = [
    '# ExcelAIVSTOAddIn 端到端测试矩阵报告',
    '',
    `**阶段**: D2  \n**时间**: ${report.meta.timestamp}  \n**运行器**: ${report.meta.runner}`,
    '',
    '## 汇总',
    '',
    `| 通过 | 失败 | 总计 |`,
    `|------|------|------|`,
    `| ${report.summary.passed} | ${report.summary.failed} | ${report.summary.total} |`,
    '',
    '## 详细结果',
    ''
];

for (const suite of report.suites) {
    mdLines.push(`### ${suite.name} ${suite.status === 'PASSED' ? '✅' : '❌'}`);
    mdLines.push('');
    for (const detail of suite.details) {
        mdLines.push(`- ${detail}`);
    }
    mdLines.push('');
}

mdLines.push('## 备注');
mdLines.push('');
mdLines.push('- L1/L2 可在 Node.js 环境下全自动运行。');
mdLines.push('- L3/L4/L5 需要 Excel + 浏览器 + AI 后端环境，需手动执行 `e2e-test-matrix.md` 中的场景。');
mdLines.push('- 下一关键任务：D1 构建流程自动化（同步 UI/js/ 到 bin/Debug/UI/js/ 并统一版本戳）。');

const mdReportPath = path.join(REPORT_DIR, `e2e-report-${Date.now()}.md`);
fs.writeFileSync(mdReportPath, mdLines.join('\n'), 'utf8');
console.log(`✅ Markdown 报告: ${mdReportPath}`);

// ============================================================
// 最终汇总
// ============================================================
logSection('🏁 最终汇总');
console.log(`通过: ${report.summary.passed} / 失败: ${report.summary.failed} / 总计: ${report.summary.total}`);

if (report.summary.failed > 0) {
    console.error('\n❌ 以下测试套件失败：');
    report.suites.filter(s => s.status === 'FAILED').forEach(s => console.error(`   • ${s.name}`));
    process.exit(1);
} else {
    console.log('\n🎉 L1/L2 端到端测试矩阵全部通过！');
    process.exit(0);
}
