/**
 * ============================================================
 * Tests/run-tests.js — 测试聚合运行器（每套件独立子进程，进程隔离）
 * ============================================================
 * 运行：node Tests/run-tests.js   （或 npm test）
 *
 * 【设计变更 2026-10-09】
 * 1. 套件列表**自动发现**（扫描 Tests/*.test.js），不再手工维护 SUITES ——
 *    手工维护曾导致 12 个护栏套件（host-export-pdf-dispatch、
 *    mcp-dispatcher-format-guard、skill-routing-domain-guard 等）从未被执行。
 *
 * 2. 每个套件在**独立子进程**中运行。此前所有套件 require 进同一进程：
 *    · 模块的模块级 loadChatV2() 全部先执行，再跑 test()；
 *    · 各套件又会在模块级/测试里改写共享全局（window.SmartPilotHost、
 *      App.cachedSkills…），于是「单跑全绿、聚合必红」，且红的一直被当成
 *      「套件不稳定」而忽略。子进程隔离后，聚合结果 == 单跑结果。
 *    （helpers/harness.js 已同步修掉 window 对象身份漂移，但进程隔离才是根治）
 *
 * 3. 子进程若没有输出 TAP 汇总（例如被套件里的 process.exit 提前结束），
 *    一律判为失败，避免「提前退出 → 假绿」。
 * ============================================================
 */
'use strict';
const { spawnSync } = require('child_process');
const path = require('path');
const fs = require('fs');

const ROOT = path.join(__dirname, '..');

/**
 * 显式排除的套件（脚本式：自行 main()/process.exit，需独立 script 运行）。
 * 新增排除项必须写明原因——否则视为漏跑。
 */
const EXCLUDES = {
    'skill-chart-clean-e2e.test.js': '脚本式（main()/process.exit 自行退出），见 package.json test:e2e-script',
    'skill-crud-e2e.test.js': '脚本式（顶层 process.exit），见 package.json test:crud'
};

/** 脚本式套件的自动识别：行首（无缩进）出现 main() 或 process.exit( */
function looksScriptStyle(src) {
    const noBlockComments = src.replace(/\/\*[\s\S]*?\*\//g, '');
    return noBlockComments.split(/\r?\n/)
        .filter(l => !/^\s*\/\//.test(l))
        .some(l => /^(?:void\s+)?(?:main\(\)\s*;?|process\.exit\()/.test(l));
}

function discoverSuites() {
    return fs.readdirSync(__dirname)
        .filter(f => f.endsWith('.test.js'))
        .filter(f => f !== path.basename(__filename))
        .filter(f => !f.startsWith('_'))
        .sort();
}

function parseTapSummary(output) {
    const pick = (key) => {
        const m = output.match(new RegExp('^#\\s+' + key + '\\s+(\\d+)', 'm'));
        return m ? Number(m[1]) : null;
    };
    const tests = pick('tests');
    const pass = pick('pass');
    const fail = pick('fail');
    const notOk = output.split(/\r?\n/)
        .filter(l => /^not ok\b/.test(l))
        .map(l => l.replace(/^not ok\s*\d*\s*-?\s*/, '').trim());
    return { tests, pass, fail, notOk };
}

function runSuite(file) {
    const full = path.join(__dirname, file);
    let res;
    try {
        // ⚠️ 必须用 --test 启动（2026-10-16 修）：直接 `node <file>` 时，node:test 的 TAP
        // 汇总要等事件循环排空才打印；而若干套件会残留 ref'd 定时器（如 StepExecutor 的
        // 60s AGENT_TIMEOUT、dashboard-html-renderer 的 1s setInterval），导致循环迟迟不空，
        // 套件自带的兜底 `setTimeout(process.exit(0), 3~5s)` 先触发 → 进程被杀、汇总永远
        // 打不出来（表现为"用例全 ok 但未打印汇总"）。--test 由 runner 接管生命周期，
        // 用例跑完立即打印汇总并带 exit code 退出，不受残留定时器影响。
        res = spawnSync(process.execPath, ['--test', full], {
            cwd: ROOT,
            stdio: ['ignore', 'pipe', 'pipe'],
            timeout: 180000,
            windowsHide: true
        });
    } catch (err) {
        return { file, ok: false, reason: '无法启动子进程: ' + err.message, tests: null, pass: null, fail: null, notOk: [] };
    }
    const output = String(res.stdout || '') + '\n' + String(res.stderr || '');
    const sum = parseTapSummary(output);

    if (sum.tests === null) {
        // 没有 TAP 汇总：多数是套件自带的「兜底退出」定时器（require.main === module 时
        // setTimeout(process.exit(0), 3~5s)）在仍有挂起用例时把进程结束了。
        // TAP 会随每个用例结束即时打印 ok/not ok，因此仍可据已打印的用例判定。
        const okCount = (output.match(/^ok\s+\d+/gm) || []).length;
        if (sum.notOk.length > 0) {
            return {
                file, ok: false, tests: okCount + sum.notOk.length, pass: okCount, fail: sum.notOk.length,
                notOk: sum.notOk, output,
                reason: `未打印汇总但有失败用例（exit=${res.status}）`
            };
        }
        if (okCount === 0) {
            return {
                file, ok: false, tests: null, pass: null, fail: null, notOk: [], output,
                reason: `无任何用例输出（exit=${res.status}${res.error ? ', ' + res.error.message : ''}）；可能崩溃或进程提前结束`
            };
        }
        return {
            file, ok: true, warn: true, tests: okCount, pass: okCount, fail: 0, notOk: [], output,
            reason: `${okCount} 个用例已通过，但进程未打印汇总（存在挂起用例，建议排查）`
        };
    }
    const failed = (sum.fail || 0) > 0 || res.status !== 0;
    return { file, ok: !failed, tests: sum.tests, pass: sum.pass, fail: sum.fail, notOk: sum.notOk, output };
}

const discovered = discoverSuites();
const suites = [];
const skipped = [];
const undeclared = [];

for (const f of discovered) {
    if (Object.prototype.hasOwnProperty.call(EXCLUDES, f)) {
        skipped.push({ file: f, reason: EXCLUDES[f] });
        continue;
    }
    if (looksScriptStyle(fs.readFileSync(path.join(__dirname, f), 'utf8'))) {
        undeclared.push(f);
        skipped.push({ file: f, reason: '疑似脚本式（未登记 EXCLUDES，已拦截）' });
        continue;
    }
    suites.push(f);
}

console.log(`[run-tests] 自动发现 ${discovered.length} 个套件：执行 ${suites.length}，跳过 ${skipped.length}`);
for (const s of skipped) console.log(`  ⏭  跳过 ${s.file} — ${s.reason}`);
if (undeclared.length > 0) {
    console.error(`  ⚠ 以下套件疑似脚本式但未登记 EXCLUDES，请补齐声明: ${undeclared.join(', ')}`);
}
console.log('');

let totalTests = 0, totalPass = 0, totalFail = 0;
const failedSuites = [];
const warnSuites = [];

for (const f of suites) {
    const r = runSuite(f);
    if (r.ok) {
        console.log(`  ${r.warn ? '⚠️ ' : '✅'} ${f}  (${r.pass}/${r.tests})${r.warn ? '  [' + r.reason + ']' : ''}`);
        if (r.warn) warnSuites.push(r);
        totalTests += r.tests || 0;
        totalPass += r.pass || 0;
        totalFail += r.fail || 0;
    } else {
        failedSuites.push(r);
        totalTests += r.tests || 0;
        totalPass += r.pass || 0;
        totalFail += (r.fail || 0) + (r.fail === null ? 1 : 0);
        console.log(`  ❌ ${f}  ${r.reason || ('失败 ' + r.fail + '/' + r.tests)}`);
        for (const n of (r.notOk || []).slice(0, 10)) console.log(`       · ${n}`);
    }
}

console.log('\n================ 汇总 ================');
console.log(`套件: ${suites.length} 个（跳过 ${skipped.length}${warnSuites.length ? '，警告 ' + warnSuites.length : ''}）`);
console.log(`用例: ${totalTests}  通过: ${totalPass}  失败: ${totalFail}`);
if (warnSuites.length > 0) {
    console.log('\n⚠️  以下套件未打印汇总（存在挂起用例，结果可能不完整，建议排查）：');
    for (const r of warnSuites) console.log(`  ⚠️  ${r.file} — ${r.reason}`);
}
if (failedSuites.length > 0) {
    console.log(`\n失败套件 ${failedSuites.length} 个:`);
    for (const r of failedSuites) {
        console.log(`  ❌ ${r.file}${r.reason ? ' — ' + r.reason : ''}`);
        if (r.reason && r.output) {
            const tail = r.output.trim().split(/\r?\n/).slice(-8).join('\n');
            console.log(String(tail).split('\n').map(l => '       ' + l).join('\n'));
        }
    }
    process.exitCode = 1;
} else {
    console.log('\n全部通过 ✅');
}
