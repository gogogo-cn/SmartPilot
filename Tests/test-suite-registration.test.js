/**
 * ============================================================
 * Tests/test-suite-registration.test.js — 测试套件注册完整性护栏
 * ============================================================
 * 背景（2026-10-09）：run-tests.js 曾手工维护 SUITES 数组，47 个 .test.js 里
 * 只有 35 个被注册，**12 个护栏套件从未被执行**——其中包括
 * host-export-pdf-dispatch、mcp-dispatcher-format-guard、
 * skill-routing-domain-guard、host-export-write-gate 等防回归套件。
 * 结果：已修复的 bug 悄悄回归却无人发现。
 *
 * 现 run-tests.js 改为自动发现（扫描 Tests/*.test.js），不在此处的套件必须
 * 在 EXCLUDES 中显式声明原因。本套件固化这两条约束。
 *
 * 运行：node Tests/test-suite-registration.test.js（随 run-tests.js 自动发现）
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const RUNNER = path.join(__dirname, 'run-tests.js');

/** 从 run-tests.js 源码中解析显式排除清单 */
function parseExcludes(source) {
    const start = source.indexOf('const EXCLUDES = {');
    assert.ok(start > -1, 'run-tests.js 应定义 EXCLUDES 显式排除清单');
    const end = source.indexOf('};', start);
    const block = source.slice(start, end);
    return [...block.matchAll(/'([^']+\.test\.js)'\s*:\s*'([^']*)'/g)]
        .map(m => ({ file: m[1], reason: m[2] }));
}

test('测试运行器: 自动发现全部 .test.js，不再依赖手工 SUITES', () => {
    const s = fs.readFileSync(RUNNER, 'utf8');
    assert.ok(!/const SUITES = \[/.test(s),
        'run-tests.js 不应再有手工维护的 SUITES 数组（历史上导致 12 个套件漏跑）');
    assert.ok(/discoverSuites\(\)/.test(s), 'run-tests.js 应通过 discoverSuites() 自动扫描');
    console.log('  [ok] 运行器为自动发现模式');
});

test('测试注册完整性: 每个 .test.js 要么被执行，要么在 EXCLUDES 中写明原因', () => {
    const s = fs.readFileSync(RUNNER, 'utf8');
    const excludes = parseExcludes(s);
    const excludedFiles = new Set(excludes.map(e => e.file));

    const onDisk = fs.readdirSync(__dirname)
        .filter(f => f.endsWith('.test.js'))
        .filter(f => f !== 'run-tests.js')
        .filter(f => !f.startsWith('_'))
        .sort();

    // EXCLUDES 里声明的文件必须真实存在（防止残留过期声明）
    for (const e of excludes) {
        assert.ok(onDisk.includes(e.file), `EXCLUDES 声明了不存在的套件: ${e.file}`);
        assert.ok(e.reason && e.reason.length > 0, `EXCLUDES 必须为 ${e.file} 写明排除原因`);
    }

    // 排除项只能是脚本式（否则就是漏跑）
    const unjustified = excludes.filter(e => !/脚本式|e2e-script|process\.exit/.test(e.reason));
    assert.deepStrictEqual(unjustified.map(e => e.file), [],
        `以下套件被排除但原因不是"脚本式"，疑似漏跑: ${unjustified.map(e => e.file).join(', ')}`);

    assert.ok(onDisk.length - excludedFiles.size >= 40,
        `被执行套件数量异常（${onDisk.length - excludedFiles.size}），疑似扫描逻辑回归`);
    console.log(`  [ok] ${onDisk.length} 个套件：执行 ${onDisk.length - excludedFiles.size}，显式排除 ${excludedFiles.size}`);
});

test('测试隔离: 运行器为每个套件起独立子进程（避免全局污染导致"单跑绿、聚合红"）', () => {
    const s = fs.readFileSync(RUNNER, 'utf8');
    assert.ok(/spawnSync\(/.test(s), 'run-tests.js 应为每个套件起独立子进程（helpers/harness 的共享全局会互相污染）');
    assert.ok(/cwd: ROOT/.test(s), '子进程应以仓库根目录为 cwd，保证相对路径行为一致');
    console.log('  [ok] 套件进程隔离就位');
});

test('测试隔离: harness 的 window/App 对象身份跨次加载保持稳定', () => {
    const harnessSrc = fs.readFileSync(path.join(__dirname, 'helpers', 'harness.js'), 'utf8');
    assert.ok(/let SHARED_WINDOW = null/.test(harnessSrc),
        'helpers/harness.js 应复用同一 window 对象（每次新建会让先加载套件的 window 变成孤儿，其 SmartPilotHost 设置失效）');
    console.log('  [ok] harness 全局身份稳定');
});
