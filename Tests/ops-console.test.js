'use strict';
/**
 * ops-console.test.js — 把 OpsConsole 子套件注册进主聚合运行器。
 *
 * 主运行器只自动发现 Tests/*.test.js，而 OpsConsole/test/ 是独立目录
 * （有自己的 npm test）。这里起一个子进程跑 `node --test OpsConsole/test`，
 * 解析 TAP 汇总，有任何失败即红——保证主套件回归能覆盖运营台。
 */
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const OPS_TEST_DIR = path.join(ROOT, 'OpsConsole', 'test');

test('OpsConsole 子套件全部通过（node --test OpsConsole/test/*.test.js）', () => {
    // 注意：本机 Node 把 `--test <目录>` 当模块加载会 MODULE_NOT_FOUND，
    // 必须显式传文件列表（与 package.json 的 npm test  Glob 等价）。
    const files = fs.readdirSync(OPS_TEST_DIR)
        .filter(f => f.endsWith('.test.js'))
        .sort()
        .map(f => path.join(OPS_TEST_DIR, f));
    assert.ok(files.length >= 5, 'OpsConsole 测试文件数量异常: ' + files.length);
    // 父进程若本身跑在 `node --test` 下会带 NODE_TEST_CONTEXT，子进程继承后会被
    // 判定「递归运行」直接跳过所有文件（只打印一条 warning，TAP 汇总为空）。
    // 必须显式删掉再 spawn。
    const env = Object.assign({}, process.env);
    delete env.NODE_TEST_CONTEXT;
    const res = spawnSync(process.execPath, ['--test', ...files], {
        cwd: ROOT,
        env,
        encoding: 'utf8',
        timeout: 120000,
        windowsHide: true
    });
    const output = String(res.stdout || '') + '\n' + String(res.stderr || '');
    const pick = (key) => {
        const m = output.match(new RegExp('^#\\s+' + key + '\\s+(\\d+)', 'm'));
        return m ? Number(m[1]) : null;
    };
    const tests = pick('tests');
    const pass = pick('pass');
    const fail = pick('fail');
    assert.ok(tests !== null, '子套件未打印 TAP 汇总（可能崩溃）：\n' + output.slice(-2000));
    assert.strictEqual(fail, 0, 'OpsConsole 有用例失败：\n' +
        output.split(/\r?\n/).filter(l => /^not ok\b/.test(l)).join('\n'));
    assert.ok(pass >= 20, '用例数量异常偏少（' + pass + '），可能没跑全');
});
