'use strict';
/**
 * module-self-sufficiency.test.js — 子模块自足性守栏。
 *
 * 背景：UI 拆分子模块（app-skill-mgmt.js → app-skill-mgmt.code-library.js 等）时，
 * 从 app.js 抄函数体偶尔会漏带被调用的私有函数定义，直到用户点击才炸：
 *   - 2026-10-21 app.excelai.local 删除代码片段 → ReferenceError:
 *     _tryAcquireActionLock is not defined（code-library.js L678）
 *   - 更早：index.html 引用未定义的 copyCode
 * 本测试对 IIFE 子模块做「裸调用差集」分析：调用一个函数但本文件既没定义、
 * 也不是内置/宿主全局 → 运行环境里必炸，直接判失败。
 * 附阳性对照：同样的分析器必须能抓住人工构造的未定义调用。
 */
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const UI_JS = path.join(__dirname, '..', 'UI', 'js');

// 语法关键字（会匹配 `name(` 模式但不是函数调用）
const KEYWORDS = new Set(['if', 'for', 'while', 'switch', 'catch', 'function', 'return', 'typeof', 'in', 'of', 'do', 'else', 'try', 'finally', 'throw', 'case', 'new', 'delete', 'void', 'instanceof', 'await', 'yield']);
// JS 内置 + 项目宿主全局（window.* / App.* 形式不算裸调用，正则已排除）
const BUILTINS = new Set([
    'confirm', 'alert', 'prompt', 'fetch', 'setTimeout', 'clearTimeout', 'setInterval',
    'clearInterval', 'queueMicrotask', 'requestAnimationFrame', 'cancelAnimationFrame',
    'JSON', 'Date', 'Object', 'Array', 'String', 'Number', 'Boolean', 'Math', 'Map', 'Set',
    'WeakMap', 'WeakSet', 'Promise', 'Error', 'TypeError', 'RangeError', 'RegExp', 'Blob',
    'File', 'FileReader', 'URL', 'btoa', 'atob', 'structuredClone', 'Function', 'Symbol',
    'parseInt', 'parseFloat', 'isNaN', 'isFinite', 'encodeURIComponent', 'decodeURIComponent',
    'BigInt', 'Reflect', 'Proxy', 'Intl',
    'require', 'module', 'exports', 'globalThis',
    'navigator', 'location', 'localStorage', 'sessionStorage', 'history', 'document',
    'window', 'console', 'App', 'ChatLog', 'ExcelBridgeConfig'
]);

/** 剥离注释与字符串字面量，降低误报（模板字符串整体剥离，宁可漏报插值内的调用） */
function stripNoise(src) {
    return src
        .replace(/\/\*[\s\S]*?\*\//g, ' ')
        .replace(/(^|[^:\\])\/\/[^\n]*/g, '$1 ')
        .replace(/'(?:[^'\\\n]|\\.)*'/g, "''")
        .replace(/"(?:[^"\\\n]|\\.)*"/g, '""')
        .replace(/`(?:[^`\\]|\\.)*`/g, '``');
}

/** 返回 [{name, line}]：本文件未定义、也非内置/宿主的裸调用 */
function findUndefinedCalls(src) {
    const code = stripNoise(src);
    const defined = new Set();
    for (const m of code.matchAll(/\bfunction\s+([A-Za-z_$][\w$]*)\s*\(/g)) defined.add(m[1]);
    for (const m of code.matchAll(/\b(?:const|let|var)\s+([A-Za-z_$][\w$]*)/g)) defined.add(m[1]);
    for (const m of code.matchAll(/function[^(]*\(([^)]*)\)/g)) {
        for (const p of m[1].split(',')) {
            const n = p.trim().split('=')[0].trim();
            if (/^[A-Za-z_$][\w$]*$/.test(n)) defined.add(n);
        }
    }
    const out = [];
    const seen = new Set();
    for (const m of code.matchAll(/(^|[^.\w$])([A-Za-z_$][\w$]*)\s*\(/g)) {
        const name = m[2];
        if (KEYWORDS.has(name) || BUILTINS.has(name) || defined.has(name) || seen.has(name)) continue;
        seen.add(name);
        out.push({ name, line: src.slice(0, m.index).split('\n').length });
    }
    return out;
}

const MODULES = ['app-skill-mgmt.js', 'app-skill-mgmt.code-library.js'];

test('skill-mgmt 子模块无「漏带定义」的裸调用', () => {
    for (const f of MODULES) {
        const p = path.join(UI_JS, f);
        assert.ok(fs.existsSync(p), '模块不存在: ' + f);
        const bad = findUndefinedCalls(fs.readFileSync(p, 'utf8'));
        assert.deepStrictEqual(bad, [], f + ' 存在未定义调用（抽模块漏带定义）');
    }
});

test('阳性对照：分析器能抓住人工构造的未定义调用', () => {
    const fake = [
        '(function(){',
        '  function ok() { return helper(); }',           // helper 未定义 → 应被抓到
        '  App.x = ok;',
        '  window.y = function(){ return ok(); };',        // ok 已定义 → 不报
        '})();'
    ].join('\n');
    const bad = findUndefinedCalls(fake);
    assert.strictEqual(bad.length, 1);
    assert.strictEqual(bad[0].name, 'helper');
});

test('_tryAcquireActionLock 在 code-library.js 中必须有定义（钉死 2026-10-21 的删除报错）', () => {
    const src = fs.readFileSync(path.join(UI_JS, 'app-skill-mgmt.code-library.js'), 'utf8');
    assert.ok(/\bfunction\s+_tryAcquireActionLock\s*\(/.test(src),
        'code-library.js 必须自带 _tryAcquireActionLock（子模块自足，不依赖主文件私有函数）');
    // 且必须真的被 deleteSnippetFromLibrary 用上
    const fnStart = src.indexOf('async function deleteSnippetFromLibrary');
    assert.ok(fnStart >= 0, '应有 deleteSnippetFromLibrary');
    const fnBody = src.slice(fnStart, src.indexOf('\n    }', fnStart));
    assert.ok(/_tryAcquireActionLock\(/.test(fnBody), 'deleteSnippetFromLibrary 应继续使用操作锁');
});
