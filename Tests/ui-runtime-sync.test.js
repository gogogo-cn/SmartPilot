/**
 * ============================================================
 * Tests/ui-runtime-sync.test.js — 运行时 UI 副本与缓存版本号一致性
 * ============================================================
 * 背景（2026-10-05 晚连续 4 轮"改了 JS 但线上没生效"）：
 *   WebView2 按 <script src="...js?v=版本号"> 缓存。我把 JS 复制到了运行时目录，
 *   却没同步 index.html、也没刷新 ?v= 参数 → **文件是新的、WebView2 加载的是旧的**。
 *   最难查的一点：agent-loop.js 的 ?v= 是我改的（看起来"更新了"），
 *   而 tool-registry.util.js 的 ?v= 还停在 9 月 26 日 → 恰好是承载真正修复的那个文件。
 *
 * 本测试锁死三件事：
 *   1) 4 处运行时目录的关键文件与源UI/ 逐字节一致（忽略 CRLF）；
 *   2) 4 处 index.html 里的 ?v= 版本号**完全相同**（不能各写各的）；
 *   3) 关键脚本的 ?v= 不为空且格式合法。
 * 运行：node Tests/ui-runtime-sync.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const SRC = path.join(ROOT, 'UI');
const TARGETS = [
    path.join(ROOT, 'bin', 'Debug', 'UI'),
    path.join(ROOT, 'bin', 'Release', 'UI'),
    path.join(ROOT, 'Hosts', 'PptAIVSTOAddIn', 'bin', 'Debug', 'UI'),
    path.join(ROOT, 'Hosts', 'WordAIVSTOAddIn', 'bin', 'Debug', 'UI'),
];

/** 曾经出过问题的关键文件（承载导出路径修复 + 失败熔断） */
const KEY_FILES = [
    'index.html',
    path.join('js', 'chat-v2', 'agent-loop.js'),
    path.join('js', 'chat-v2', 'services', 'tool-registry.js'),
    path.join('js', 'chat-v2', 'services', 'tool-registry.util.js'),
];

const read = (p) => fs.readFileSync(p, 'utf8').replace(/\r\n/g, '\n');
const exists = (p) => fs.existsSync(p);

test('4 处运行时目录都存在', () => {
    for (const t of TARGETS) {
        assert.ok(exists(t), '运行时目录不存在（宿主 DLL 未构建？）: ' + t);
    }
});

test('关键文件与源 UI 逐字节一致（否则改了也不生效）', () => {
    for (const t of TARGETS) {
        for (const f of KEY_FILES) {
            const rel = f.split(path.sep).join('/');
            assert.ok(exists(path.join(t, f)), '缺失文件: ' + path.join(t, f));
            assert.strictEqual(
                read(path.join(t, f)),
                read(path.join(SRC, f)),
                `内容不一致: ${rel} @ ${t}\n` +
                '→ 跑 bash Scripts/sync-ui-runtime.sh 同步'
            );
        }
    }
});

test('4 处 index.html 的缓存版本号完全相同（防某处漏刷导致加载旧 JS）', () => {
    const stamps = TARGETS.map((t) => {
        const html = read(path.join(t, 'index.html'));
        const m = html.match(/tool-registry\.util\.js\?v=(v[0-9]{12})/);
        assert.ok(m, 'index.html 里找不到 tool-registry.util.js 的版本号 @ ' + t);
        return m[1];
    });
    const uniq = [...new Set(stamps)];
    assert.strictEqual(uniq.length, 1,
        '4 处 index.html 版本号不一致，WebView2 会加载到旧版本: ' + JSON.stringify(stamps) +
        '\n→ 跑 bash Scripts/sync-ui-runtime.sh');
});

test('承载修复的关键脚本都带合法版本号', () => {
    const html = read(path.join(SRC, 'index.html'));
    for (const f of ['js/chat-v2/agent-loop.js', 'js/chat-v2/services/tool-registry.js',
        'js/chat-v2/services/tool-registry.util.js']) {
        const re = new RegExp(f.replace(/[/.]/g, '\\$&') + '\\?v=(v[0-9]{12})');
        const m = html.match(re);
        assert.ok(m, `index.html 里 ${f} 缺少 ?v= 版本号（WebView2 会用缓存）`);
    }
});

test('内联 __APP_VERSION__ 与 ?v= 版本号一致（源 + 4 处运行时，防只刷一半）', () => {
    // 背景：__APP_VERSION__ 不随 ?v= 查询参数走缓存，状态栏/关于面板直接读它。
    // .ps1 同步脚本 2026-10-12 就刷了它，.sh 一直漏 → 同步后 ?v= 是新的、
    // __APP_VERSION__ 停在几个月前，误导排查。两个脚本必须同刷，这里守住。
    const stampOf = (html) => {
        const qv = html.match(/tool-registry\.util\.js\?v=(v[0-9]{12})/);
        const av = html.match(/window\.__APP_VERSION__\s*=\s*'(v[0-9]{12})'/);
        return { qv: qv && qv[1], av: av && av[1] };
    };
    const src = stampOf(read(path.join(SRC, 'index.html')));
    assert.ok(src.av, '源 index.html 缺少 __APP_VERSION__');
    assert.strictEqual(src.av, src.qv,
        `源 __APP_VERSION__（${src.av}）与 ?v=（${src.qv}）不一致：同步脚本只刷了一半`);
    for (const t of TARGETS) {
        const s = stampOf(read(path.join(t, 'index.html')));
        assert.ok(s.av, '运行时副本缺少 __APP_VERSION__ @ ' + t);
        assert.strictEqual(s.av, src.qv,
            `运行时 __APP_VERSION__（${s.av}）与源 ?v=（${src.qv}）不一致 @ ${t} → 跑 bash Scripts/sync-ui-runtime.sh`);
    }
});

test('同步脚本存在且覆盖 4 处目录与 index.html', () => {
    const sh = path.join(ROOT, 'Scripts', 'sync-ui-runtime.sh');
    assert.ok(exists(sh), '缺少 Scripts/sync-ui-runtime.sh');
    const body = read(sh);
    for (const needle of ['index.html', 'bin/Debug/UI', 'bin/Release/UI',
        'PptAIVSTOAddIn', 'WordAIVSTOAddIn']) {
        assert.ok(body.includes(needle), '同步脚本未覆盖: ' + needle);
    }
    assert.ok(/index\.html/.test(body) && /\?v=/.test(body),
        '同步脚本必须同时刷新 index.html 的 ?v= 版本号');
    assert.ok(/__APP_VERSION__/.test(body),
        '同步脚本必须刷新内联 __APP_VERSION__（.ps1 2026-10-12 已有，.sh 曾漏导致两边行为分叉）');
});
