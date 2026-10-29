/**
 * ============================================================
 * sync-ui-to-debug.js — 单目标 UI 同步（D1 / --dest）
 * ============================================================
 * 【日常开发请优先用】bash Scripts/sync-ui-runtime.sh
 *   一次覆盖 4 处运行时目录 + 统一刷新 ?v= + 逐字节校验。
 *
 * 本脚本适用场景：只同步到单个目标（默认 bin/Debug/UI，或 --dest）。
 * 会生成 .bak 备份；若只需清垃圾备份，可手工删各输出目录下 UI 里的 .bak 文件。
 * （勿在块注释里写 star-star-slash 通配路径，会提前结束注释。）
 *
 * 运行方式：
 *   node Tools/sync-ui-to-debug.js
 *
 * 可选参数：
 *   --stamp v2026072355   强制使用指定版本戳
 *   --dest <dir>          指定目标 UI 目录（绝对或相对 cwd），用于 PPT/Word 等非默认输出
 *   --dry-run             只打印计划，不实际写入
 *   --prune               删除目标中不在源目录的文件（默认保留 .bak）
 * ============================================================
 */

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const SRC_ROOT = path.join(ROOT, 'UI');
// 默认目标为 Excel 的 bin/Debug/UI；PPT/Word 通过 --dest 指向各自输出目录，
// 这样同一脚本即可统一刷新任意宿主的 HTML 版本戳（缓存 busting）。
const args = process.argv.slice(2);
const destArg = args.find((_, i) => args[i - 1] === '--dest');
const DEST_ROOT = destArg ? path.resolve(destArg) : path.join(ROOT, 'bin', 'Debug', 'UI');

const isDryRun = args.includes('--dry-run');
const doPrune = args.includes('--prune');
const stampArg = args.find((_, i) => args[i - 1] === '--stamp');

function nowStamp() {
    const d = new Date();
    const pad = n => String(n).padStart(2, '0');
    return `v${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}${pad(d.getHours())}${pad(d.getMinutes())}`;
}

/**
 * 取源 UI/index.html 当前的缓存版本号 —— 源是版本号的唯一真源。
 *
 * 为什么未显式传 --stamp 时必须沿用源戳、而不是新生成时间戳：
 *   Scripts/sync-ui-runtime.sh 的流程是「先刷新源 index.html 的 ?v=，再原样复制到 4 处
 *   运行时目录」，且明确要求「不要在各目标再单独刷一次 —— 会产生不同版本号，既让
 *   『4 处一致性』检查永远失败，也会让 WebView2 按 query 缓存出多份副本」。
 *   本脚本若按当前时间另生成戳，则每次运行（包括被文件监听/自动同步反复触发时）
 *   都会让该目标与其它 3 处及源产生分歧，ui-runtime-sync 的一致性闸门必然失败。
 *   实测 2026-10-09：bin/Debug 被反复刷成 v202610092259 / v202610092317，
 *   而源与其它 3 处停在 v202610092200。
 *   需要真正升版（改了 JS 要 bust 缓存）时：显式传 --stamp，
 *   或跑 bash Scripts/sync-ui-runtime.sh（它会统一刷新源 + 4 处）。
 */
function sourceStamp() {
    try {
        const srcHtml = fs.readFileSync(path.join(SRC_ROOT, 'index.html'), 'utf8');
        // ① 权威常量：window.__APP_VERSION__ 由规范脚本统一刷新，是版本号的单一真源。
        //    不能只取首个 ?v= —— 源里可能残留个别旧戳（历史 sed 未覆盖到的引用），
        //    取到旧值会让目标与源的 __APP_VERSION__ 仍然不一致。
        const appVer = srcHtml.match(/window\.__APP_VERSION__\s*=\s*['"](v[0-9]{12})['"]/);
        if (appVer) return appVer[1];
        // ② 兜底：取出现次数最多的 ?v= 戳（多数引用一致时才可信）
        const all = Array.from(srcHtml.matchAll(/\?v=(v[0-9]{12})/g)).map(m => m[1]);
        if (!all.length) return null;
        const freq = new Map();
        for (const s of all) freq.set(s, (freq.get(s) || 0) + 1);
        return Array.from(freq.entries()).sort((a, b) => b[1] - a[1])[0][0];
    } catch (_) {
        return null;
    }
}

const reusedSourceStamp = !stampArg ? sourceStamp() : null;
const newStamp = stampArg || reusedSourceStamp || nowStamp();

// 目录同步映射：源相对 UI/ 的路径 → 目标相对 bin/Debug/UI/ 的路径
const DIR_SYNCS = [
    { src: 'js', dest: 'js' },
    { src: 'css', dest: 'css' },
    { src: 'images', dest: 'images' },
    { src: 'lib', dest: 'lib' },
    { src: 'config', dest: 'config' },
    { src: 'prompts', dest: 'prompts' }
];

// 根目录下单文件同步（通配所有根目录文件，确保 md/txt/js/json/html 等文档与入口都被同步）
const ROOT_FILE_EXCLUDES = [
    /\.bak-\d+$/  // 排除所有同步脚本生成的备份文件
];

// 运行时需要同步的 HTML 文件（目标路径与源路径相同）
const HTML_FILES = [
    'index.html',
    'ai-settings-panel.html',
    'debug-window.html',
    'mobile-control-panel.html',
    'mobile-ocr.html',
    'ocr-panel.html',
    'ocr-settings-panel.html'
];

// bin/Debug 端允许保留的独有文件（不会被 prune 删除）
const PRUNE_PRESERVE_PATTERNS = [
    /\.bak-\d+$/,
    /PHASE\d+_.*\.md$/,
    /IMPROVEMENT_SUMMARY\.md$/
];

function ensureDir(dir) {
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
    }
}

function isPreserved(relPath) {
    return PRUNE_PRESERVE_PATTERNS.some(p => p.test(relPath));
}

function copyFile(srcPath, destPath) {
    const srcStat = fs.statSync(srcPath);
    let needCopy = true;
    if (fs.existsSync(destPath)) {
        const destStat = fs.statSync(destPath);
        if (srcStat.size === destStat.size && srcStat.mtimeMs <= destStat.mtimeMs) {
            needCopy = false;
        }
    }

    if (needCopy) {
        if (!isDryRun) {
            ensureDir(path.dirname(destPath));
            fs.copyFileSync(srcPath, destPath);
            fs.utimesSync(destPath, srcStat.atime, srcStat.mtime);
        }
        return true;
    }
    return false;
}

function copyDir(src, dest) {
    ensureDir(dest);
    const items = fs.readdirSync(src, { withFileTypes: true });
    let copied = 0;
    let skipped = 0;

    for (const item of items) {
        const srcPath = path.join(src, item.name);
        const destPath = path.join(dest, item.name);

        if (item.isDirectory()) {
            const sub = copyDir(srcPath, destPath);
            copied += sub.copied;
            skipped += sub.skipped;
        } else {
            if (copyFile(srcPath, destPath)) {
                copied++;
                console.log(`  + ${path.relative(ROOT, srcPath)} -> ${path.relative(ROOT, destPath)}`);
            } else {
                skipped++;
            }
        }
    }

    return { copied, skipped };
}

function backupFile(filePath) {
    const backupPath = `${filePath}.bak-${Date.now()}`;
    fs.copyFileSync(filePath, backupPath);
    return backupPath;
}

// 版本戳只写入目标副本（返回带新戳的内容），源 HTML 保持不被污染
function stampHtmlContent(content) {
    const oldStamps = [...content.matchAll(/\?v=([^"'\s>]+)/g)].map(m => m[1]);
    const uniqueOldStamps = [...new Set(oldStamps)];

    if (uniqueOldStamps.length === 0) {
        return { content, oldStamps: [], appVersionUpdated: false, updated: false, reason: 'no version stamp' };
    }

    let newContent = content.replace(/\?v=([^"'\s>]+)/g, `?v=${newStamp}`);

    // 同步更新内联版本常量 window.__APP_VERSION__，确保动态加载器使用同一版本戳
    const appVersionMatch = newContent.match(/window\.__APP_VERSION__\s*=\s*['"]([^'"]+)['"]/);
    if (appVersionMatch) {
        newContent = newContent.replace(/window\.__APP_VERSION__\s*=\s*['"]([^'"]+)['"]/g, `window.__APP_VERSION__ = '${newStamp}'`);
    }

    return { content: newContent, oldStamps: uniqueOldStamps, appVersionUpdated: !!appVersionMatch, updated: true };
}

function validateHtml(destHtmlPath) {
    const content = fs.readFileSync(destHtmlPath, 'utf8');
    const issues = [];

    if (content.trim().length < 100) issues.push('content too short');
    if (!content.includes('</html>')) issues.push('missing </html>');

    const stamps = [...content.matchAll(/\?v=([^"'\s>]+)/g)].map(m => m[1]);
    const unique = [...new Set(stamps)];

    if (unique.length > 0) {
        // 只有文件本身包含版本戳时才要求统一且更新
        if (!content.includes(`?v=${newStamp}`)) issues.push(`new version stamp ?v=${newStamp} not found`);
        if (unique.length > 1) issues.push(`inconsistent version stamps: ${unique.join(', ')}`);
    }

    return { valid: issues.length === 0, issues, stamps: unique };
}

function collectDestFiles(dir, baseRel = '') {
    const result = [];
    if (!fs.existsSync(dir)) return result;

    for (const item of fs.readdirSync(dir, { withFileTypes: true })) {
        const rel = path.posix.join(baseRel, item.name).replace(/\\/g, '/');
        const full = path.join(dir, item.name);
        if (item.isDirectory()) {
            result.push(...collectDestFiles(full, rel));
        } else {
            result.push({ rel, full });
        }
    }
    return result;
}

function pruneDest(expectedRels) {
    const destFiles = collectDestFiles(DEST_ROOT);
    let removed = 0;
    let preserved = 0;

    for (const { rel, full } of destFiles) {
        if (expectedRels.has(rel)) continue;

        if (isPreserved(rel)) {
            preserved++;
            continue;
        }

        if (!isDryRun) {
            fs.rmSync(full, { force: true });
        }
        removed++;
        console.log(`  - ${rel}`);
    }

    return { removed, preserved };
}

// ============================================================
// 主流程
// ============================================================
console.log('='.repeat(70));
console.log(`D1 build sync: UI -> ${path.relative(ROOT, DEST_ROOT) || DEST_ROOT} (full resource copy)`);
console.log('='.repeat(70));
console.log(`mode: ${isDryRun ? 'dry-run (preview)' : 'actual write'}`);
console.log(`prune: ${doPrune ? 'on (remove bin/Debug-only files)' : 'off'}`);
console.log(`new version stamp: ${newStamp}${reusedSourceStamp
    ? ' (reused from source UI/index.html, not bumped; pass --stamp or run Scripts/sync-ui-runtime.sh to bump)'
    : ''}`);
console.log('');

let totalCopied = 0;
let totalSkipped = 0;
const expectedRels = new Set();

// 同步目录
for (const { src: srcRel, dest: destRel } of DIR_SYNCS) {
    const src = path.join(SRC_ROOT, srcRel);
    const dest = path.join(DEST_ROOT, destRel);

    if (!fs.existsSync(src)) {
        console.log(`  - skip ${srcRel}: source dir not found`);
        continue;
    }

    console.log(`>> sync UI/${srcRel} -> bin/Debug/UI/${destRel}`);
    const result = copyDir(src, dest);
    totalCopied += result.copied;
    totalSkipped += result.skipped;
    console.log(`   copied: ${result.copied}, skipped: ${result.skipped}`);
    console.log('');
}

// 同步根目录下所有文件
console.log('>> sync UI/ root files -> bin/Debug/UI/');
let rootCopied = 0;
let rootSkipped = 0;
for (const item of fs.readdirSync(SRC_ROOT, { withFileTypes: true })) {
    if (item.isDirectory()) continue;

    const srcRel = item.name;
    if (ROOT_FILE_EXCLUDES.some(p => p.test(srcRel))) continue;

    const src = path.join(SRC_ROOT, srcRel);
    const dest = path.join(DEST_ROOT, srcRel);

    if (copyFile(src, dest)) {
        rootCopied++;
        console.log(`  + UI/${srcRel} -> bin/Debug/UI/${srcRel}`);
    } else {
        rootSkipped++;
    }
    expectedRels.add(srcRel.replace(/\\/g, '/'));
}
totalCopied += rootCopied;
totalSkipped += rootSkipped;
console.log(`   copied: ${rootCopied}, skipped: ${rootSkipped}`);
console.log('');

// 同步 HTML 文件并刷新版本戳（只写目标副本，不污染源文件）
console.log('>> stamp HTML version (dest only)');
const htmlResults = [];
for (const htmlRel of HTML_FILES) {
    const srcHtml = path.join(SRC_ROOT, htmlRel);
    const destHtml = path.join(DEST_ROOT, htmlRel);
    expectedRels.add(htmlRel);

    try {
        let content;
        if (fs.existsSync(srcHtml)) {
            content = fs.readFileSync(srcHtml, 'utf8');
        } else if (fs.existsSync(destHtml)) {
            content = fs.readFileSync(destHtml, 'utf8');
        } else {
            console.error(`  [ERR] ${htmlRel}: neither source nor dest exists`);
            htmlResults.push({ path: htmlRel, error: 'file missing' });
            continue;
        }

        const stamped = stampHtmlContent(content);
        htmlResults.push({ path: `UI/${htmlRel}`, ...stamped });

        if (!stamped.updated) {
            console.log(`  [i] UI/${htmlRel}: ${stamped.reason}`);
            if (fs.existsSync(srcHtml)) {
                if (copyFile(srcHtml, destHtml)) totalCopied++; else totalSkipped++;
            }
            continue;
        }

        if (!isDryRun) {
            if (fs.existsSync(destHtml)) backupFile(destHtml);
            ensureDir(path.dirname(destHtml));
            fs.writeFileSync(destHtml, stamped.content, 'utf8');
            totalCopied++;
            console.log(`  [OK] UI/${htmlRel}: ${(stamped.oldStamps || []).join(', ')} -> ${newStamp}`);
        } else {
            totalSkipped++;
            console.log(`  [note] UI/${htmlRel}: ${(stamped.oldStamps || []).join(', ')} -> ${newStamp} (dry-run)`);
        }
    } catch (e) {
        console.error(`  [ERR] ${htmlRel}: ${e.message}`);
        htmlResults.push({ path: htmlRel, error: e.message });
    }
}

// 收集所有期望存在的相对路径（用于 prune）
for (const { src: srcRel, dest: destRel } of DIR_SYNCS) {
    const src = path.join(SRC_ROOT, srcRel);
    if (!fs.existsSync(src)) continue;
    const walk = (dir, base) => {
        for (const item of fs.readdirSync(dir, { withFileTypes: true })) {
            const rel = path.posix.join(base, item.name).replace(/\\/g, '/');
            if (item.isDirectory()) {
                walk(path.join(dir, item.name), rel);
            } else {
                expectedRels.add(rel);
            }
        }
    };
    walk(src, destRel);
}
for (const htmlRel of HTML_FILES) {
    expectedRels.add(htmlRel);
}

// prune
if (doPrune) {
    console.log('');
    console.log('>> prune unneeded files in bin/Debug');
    const pruneResult = pruneDest(expectedRels);
    console.log(`   removed: ${pruneResult.removed}, preserved: ${pruneResult.preserved}`);
}

// 验证 HTML 完整性
console.log('');
console.log('>> validate HTML integrity');
let allValid = true;
if (isDryRun) {
    console.log('  - dry-run: skip validation');
} else {
    for (const htmlRel of HTML_FILES) {
        const destHtml = path.join(DEST_ROOT, htmlRel);
        if (!fs.existsSync(destHtml)) {
            console.error(`  [ERR] ${htmlRel}: file missing`);
            allValid = false;
            continue;
        }
        try {
            const validation = validateHtml(destHtml);
            if (validation.valid) {
                console.log(`  [OK] ${htmlRel}: single version stamp (${validation.stamps.join(', ')})`);
            } else {
                console.error(`  [ERR] ${htmlRel}: ${validation.issues.join('; ')}`);
                allValid = false;
            }
        } catch (e) {
            console.error(`  [ERR] ${htmlRel}: validation error - ${e.message}`);
            allValid = false;
        }
    }
}

console.log('');
console.log('='.repeat(70));
console.log(`>> sync end: copied ${totalCopied}, skipped ${totalSkipped}`);
console.log(`>> single version stamp: ${newStamp}`);

if (!allValid) {
    console.error('[ERR] HTML validation failed; check backups and fix manually');
    process.exit(1);
}

if (isDryRun) {
    console.log('[note] this was a dry-run; nothing written');
} else {
    console.log('sync success!');
}
