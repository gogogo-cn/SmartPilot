'use strict';
/**
 * 服务端上传包打包工具 —— 把「要上传的服务器文件」和「留在客户端的文件」彻底分开。
 *
 * 背景：仓库里 LicenseServer/ 既含可公开的代码，也含绝不能上传的东西：
 *   .env（真实密钥）、config.json（运营台写入的支付配置）、data/（用户库 + 密码哈希 + 台账）、
 *   logs/、node_modules/、test/。手工挑文件极易漏，故用白名单打包 + 出库前密钥扫描。
 *
 * 用法：
 *   node Tools/package-server.js                     # 默认输出到 dist/server/
 *   node Tools/package-server.js --out build/server  # 指定输出目录
 *   node Tools/package-server.js --include-data      # 迁移台账/用户库时（谨慎，含用户数据）
 *   node Tools/package-server.js --dry-run           # 只打印清单，不写盘
 *
 * 产物（dist/server/ 的内容可直接整个上传到服务器目录）：
 *   server.js  package.json  package-lock.json  README.md  .env.example
 *   lib/  public/  tools/  [data/ 仅 --include-data]
 * 上传到服务器后执行：npm install --omit=dev && 按 .env.example 写 .env && npm start
 *
 * 安全：出库前扫描真实密钥形态（sk_live_/sk_test_长串/whsec_长串/PEM 私钥），命中即退出码 1。
 *       .env.example 里的占位值（sk_test_xxx / whsec_xxx）不算命中。
 */
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const SRC = path.join(ROOT, 'LicenseServer');

// 白名单：只有这些会被打进上传包
const INCLUDE_FILES = ['server.js', 'package.json', 'package-lock.json', 'README.md', '.env.example'];
const INCLUDE_DIRS = ['lib', 'public', 'tools'];

// 永不进包（即使存在）
const EXCLUDE_NAMES = new Set(['node_modules', 'data', 'logs', 'test', 'config.json', '.env', '.git', 'obj', 'bin']);
const EXCLUDE_SUFFIX = ['.log', '.tmp', '.old'];

// 真实密钥形态（占位值 sk_test_xxx / whsec_xxx 因长度不足不会命中）
const SECRET_PATTERNS = [
    { name: 'Stripe live key', re: /sk_live_[A-Za-z0-9]{10,}/ },
    { name: 'Stripe test key', re: /sk_test_[A-Za-z0-9]{20,}/ },
    { name: 'Stripe webhook secret', re: /whsec_[A-Za-z0-9]{20,}/ },
    { name: 'PEM private key', re: /-----BEGIN (RSA |EC |OPENSSH |)?PRIVATE KEY-----/ }
];

function isExcludedName(name) {
    if (EXCLUDE_NAMES.has(name)) return true;
    const lower = name.toLowerCase();
    // 备份文件形如 xxx.bak 或 xxx.bak-20261020-xxx（后缀还有内容，endsWith 抓不到）
    if (lower.includes('.bak')) return true;
    return EXCLUDE_SUFFIX.some((s) => lower.endsWith(s));
}

function copyRecursive(src, dest, out) {
    const stat = fs.statSync(src);
    if (stat.isDirectory()) {
        fs.mkdirSync(dest, { recursive: true });
        for (const entry of fs.readdirSync(src)) {
            if (isExcludedName(entry)) { out.excluded.push(path.relative(src, entry)); continue; }
            copyRecursive(path.join(src, entry), path.join(dest, entry), out);
        }
        return;
    }
    fs.mkdirSync(path.dirname(dest), { recursive: true });
    fs.copyFileSync(src, dest);
    out.files.push(path.relative(ROOT, dest));
}

/** 扫描产物中是否夹带真实密钥。返回命中列表。 */
function scanForSecrets(dir) {
    const findings = [];
    if (!fs.existsSync(dir)) return findings;
    const walk = (d) => {
        for (const entry of fs.readdirSync(d)) {
            const p = path.join(d, entry);
            const st = fs.statSync(p);
            if (st.isDirectory()) { walk(p); continue; }
            if (st.size > 2 * 1024 * 1024) continue; // 跳过超大文件
            let text = '';
            try { text = fs.readFileSync(p, 'utf8'); } catch (_) { continue; }
            for (const pat of SECRET_PATTERNS) {
                if (pat.re.test(text)) {
                    findings.push({ file: path.relative(dir, p), pattern: pat.name });
                }
            }
        }
    };
    walk(dir);
    return findings;
}

/**
 * 构建服务端上传包。
 * @param {{outDir?: string, includeData?: boolean, dryRun?: boolean, allowOutside?: boolean}} opts
 */
function buildServerPackage(opts) {
    const o = opts || {};
    const includeData = !!o.includeData;
    const dryRun = !!o.dryRun;
    let outDir = path.resolve(o.outDir || path.join(ROOT, 'dist', 'server'));

    // 安全闸：默认只允许写到仓库内，避免误删/覆盖仓库外目录
    if (!o.allowOutside && !outDir.startsWith(ROOT + path.sep) && outDir !== ROOT) {
        throw new Error('拒绝写入仓库外的目录：' + outDir + '（确需请用 allowOutside: true）');
    }
    if (!fs.existsSync(SRC)) throw new Error('找不到源目录：' + SRC);

    // 只清理本脚本管理的顶层条目，不做 rm -rf
    if (!dryRun && fs.existsSync(outDir)) {
        const managed = INCLUDE_FILES.concat(INCLUDE_DIRS);
        if (includeData) managed.push('data');
        for (const name of managed) {
            const p = path.join(outDir, name);
            if (fs.existsSync(p)) fs.rmSync(p, { recursive: true, force: true });
        }
    }

    const out = { outDir, files: [], excluded: [], secretFindings: [], includeData };

    for (const f of INCLUDE_FILES) {
        const src = path.join(SRC, f);
        if (!fs.existsSync(src)) continue;
        if (!dryRun) copyRecursive(src, path.join(outDir, f), out);
        else out.files.push('LicenseServer/' + f);
    }
    for (const d of INCLUDE_DIRS) {
        const src = path.join(SRC, d);
        if (!fs.existsSync(src)) continue;
        if (!dryRun) copyRecursive(src, path.join(outDir, d), out);
        else out.files.push('LicenseServer/' + d + '/');
    }
    if (includeData) {
        const src = path.join(SRC, 'data');
        if (fs.existsSync(src) && !dryRun) copyRecursive(src, path.join(outDir, 'data'), out);
    }

    if (!dryRun) out.secretFindings = scanForSecrets(outDir);
    return out;
}

function main() {
    const argv = process.argv.slice(2);
    const get = (flag) => {
        const i = argv.indexOf(flag);
        return i >= 0 ? argv[i + 1] : null;
    };
    const dryRun = argv.includes('--dry-run');
    const includeData = argv.includes('--include-data');
    const outArg = get('--out');

    const res = buildServerPackage({ outDir: outArg || undefined, includeData, dryRun });

    console.log('=========== 服务端上传包 ===========');
    console.log('输出目录: ' + res.outDir + (dryRun ? '  (dry-run，未写盘)' : ''));
    console.log('包含文件/目录:');
    // 打印顶层清单
    const top = INCLUDE_FILES.concat(INCLUDE_DIRS);
    if (includeData) top.push('data');
    top.forEach((t) => {
        const src = path.join(SRC, t);
        console.log('  ' + (fs.existsSync(src) ? '✓' : '-') + ' ' + t);
    });
    console.log('已排除(永不进包): ' + Array.from(EXCLUDE_NAMES).join(', '));
    console.log('');
    console.log('密钥扫描: ' + (res.secretFindings.length === 0 ? '通过（未发现真实密钥）' : '发现 ' + res.secretFindings.length + ' 处!'));
    res.secretFindings.forEach((f) => console.log('  ⚠ ' + f.file + ' → ' + f.pattern));
    console.log('');
    console.log('服务器上执行：npm install --omit=dev 后按 .env.example 写 .env，再 npm start');
    if (res.secretFindings.length) process.exit(1);
}

if (require.main === module) main();

module.exports = {
    buildServerPackage,
    scanForSecrets,
    INCLUDE_FILES,
    INCLUDE_DIRS,
    EXCLUDE_NAMES,
    SECRET_PATTERNS,
    ROOT,
    SRC
};
