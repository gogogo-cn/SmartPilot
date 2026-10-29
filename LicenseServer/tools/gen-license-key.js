'use strict';
/**
 * gen-license-key.js — 离线激活码生成工具（路线①：人工收款后发 key）。
 *
 * 用法：
 *   node tools/gen-license-key.js --login user@x.com --plan pro_yearly --days 365
 *   node tools/gen-license-key.js --batch 20 --plan pro_monthly --days 31 --out keys.csv
 *   node tools/gen-license-key.js --verify SPK1-xxxx.yyyy
 *
 * 密钥来源：环境变量 LICENSE_KEY_SECRET，未设置时用 lib/licenseKey.js 的默认值。
 * 注意：默认 secret 同时内置于 Bridges/LicenseKey.vb 与 Hosts/Shared/HostLicense.vb，
 * 换 secret 必须三处同步并重新编译发布，否则旧 key 全部失效。
 *
 * CSV 列：key,login,plan,seats,expiresAt —— 可直接导入自动发货发卡平台。
 *
 * 发货台账：传 --out 时会同时把每个 key 追加到 data/issued-keys.csv
 * （issuedAt,key,login,plan,seats,expiresAt,note），用于对账、防重复发货、
 * 售后查证。--note 建议写客户名/订单号。
 */
const path = require('path');
const fs = require('fs');
const lk = require(path.join(__dirname, '..', 'lib', 'licenseKey.js'));

function parseArgs(argv) {
    const args = { days: 365, seats: 1, plan: 'pro_yearly' };
    for (let i = 0; i < argv.length; i++) {
        const a = argv[i];
        const next = () => argv[++i];
        if (a === '--login') args.login = next();
        else if (a === '--plan') args.plan = next();
        else if (a === '--seats') args.seats = parseInt(next(), 10) || 1;
        else if (a === '--days') args.days = parseInt(next(), 10) || 365;
        else if (a === '--count' || a === '--batch') args.count = parseInt(next(), 10) || 1;
        else if (a === '--out') args.out = next();
        else if (a === '--note') args.note = next();
        else if (a === '--exp') args.exp = next();
        else if (a === '--verify') args.verify = next();
        else if (a === '--help' || a === '-h') args.help = true;
        else if (!a.startsWith('--') && !args._positional) args._positional = a;
    }
    return args;
}

function csvCell(s) {
    const t = String(s == null ? '' : s);
    return /[",\r\n]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t;
}

const LEDGER_PATH = path.join(__dirname, '..', 'data', 'issued-keys.csv');
const LEDGER_HEADER = 'issuedAt,key,login,plan,seats,expiresAt,note';

/** 追加发货台账（append-only，人工收款/发卡平台补货都走这里）。 */
function appendLedger(rows, note) {
    const dir = path.dirname(LEDGER_PATH);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    const isNew = !fs.existsSync(LEDGER_PATH);
    const lines = rows.map(r => [
        new Date().toISOString(), r.key, r.login, r.plan, r.seats, r.expiresAt, note || ''
    ].map(csvCell).join(','));
    fs.appendFileSync(LEDGER_PATH, (isNew ? LEDGER_HEADER + '\r\n' : '') + lines.join('\r\n') + '\r\n', 'utf8');
    return LEDGER_PATH;
}

function main() {
    const args = parseArgs(process.argv.slice(2));

    if (args.help) {
        console.log('用法: node tools/gen-license-key.js [--login 账号] [--plan 方案] [--seats 席位] [--days 天数|--exp Unix秒] [--count N] [--out keys.csv] [--note 备注] [--verify key]');
        console.log('plan: pro_monthly | pro_yearly | team');
        return 0;
    }

    if (args.verify) {
        const r = lk.verify(args.verify);
        console.log(JSON.stringify(r, null, 2));
        return r.valid ? 0 : 1;
    }

    const count = Math.max(1, args.count || 1);
    const rows = [];
    for (let i = 0; i < count; i++) {
        const login = args.login
            ? (count > 1 ? args.login.replace(/@.*$/, '') + '+' + (i + 1) + '@' + (args.login.split('@')[1] || 'local') : args.login)
            : '-';
        const key = lk.createKey({
            login,
            plan: args.plan,
            seats: args.seats,
            days: args.days,
            expUnix: args.exp
        });
        const v = lk.verify(key);
        if (!v.valid) {
            console.error('自检失败：生成的 key 无法通过验签 ->', v.reason);
            return 2;
        }
        rows.push({ key, login: v.login, plan: v.plan, seats: v.seats, expiresAt: v.periodEnd });
    }

    if (args.out) {
        const header = 'key,login,plan,seats,expiresAt';
        const lines = rows.map(r =>
            [r.key, r.login, r.plan, r.seats, r.expiresAt].map(csvCell).join(','));
        fs.writeFileSync(path.resolve(args.out), header + '\r\n' + lines.join('\r\n') + '\r\n', 'utf8');
        const ledger = appendLedger(rows, args.note);
        console.log('已写入 ' + rows.length + ' 个激活码 -> ' + path.resolve(args.out));
        console.log('台账已登记 -> ' + ledger + (args.note ? '（备注: ' + args.note + '）' : ''));
    } else {
        rows.forEach(r => console.log(r.key));
        if (rows.length === 1) {
            console.error('方案=' + rows[0].plan + ' 席位=' + rows[0].seats + ' 到期=' + rows[0].expiresAt);
        }
    }
    return 0;
}

if (require.main === module) process.exit(main());
module.exports = { main, parseArgs };
