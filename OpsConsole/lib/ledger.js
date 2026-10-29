'use strict';
/**
 * ledger.js — 发货台账 CSV 读写（OpsConsole 数据层之一）。
 *
 * 格式与 LicenseServer/tools/gen-license-key.js 逐字节对齐（改格式必须两边同步，
 * OpsConsole/test/compat.test.js 有互认守栏）：
 *   表头：issuedAt,key,login,plan,seats,expiresAt,note
 *   行尾：CRLF；新文件带表头；单元格含 " , \r \n 时加双引号，内部 " 双写。
 *
 * CSV 是权威账本：gen-key 命令行与运营台都只追加、不改写。
 * 运营状态（已发货/已退款等）不进 CSV，走 ops-state.json 覆盖层（见 store.js）。
 */
const fs = require('fs');
const path = require('path');

const HEADER = 'issuedAt,key,login,plan,seats,expiresAt,note';

function csvCell(s) {
    const t = String(s == null ? '' : s);
    return /[",\r\n]/.test(t) ? '"' + t.replace(/"/g, '""') + '"' : t;
}

/**
 * 解析整个 CSV 文本为记录数组。支持引号内的逗号、换行、双写引号。
 * @returns {string[][]} 每行一个单元格数组
 */
function parseCsv(text) {
    const t = String(text == null ? '' : text);
    const records = [];
    let field = '';
    let record = [];
    let inQuotes = false;
    let i = 0;
    while (i < t.length) {
        const ch = t[i];
        if (inQuotes) {
            if (ch === '"') {
                if (t[i + 1] === '"') { field += '"'; i += 2; continue; }
                inQuotes = false; i++; continue;
            }
            field += ch; i++; continue;
        }
        if (ch === '"') { inQuotes = true; i++; continue; }
        if (ch === ',') { record.push(field); field = ''; i++; continue; }
        if (ch === '\r') { i++; continue; }
        if (ch === '\n') {
            // 空行不产生记录（幻影行会让台账行数虚高）
            if (field !== '' || record.length > 0) {
                record.push(field); field = ''; records.push(record); record = [];
            } else {
                field = '';
            }
            i++; continue;
        }
        field += ch; i++;
    }
    if (field !== '' || record.length > 0) { record.push(field); records.push(record); }
    return records;
}

/**
 * 读台账 → 行对象数组。跳过表头与空行；文件不存在返回 []。
 * 列数不足补空串；note 列含未加引号的逗号时可还原（容错，正常写入不会出现）。
 */
function readLedger(file) {
    if (!file || !fs.existsSync(file)) return [];
    const records = parseCsv(fs.readFileSync(file, 'utf8'));
    const rows = [];
    for (let n = 0; n < records.length; n++) {
        const r = records[n];
        if (r.length === 1 && String(r[0]).trim() === '') continue;
        if (r[0] === 'issuedAt') continue; // 表头（含容错：出现在任意行都跳过）
        rows.push({
            issuedAt: r[0] || '',
            key: r[1] || '',
            login: r[2] || '',
            plan: r[3] || '',
            seats: r[4] || '',
            expiresAt: r[5] || '',
            note: r.length > 6 ? r.slice(6).join(',') : (r[6] || '')
        });
    }
    return rows;
}

/**
 * 追加行（与 gen-key 的 appendLedger 同语义：新文件写表头、CRLF、整批同一时间戳）。
 * @param {string} file 台账路径
 * @param {Array<{key:string,login:string,plan:string,seats:number,expiresAt:string}>} rows
 * @param {string} [note] 本批备注（客户名/订单号）
 */
function appendRows(file, rows, note) {
    const dir = path.dirname(file);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    const isNew = !fs.existsSync(file);
    const stamp = new Date().toISOString();
    const lines = rows.map(r => [
        stamp, r.key, r.login, r.plan, r.seats, r.expiresAt, note || ''
    ].map(csvCell).join(','));
    fs.appendFileSync(file, (isNew ? HEADER + '\r\n' : '') + lines.join('\r\n') + '\r\n', 'utf8');
    return file;
}

module.exports = { HEADER, csvCell, parseCsv, readLedger, appendRows };
