'use strict';
/**
 * 本地 OCR 每日免费额度。
 *
 * 2026-10-21 新增。产品口径（用户拍板）：
 *   未激活 20 次/天 → 超限硬拦截并引导购买；云 OCR（用户自带 Key）不计数；已激活不限。
 *
 * 诚实边界：计数在 localStorage，是**软限制**（清站点数据即清零），
 * 与整条离线授权链一致——挡白嫖，不挡刻意破解。
 *
 * 2026-10-21 补丁（用户反馈"识别次数好像是假的"）：
 *   1. OCR 工作台 recognizeImage 直接调桥绕过闸门 → 工作台所有识别
 *      （单张/区域/PDF 逐页/截图）完全不计数，是"次数不涨"的根因；
 *   2. "放行即扣"导致启动失败也白扣 → 增 refund()，失败路径全额退还；
 *   3. 计数清单删掉两个零调用点的死条目（StartPdfOcrRecognition /
 *      RecognizeCurrentImage），PDF 实际是渲染成图后逐页识别。
 */
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const MODULE_PATH = path.join(ROOT, 'UI', 'js', 'ocr-quota.js');

/** 造一个干净的模块实例 + 假 localStorage */
function fresh(initial) {
    const store = new Map();
    const ls = {
        getItem: (k) => (store.has(k) ? store.get(k) : null),
        setItem: (k, v) => store.set(k, String(v)),
        removeItem: (k) => store.delete(k)
    };
    const prevLs = globalThis.localStorage;
    const prevSp = globalThis.SmartPilotOcrQuota;
    Object.defineProperty(globalThis, 'localStorage', { value: ls, configurable: true, writable: true });
    delete require.cache[MODULE_PATH];
    const q = require(MODULE_PATH);
    if (initial !== undefined) store.set(q.STORAGE_KEY, initial);
    return {
        q,
        store,
        restore() {
            if (prevLs === undefined) delete globalThis.localStorage;
            else Object.defineProperty(globalThis, 'localStorage', { value: prevLs, configurable: true, writable: true });
            globalThis.SmartPilotOcrQuota = prevSp;
            delete require.cache[MODULE_PATH];
        },
        readRaw() { return store.get(q.STORAGE_KEY); }
    };
}

function runN(q, n, engine) {
    let ok = 0;
    for (let i = 0; i < n; i++) {
        if (q.gate({ method: 'StartOcrRecognition', engine: engine || '' })) ok++;
    }
    return ok;
}

test('未激活：额度内全部放行，用完后硬拦截且不继续累加', () => {
    const f = fresh();
    try {
        const q = f.q;
        assert.strictEqual(q.LIMIT, 20, '免费额度应为 20/天');
        assert.strictEqual(q.isMetered(), true, '默认未激活 → 计费');
        assert.strictEqual(runN(q, q.LIMIT), q.LIMIT, '额度内应全部放行');
        assert.strictEqual(q.status().remaining, 0);
        assert.strictEqual(q.gate({ method: 'StartOcrRecognition', engine: '' }), false, '第 21 次必须拦截');
        // 拦截后不得继续累加（否则用户明天一上来就被历史超额吃掉）
        assert.strictEqual(q.status().used, q.LIMIT);
    } finally { f.restore(); }
});

test('跨天自动重置', () => {
    const f = fresh(JSON.stringify({ date: '2020-01-01', used: 999 }));
    try {
        const q = f.q;
        assert.strictEqual(q.status().used, 0, '日期变了应从 0 开始');
        assert.strictEqual(q.gate({ method: 'StartOcrRecognition', engine: '' }), true, '新的一天应有额度');
        assert.strictEqual(q.status().used, 1);
    } finally { f.restore(); }
});

test('云 OCR（用户自带 Key）不占免费额度', () => {
    const f = fresh();
    try {
        const q = f.q;
        for (let i = 0; i < 50; i++) {
            assert.strictEqual(q.gate({ method: 'StartOcrRecognition', engine: 'cloud' }), true, '云通道应始终放行');
        }
        assert.strictEqual(q.status().used, 0, '云通道不得计数');
    } finally { f.restore(); }
});

test('轮询类方法（Get*）不计数，避免把一次识别算成几十次', () => {
    const f = fresh();
    try {
        const q = f.q;
        const polls = ['GetOcrResult', 'GetOcrResultWithFormat', 'GetOcrStructuredResult', 'GetOcrElementsJson', 'SaveBase64Image'];
        polls.forEach((m) => {
            for (let i = 0; i < 30; i++) {
                assert.strictEqual(q.gate({ method: m, engine: '' }), true, m + ' 不应受限');
            }
        });
        assert.strictEqual(q.status().used, 0);
    } finally { f.restore(); }
});

test('已激活：不限次', () => {
    const f = fresh();
    try {
        const q = f.q;
        q.setLicensed(true);
        assert.strictEqual(q.isMetered(), false);
        assert.strictEqual(runN(q, 200), 200, '已激活不应有上限');
        assert.strictEqual(q.status().remaining, Infinity);
    } finally { f.restore(); }
});

test('退还额度：只退真扣过的（云通道/未登记方法不乱退，不会退成负数）', () => {
    const f = fresh();
    try {
        const q = f.q;
        q.gate({ method: 'StartOcrRecognition', engine: '' });
        q.gate({ method: 'StartOcrRecognition', engine: '' });
        assert.strictEqual(q.status().used, 2);
        // 启动失败退一次真扣过的
        q.refund({ method: 'StartOcrRecognition', engine: '' });
        assert.strictEqual(q.status().used, 1, '启动失败应退还');
        // 云通道没扣过，退它必须是 no-op（不能把本地余额扣虚）
        q.refund({ method: 'StartOcrRecognition', engine: 'cloud' });
        assert.strictEqual(q.status().used, 1, '云通道乱退不得影响余额');
        // 未登记方法同理
        q.refund({ method: 'GetOcrResult', engine: '' });
        assert.strictEqual(q.status().used, 1, '未登记方法乱退必须是 no-op');
        // 连退多次不会退成负数
        q.refund({ method: 'StartOcrRecognition', engine: '' });
        q.refund({ method: 'StartOcrRecognition', engine: '' });
        q.refund({ method: 'StartOcrRecognition', engine: '' });
        assert.strictEqual(q.status().used, 0, '余额不会退成负数');
    } finally { f.restore(); }
});

test('已激活时退还为 no-op', () => {
    const f = fresh();
    try {
        const q = f.q;
        q.setLicensed(true);
        q.gate({ method: 'StartOcrRecognition', engine: '' });
        q.refund({ method: 'StartOcrRecognition', engine: '' });
        assert.strictEqual(q.status().used, 0, '已激活本就没计数，退什么都是 0');
    } finally { f.restore(); }
});

test('计数清单只含真实调用点（删死条目，防止计费口径撒谎）', () => {
    const f = fresh();
    try {
        const q = f.q;
        // 全 UI 真实调用只有这两个"启动识别"方法
        assert.deepStrictEqual(q.COUNTED_METHODS.slice().sort(),
            ['StartOcrRecognition', 'StartOcrRecognitionWithFormat']);
        // 死条目不得回归（2026-10-21 删：UI 零调用点）
        assert.ok(q.COUNTED_METHODS.indexOf('StartPdfOcrRecognition') < 0,
            'PDF 实际是渲染成图后逐页识别，没有这个方法调用');
        assert.ok(q.COUNTED_METHODS.indexOf('RecognizeCurrentImage') < 0, 'UI 无此调用点');
    } finally { f.restore(); }
});

test('localStorage 损坏/不可用时不崩，且降级到内存计数', () => {
    const prev = globalThis.localStorage;
    // 1) 坏 JSON
    Object.defineProperty(globalThis, 'localStorage', {
        value: { getItem: () => '{not json', setItem: () => {}, removeItem: () => {} },
        configurable: true, writable: true
    });
    delete require.cache[MODULE_PATH];
    let q = require(MODULE_PATH);
    assert.strictEqual(q.status().used, 0, '坏数据应按空计数处理');
    assert.strictEqual(q.gate({ method: 'StartOcrRecognition', engine: '' }), true, '坏数据不应阻塞用户');

    // 2) 完全没有 localStorage
    delete globalThis.localStorage;
    delete require.cache[MODULE_PATH];
    q = require(MODULE_PATH);
    assert.strictEqual(q.gate({ method: 'StartOcrRecognition', engine: '' }), true, '无 localStorage 也应能用');
    assert.strictEqual(q.status().used, 1, '应降级到内存计数');

    if (prev === undefined) delete globalThis.localStorage;
    else Object.defineProperty(globalThis, 'localStorage', { value: prev, configurable: true, writable: true });
    delete require.cache[MODULE_PATH];
});

test('余额角标：本机显示剩余次数，云通道 / 已激活显示免计', () => {
    const f = fresh();
    try {
        const q = f.q;
        assert.strictEqual(q.badgeText(''), '今日剩余 20/20');
        q.gate({ method: 'StartOcrRecognition', engine: '' });
        q.gate({ method: 'StartOcrRecognition', engine: '' });
        assert.strictEqual(q.badgeText(''), '今日剩余 18/20', '用一次少一次');
        q.setLicensed(true);
        assert.strictEqual(q.badgeText(''), '已激活 · 不限次数');
        q.setLicensed(false);
        assert.strictEqual(q.badgeText('cloud'), '云端识别不计额度');
    } finally { f.restore(); }
});

test('余额角标渲染：剩 3 次起警示色，免计显绿色', () => {
    const f = fresh();
    const prevDoc = globalThis.document;
    const el = { textContent: '', className: '', style: {} };
    // 注意：要在 require 之后挂 document，否则模块加载时的定时器会让 Node 事件循环挂着不放
    globalThis.document = {
        readyState: 'complete',
        getElementById: (id) => (id === 'ocr-quota-badge' ? el : null),
        addEventListener: () => {}
    };
    try {
        const q = f.q;
        q.renderBadge();
        assert.strictEqual(el.textContent, '今日剩余 20/20');
        assert.strictEqual(el.className, 'ocr-quota-badge', '初始应无附加样式');
        for (let i = 0; i < 17; i++) q.gate({ method: 'StartOcrRecognition', engine: '' });
        assert.strictEqual(el.textContent, '今日剩余 3/20');
        assert.ok(el.className.indexOf('is-low') >= 0, '剩 3 次应变警示色');
        q.setLicensed(true);
        assert.ok(el.className.indexOf('is-free') >= 0, '已激活应为绿色（setLicensed 自动重渲染）');
        assert.strictEqual(el.textContent, '已激活 · 不限次数');
    } finally {
        f.restore();
        if (prevDoc === undefined) delete globalThis.document;
        else globalThis.document = prevDoc;
    }
});

test('守栏：四条识别路径全部过闸门，且失败路径退还额度', () => {
    // 2026-10-21 修复"识别次数是假的"：工作台 recognizeImage 直接调桥绕过闸门
    // （计数主路径没接，用户看到次数不涨）；三个包装层只接闸门不接退还
    // （启动失败也白扣）。四个文件都必须既有闸门又有退还。
    const files = [
        path.join(ROOT, 'UI', 'js', 'app-ocr.js'),
        path.join(ROOT, 'UI', 'js', 'ocr-batch-ui.js'),
        path.join(ROOT, 'UI', 'js', 'chat-v2', 'pipelines', 'image-pipeline.js'),
        path.join(ROOT, 'UI', 'ocr-panel.html')
    ];
    files.forEach((file) => {
        const src = fs.readFileSync(file, 'utf8');
        assert.ok(src.indexOf('SmartPilotOcrQuota') >= 0,
            path.basename(file) + ' 必须接入 OCR 额度闸门');
        assert.ok(src.indexOf('OCR_QUOTA_EXCEEDED') >= 0,
            path.basename(file) + ' 拦截时应抛出统一错误码');
        assert.ok(src.indexOf('.refund(') >= 0,
            path.basename(file) + ' 失败路径必须退还额度（只为真正启动的识别计费）');
    });

    // 闸门必须出现在 recognizeImage 里、且早于真正的桥调用
    const panel = fs.readFileSync(path.join(ROOT, 'UI', 'ocr-panel.html'), 'utf8');
    const fnStart = panel.indexOf('async function recognizeImage');
    assert.ok(fnStart >= 0, '应有 recognizeImage');
    const nextAsync = panel.indexOf('\n        async function ', fnStart + 10);
    const nextSync = panel.indexOf('\n        function ', fnStart + 10);
    let fnEnd = panel.length;
    if (nextAsync > 0) fnEnd = Math.min(fnEnd, nextAsync);
    if (nextSync > 0) fnEnd = Math.min(fnEnd, nextSync);
    const body = panel.slice(fnStart, fnEnd);
    const gateIdx = body.indexOf('SmartPilotOcrQuota.gate');
    const callIdx = body.indexOf('StartOcrRecognitionWithFormat(saveObj.path');
    assert.ok(gateIdx >= 0, 'recognizeImage 必须过闸门（工作台所有识别的总入口）');
    assert.ok(callIdx > gateIdx, '闸门必须在启动识别之前，否则等于没拦');
    assert.ok(/refund\(\{/.test(body), 'recognizeImage 启动失败必须退还额度');
});

test('守栏：加载 OCR 的两个页面都得先引 ocr-quota.js', () => {
    ['UI/index.html', 'UI/ocr-panel.html'].forEach((rel) => {
        const html = fs.readFileSync(path.join(ROOT, rel), 'utf8');
        assert.ok(html.indexOf('js/ocr-quota.js') >= 0, rel + ' 应引入 ocr-quota.js');
        // 必须在调用它的模块之前出现
        const quotaIdx = html.indexOf('js/ocr-quota.js');
        const consumerIdx = html.indexOf('js/app-ocr.js') >= 0
            ? html.indexOf('js/app-ocr.js')
            : html.indexOf('js/ocr-batch-ui.js');
        assert.ok(consumerIdx >= 0, rel + ' 应有 OCR 消费模块');
        assert.ok(quotaIdx < consumerIdx, rel + ' 中 ocr-quota.js 必须先加载');
    });
});

test('守栏：OCR 工作台页面也必须能拿到激活状态', () => {
    // 2026-10-21：OCR 工作台（ocr-panel.html）没有账户面板 DOM，而 license.js 原先只在
    // 有 #license-card / #side-account-btn 时才刷新权益 → 已付费用户在那里会被当未激活白白限额。
    // 必须有 syncOcrQuota()，且它不能被包在那个 if 里。
    const lic = fs.readFileSync(path.join(ROOT, 'UI', 'js', 'license.js'), 'utf8');
    assert.ok(lic.indexOf('async function syncOcrQuota') >= 0, 'license.js 应提供 syncOcrQuota');
    const idxInit = lic.indexOf("document.addEventListener('DOMContentLoaded'");
    assert.ok(idxInit > 0, '应有 DOMContentLoaded 初始化');
    const init = lic.slice(idxInit);
    assert.ok(/syncOcrQuota\(\)\.catch/.test(init), '初始化时应无条件调用 syncOcrQuota');
    const ifIdx = init.indexOf('if (staticCard || document.getElementById');
    const callIdx = init.indexOf('syncOcrQuota()');
    const ifEnd = init.indexOf('}', init.indexOf('ChatLog.warn?.(\'LICENSE\'', ifIdx));
    assert.ok(callIdx > ifEnd, 'syncOcrQuota 必须在 DOM 判断之外，否则 OCR 页拿不到激活状态');

    // OCR 工作台的加载顺序：额度模块先于消费模块，也先于 license.js
    const panel = fs.readFileSync(path.join(ROOT, 'UI', 'ocr-panel.html'), 'utf8');
    assert.ok(panel.indexOf('ocr-quota-badge') >= 0, 'OCR 工作台应有余额角标容器');
    assert.ok(panel.indexOf('js/license.js') >= 0, 'OCR 工作台应加载 license.js 才能同步激活状态');
    assert.ok(panel.indexOf('js/ocr-quota.js') < panel.indexOf('js/license.js'),
        'ocr-quota.js 必须比 license.js 先加载');
});

test('守栏：文案中英齐平，且 license.js 内置字典里有备份', () => {
    const i18n = fs.readFileSync(path.join(ROOT, 'UI', 'js', 'i18n.js'), 'utf8');
    const zhCount = (i18n.match(/'ocr\.quota\.exhausted'/g) || []).length;
    assert.ok(zhCount >= 2, 'i18n 中英各一处必须有 ocr.quota.exhausted，实际 ' + zhCount);
    const lic = fs.readFileSync(path.join(ROOT, 'UI', 'js', 'license.js'), 'utf8');
    assert.ok(lic.indexOf('ocr.quota.exhausted') >= 0, 'license.js 内置字典应备份该文案');
});
