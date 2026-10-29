/**
 * 本地 OCR 每日免费额度。
 *
 * 背景（2026-10-21 用户决定）：OCR 原先完全无限制。两条通道其实都不花作者的钱
 * —— 本地引擎跑在用户机器上，云 OCR 用的是用户自己填的百度/腾讯 Key（BYOK）——
 * 所以限额不是为了止损，而是给付费一个理由：不设额度，付费版就少了最硬的差别。
 *
 * 规则：
 *   1. 仅对「未激活」用户计数；已激活不限。
 *   2. 云 OCR（用户自费调用自己的 Key）不计数 —— 见 CLOUD_ENGINES。
 *   3. 按每日 0 点重置（本地日期），次数记在 localStorage，不联网、不需要服务器。
 *   4. 额度耗尽直接拦截 OCR 调用，并引导购买。
 *
 * ⚠️ 诚实边界：这是**软限制**。计数在前端 localStorage，清站点数据即清零。
 *    与整条离线授权链一致——目的是挡住"顺手白嫖"，不是挡刻意破解。
 *    真要硬限制得像 activation 那样落到 VB / DPAPI，那需要重新编译 VSTO。
 */
(function (global) {
    'use strict';

    // ==================== 可调参数（只改这里） ====================
    var FREE_DAILY_LIMIT = 20;      // 未激活用户每日本地 OCR 次数；0 = 未激活完全不能用
    var STORAGE_KEY = 'SmartPilot.OcrQuota.v1';
    var CLOUD_ENGINES = ['cloud'];  // 用户自费的自带 Key 通道，不占免费额度
    // 计入额度的桥方法：一次「启动识别」= 一次。
    // 注意 PDF 不按文档算——面板把 PDF 渲染成图片后逐页走 recognizeImage
    // （每页一次 StartOcrRecognitionWithFormat）；批量识别每文件一次。
    // （StartPdfOcrRecognition / RecognizeCurrentImage 已于 2026-10-21 删除：
    //   全 UI 零调用点，留在清单里只会让"计费口径"撒谎。）
    var COUNTED_METHODS = [
        'StartOcrRecognition',             // 聊天发图（image-pipeline.js）
        'StartOcrRecognitionWithFormat'    // OCR 工作台单张/区域/PDF 逐页 + 批量识别
    ];
    // ===========================================================

    var licensed = false;   // 由 license.js 在刷新权益时写入
    var memFallback = null; // localStorage 不可用时的内存兜底

    function todayStr(d) {
        var dt = d || new Date();
        var m = String(dt.getMonth() + 1);
        var day = String(dt.getDate());
        return dt.getFullYear() + '-' + (m.length < 2 ? '0' + m : m) + '-' + (day.length < 2 ? '0' + day : day);
    }

    function readRaw() {
        try {
            if (memFallback) return memFallback;
            var raw = global.localStorage && global.localStorage.getItem(STORAGE_KEY);
            return raw ? JSON.parse(raw) : null;
        } catch (_) {
            return null;
        }
    }

    function writeRaw(obj) {
        try {
            if (global.localStorage) global.localStorage.setItem(STORAGE_KEY, JSON.stringify(obj));
            else memFallback = obj;
        } catch (_) {
            memFallback = obj;
        }
    }

    /** 取今日计数；跨天自动归零 */
    function state() {
        var today = todayStr();
        var s = readRaw();
        if (!s || s.date !== today) return { date: today, used: 0 };
        var used = parseInt(s.used, 10);
        return { date: today, used: Number.isFinite(used) && used > 0 ? used : 0 };
    }

    function saveState(s) { writeRaw({ date: s.date, used: s.used }); }

    /** 该方法是否计入配额：非云引擎 + 是启动识别类方法 */
    function countFor(method, engine) {
        if (COUNTED_METHODS.indexOf(String(method || '')) < 0) return 0;
        if (engine && CLOUD_ENGINES.indexOf(String(engine).toLowerCase()) >= 0) return 0;
        return 1;
    }

    function setLicensed(v) {
        var changed = licensed !== (v === true);
        licensed = v === true;
        if (changed) renderBadge();
    }

    function status() {
        var s = state();
        var metered = !licensed;
        var remaining = metered ? Math.max(0, FREE_DAILY_LIMIT - s.used) : Infinity;
        return {
            metered: metered,
            used: s.used,
            limit: FREE_DAILY_LIMIT,
            remaining: remaining,
            date: s.date
        };
    }

    /** 能否再跑 n 次（只探测，不计数） */
    function canRun(n) {
        n = parseInt(n, 10) || 1;
        if (!licensed) {
            var s = state();
            if (FREE_DAILY_LIMIT <= 0 || s.used + n > FREE_DAILY_LIMIT) return false;
        }
        return true;
    }

    /** 真正扣额度 */
    function commit(n) {
        n = parseInt(n, 10) || 1;
        if (licensed) return status();
        var s = state();
        s.used += n;
        saveState(s);
        return status();
    }

    /**
     * 退还额度：识别启动失败或桥调用异常时调用。
     * 语义：只有「真正启动成功的识别」才计 1 次——失败不该花用户的额度。
     * 传参与 gate 相同（method + engine），内部重判这次是否真扣过：
     * 云通道 / 未登记方法本来就没扣，乱退会把本地余额扣虚，所以是 no-op。
     */
    function refund(opts) {
        opts = opts || {};
        var n = countFor(opts.method, opts.engine);
        if (n <= 0 || licensed) return status();
        var s = state();
        s.used = Math.max(0, s.used - n);
        saveState(s);
        renderBadge();   // 余额实时回跳
        return status();
    }

    /** 一条龙：判断是否放行（含一次性的超限提示）。返回 true = 允许 */
    function gate(opts) {
        opts = opts || {};
        var n = countFor(opts.method, opts.engine);
        if (n <= 0) return true;              // 不计入的方法/云通道
        if (!canRun(n)) {
            notifyExhausted();
            return false;
        }
        commit(n);
        renderBadge();   // 余额实时跟着变
        return true;
    }

    var LAST_NOTIFY_AT = 0;

    /** 超限提示（4 秒内防抖，避免批量任务刷屏） */
    function notifyExhausted() {
        if (Date.now() - LAST_NOTIFY_AT < 4000) return;
        LAST_NOTIFY_AT = Date.now();
        var msg = '今日免费额度已用完（' + FREE_DAILY_LIMIT + ' 次/天），明天自动重置；激活后不限次数。';
        try {
            if (global.App && global.App.i18n && typeof global.App.i18n.t === 'function') {
                var got = global.App.i18n.t('ocr.quota.exhausted', { limit: FREE_DAILY_LIMIT });
                if (got && got !== 'ocr.quota.exhausted') msg = got;
            }
        } catch (_) { /* ignore */ }
        try {
            if (global.App && typeof global.App.showToast === 'function') global.App.showToast(msg, 'warning');
            else if (typeof global.showToast === 'function') global.showToast(msg, 'warning');
        } catch (_) { /* ignore */ }
        try {
            if (typeof global.openAccountPanel === 'function') {
                setTimeout(global.openAccountPanel, 120);
            }
        } catch (_) { /* ignore */ }
    }

    function reset() {
        try {
            if (global.localStorage) global.localStorage.removeItem(STORAGE_KEY);
        } catch (_) { /* ignore */ }
        memFallback = null;
    }

    // ---------------- 余额展示（未激活时才计数） ----------------

    function pick(key, fallback, vars) {
        try {
            if (global.App && global.App.i18n && typeof global.App.i18n.t === 'function') {
                var got = global.App.i18n.t(key, vars);
                if (got && got !== key) return got;
            }
        } catch (_) { /* ignore */ }
        return fallback.replace(/\{(\w+)\}/g, function (m, k) {
            return vars && vars[k] !== undefined ? vars[k] : m;
        });
    }

    /** 当前引擎：优先缓存，其次设置页下拉 */
    function currentEngine() {
        try {
            if (global.__currentOcrEngine) return String(global.__currentOcrEngine);
            var sel = global.document && global.document.getElementById('ocr-engine-select');
            if (sel && sel.value) return String(sel.value);
        } catch (_) { /* ignore */ }
        return '';
    }

    function badgeText(engine) {
        if (licensed) return pick('ocr.quota.unlimited', '已激活 · 不限次数');
        if (engine && CLOUD_ENGINES.indexOf(engine.toLowerCase()) >= 0) {
            return pick('ocr.quota.cloudFree', '云端识别不计额度');
        }
        var s = status();
        return pick('ocr.quota.badge', '今日剩余 {remaining}/{limit}', { remaining: s.remaining, limit: FREE_DAILY_LIMIT });
    }

    /** 刷新页面上所有 id=ocr-quota-badge 的元素；没这个元素就什么都不做 */
    function renderBadge() {
        var doc = global.document;
        if (!doc) return null;
        var el = doc.getElementById('ocr-quota-badge');
        if (!el) return null;
        var engine = currentEngine();
        var metered = !licensed && !(engine && CLOUD_ENGINES.indexOf(engine.toLowerCase()) >= 0);
        var s = metered ? status() : null;
        el.textContent = badgeText(engine);
        el.className = 'ocr-quota-badge' + (metered && s && s.remaining <= 3 ? ' is-low' : (metered ? '' : ' is-free'));
        el.style.display = '';
        return el.textContent;
    }

    var api = {
        LIMIT: FREE_DAILY_LIMIT,
        STORAGE_KEY: STORAGE_KEY,
        COUNTED_METHODS: COUNTED_METHODS,
        CLOUD_ENGINES: CLOUD_ENGINES,
        todayStr: todayStr,
        setLicensed: setLicensed,
        isMetered: function () { return !licensed; },
        countFor: countFor,
        status: status,
        canRun: canRun,
        commit: commit,
        refund: refund,
        gate: gate,
        reset: reset,
        renderBadge: renderBadge,
        badgeText: badgeText
    };

    global.SmartPilotOcrQuota = api;

    // 页面就绪后刷新余额；每分钟重刷一次（跨零点/跨天后自动回到满额）
    function startBadgeClock() {
        renderBadge();
        try { global.setInterval(renderBadge, 60000); } catch (_) { /* ignore */ }
    }
    try {
        if (global.document) {
            if (global.document.readyState === 'loading') {
                global.document.addEventListener('DOMContentLoaded', startBadgeClock);
            } else {
                startBadgeClock();
            }
            // 切引擎时立刻重算（云通道不计额度）
            setTimeout(function () {
                var sel = global.document.getElementById('ocr-engine-select');
                if (sel) sel.addEventListener('change', function () { renderBadge(); });
            }, 0);
        }
    } catch (_) { /* ignore */ }

    if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
