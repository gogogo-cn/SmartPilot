/**
 * ============================================================
 * pomodoro-app.js — 独立番茄钟应用（HTTP 双向同步 / 本地兜底）
 * ============================================================
 * 功能：
 *   1. 与 Excel 内嵌番茄钟通过 /api/pomodoro/* 双向同步（session token 鉴权）
 *   2. Excel 服务不可用时降级 localStorage 离线模式
 *   3. 完整计时逻辑（专注/短休/长休）、设置面板、提示音、桌面通知
 *
 * 同步协议（与 app-history.js SharedPomodoro 保持一致）：
 *   data = { runtime, counters, config }
 *   runtime = { isRunning, isPaused, isBreak, minutes, seconds, remainingSeconds,
 *               sessionCount, maxSessions, updatedAt }
 * ============================================================
 */
(function () {
    'use strict';

    function tt(key, vars) {
        try {
            if (window.App && App.i18n && typeof App.i18n.t === 'function') return App.i18n.t(key, vars);
            if (typeof window.t === 'function') return window.t(key, vars);
        } catch (e) { /* ignore */ }
        return key;
    }
    function applyI18n() {
        try {
            if (window.App && App.i18n) {
                App.i18n.apply(document);
                document.title = tt('pomo.titleStandalone');
            }
        } catch (e) { /* ignore */ }
    }

    function knownUiLocale(loc) {
        try {
            if (window.App && App.i18n && typeof App.i18n.resolveLocale === 'function') return App.i18n.resolveLocale(loc);
        } catch (e) { /* ignore */ }
        return (loc === 'en' || loc === 'zh') ? loc : null;
    }

    async function syncLocaleFromHost() {
        try {
            if (window.opener && window.opener.App && window.opener.App.i18n && typeof window.opener.App.i18n.getLocale === 'function') {
                var fromOpener = knownUiLocale(window.opener.App.i18n.getLocale());
                if (fromOpener && window.App && App.i18n) {
                    App.i18n.setLocale(fromOpener, { syncHost: false, reloadWelcome: false, reloadHistory: false });
                    return;
                }
            }
        } catch (e1) { /* ignore */ }
        try {
            var bridge = null;
            if (window.chrome && chrome.webview && chrome.webview.hostObjects) {
                bridge = chrome.webview.hostObjects.excelBridge;
            }
            if (!bridge && window.opener && window.opener.chrome && window.opener.chrome.webview &&
                window.opener.chrome.webview.hostObjects) {
                bridge = window.opener.chrome.webview.hostObjects.excelBridge;
            }
            if (!bridge || typeof bridge.GetUiLocale !== 'function') return;
            var raw = await bridge.GetUiLocale();
            var j = typeof raw === 'string' ? JSON.parse(raw) : raw;
            var hostLoc = j && knownUiLocale(j.locale);
            if (hostLoc && window.App && App.i18n) {
                App.i18n.setLocale(hostLoc, { syncHost: false, reloadWelcome: false, reloadHistory: false });
            }
        } catch (e2) { /* ignore */ }
    }


    var HTTP = (typeof location !== 'undefined' && location.protocol === 'http:')
        ? location.origin
        : 'http://localhost:9528';
    var WRITER_ID = 'standalone-' + Math.random().toString(36).slice(2, 10);
    var KEY = 'shared_pomodoro_v1';
    var CFG_KEY = 'pomodoro_config_v1';
    var COUNTERS_KEY = 'pomodoro_counters_v1';

    var token = '';
    var knownVersion = -1;
    var httpOk = false;
    var pollTimer = null;

    var cfg = loadCfg();
    var counters = loadCounters();
    var state = {
        minutes: cfg.workDur | 0,
        seconds: 0,
        isRunning: false,
        isBreak: false,
        sessionCount: 0,
        maxSessions: cfg.longEvery | 0,
        timerId: null,
        breakTotal: 0
    };

    // ---------- 基础工具 ----------
    function $(id) { try { return document.getElementById(id); } catch (_) { return null; } }
    function todayKey() { return new Date().toISOString().slice(0, 10); }

    function loadCfg() {
        var def = { workDur: 25, shortDur: 5, longDur: 15, longEvery: 4, sound: true, notif: true };
        try {
            var s = localStorage.getItem(CFG_KEY);
            if (s) { var c = JSON.parse(s); if (c && typeof c === 'object') return Object.assign(def, c); }
        } catch (_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('POMODORO', '操作失败:', _ && _.message); }
        return def;
    }
    function saveCfg() { try { localStorage.setItem(CFG_KEY, JSON.stringify(cfg)); } catch (_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('POMODORO', '操作失败:', _ && _.message); } }

    function loadCounters() {
        var c = { completedPomodoros: 0, completedMinutes: 0, todayPomodoros: 0, currentCycle: 0, todayKey: todayKey() };
        try {
            var s = localStorage.getItem(COUNTERS_KEY);
            if (s) c = Object.assign(c, JSON.parse(s) || {});
        } catch (_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('POMODORO', '操作失败:', _ && _.message); }
        if (c.todayKey !== todayKey()) { c.todayPomodoros = 0; c.todayKey = todayKey(); }
        return c;
    }
    function saveCounters() { try { localStorage.setItem(COUNTERS_KEY, JSON.stringify(counters)); } catch (_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('POMODORO', '操作失败:', _ && _.message); } }

    // ---------- session token ----------
    function getToken() {
        return new Promise(function (resolve) {
            if (token) return resolve(token);
            try {
                if (window.ExcelBridge && typeof window.ExcelBridge.GetMobileControlStatus === 'function') {
                    var r = window.ExcelBridge.GetMobileControlStatus();
                    if (r) {
                        var j = JSON.parse(r);
                        if (j && j.sessionToken) { token = j.sessionToken; return resolve(token); }
                    }
                }
            } catch (_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('POMODORO', '操作失败:', _ && _.message); }
            try {
                var ctrl = new AbortController();
                var t = setTimeout(function () { ctrl.abort(); }, 1500);
                fetch(HTTP + '/api/status', { cache: 'no-store', signal: ctrl.signal })
                    .then(function (res) { clearTimeout(t); return res.json(); })
                    .then(function (j) { if (j && j.sessionToken) token = j.sessionToken; resolve(token); })
                    .catch(function () { clearTimeout(t); resolve(''); });
            } catch (_) { resolve(''); }
        });
    }
    function apiUrl(path) { return HTTP + path + '?token=' + encodeURIComponent(token || ''); }

    // ---------- HTTP API ----------
    async function httpGet() {
        try {
            await getToken();
            var r = await fetch(apiUrl('/api/pomodoro/get'), { cache: 'no-store' });
            if (!r.ok) return null;
            return await r.json();
        } catch (_) { return null; }
    }

    async function httpSet(payload) {
        try {
            await getToken();
            var body = JSON.stringify(Object.assign({ writerId: WRITER_ID, fromVersion: knownVersion }, payload));
            var r = await fetch(apiUrl('/api/pomodoro/set'), {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: body,
                cache: 'no-store'
            });
            if (!r.ok) return null;
            return await r.json();
        } catch (_) { return null; }
    }

    async function httpPoll() {
        try {
            await getToken();
            var q = knownVersion < 0 ? -1 : knownVersion;
            var r = await fetch(apiUrl('/api/pomodoro/poll?sinceVersion=' + q), { cache: 'no-store' });
            if (!r.ok) return null;
            return await r.json();
        } catch (_) { return null; }
    }

    // ---------- 本地兜底存储 ----------
    function readLocal() {
        try {
            var s = localStorage.getItem(KEY);
            if (!s) return null;
            var j = JSON.parse(s);
            return (j && typeof j === 'object') ? j : null;
        } catch (_) { return null; }
    }
    function writeLocal(payload) {
        try {
            var cur = readLocal() || {};
            var next = Object.assign({}, cur, payload, {
                version: (cur.version | 0) + 1,
                lastWriterId: payload.writerId || WRITER_ID,
                updatedAt: Date.now()
            });
            localStorage.setItem(KEY, JSON.stringify(next));
            knownVersion = next.version | 0;
        } catch (_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('POMODORO', '操作失败:', _ && _.message); }
    }
    function writeRemoteToLocal(resp) {
        try {
            localStorage.setItem(KEY, JSON.stringify({
                version: resp.version | 0,
                lastWriterId: resp.lastWriterId || '',
                data: resp.data,
                updatedAt: Date.now()
            }));
        } catch (_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('POMODORO', '操作失败:', _ && _.message); }
    }

    // ---------- 共享数据序列化（与内嵌端一致） ----------
    function toShared() {
        var runtime = {
            isRunning: !!state.isRunning,
            isPaused: !state.isRunning && ((state.minutes | 0) !== (cfg.workDur | 0) || (state.seconds | 0) !== 0),
            isBreak: !!state.isBreak,
            minutes: state.minutes | 0,
            seconds: state.seconds | 0,
            remainingSeconds: (state.minutes | 0) * 60 + (state.seconds | 0),
            sessionCount: state.sessionCount | 0,
            maxSessions: state.maxSessions | 0,
            updatedAt: Date.now()
        };
        var counters2 = {
            completedPomodoros: counters.completedPomodoros | 0,
            completedMinutes: counters.completedMinutes | 0,
            todayPomodoros: counters.todayPomodoros | 0,
            currentCycle: counters.currentCycle | 0,
            todayKey: counters.todayKey || todayKey()
        };
        return { runtime: runtime, counters: counters2, config: Object.assign({}, cfg) };
    }

    function applyShared(shared) {
        try {
            if (!shared) return;
            if (shared.config) {
                var merged = Object.assign({}, cfg, shared.config);
                var changed = JSON.stringify(merged) !== JSON.stringify(cfg);
                cfg = merged;
                if (changed) { saveCfg(); syncSettingsUI(); }
            }
            if (shared.counters) {
                counters.completedPomodoros = Math.max(counters.completedPomodoros | 0, shared.counters.completedPomodoros | 0);
                counters.completedMinutes = Math.max(counters.completedMinutes | 0, shared.counters.completedMinutes | 0);
                counters.todayKey = shared.counters.todayKey || counters.todayKey || todayKey();
                counters.todayPomodoros = Math.max(counters.todayPomodoros | 0, shared.counters.todayPomodoros | 0);
                counters.currentCycle = Math.max(counters.currentCycle | 0, shared.counters.currentCycle | 0);
                saveCounters();
            }
            if (shared.runtime) {
                var remRun = !!shared.runtime.isRunning && !shared.runtime.isPaused;
                var locRun = state.isRunning;
                if (!locRun) {
                    state.minutes = shared.runtime.minutes | 0;
                    state.seconds = shared.runtime.seconds | 0;
                    state.isBreak = !!shared.runtime.isBreak;
                    state.sessionCount = shared.runtime.sessionCount | 0;
                    state.maxSessions = shared.runtime.maxSessions | 0;
                    state.breakTotal = state.isBreak ? (state.minutes * 60 + state.seconds) : 0;
                    if (remRun) startTick();
                } else {
                    var rem = shared.runtime.remainingSeconds | 0;
                    var loc = state.minutes * 60 + state.seconds;
                    if (rem < loc) { state.minutes = Math.floor(rem / 60); state.seconds = rem % 60; }
                }
            }
            updateUI();
        } catch (_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('POMODORO', '操作失败:', _ && _.message); }
    }

    async function push(evt) {
        var data = toShared();
        writeLocal({ data: data, event: evt || '' });
        saveCounters();
        if (httpOk) {
            var resp = await httpSet({ data: data, event: evt || '' });
            if (resp && typeof resp.version === 'number') {
                knownVersion = resp.version | 0;
                if (resp.conflict && resp.data && resp.data.runtime) applyShared(resp.data);
            }
        }
    }

    // ---------- 轮询同步 ----------
    async function tickSync() {
        if (httpOk) {
            var resp = await httpPoll();
            if (!resp || resp.success === false) { setHttp(false); return; }
            if (resp.changed) {
                knownVersion = resp.version | 0;
                if (resp.lastWriterId !== WRITER_ID) {
                    if (resp.data && resp.data.runtime) applyShared(resp.data);
                    writeRemoteToLocal(resp);
                }
            }
        } else {
            var cur = readLocal();
            if (!cur) return;
            var v = cur.version | 0;
            if (v === knownVersion) return;
            knownVersion = v;
            if (cur.lastWriterId === WRITER_ID) return;
            if (cur.data && cur.data.runtime) applyShared(cur.data);
        }
    }

    async function initSync() {
        var tok = await getToken();
        var remote = null;
        if (tok) {
            try {
                var ctrl = new AbortController();
                var t = setTimeout(function () { ctrl.abort(); }, 2000);
                var r = await fetch(apiUrl('/api/pomodoro/get'), { cache: 'no-store', signal: ctrl.signal });
                clearTimeout(t);
                if (r.ok) remote = await r.json();
            } catch (_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('POMODORO', '操作失败:', _ && _.message); }
        }
        if (remote && remote.success !== false && remote.data && Object.keys(remote.data).length > 0) {
            setHttp(true);
            knownVersion = remote.version | 0;
            applyShared(remote.data);
            writeRemoteToLocal(remote);
        } else {
            var local = readLocal();
            if (local && local.data && Object.keys(local.data).length > 0) {
                knownVersion = local.version | 0;
                applyShared(local.data);
                if (tok) { try { await httpSet({ data: local.data, event: 'sync_from_local' }); } catch (_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('POMODORO', '操作失败:', _ && _.message); } }
            }
        }
        pollTimer = setInterval(tickSync, 2000);
    }

    // ---------- 计时控制 ----------
    function startTick() {
        state.isRunning = true;
        if (state.timerId) clearInterval(state.timerId);
        state.timerId = setInterval(tick, 1000);
        var pb = $('pauseBtn');
        if (pb) { pb.textContent = tt('pomo.pause'); pb.disabled = false; }
        var sb = $('startBtn');
        if (sb) sb.disabled = true;
    }
    function stopTick() {
        state.isRunning = false;
        if (state.timerId) { clearInterval(state.timerId); state.timerId = null; }
    }

    function doStart() {
        if (state.isRunning) { toast(tt('pomo.toast.running'), 'warn'); return; }
        if (state.minutes === 0 && state.seconds === 0) {
            state.minutes = state.isBreak ? (cfg.shortDur | 0) : (cfg.workDur | 0);
            state.seconds = 0;
            state.isBreak = false;
            toast(tt('pomo.toast.resetStart'), 'info');
            updateUI();
            return;
        }
        startTick();
        push('start');
        toast(tt('pomo.toast.focusStart'), 'ok');
        updateUI();
    }

    function doPause() {
        if (!state.isRunning) { toast(tt('pomo.toast.notRunning'), 'warn'); return; }
        stopTick();
        var pb = $('pauseBtn');
        if (pb) { pb.textContent = tt('pomo.resume'); pb.disabled = true; }
        var sb = $('startBtn');
        if (sb) sb.disabled = false;
        push('pause');
        toast(tt('pomo.toast.paused'), 'info');
        updateUI();
    }

    function doReset() {
        stopTick();
        state.minutes = state.isBreak ? (cfg.shortDur | 0) : (cfg.workDur | 0);
        state.seconds = 0;
        state.isBreak = false;
        var pb = $('pauseBtn');
        if (pb) { pb.textContent = tt('pomo.pause'); pb.disabled = true; }
        var sb = $('startBtn');
        if (sb) sb.disabled = false;
        push('reset');
        toast(tt('pomo.toast.resetTo', { time: String(state.minutes).padStart(2, '0') + ':00' }), 'ok');
        updateUI();
    }

    function tick() {
        if (state.seconds > 0) {
            state.seconds--;
        } else if (state.minutes > 0) {
            state.minutes--;
            state.seconds = 59;
        } else {
            if (!state.isBreak) {
                state.sessionCount++;
                counters.currentCycle = Math.max(counters.currentCycle | 0, state.sessionCount);
                counters.completedPomodoros++;
                counters.todayPomodoros++;
                counters.completedMinutes += (cfg.workDur | 0) || 25;
                counters.todayKey = todayKey();
                state.isBreak = true;
                if (state.sessionCount >= state.maxSessions) {
                    state.minutes = (cfg.longDur | 0) || 15;
                    state.sessionCount = 0;
                    counters.currentCycle = 0;
                } else {
                    state.minutes = (cfg.shortDur | 0) || 5;
                }
                state.seconds = 0;
                state.breakTotal = state.minutes * 60;
                saveCounters();
                push('complete_focus');
                playSound();
                notify(tt('pomo.notify.focusTitle'), tt('pomo.notify.focusBody'));
                startTick();
            } else {
                state.isBreak = false;
                state.minutes = (cfg.workDur | 0) || 25;
                state.seconds = 0;
                push('complete_break');
                playSound();
                notify(tt('pomo.notify.breakTitle'), tt('pomo.notify.breakBody'));
                startTick();
            }
        }
        var remaining = state.minutes * 60 + state.seconds;
        if (remaining % 10 === 0) push('tick');
        updateUI();
    }

    // ---------- 提示音 & 桌面通知 ----------
    var audioCtx = null;
    function playSound() {
        try {
            if (!cfg.sound) return;
            if (!audioCtx) audioCtx = new (window.AudioContext || window.webkitAudioContext)();
            var ctx = audioCtx;
            var now = ctx.currentTime;
            [[880, 0], [660, 0.2]].forEach(function (pair) {
                var o = ctx.createOscillator();
                var g = ctx.createGain();
                o.type = 'sine';
                o.frequency.value = pair[0];
                g.gain.setValueAtTime(0.001, now + pair[1]);
                g.gain.exponentialRampToValueAtTime(0.28, now + pair[1] + 0.02);
                g.gain.exponentialRampToValueAtTime(0.001, now + pair[1] + 0.3);
                o.connect(g); g.connect(ctx.destination);
                o.start(now + pair[1]); o.stop(now + pair[1] + 0.32);
            });
        } catch (_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('POMODORO', '操作失败:', _ && _.message); }
    }

    function notify(title, body) {
        try {
            if (!cfg.notif) return;
            if (typeof Notification === 'undefined') return;
            if (Notification.permission === 'granted') {
                new Notification(title, { body: body });
            } else if (Notification.permission !== 'denied') {
                Notification.requestPermission().then(function (p) {
                    if (p === 'granted') new Notification(title, { body: body });
                });
            }
        } catch (_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('POMODORO', '操作失败:', _ && _.message); }
    }

    // ---------- UI ----------
    function updateUI() {
        var timerEl = $('timer');
        var statusEl = $('status');
        var sessionEl = $('session');
        if (timerEl) timerEl.textContent = String(state.minutes).padStart(2, '0') + ':' + String(state.seconds).padStart(2, '0');
        if (statusEl) statusEl.textContent = state.isBreak ? tt('pomo.break') : tt('pomo.focus');
        if (sessionEl) sessionEl.textContent = tt('pomo.session', { n: state.sessionCount, max: state.maxSessions });

        [timerEl, statusEl, $('startBtn'), $('pauseBtn'), $('resetBtn'), $('progress')].forEach(function (el) {
            if (el) el.classList.toggle('break-mode', !!state.isBreak);
        });

        var total = state.isBreak
            ? (state.breakTotal || (cfg.shortDur | 0) * 60)
            : ((cfg.workDur | 0) * 60);
        var remaining = state.minutes * 60 + state.seconds;
        var pct = total > 0 ? Math.max(0, Math.min(100, Math.round((1 - remaining / total) * 100))) : 0;
        var prog = $('progress');
        if (prog) prog.style.width = pct + '%';

        var st = $('statTotal'); if (st) st.textContent = counters.completedPomodoros | 0;
        var sm = $('statMinutes'); if (sm) sm.textContent = counters.completedMinutes | 0;
        var sd = $('statToday'); if (sd) sd.textContent = counters.todayPomodoros | 0;
    }

    function setHttp(ok) {
        httpOk = ok;
        var sb = $('syncBtn'), label = $('syncLabel');
        if (sb) { sb.classList.toggle('synced', !!ok); sb.classList.toggle('offline', !ok); }
        if (label) label.textContent = ok ? tt('pomo.synced') : tt('pomo.offline');
        setBanner(ok ? '' : tt('pomo.offlineBanner'), 'warn');
    }

    function setBanner(text, kind) {
        var b = $('banner');
        if (!b) return;
        if (!text) { b.className = 'banner'; b.textContent = ''; return; }
        b.className = 'banner show ' + (kind || 'warn');
        b.textContent = text;
    }

    async function manualSync() {
        var label = $('syncLabel'), sb = $('syncBtn');
        if (sb) { sb.classList.remove('synced', 'offline'); }
        if (label) label.textContent = tt('pomo.syncing');
        var tok = await getToken();
        if (tok) {
            var resp = await httpGet();
            if (resp && resp.success !== false) {
                setHttp(true);
                if (resp.lastWriterId !== WRITER_ID && resp.data && resp.data.runtime) {
                    knownVersion = resp.version | 0;
                    applyShared(resp.data);
                    writeRemoteToLocal(resp);
                }
                toast(tt('pomo.toast.synced'), 'ok');
                return;
            }
        }
        setHttp(false);
        toast(tt('pomo.toast.offlineMode'), 'warn');
    }

    // ---------- 设置面板 ----------
    function openSettings() {
        var w = $('workDur'), s = $('shortDur'), l = $('longDur'), e = $('longEvery');
        if (w) w.value = cfg.workDur | 0;
        if (s) s.value = cfg.shortDur | 0;
        if (l) l.value = cfg.longDur | 0;
        if (e) e.value = cfg.longEvery | 0;
        var sc = $('soundChk'); if (sc) sc.checked = !!cfg.sound;
        var nc = $('notifChk'); if (nc) nc.checked = !!cfg.notif;
        syncSettingsUI();
        var ov = $('overlay');
        if (ov) ov.classList.add('show');
    }
    function closeSettings() {
        var ov = $('overlay');
        if (ov) ov.classList.remove('show');
    }
    function syncSettingsUI() {
        var v1 = $('workDurVal'); if (v1) v1.textContent = tt('pomo.min', { n: (cfg.workDur | 0) });
        var v2 = $('shortDurVal'); if (v2) v2.textContent = tt('pomo.min', { n: (cfg.shortDur | 0) });
        var v3 = $('longDurVal'); if (v3) v3.textContent = tt('pomo.min', { n: (cfg.longDur | 0) });
        var v4 = $('longEveryVal'); if (v4) v4.textContent = tt('pomo.count', { n: (cfg.longEvery | 0) });
    }
    function saveSettings() {
        var oldWork = cfg.workDur | 0;
        cfg.workDur = parseInt($('workDur').value, 10) || 25;
        cfg.shortDur = parseInt($('shortDur').value, 10) || 5;
        cfg.longDur = parseInt($('longDur').value, 10) || 15;
        cfg.longEvery = parseInt($('longEvery').value, 10) || 4;
        cfg.sound = !!$('soundChk').checked;
        cfg.notif = !!$('notifChk').checked;
        saveCfg();
        state.maxSessions = cfg.longEvery | 0;
        if (!state.isRunning && !state.isBreak && state.seconds === 0 && state.minutes === oldWork) {
            state.minutes = cfg.workDur | 0;
        }
        closeSettings();
        updateUI();
        push('settings');
        toast(tt('pomo.toast.saved'), 'ok');
    }

    // ---------- Toast ----------
    function toast(text, kind) {
        try {
            var d = document.createElement('div');
            d.className = 'pom-toast ' + (kind || 'info');
            d.textContent = text;
            document.body.appendChild(d);
            requestAnimationFrame(function () { d.classList.add('show'); });
            setTimeout(function () {
                d.classList.remove('show');
                setTimeout(function () { try { d.remove(); } catch (_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('POMODORO', '操作失败:', _ && _.message); } }, 220);
            }, 1800);
        } catch (_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('POMODORO', '操作失败:', _ && _.message); }
    }

    // ---------- 初始化 ----------
    function bindRange(id, valId, unit) {
        var el = $(id);
        if (!el) return;
        el.addEventListener('input', function () {
            var v = $(valId);
            if (v) {
                if (unit === tt('pomo.unitMin') || unit === '分' || unit === 'min') v.textContent = tt('pomo.min', { n: el.value });
                else if (unit === tt('pomo.unitCount') || unit === '个' || unit === '') v.textContent = tt('pomo.count', { n: el.value });
                else v.textContent = el.value + (unit ? (' ' + unit) : '');
            }
        });
    }

    function init() {
        applyI18n();
        syncLocaleFromHost().then(function () {
            applyI18n();
            try { updateUI(); syncSettingsUI(); } catch (e) { /* ignore */ }
        });
        var sb = $('startBtn'); if (sb) sb.addEventListener('click', doStart);
        var pb = $('pauseBtn'); if (pb) pb.addEventListener('click', doPause);
        var rb = $('resetBtn'); if (rb) rb.addEventListener('click', doReset);
        var gb = $('gearBtn'); if (gb) gb.addEventListener('click', openSettings);
        var cs = $('cancelSettings'); if (cs) cs.addEventListener('click', closeSettings);
        var sv = $('saveSettings'); if (sv) sv.addEventListener('click', saveSettings);
        var syncB = $('syncBtn'); if (syncB) syncB.addEventListener('click', manualSync);

        bindRange('workDur', 'workDurVal', tt('pomo.unitMin'));
        bindRange('shortDur', 'shortDurVal', tt('pomo.unitMin'));
        bindRange('longDur', 'longDurVal', tt('pomo.unitMin'));
        bindRange('longEvery', 'longEveryVal', tt('pomo.unitCount'));

        var sc = $('soundChk'); if (sc) sc.addEventListener('change', function () { cfg.sound = sc.checked; });
        var nc = $('notifChk'); if (nc) nc.addEventListener('change', function () { cfg.notif = nc.checked; });

        updateUI();
        syncSettingsUI();
        initSync();
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
    window.addEventListener('smartpilot:locale-changed', function () { applyI18n(); try { updateUI(); } catch (e) {} });
})();