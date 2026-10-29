// 侧边可折叠工具栏：折叠/展开、视图联动、全局操作、快捷提示词
(function () {
    'use strict';

    var COLLAPSE_KEY = 'smartpilot.sideToolbar.collapsed';

    // ===== Excel 9528 API session token（/api/* 鉴权；9529 独立服务无需 token）=====
    // 优先 COM 桥（内嵌 WebView2），其次公开端点 /api/status；缓存后复用
    var excelToken = '';
    function getSessionToken(cb) {
        if (excelToken) { cb(excelToken); return; }
        try {
            if (window.ExcelBridge && typeof window.ExcelBridge.GetMobileControlStatus === 'function') {
                var r = window.ExcelBridge.GetMobileControlStatus();
                if (r) { var j = JSON.parse(r); if (j && j.sessionToken) { excelToken = j.sessionToken; cb(excelToken); return; } }
            }
        } catch(e) {
            if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', 'GetMobileControlStatus 失败:', e && e.message);
        }
        // 【守卫修正·B】VSTO(WebView2) 已走 COM 桥取 token，本地 9528 不存在，跳过 HTTP 探测。
        if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects) { cb(''); return; }
        try {
            if (!window.fetch) { cb(''); return; }
            var c = new AbortController();
            var t = setTimeout(function(){ c.abort(); }, 1500);
            fetch((App.Endpoints ? App.Endpoints.excelVstoHttp.status : 'http://localhost:9528/api/status'), { cache:'no-store', signal: c.signal })
                .then(function(res){ clearTimeout(t); return res.json(); })
                .then(function(j){ if (j && j.sessionToken) excelToken = j.sessionToken; cb(excelToken); })
                .catch(function(){ clearTimeout(t); cb(''); });
        } catch(e) {
            if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', 'fetch 9528 token 失败:', e && e.message);
            cb('');
        }
    }
    function excelApiUrl(url) {
        return url + (url.indexOf('?') >= 0 ? '&' : '?') + 'token=' + encodeURIComponent(excelToken || '');
    }

    function safeCall(name) {
        if (typeof window[name] === 'function') {
            window[name]();
            return true;
        }
        if (window.App && typeof window.App.showToast === 'function') {
            var loadingMsg = '功能加载中，请稍候…';
            try {
                if (window.App && App.i18n && typeof App.i18n.t === 'function') loadingMsg = App.i18n.t('toast.loading');
            } catch (_) { /* ignore */ }
            window.App.showToast(loadingMsg, 'info');
        }
        return false;
    }

    function syncActiveView(tab) {
        document.querySelectorAll('#side-toolbar .side-tool-btn[data-nav]').forEach(function (b) {
            b.classList.toggle('active', b.getAttribute('data-nav') === tab);
        });
    }

    // 将左侧「对话」按钮置为活动状态。
    // 任务规划 / 代码编辑器 / 快捷提示词均属于对话视图下的子面板，
    // 切换这些面板时对话按钮应保持活动高亮。
    window.activateChatNav = function () {
        activateChatView();
    };

    // 切回对话视图：点亮左侧「对话」按钮，并把主区域从其它标签（历史/扩展/知识库）切回对话。
    // 供任务规划 / 代码编辑器 / 快捷提示词等对话子面板入口调用。
    function activateChatView() {
        syncActiveView('chat');
        // 当前不在对话视图时，模拟点击顶部「💬对话」标签触发完整 tab 切换
        var activeTab = document.querySelector('.tab-btn.active');
        if (activeTab && activeTab.getAttribute('data-tab') !== 'chat') {
            var chatBtn = document.querySelector('.tab-btn[data-tab="chat"]');
            if (chatBtn) chatBtn.click();
        }
    }
    window.activateChatView = activateChatView;

    // 完全回到对话状态：收起左侧代码编辑器 / 任务规划面板。
    // 供新建对话、快捷提示词等对话子面板入口复用（仅在面板展开时生效）。
    window.collapseSidePanels = function () {
        var leftPanel = document.getElementById('left-panel');
        if (!leftPanel || !leftPanel.classList.contains('expanded')) return;
        // 先关闭任务规划面板（如打开）
        var tpPanel = document.getElementById('task-planner-panel');
        if (tpPanel && tpPanel.classList.contains('active')) {
            tpPanel.classList.remove('active');
            tpPanel.style.display = 'none';
        }
        // 再调用 toggleEditor 收起编辑器（此时已非任务规划态，走收起分支）
        if (typeof window.toggleEditor === 'function') window.toggleEditor();
    };

    function setToggleIcon() {
        var bar = document.getElementById('side-toolbar');
        var btn = document.getElementById('side-toolbar-toggle');
        if (bar && btn) btn.innerHTML = bar.classList.contains('collapsed') ? '›' : '‹';
    }

    // ============ 工具栏中部的上下滚动（箭头 + 鼠标滚轮） ============
    function getToolbarScroll() {
        return document.getElementById('side-toolbar')
            ? document.querySelector('#side-toolbar .side-tool-scroll')
            : null;
    }

    function updateToolbarArrows() {
        var sc = getToolbarScroll();
        var up = document.getElementById('side-tool-up');
        var down = document.getElementById('side-tool-down');
        if (!sc) return;
        // 仅在该方向仍有可滚动内容时显示对应箭头，否则隐藏
        var canUp = sc.scrollTop > 0;
        var canDown = sc.scrollTop + sc.clientHeight < sc.scrollHeight - 1;
        if (up) up.style.display = canUp ? 'flex' : 'none';
        if (down) down.style.display = canDown ? 'flex' : 'none';
    }

    // 根据方向滚动（dir: -1 向上 / 1 向下）
    window.toolbarScroll = function (dir) {
        var sc = getToolbarScroll();
        if (!sc) return;
        sc.scrollTop += Math.round(sc.clientHeight * 0.8) * dir;
        updateToolbarArrows();
    };

    // 折叠 / 展开
    window.toggleSideToolbar = function () {
        var bar = document.getElementById('side-toolbar');
        if (!bar) return;
        bar.classList.toggle('collapsed');
        var collapsed = bar.classList.contains('collapsed');
        setToggleIcon();
        if (collapsed) {
            closeQuick();
            if (typeof window.closeSideMore === 'function') window.closeSideMore();
        }
        try { localStorage.setItem(COLLAPSE_KEY, collapsed ? '1' : '0'); } catch (e) {
            if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', 'localStorage 保存折叠状态失败:', e && e.message);
        }
        window.applySideToolAdapt && window.applySideToolAdapt();
    };

    // 自适应：按可用高度统一缩放按钮/图标（--tool-scale），屏幕矮时整体等比缩小，
    // 避免图标被固有尺寸撑开、按钮被压缩变形或大量依赖滚动。
    window.applySideToolAdapt = function () {
        var bar = document.getElementById('side-toolbar');
        if (!bar) return;
        var scroll = bar.querySelector('.side-tool-scroll');
        var scale = 1;
        if (scroll) {
            var sh = scroll.clientHeight;
            var total = scroll.scrollHeight;
            var collapsed = bar.classList.contains('collapsed');
            if (sh > 0 && total > sh && !collapsed) {
                scale = Math.max(0.7, (sh / total) * 1.0);
            }
        }
        bar.style.setProperty('--tool-scale', scale.toFixed(3));
        rasterizeToolbarIcons(bar, scale);
    };

    // 窗口未最大化时，任务窗格原点经常落在半个物理像素上。
    // 彩色 emoji 若按文字合成，DirectWrite 会在旁边再画一层淡副本（虚影）。
    // 先在整数设备像素的画布上画好，再当图片贴上去，字形只栅格化一次。
    var glyphCache = {};
    function paintGlyph(text, devicePx) {
        var key = text + '\0' + devicePx;
        if (glyphCache[key]) return glyphCache[key];
        var canvas = document.createElement('canvas');
        canvas.width = devicePx;
        canvas.height = devicePx;
        var ctx = canvas.getContext('2d', { alpha: true });
        if (!ctx) return '';
        ctx.clearRect(0, 0, devicePx, devicePx);
        var fontPx = Math.max(8, Math.round(devicePx * 0.86));
        ctx.font = fontPx + 'px "Segoe UI Emoji","Segoe UI Symbol",sans-serif';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        ctx.fillText(text, devicePx / 2, devicePx / 2 + devicePx * 0.02);
        var url = '';
        try { url = canvas.toDataURL('image/png'); } catch (e) { url = ''; }
        if (url) glyphCache[key] = url;
        return url;
    }

    function glyphText(node) {
        return String(node.nodeValue || '').replace(/\s+/g, '');
    }

    function rasterizeToolbarIcons(bar, scale) {
        if (!bar) return;
        var dpr = window.devicePixelRatio || 1;
        var devicePx = Math.max(12, Math.round(20 * (scale || 1) * dpr));
        var displayPx = devicePx / dpr;
        bar.style.setProperty('--emoji-size', displayPx + 'px');
        var buttons = bar.querySelectorAll('.side-tool-btn');
        for (var i = 0; i < buttons.length; i++) {
            var btn = buttons[i];
            var texts = [];
            var kids = btn.childNodes;
            for (var k = 0; k < kids.length; k++) {
                var n = kids[k];
                if (n.nodeType === 3) {
                    var t = glyphText(n);
                    if (t && /[^\u0000-\u007F]/.test(t)) texts.push(n);
                }
            }
            for (var tix = 0; tix < texts.length; tix++) {
                var node = texts[tix];
                var text = glyphText(node);
                var img = document.createElement('img');
                img.className = 'side-tool-emoji';
                img.alt = '';
                img.setAttribute('aria-hidden', 'true');
                img.draggable = false;
                img.dataset.glyph = text;
                img.dataset.px = String(devicePx);
                img.src = paintGlyph(text, devicePx);
                if (node.parentNode) node.parentNode.replaceChild(img, node);
            }
            var imgs = btn.querySelectorAll('img.side-tool-emoji');
            for (var j = 0; j < imgs.length; j++) {
                var exist = imgs[j];
                if (exist.dataset.px === String(devicePx)) continue;
                var g = exist.dataset.glyph;
                if (!g) continue;
                var next = paintGlyph(g, devicePx);
                if (!next) continue;
                exist.src = next;
                exist.dataset.px = String(devicePx);
            }
        }
    }

    // 视图切换：复用顶部标签的既有逻辑
    window.sideToolbarNav = function (tab) {
        if (typeof window.closeSideMore === 'function') window.closeSideMore();
        syncActiveView(tab);
        var topBtn = document.querySelector('.tab-btn[data-tab="' + tab + '"]');
        if (topBtn) topBtn.click();
        // 切到对话视图时确保欢迎面板存在（空聊天区）
        if (tab === 'chat') {
            window.ensurePixelWelcome && window.ensurePixelWelcome();
        }
    };

    // 确保聊天区显示产品欢迎页（仅当聊天区为空时）
    window.ensurePixelWelcome = function () {
        var c = document.getElementById('chat-messages');
        if (!c || c.children.length > 0) return;
        try {
            if (typeof window.showWelcomePanel === 'function') window.showWelcomePanel();
            else if (typeof window.addWelcomeMessage === 'function') window.addWelcomeMessage();
        } catch (e) {
            if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', '显示欢迎页失败:', e && e.message);
        }
    };

    // 通用操作调用
    window.sideToolbarCall = function (name) {
        safeCall(name);
    };

    // 调试器：优先走 App.openDebugWindow
    window.openDebugTool = function () {
        if (window.App && typeof window.App.openDebugWindow === 'function') {
            window.App.openDebugWindow();
        } else if (typeof window.openDebugWindow === 'function') {
            window.openDebugWindow();
        } else {
            safeCall('openDebugWindow');
        }
    };

    // ⚙️ 打开 AI 模型设置（内嵌抽屉，避免 window.open 丢上下文）
    function resolveAiSettingsUrl() {
        var url = 'ai-settings-panel.html';
        try {
            if (window.ExcelBridge && typeof window.ExcelBridge.GetUIPath === 'function') {
                var p = window.ExcelBridge.GetUIPath();
                if (p) {
                    if (p.indexOf('file://') !== 0) p = 'file:///' + p.replace(/\\/g, '/').replace(/^\/+/, '');
                    if (p.charAt(p.length - 1) !== '/') p += '/';
                    url = p + 'ai-settings-panel.html';
                }
            }
        } catch (e) {
            if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', 'GetUIPath 失败:', e && e.message);
        }
        // embed=1 供面板区分内嵌模式；附带版本戳，避免 WebView2 缓存旧的 ai-settings-panel.html
        // （账户/订阅等新卡片若无戳会一直看不到）
        var ver = '';
        try {
            ver = (typeof window.__APP_VERSION__ === 'string' && window.__APP_VERSION__) || '';
        } catch (_) { /* ignore */ }
        if (!ver) {
            try { ver = 't' + Date.now(); } catch (_) { ver = '1'; }
        }
        var sep = url.indexOf('?') >= 0 ? '&' : '?';
        return url + sep + 'embed=1&v=' + encodeURIComponent(ver);
    }

    function ensureAiSettingsDrawer() {
        var existing = document.getElementById('ai-settings-drawer');
        if (existing) return existing;
        var overlay = document.createElement('div');
        overlay.id = 'ai-settings-drawer';
        overlay.className = 'ai-settings-drawer';
        overlay.setAttribute('aria-hidden', 'true');
        overlay.innerHTML =
            '<div class="ai-settings-drawer-backdrop" data-as-close="1"></div>' +
            '<div class="ai-settings-drawer-panel" role="dialog" aria-modal="true" aria-label="' + (function(){try{return App.i18n.t('nav.settings');}catch(_){return 'AI 模型设置';}})() + '">' +
            '  <div class="ai-settings-drawer-header">' +
            '    <span>' + (function(){try{return App.i18n.t('nav.settings');}catch(_){return 'AI 模型设置';}})() + '</span>' +
            '    <button type="button" class="ai-settings-drawer-close" data-as-close="1" title="' + (function(){try{return App.i18n.t('settings.closeEsc');}catch(_){return '关闭 (Esc)';}})() + '">×</button>' +
            '  </div>' +
            '  <iframe class="ai-settings-drawer-frame" title="' + (function(){try{return App.i18n.t('nav.settings');}catch(_){return 'AI 模型设置';}})() + '"></iframe>' +
            '</div>';
        document.body.appendChild(overlay);
        overlay.addEventListener('click', function (e) {
            var t = e.target;
            if (t && (t.getAttribute('data-as-close') === '1' || t.classList.contains('ai-settings-drawer-backdrop'))) {
                window.closeAiSettingsDrawer();
            }
        });
        return overlay;
    }

    window.closeAiSettingsDrawer = function () {
        var drawer = document.getElementById('ai-settings-drawer');
        if (!drawer) return;
        drawer.classList.remove('open');
        drawer.setAttribute('aria-hidden', 'true');
        var frame = drawer.querySelector('.ai-settings-drawer-frame');
        if (frame) {
            try { frame.src = 'about:blank'; } catch (_) { /* ignore */ }
        }
    };

    function hostBridge() {
        try {
            if (window.BridgeAccessor && typeof BridgeAccessor.getBridge === 'function') {
                var found = BridgeAccessor.getBridge();
                if (found) return found;
            }
        } catch (_) { /* ignore */ }
        try {
            return window.chrome && chrome.webview && chrome.webview.hostObjects && chrome.webview.hostObjects.excelBridge;
        } catch (_) {
            return null;
        }
    }

    function callHost(method, arg) {
        var b = hostBridge();
        // BridgeAccessor 对任意方法名都返回 function。这里只跳过没有桥或取不到成员的情况；
        // 宿主未实现时，调用结果里会带 bridge-call-failed。
        if (!b || b[method] == null) return null;
        try {
            return arg === undefined ? b[method]() : b[method](arg);
        } catch (e) {
            if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', method + ' 失败:', e && e.message);
            return null;
        }
    }

    window.closeSideMore = function () {
        var pop = document.getElementById('side-more-popover');
        var btn = document.getElementById('side-more-btn');
        if (pop) pop.classList.remove('open');
        if (btn) btn.classList.remove('active');
    };

    window.toggleSideMore = function (evt) {
        if (evt && evt.stopPropagation) evt.stopPropagation();
        var pop = document.getElementById('side-more-popover');
        var btn = document.getElementById('side-more-btn');
        if (!pop || !btn) return;
        var willOpen = !pop.classList.contains('open');
        window.closeSideMore();
        if (!willOpen) return;
        if (pop.parentNode !== document.body) document.body.appendChild(pop);
        pop.classList.add('open');
        btn.classList.add('active');
        var rect = btn.getBoundingClientRect();
        var popH = Math.min(480, window.innerHeight - 16);
        var top = rect.top;
        if (top + popH > window.innerHeight - 8) top = Math.max(8, window.innerHeight - popH - 8);
        var left = rect.right + 8;
        if (left + 220 > window.innerWidth - 8) left = Math.max(8, rect.left - 220 - 8);
        pop.style.top = top + 'px';
        pop.style.left = left + 'px';
    };

    window.openOcrSettings = async function () {
        window.closeSideMore();
        var _r = callHost('OpenOcrSettingsPanel');
        // 桌面版宿主（Mock 桥）没有 OpenOcrSettingsPanel：resolve 成 {success:false} 或失败 JSON 字符串，
        // 不能当成功——降级直接打开同源 OCR 设置页（ocr-settings-panel.html 走 9527 桥，与 OCR 面板同链路）。
        try {
            if (_r && typeof _r.then === 'function') _r = await _r;
        } catch (_e) { /* 宿主失败，走降级 */ }
        var _fail = (_r == null)
            || (_r && typeof _r === 'object' && _r.success === false)
            || (typeof _r === 'string'
                && /bridge-call-failed|not supported|"success"\s*:\s*false|未实现/i.test(_r));
        if (!_fail) return true;
        if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', 'OpenOcrSettingsPanel 不可用，降级打开网页 OCR 设置页');
        try {
            var _w = window.open('ocr-settings-panel.html', '_blank',
                'width=980,height=720,menubar=no,toolbar=no,location=no,status=no,resizable=yes,scrollbars=yes');
            if (_w) return true;
            window.location.href = 'ocr-settings-panel.html';
            return true;
        } catch (_e2) {
            if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', '打开 OCR 网页设置页失败:', _e2 && _e2.message);
        }
        return false;
    };

    window.openMobileFromMenu = async function () {
        window.closeSideMore();
        var raw = callHost('StartMobileControl', -1);
        try {
            if (raw && typeof raw.then === 'function') raw = await raw;
        } catch (_) {
            raw = null;
        }
        var data = raw;
        if (typeof raw === 'string') {
            try { data = JSON.parse(raw); } catch (_) { data = null; }
        }
        // VSTO 宿主启动手机服务成功 → 打开内置面板
        if (data && data.success === true) {
            callHost('OpenMobilePanel');
            return;
        }
        // 桌面版没有手机遥控服务（桥无 StartMobileControl）：降级打开独立页，
        // 页面会如实提示"需在 Excel 插件版启动手机服务"（不再只弹 toast 不打开）。
        try {
            if (window.App && typeof App.showToast === 'function') {
                App.showToast('手机遥控需 Excel 插件版启动手机服务，已打开独立页查看', 'info');
            }
        } catch (_) { /* ignore */ }
        var w = window.open('mobile-control-panel.html', '_blank',
            'width=900,height=760,menubar=no,toolbar=no,location=no,status=no,resizable=yes,scrollbars=yes');
        if (!w) window.location.href = 'mobile-control-panel.html';
    };

    window.ensureStandaloneServices = function () {
        callHost('StartMobileControl', -1);
        callHost('EnsureLocalBridge');
    };

    window.openOcrPanel = async function () {
        syncActiveView('ocr');
        window.closeSideMore();
        var b = null;
        try {
            if (window.BridgeAccessor && typeof BridgeAccessor.getBridge === 'function') b = BridgeAccessor.getBridge();
        } catch (_) { /* ignore */ }
        if (!b) {
            try {
                b = window.chrome && chrome.webview && chrome.webview.hostObjects && chrome.webview.hostObjects.excelBridge;
            } catch (_) { /* ignore */ }
        }
        if (b && b.OpenOcrPanel != null) {
            try {
                var _ocrRes = await b.OpenOcrPanel();
                // 宿主/Mock 对未实现方法会 resolve 成 {success:false} 对象或失败 JSON 字符串，
                // 不能当成功——否则桌面版点了没任何反应。识别后走下方网页降级。
                var _ocrFail = (_ocrRes && typeof _ocrRes === 'object' && _ocrRes.success === false)
                    || (typeof _ocrRes === 'string'
                        && /bridge-call-failed|not supported|"success"\s*:\s*false|未实现/i.test(_ocrRes));
                if (!_ocrFail) return true;
                if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', 'OpenOcrPanel 不可用，降级打开网页 OCR 面板');
            } catch (e) {
                if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', 'OpenOcrPanel 失败:', e && e.message);
            }
        }
        // 桌面版/网页版降级：直接打开 OCR 面板独立页（与主页同源，可直连本地桥 9527）
        try {
            var _ocrW = window.open('ocr-panel.html', '_blank',
                'width=1100,height=800,menubar=no,toolbar=no,location=no,status=no,resizable=yes,scrollbars=yes');
            if (_ocrW) return true;
            window.location.href = 'ocr-panel.html';
            return true;
        } catch (_ocrE) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', '打开 OCR 网页面板失败:', _ocrE && _ocrE.message); }
        try {
            if (window.App && typeof App.showToast === 'function') {
                App.showToast('请点功能区「识图入表」打开识别面板。', 'info');
            }
        } catch (_) { /* ignore */ }
        return false;
    };

    window.openAiSettings = function () {
        var drawer = ensureAiSettingsDrawer();
        var frame = drawer.querySelector('.ai-settings-drawer-frame');
        var url = resolveAiSettingsUrl();
        if (frame) frame.src = url;
        drawer.classList.add('open');
        drawer.setAttribute('aria-hidden', 'false');
        // Esc 关闭
        if (!drawer._escBound) {
            drawer._escBound = true;
            document.addEventListener('keydown', function (e) {
                if (e.key === 'Escape' && drawer.classList.contains('open')) {
                    window.closeAiSettingsDrawer();
                }
            });
        }
    };

    // 内嵌设置面板 ↔ 主界面消息
    if (!window._aiSettingsMsgBound) {
        window._aiSettingsMsgBound = true;
        window.addEventListener('message', function (e) {
            var data = e && e.data;
            if (!data || typeof data !== 'object') return;
            if (data.type === 'smartpilot:ai-settings-close') {
                window.closeAiSettingsDrawer();
                return;
            }
            if (data.type === 'smartpilot:open-account') {
                try {
                    if (typeof window.openAccountPanel === 'function') window.openAccountPanel();
                } catch (_) { /* ignore */ }
                return;
            }
            if (data.type === 'smartpilot:ai-settings-saved') {
                try {
                    if (window.App && typeof App.loadConfig === 'function') App.loadConfig();
                } catch (_) { /* ignore */ }
                try {
                    if (typeof window.refreshWelcomeSetupCta === 'function') window.refreshWelcomeSetupCta();
                } catch (_) { /* ignore */ }
                try {
                    if (window.App && typeof App.showToast === 'function') {
                        App.showToast('设置已保存，可以开始对话', 'success');
                    }
                } catch (_) { /* ignore */ }
                // 保存后稍等再关，让用户看到成功反馈
                setTimeout(function () { window.closeAiSettingsDrawer(); }, 400);
                return;
            }
            // 设置面板（file:// iframe，拿不到宿主桥）请求代开外部下载链接：
            // 主页面在虚拟主机源下才有 WebView2 宿主桥，这里代调 OpenExternalUrl 并回包结果。
            if (data.type === 'smartpilot:open-external-url' && data.url) {
                var openUrl = String(data.url);
                var replyOpen = function (ok, error) {
                    try {
                        if (e.source) e.source.postMessage({
                            type: 'smartpilot:open-external-url-result',
                            token: data.token, url: openUrl, ok: !!ok, error: error || ''
                        }, '*');
                    } catch (_) { /* ignore */ }
                };
                try {
                    var res = callHost('OpenExternalUrl', openUrl);
                    Promise.resolve(res).then(function (raw) {
                        var ok = false, err = '';
                        try {
                            var r = typeof raw === 'string' ? JSON.parse(raw) : raw;
                            ok = !!(r && r.success);
                            err = (r && r.error) || '';
                        } catch (_) { ok = true; }
                        replyOpen(ok, err);
                    }).catch(function (err) {
                        replyOpen(false, err && err.message);
                    });
                } catch (err) {
                    replyOpen(false, err && err.message);
                }
                return;
            }
            // 设置面板请求代复制（iframe 剪贴板可能被权限策略拦）：主页面代复制并回包。
            if (data.type === 'smartpilot:copy-text') {
                var copyText = String(data.text == null ? '' : data.text);
                var replyCopy = function (ok) {
                    try {
                        if (e.source) e.source.postMessage({
                            type: 'smartpilot:copy-text-result',
                            token: data.token, ok: !!ok
                        }, '*');
                    } catch (_) { /* ignore */ }
                };
                var execCopy = function () {
                    try {
                        var ta = document.createElement('textarea');
                        ta.value = copyText;
                        ta.style.position = 'fixed';
                        ta.style.opacity = '0';
                        document.body.appendChild(ta);
                        ta.select();
                        var ok = document.execCommand('copy');
                        document.body.removeChild(ta);
                        return !!ok;
                    } catch (_) { return false; }
                };
                if (navigator.clipboard && navigator.clipboard.writeText) {
                    navigator.clipboard.writeText(copyText).then(function () { replyCopy(true); }).catch(function () { replyCopy(execCopy()); });
                } else {
                    replyCopy(execCopy());
                }
                return;
            }
        });
    }

    // 🎨 PPT 主题模板编辑器：统一打开独立页（ppt-theme-editor.html 自带 callHost，
    // 在 VSTO 宿主走 hostObjects、桌面版走 9527 桥，纯前端可用）。不再调 VSTO 窗格方法，
    // 避免桌面版 Mock 桥把"未实现"误判为成功导致页面打不开（归一）。
    window.openPptThemeEditor = function () {
        window.closeSideMore();
        var w = window.open('ppt-theme-editor.html', '_blank',
            'width=1180,height=840,menubar=no,toolbar=no,location=no,status=no,resizable=yes,scrollbars=yes');
        if (!w) window.location.href = 'ppt-theme-editor.html';
    };

    // 🍅 打开独立番茄钟页面：优先 HTTP 9528 服务（双向同步），失败降级 file://
    window.openStandalonePomodoro = function (evt) {
        if (typeof window.ensureStandaloneServices === 'function') window.ensureStandaloneServices();
        try { if (evt) { evt.stopPropagation && evt.stopPropagation(); evt.preventDefault && evt.preventDefault(); } } catch(e) {
            if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', '阻止事件传播失败:', e && e.message);
        }
        var HTTP = (App.Endpoints ? App.Endpoints.excelVstoHttp.pomodoro : 'http://localhost:9528/pomodoro');
        var via = 'http';
        var openUrl = function(u){
            try {
                var w = window.open(u, '_blank', 'width=520,height=780,menubar=no,toolbar=no,location=no,status=no,resizable=yes,scrollbars=no');
                if (!w) window.location.href = u;
            } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', 'openUrl 失败:', e && e.message); if (u && u.substr(0,4) === 'http') { try { window.location.href = u; } catch(x){ if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', '降级跳转也失败:', x && x.message); } } }
        };
        // 先用 HTTP 探测；如果成功说明 Excel 服务在运行 → 走共享权威同步
        var fallbackFile = function(){
            var url = 'pomodoro.html';
            try {
                if (window.ExcelBridge && typeof window.ExcelBridge.GetUIPath === 'function') {
                    var p = window.ExcelBridge.GetUIPath();
                    if (p) {
                        if (p.indexOf('file://') !== 0) p = 'file:///' + p.replace(/\\/g,'/').replace(/^\/+/,'');
                        if (p.charAt(p.length-1) !== '/') p += '/';
                        url = p + 'pomodoro.html';
                        via = 'file-excelbridge';
                    }
                }
            } catch(e){ via = 'file-relative'; if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', 'GetUIPath(pomodoro) 失败:', e && e.message); }
            openUrl(url);
        };
        // 获取 session token：优先 COM 桥（内嵌 WebView2），其次公开端点 /api/status（模块级 getSessionToken 缓存复用）
        try {
            var ctrl, t;
            try {
                ctrl = new AbortController();
                t = setTimeout(function(){ ctrl.abort(); }, 1500);
            } catch(e){ if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', '创建 AbortController 失败:', e && e.message); }
            if (!window.fetch) throw new Error('no fetch');
            getSessionToken(function(tok){
                var probe = excelApiUrl((App.Endpoints ? App.Endpoints.excelVstoHttp.pomodoroGet : 'http://localhost:9528/api/pomodoro/get'));
                fetch(probe, { method:'GET', signal: ctrl ? ctrl.signal : undefined })
                  .then(function(res){ if (t) clearTimeout(t); return res.json(); })
                  .then(function(r){ if (r && r.success) { via = 'http'; openUrl(HTTP); } else { fallbackFile(); } })
                  .catch(function(){ if (t) clearTimeout(t); fallbackFile(); });
            });
        } catch(e){ if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', '番茄钟 HTTP 探测失败，降级 file://:', e && e.message); fallbackFile(); }
    };

    // 🗂 打开独立任务列表页面：优先 HTTP 9528（带配置/API），失败降级 file://
    // ============ 🗂 打开独立任务列表页面（一键启动独立服务 9529 + 打开页面，共用同一个按钮） ============
    window.openStandaloneTaskPlanner = function () {
        if (typeof window.ensureStandaloneServices === 'function') window.ensureStandaloneServices();
        // 第一步：先尝试启动（或确认）独立服务 9529（和「一键启动」按钮共用后端 StartStandaloneAgent，幂等）
        //   - 成功/已在运行：刚启动就会把 Excel 里的所有任务灌过去
        //   - 失败/Excel 未运行：依然能打开本地 file:// 离线页，不阻塞
        var ensureServiceDone = false;
        try {
            if (window.fetch) {
                getSessionToken(function(){
                    try {
                        var c1 = new AbortController();
                        var t1 = setTimeout(function(){ c1.abort(); }, 6000); // 启动最多等 6s（启动需耗时加载 python）
                        fetch(excelApiUrl((App.Endpoints ? App.Endpoints.excelVstoHttp.standaloneStart : 'http://localhost:9528/api/standalone/start')), {
                            method: 'POST',
                            signal: c1.signal
                        }).then(function(r){ clearTimeout(t1); ensureServiceDone = true; })
                          .catch(function(){ clearTimeout(t1); /* 连不上 Excel 9528：不报错，继续开页面 */ });
                        // 🆕 顺便把 AI 模型配置也重新导出到权威位置（Excel 在线时，fire-and-forget；独立页打开也不用再手动点「修复配置」）
                        //      失败静默跳过：即使导出不了，task-planner.html 初始化时也会自己探测并兜底弹表单
                        try {
                            var c2 = new AbortController();
                            var t2 = setTimeout(function(){ c2.abort(); }, 4000);
                            fetch(excelApiUrl((App.Endpoints ? App.Endpoints.excelVstoHttp.llmExportConfig : 'http://localhost:9528/api/llm/export-config')), { method: 'POST', signal: c2.signal })
                                .then(function(){ clearTimeout(t2); })
                                .catch(function(){ clearTimeout(t2); });
                        } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', '导出 LLM 配置失败:', e && e.message); }
                    } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', '启动独立服务失败:', e && e.message); }
                });
            }
        } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', '确保服务运行失败:', e && e.message); }

        // 优先独立服务 9529（Excel 关闭后仍可用且托管页面），其次 Excel 内嵌 9528（在线同步），最后降级本地 file://
        var HTTP = (App.Endpoints ? App.Endpoints.standaloneAgent.taskPlanner : 'http://localhost:9529/task-planner');
        var openUrl = function(u){
            try {
                var w = window.open(u, '_blank', 'width=1000,height=760,menubar=no,toolbar=no,location=no,status=no,resizable=yes,scrollbars=yes');
                if (!w) window.location.href = u;
            } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', 'openUrl 失败:', e && e.message); if (u && u.substr(0,4) === 'http') { try { window.location.href = u; } catch(x){ if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', '降级跳转也失败:', x && x.message); } } }
        };
        // 先用 HTTP 探测；成功 → 走 HTTP 版（独立服务可离线用），失败依次降级 9528 → 本地文件
        var fallbackFile = function(){
            var url = 'task-planner.html';
            try {
                if (window.ExcelBridge && typeof window.ExcelBridge.GetUIPath === 'function') {
                    var p = window.ExcelBridge.GetUIPath();
                    if (p) {
                        if (p.indexOf('file://') !== 0) p = 'file:///' + p.replace(/\\/g,'/').replace(/^\/+/,'');
                        if (p.charAt(p.length-1) !== '/') p += '/';
                        url = p + 'task-planner.html';
                    }
                }
            } catch(e){ url = 'task-planner.html'; if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', 'GetUIPath(task-planner) 失败:', e && e.message); }
            openUrl(url);
        };
        var tryHttpThenOpen = function(port){
            try {
                var ctrl, t;
                try { ctrl = new AbortController(); t = setTimeout(function(){ ctrl.abort(); }, 1500); } catch(e){ if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', '创建 AbortController 失败:', e && e.message); }
                if (!window.fetch) throw new Error('no fetch');
                fetch('http://localhost:' + port + '/task-planner', { method: 'GET', signal: ctrl ? ctrl.signal : undefined })
                  .then(function(res){ if (t) clearTimeout(t); if (res && res.ok) { openUrl('http://localhost:' + port + '/task-planner'); } else { tryNext(port); } })
                  .catch(function(){ if (t) clearTimeout(t); tryNext(port); });
            } catch(e){ tryNext(port); }
        };
        var tryNext = function(port){
            if (port === 9529) { tryHttpThenOpen(9528); return; }
            fallbackFile();
        };
        // 如果服务启动请求还在飞，稍等它完成（避免页面刚打开还没拉到任务）
        var waitMs = ensureServiceDone ? 0 : 1200;
        setTimeout(function(){ tryHttpThenOpen(9529); }, waitMs);
    };

    // ============ ⚙️ 一键启动/停止独立服务（9529） ============
    // 独立服务在 Excel 关闭后仍可执行任务；运行中点击即停止，未运行点击即启动。
    function sarToast(msg, type) {
        if (window.App && typeof window.App.showToast === 'function') {
            window.App.showToast(msg, type || 'info');
        } else {
            try { alert(msg); } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', 'alert 也失败:', e && e.message); }
        }
    }
    function probeSar() {
        return new Promise(function (resolve) {
            if (!window.fetch) { resolve(false); return; }
            var ctrl, t;
            try { ctrl = new AbortController(); t = setTimeout(function(){ ctrl.abort(); }, 1500); } catch(e){ if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', '创建 AbortController 失败:', e && e.message); }
            fetch((App.Endpoints ? App.Endpoints.standaloneAgent.health : 'http://localhost:9529/health'), { method: 'GET', signal: ctrl ? ctrl.signal : undefined })
                .then(function (r) { if (t) clearTimeout(t); resolve(true); })
                .catch(function () { if (t) clearTimeout(t); resolve(false); });
        });
    }
    function waitSarDown(rounds) {
        var i = 0;
        return new Promise(function (resolveFn) {
            (function attempt() {
                i++;
                if (i >= rounds) { resolveFn(false); return; }
                probeSar().then(function (running) {
                    if (!running) resolveFn(true);
                    else setTimeout(attempt, 500);
                });
            })();
        });
    }
    window.toggleStandaloneServiceEntry = function (evt) {
        try { if (evt) { evt.stopPropagation && evt.stopPropagation(); evt.preventDefault && evt.preventDefault(); } } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', '阻止事件传播失败:', e && e.message); }
        var SAR = (App.Endpoints ? App.Endpoints.standaloneAgent.baseUrl : 'http://localhost:9529');
        probeSar().then(function (running) {
            if (!running) {
                // 未运行 → 通过 Excel(9528) 拉起 python 进程
                sarToast('正在启动独立服务…', 'info');
                return new Promise(function (resolve) {
                    getSessionToken(function () {
                        fetch(excelApiUrl((App.Endpoints ? App.Endpoints.excelVstoHttp.standaloneStart : 'http://localhost:9528/api/standalone/start')), { method: 'POST' })
                            .then(function (r) { return r.json(); })
                            .then(function (d) {
                                if (d && d.success) { sarToast(d.message || '独立服务已启动', 'success'); return; }
                                sarToast((d && d.error) || '独立服务启动失败', 'error');
                            })
                            .catch(function () {
                                sarToast('无法通过 Excel 启动独立服务（Excel 可能未运行）。可手动运行 StandaloneAgent\\standalone_agent.py', 'error');
                            });
                    });
                });
            }
            // 运行中 → 停止（走 /shutdown 干净退出并保存任务）
            sarToast('正在停止独立服务…', 'info');
            return fetch(SAR + '/shutdown', { method: 'POST' })
                .then(function () { return waitSarDown(10); })
                .then(function (down) {
                    sarToast(down ? '独立服务已停止' : '独立服务停止中，请稍候…', down ? 'success' : 'info');
                })
                .catch(function () { sarToast('停止独立服务失败', 'error'); });
        });
    };

    // 占位工具
    window.sideToolbarPlaceholder = function () {
        if (window.App && typeof window.App.showToast === 'function') {
            window.App.showToast('该功能正在规划中', 'info');
        }
    };

    // 切换任务规划面板（与编辑器互斥切换）
    window.toggleTaskPlanner = function () {
        // 任务规划属于对话视图的子面板：切回对话视图并点亮「对话」按钮
        activateChatView();
        var leftPanel = document.getElementById('left-panel');
        var tpPanel = document.getElementById('task-planner-panel');
        var editorPanel = document.getElementById('editor-panel');
        if (!leftPanel || !tpPanel) return;

        var isTpActive = tpPanel.classList.contains('active');

        if (isTpActive) {
            // 关闭任务规划面板
            tpPanel.classList.remove('active');
            tpPanel.style.display = 'none';
            editorPanel.style.display = '';
            // 如果编辑器不是展开状态，收起面板
            var codeContainer = document.getElementById('code-with-lines');
            if (!codeContainer || codeContainer.classList.contains('is-hidden')) {
                leftPanel.classList.remove('expanded');
            }
        } else {
            // 隐藏编辑器内容
            editorPanel.style.display = 'none';
            // 显示任务规划面板
            tpPanel.classList.add('active');
            tpPanel.style.display = 'flex';
            leftPanel.classList.add('expanded');
            if (window.TaskPlanner) {
                window.TaskPlanner.init();
                window.TaskPlanner.loadTasks();
                try { window.TaskPlanner.loadLLMBar(); } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', 'TaskPlanner.loadLLMBar 失败:', e && e.message); }
            }
        }
    };

    // 隐藏内联表单
    window.hideInlineForm = function () {
        var form = document.getElementById('tp-inline-form');
        if (form) form.style.display = 'none';
        var saveBtn = document.getElementById('tp-save-btn');
        if (saveBtn) { saveBtn.disabled = false; saveBtn.textContent = '保存'; saveBtn.style.opacity = '1'; }
        window.TaskPlanner && (window.TaskPlanner._editingId = null);
    };

    // 频率切换时隐藏/显示时间
    window.onTpFreqChange = function () {
        var sel = document.getElementById('tp-frequency');
        var grp = document.getElementById('tp-time-group');
        if (sel && grp) {
            grp.style.display = sel.value === 'hourly' ? 'none' : '';
        }
    };

    // ============ 快捷提示词（统一配置模块 App.QuickPrompts） ============
    function getPrompts() {
        if (window.App && window.App.QuickPrompts) {
            return window.App.QuickPrompts.getItems();
        }
        return [];
    }

    // 按关键字过滤提示词（label / prompt / group，大小写不敏感）
    function filterPrompts(items, kw) {
        kw = (kw || '').trim().toLowerCase();
        if (!kw) return items;
        return items.filter(function (it) {
            return (it.label || '').toLowerCase().indexOf(kw) >= 0 ||
                (it.prompt || '').toLowerCase().indexOf(kw) >= 0 ||
                (it.group || '').toLowerCase().indexOf(kw) >= 0;
        });
    }

    // 浏览视图：按分组渲染提示词列表（支持搜索过滤）
    function renderQuickList() {
        var wrap = document.getElementById('side-quick-list');
        if (!wrap) return;
        var inp = document.getElementById('side-quick-input');
        var kw = inp ? inp.value : '';
        var items = filterPrompts(getPrompts(), kw);
        wrap.innerHTML = '';
        if (!items.length) {
            var empty = document.createElement('div');
            empty.style.cssText = 'color:#999;font-size:12px;padding:6px;';
            empty.textContent = kw ? '未找到匹配的提示词' : '暂无快捷提示词';
            wrap.appendChild(empty);
            return;
        }
        var groups = [];
        items.forEach(function (it) {
            if (groups.indexOf(it.group) === -1) groups.push(it.group);
        });
        groups.forEach(function (grp) {
            var g = document.createElement('div');
            g.className = 'sq-group';
            var title = document.createElement('div');
            title.className = 'sq-group-title';
            title.textContent = grp;
            g.appendChild(title);
            items.forEach(function (it) {
                if (it.group !== grp) return;
                var item = document.createElement('button');
                item.type = 'button';
                item.className = 'side-quick-item';
                item.setAttribute('data-id', it.id);
                if (it.icon) {
                    var ic = document.createElement('span');
                    ic.className = 'sq-item-icon';
                    ic.textContent = it.icon;
                    item.appendChild(ic);
                }
                var tx = document.createElement('span');
                tx.textContent = it.prompt;
                item.appendChild(tx);
                item.addEventListener('click', function () { useQuickPrompt(it); });
                g.appendChild(item);
            });
            wrap.appendChild(g);
        });
    }

    // 管理视图：按分组渲染可编辑行（支持搜索过滤）
    function renderQuickMgmt() {
        var wrap = document.getElementById('side-quick-mgmt-list');
        if (!wrap) return;
        var inp = document.getElementById('side-quick-mgmt-input');
        var kw = inp ? inp.value : '';
        var items = filterPrompts(getPrompts(), kw);
        wrap.innerHTML = '';
        if (!items.length) {
            var empty = document.createElement('div');
            empty.style.cssText = 'color:#999;font-size:12px;padding:6px;';
            empty.textContent = kw ? '未找到匹配的提示词' : '暂无快捷提示词，可点击上方「＋ 新增」';
            wrap.appendChild(empty);
            return;
        }
        var groups = [];
        items.forEach(function (it) {
            if (groups.indexOf(it.group) === -1) groups.push(it.group);
        });
        groups.forEach(function (grp) {
            var g = document.createElement('div');
            g.className = 'sq-group';
            var title = document.createElement('div');
            title.className = 'sq-group-title';
            title.textContent = grp;
            g.appendChild(title);
            items.forEach(function (it) {
                if (it.group !== grp) return;
                g.appendChild(buildMgmtRow(it));
            });
            wrap.appendChild(g);
        });
    }

    function buildMgmtRow(it) {
        var row = document.createElement('div');
        row.className = 'sq-mgmt-row';
        row.setAttribute('data-id', it.id);
        row.innerHTML =
            '<span class="sq-row-icon">' + (it.icon || '') + '</span>' +
            '<div class="sq-row-text">' +
            '<div class="sq-row-label"></div>' +
            '<div class="sq-row-prompt"></div>' +
            '</div>' +
            '<div class="sq-row-ops">' +
            '<button type="button" class="sq-op" data-act="up" title="上移">↑</button>' +
            '<button type="button" class="sq-op" data-act="down" title="下移">↓</button>' +
            '<button type="button" class="sq-op" data-act="edit" title="编辑">✎</button>' +
            '<button type="button" class="sq-op danger" data-act="del" title="删除">✕</button>' +
            '</div>';
        row.querySelector('.sq-row-label').textContent = it.label;
        row.querySelector('.sq-row-prompt').textContent = it.prompt;
        return row;
    }

    // 行内编辑表单
    function buildEditForm(it) {
        var groups = [];
        try { groups = window.App.QuickPrompts.getGroups(); } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', '获取快捷提示词分组失败:', e && e.message); }
        var opts = groups.map(function (g) {
            return '<option value="' + g.replace(/"/g, '&quot;') + '"' + (g === it.group ? ' selected' : '') + '>' + g + '</option>';
        }).join('');
        var box = document.createElement('div');
        box.className = 'sq-edit-box';
        box.setAttribute('data-id', it.id);
        box.innerHTML =
            '<div class="sq-edit-fields">' +
            '<input type="text" class="sq-edit-group" list="sq-group-list" placeholder="分组">' +
            '<datalist id="sq-group-list">' + opts + '</datalist>' +
            '<input type="text" class="sq-edit-icon" maxlength="4" placeholder="图标">' +
            '<input type="text" class="sq-edit-label" placeholder="显示名">' +
            '<textarea class="sq-edit-prompt" rows="2" placeholder="提示词"></textarea>' +
            '</div>' +
            '<div class="sq-edit-ops">' +
            '<button type="button" class="sq-op primary" data-act="save">保存</button>' +
            '<button type="button" class="sq-op" data-act="cancel">取消</button>' +
            '</div>';
        box.querySelector('.sq-edit-group').value = it.group;
        box.querySelector('.sq-edit-icon').value = it.icon || '';
        box.querySelector('.sq-edit-label').value = it.label;
        box.querySelector('.sq-edit-prompt').value = it.prompt;
        return box;
    }

    function toast(msg, type) {
        if (window.App && typeof window.App.showToast === 'function') {
            window.App.showToast(msg, type || 'info');
        }
    }

    // 管理视图事件：委托处理行操作
    function onMgmtClick(e) {
        var btn = e.target && e.target.closest ? e.target.closest('.sq-op') : null;
        if (!btn || !btn.dataset || !btn.dataset.act) return;
        var act = btn.dataset.act;
        var row = btn.closest('.sq-mgmt-row');
        var box = btn.closest('.sq-edit-box');
        if (act === 'edit') {
            if (!row) return;
            var it = getPrompts().filter(function (x) { return x.id === row.getAttribute('data-id'); })[0];
            if (!it) return;
            row.style.display = 'none';
            row.insertAdjacentElement('afterend', buildEditForm(it));
            return;
        }
        if (act === 'cancel') {
            if (!box) return;
            var prev = box.previousElementSibling;
            box.parentNode.removeChild(box);
            if (prev && prev.classList.contains('sq-mgmt-row')) prev.style.display = '';
            return;
        }
        if (act === 'save') {
            if (!box) return;
            var id = box.getAttribute('data-id');
            var grp = box.querySelector('.sq-edit-group').value.trim();
            var icon = box.querySelector('.sq-edit-icon').value.trim();
            var label = box.querySelector('.sq-edit-label').value.trim();
            var prompt = box.querySelector('.sq-edit-prompt').value.trim();
            if (!prompt) { toast('提示词不能为空', 'error'); return; }
            var patch = { group: grp || '常用', icon: icon, label: label || prompt, prompt: prompt };
            if (id === 'new') {
                window.App.QuickPrompts.add(patch);
                toast('已添加', 'success');
            } else {
                window.App.QuickPrompts.update(id, patch);
            }
            return;
        }
        var target = row;
        if (!target) return;
        var id2 = target.getAttribute('data-id');
        if (act === 'up') { window.App.QuickPrompts.moveInGroup(id2, -1); return; }
        if (act === 'down') { window.App.QuickPrompts.moveInGroup(id2, 1); return; }
        if (act === 'del') {
            var it2 = getPrompts().filter(function (x) { return x.id === id2; })[0];
            if (!it2) return;
            if (window.confirm('删除快捷提示词「' + (it2.label || it2.prompt) + '」？')) {
                window.App.QuickPrompts.remove(id2);
                toast('已删除', 'success');
            }
        }
    }

    // 视图切换
    function showQuickMgmt() {
        var browse = document.getElementById('side-quick-browse');
        var mgmt = document.getElementById('side-quick-mgmt');
        if (!browse || !mgmt) return;
        renderQuickMgmt();
        browse.style.display = 'none';
        mgmt.style.display = 'flex';
        var input = document.getElementById('side-quick-mgmt-input');
        if (input) { try { input.focus(); } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', 'input.focus 失败:', e && e.message); } }
    }

    function showQuickBrowse() {
        var browse = document.getElementById('side-quick-browse');
        var mgmt = document.getElementById('side-quick-mgmt');
        if (!browse || !mgmt) return;
        renderQuickList();
        mgmt.style.display = 'none';
        browse.style.display = 'flex';
    }
    window.showQuickBrowse = showQuickBrowse;

    function useQuickPrompt(it) {
        closeQuick();
        if (it && typeof it === 'object' && it.action === 'ocr') {
            if (typeof window.openOcrPanel === 'function') window.openOcrPanel();
            return;
        }
        var text = typeof it === 'string' ? it : (it && it.prompt);
        var input = document.getElementById('message-input');
        if (input) {
            input.value = text;
            try { input.focus(); } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', 'message-input.focus 失败:', e && e.message); }
            window.App.showToast('已填入输入框，点击发送即可', 'info');
        }
    }

    window.toggleQuickPrompts = function (e) {
        if (e && e.stopPropagation) e.stopPropagation();
        // 快捷提示词属于对话视图的子面板：完全回到对话状态（切回对话视图 + 收起左侧面板）
        activateChatView();
        if (typeof window.collapseSidePanels === 'function') window.collapseSidePanels();
        var pop = document.getElementById('side-quick-popover');
        if (!pop) return;
        var willOpen = pop.style.display === 'none';
        closeQuick();
        if (willOpen) {
            if (pop.parentNode && pop.parentNode !== document.body) {
                document.body.appendChild(pop);
            }
            var anchor = document.getElementById('side-quick-btn') || document.getElementById('side-more-btn');
            var rect = anchor ? anchor.getBoundingClientRect() : null;
            if (rect) {
                var popH = 420;
                var spaceBelow = window.innerHeight - rect.bottom - 12;
                var top = rect.top;
                var left = rect.right + 8;
                if (left + 300 > window.innerWidth - 8) {
                    left = Math.max(8, rect.left - 300 - 8);
                }
                if (spaceBelow < popH + 12) {
                    top = Math.max(12, window.innerHeight - popH - 12);
                }
                pop.style.left = left + 'px';
                pop.style.top = top + 'px';
            } else {
                pop.style.left = '12px';
                pop.style.top = '12px';
            }
            showQuickBrowse();
            pop.style.display = 'flex';
        }
    };

    window.toggleQuickMgmt = function () {
        if (window.App && window.App.QuickPrompts) {
            showQuickMgmt();
        } else {
            toast('配置模块未加载', 'error');
        }
    };

    function closeQuick() {
        var pop = document.getElementById('side-quick-popover');
        if (pop) pop.style.display = 'none';
    }

    // 浏览视图搜索：实时过滤列表；回车 → 使用第一个匹配项
    function setupQuickInput() {
        var input = document.getElementById('side-quick-input');
        if (!input) return;
        input.addEventListener('input', function () { renderQuickList(); });
        input.addEventListener('keydown', function (e) {
            if (e.key !== 'Enter') return;
            var kw = input.value.trim();
            if (!kw) return;
            var m = filterPrompts(getPrompts(), kw)[0];
            if (m) { useQuickPrompt(m.prompt); }
        });
    }

    // 管理视图搜索：实时过滤可编辑列表（新增统一走「＋ 新增」按钮）
    function setupQuickMgmtInput() {
        var input = document.getElementById('side-quick-mgmt-input');
        if (!input) return;
        input.addEventListener('input', function () { renderQuickMgmt(); });
    }

    // 恢复默认配置
    window.resetQuickPrompts = function () {
        if (!window.confirm('恢复默认快捷提示词？当前所有自定义内容将被清除。')) return;
        window.App.QuickPrompts.reset();
        toast('已恢复默认配置', 'success');
    };

    // 管理视图顶部「＋ 新增」：清空搜索过滤并插入空白编辑表单
    window.addQuickPrompt = function () {
        var list = document.getElementById('side-quick-mgmt-list');
        if (!list) return;
        var existing = list.querySelector('.sq-edit-box[data-id="new"]');
        if (existing) {
            var p = existing.querySelector('.sq-edit-prompt');
            if (p) { try { p.focus(); } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', 'textarea.focus 失败:', e && e.message); } }
            return;
        }
        var searchInput = document.getElementById('side-quick-mgmt-input');
        if (searchInput && searchInput.value) {
            searchInput.value = '';
            renderQuickMgmt();
        }
        var groups = [];
        try { groups = window.App.QuickPrompts.getGroups(); } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', '获取快捷提示词分组失败:', e && e.message); }
        var box = buildEditForm({ id: 'new', group: groups[0] || '常用', icon: '', label: '', prompt: '' });
        list.insertBefore(box, list.firstChild);
        var input = box.querySelector('.sq-edit-prompt');
        if (input) { try { input.focus(); } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', 'input.focus 失败:', e && e.message); } }
    };

    // 顶部标签切换时同步视图高亮
    document.addEventListener('click', function (e) {
        var tabBtn = e.target && e.target.closest ? e.target.closest('.tab-btn') : null;
        if (tabBtn && tabBtn.getAttribute('data-tab')) {
            syncActiveView(tabBtn.getAttribute('data-tab'));
        }
    });

    // 点击工具栏外部关闭快捷弹层（弹层挂在 body 下，需排除弹层内部点击，否则管理按钮/行操作会被立即关闭）
    document.addEventListener('click', function (e) {
        var bar = document.getElementById('side-toolbar');
        var pop = document.getElementById('side-quick-popover');
        var more = document.getElementById('side-more-popover');
        var toolBtn = e.target && e.target.closest ? e.target.closest('#side-toolbar .side-tool-btn') : null;
        if (toolBtn && toolBtn.id !== 'side-more-btn' && typeof window.closeSideMore === 'function') {
            window.closeSideMore();
        }
        if (bar && !bar.contains(e.target) && !(pop && pop.contains(e.target)) && !(more && more.contains(e.target))) {
            closeQuick();
            if (typeof window.closeSideMore === 'function') window.closeSideMore();
        }
    });

    // 初始化
    function init() {
        var bar = document.getElementById('side-toolbar');
        if (!bar) return;
        try {
            if (localStorage.getItem(COLLAPSE_KEY) === '1') bar.classList.add('collapsed');
        } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', '读取折叠状态失败:', e && e.message); }
        setToggleIcon();
        var active = document.querySelector('.tab-btn.active');
        if (active && active.getAttribute('data-tab')) syncActiveView(active.getAttribute('data-tab'));
        setupQuickInput();
        setupQuickMgmtInput();
        var pop = document.getElementById('side-quick-popover');
        if (pop && !pop._mgmtBound) {
            pop._mgmtBound = true;
            pop.addEventListener('click', onMgmtClick);
        }
        // 配置变更 → 同步刷新浏览/管理列表（订阅一次，防止重复）
        if (window.App && window.App.QuickPrompts && !window._sqUpdateBound) {
            window._sqUpdateBound = true;
            window.App.QuickPrompts.onUpdate(function () {
                var mgmt = document.getElementById('side-quick-mgmt');
                var browse = document.getElementById('side-quick-browse');
                if (mgmt && mgmt.style.display !== 'none') {
                    renderQuickMgmt();
                } else {
                    renderQuickList();
                }
            });
        }
        renderQuickList();
        var sc = getToolbarScroll();
        if (sc) {
            sc.addEventListener('scroll', updateToolbarArrows);
            updateToolbarArrows();
        }
        window.applySideToolAdapt && window.applySideToolAdapt();
        // 窗口 / 滚动区尺寸变化时重新自适应
        if (window.ResizeObserver) {
            try {
                var sc2 = getToolbarScroll();
                if (sc2) new ResizeObserver(function () { window.applySideToolAdapt && window.applySideToolAdapt(); }).observe(sc2);
            } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', 'ResizeObserver 失败:', e && e.message); }
        }
        window.removeEventListener('resize', onWinResize);
        window.addEventListener('resize', onWinResize);
    }

    function onWinResize() {
        window.applySideToolAdapt && window.applySideToolAdapt();
    }

    // ============ 通用文件选择（目标 Excel 文件） ============
    var _hostCache2 = null;
    function fetchHostWorkbookSyncLike(forceRefresh, cb) {
        if (!forceRefresh && _hostCache2) { cb(_hostCache2); return; }
        var DEF = { success: false, hasHost: false, path: '', directory: '', name: '' };
        var done = function (d) {
            var r = Object.assign({}, DEF, d || {});
            _hostCache2 = r;
            cb(r);
        };
        // 1) COM Bridge
        try {
            if (window.ExcelBridge && typeof window.ExcelBridge.GetHostWorkbook === 'function') {
                var raw = window.ExcelBridge.GetHostWorkbook();
                if (raw) {
                    try { done((typeof raw === 'string') ? JSON.parse(raw) : raw); return; } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', '解析 GetHostWorkbook 返回失败:', e && e.message); }
                }
            }
        } catch(e){ if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', 'COM Bridge 获取 HostWorkbook 失败:', e && e.message); }
        // 2) HTTP fallback
        if (!window.fetch) { done(null); return; }
        try {
            var ctrl, t;
            try { ctrl = new AbortController(); t = setTimeout(function(){ ctrl.abort(); }, 2000); } catch(e){ if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', '创建 AbortController 失败:', e && e.message); }
            window.fetch((App.Endpoints ? App.Endpoints.excelVstoHttp.hostWorkbook : 'http://localhost:9528/api/host/workbook'), { method: 'GET', signal: ctrl ? ctrl.signal : undefined })
                .then(function (res) { if (t) clearTimeout(t); if (res && res.ok) return res.json(); throw new Error('http'); })
                .then(function (d) { done(d); })
                .catch(function () { if (t) clearTimeout(t); done(null); });
        } catch(e) { done(null); }
    }
    function doToast(msg, type) {
        if (window.App && typeof window.App.showToast === 'function') {
            window.App.showToast(msg, type || 'info');
        } else if (window.TaskPlanner && typeof window.TaskPlanner._toast === 'function') {
            window.TaskPlanner._toast(msg, type || 'info');
        } else {
            try { console.info(msg); } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', 'console.info 也失败:', e && e.message); }
        }
    }
    window.pickInlineTargetFile = function () {
        fetchHostWorkbookSyncLike(true, function (host) {
            var dir = (host && (host.hasHost || host.directory)) ? (host.directory || '') : '';
            (async function () {
                var picked = null;
                // 1) Excel COM Bridge
                try {
                    if (window.ExcelBridge && typeof window.ExcelBridge.PickExcelFile === 'function') {
                        var raw = window.ExcelBridge.PickExcelFile('', '选择目标 Excel 文件（独立执行操作此文件）', dir);
                        if (raw) {
                            var d = (typeof raw === 'string') ? JSON.parse(raw) : raw;
                            if (d && d.success && d.filePath) picked = d.filePath;
                        }
                    }
                } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', 'COM Bridge PickExcelFile 失败:', e && e.message); }
                // 2) Excel HTTP fallback (如果 COM 没注入，试试 HTTP API 9528)
                if (!picked && window.fetch) {
                    try {
                        await new Promise(function (resolve) { getSessionToken(resolve); });
                        var res = await fetch(excelApiUrl((App.Endpoints ? App.Endpoints.excelVstoHttp.browseExcelFile : 'http://localhost:9528/api/browse/excel-file')), {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({ title: '选择目标 Excel 文件（独立执行操作此文件）', initialDir: dir })
                        });
                        if (res.ok) {
                            var j = await res.json();
                            if (j && j.success && j.filePath) picked = j.filePath;
                        }
                    } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', 'HTTP 浏览 Excel 文件失败:', e && e.message); }
                }
                // 3) WebView2 下的原生 input[type=file]
                if (!picked) {
                    picked = await new Promise(function (resolve) {
                        try {
                            var input = document.createElement('input');
                            input.type = 'file';
                            input.accept = '.xlsx,.xls,.xlsm,.csv';
                            input.multiple = false;
                            input.style.cssText = 'position:absolute;left:-9999px;top:-9999px;opacity:0;';
                            var cleanup = function () { try { input && input.parentNode && input.parentNode.removeChild(input); } catch(e){ if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOLBAR', '清理 input 元素失败:', e && e.message); } };
                            input.addEventListener('change', function () {
                                var f = input.files && input.files[0];
                                cleanup();
                                if (f && f.path) { resolve(f.path); return; }
                                if (f && f.name) {
                                    doToast('浏览器沙箱无法获取本地路径，请手动粘贴路径。文件：' + f.name, 'info');
                                }
                                resolve(null);
                            }, { once: true });
                            document.body.appendChild(input);
                            setTimeout(function () { try { input.click(); } catch(e){ resolve(null); } }, 10);
                        } catch (e) { resolve(null); }
                    });
                }
                if (picked) {
                    var el = document.getElementById('tp-target-file');
                    if (el) el.value = picked;
                }
            })();
        });
    };

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();