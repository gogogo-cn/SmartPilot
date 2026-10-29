/**
 * ============================================================
 * bridge-accessor.js — 统一 Bridge 访问层
 * ============================================================
 * 功能：为前端提供统一的 excelBridge 访问接口，
 *       消除 4 种不同访问方式导致的混乱。
 *
 * 使用方式：
 *   const bridge = BridgeAccessor.getBridge();
 *   const syncBridge = BridgeAccessor.getSyncBridge();
 *   const result = await bridge.SomeMethod();
 *   const syncResult = syncBridge.SomeMethod();
 *
 * 依赖：
 *   - app-bootstrap.js（App 命名空间）
 * ============================================================
 */

(function() {
    'use strict';

    var BridgeAccessor = {
        _bridge: null,
        _syncBridge: null,
        _initialized: false,
        _wrappedSource: null,

        /**
         * 初始化 Bridge 访问器
         * 优先使用 App.excelBridge（已由 WebView2 初始化完成），
         * 回退到 window.chrome.webview.hostObjects
         */
        /** WebView2：必须尽早 defaultSyncMethod='none'，否则桥方法会抛 e is not a function */
        _ensureAsyncHostObjects: function() {
            try {
                var opts = window.chrome && window.chrome.webview && window.chrome.webview.hostObjects
                    && window.chrome.webview.hostObjects.options;
                if (opts) opts.defaultSyncMethod = 'none';
            } catch (_e) { /* ignore */ }
        },

        /** 每次调用取新鲜 hostObjects.excelBridge（缓存引用在 WebView2 上会让部分方法变 e is not a function） */
        _rawHostBridge: function() {
            try {
                if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects) {
                    return window.chrome.webview.hostObjects.excelBridge || null;
                }
            } catch (_e) { /* ignore */ }
            return null;
        },

        /**
         * 用「新鲜裸桥」再包一层 Proxy，并写回 App/window.excelBridge。
         * 注意：绝不能把 Proxy 再包一层，也不能把缓存的 host 引用当 target。
         */
        adoptHostBridge: function(raw) {
            this._ensureAsyncHostObjects();
            var host = this._rawHostBridge() || raw;
            if (!host) return null;
            this._wrappedSource = null;
            this._wrappedBridge = null;
            this._bridge = this._wrapHostBridge(host);
            if (typeof App !== 'undefined') {
                App.excelBridge = this._bridge;
                if (App.state) App.state.excelBridge = this._bridge;
            }
            window.excelBridge = this._bridge;
            this._initialized = true;
            return this._bridge;
        },

        init: function() {
            if (this._initialized) return;

            try {
                this._ensureAsyncHostObjects();
                var host = this._rawHostBridge();
                if (host) {
                    this.adoptHostBridge(host);
                    try {
                        this._syncBridge = window.chrome.webview.hostObjects.sync.excelBridge;
                    } catch (_s) { this._syncBridge = null; }
                } else if (typeof App !== 'undefined' && App.excelBridge) {
                    this.adoptHostBridge(App.excelBridge);
                } else if (window.excelBridge) {
                    this.adoptHostBridge(window.excelBridge);
                }

                if (this._bridge) {
                    if (typeof ChatLog !== 'undefined') {
                        ChatLog.info('BRIDGE', '🔗 BridgeAccessor 初始化完成 (WebView2/VSTO)');
                    }
                    return;
                }
            } catch (e) {
                if (typeof ChatLog !== 'undefined') {
                    ChatLog.error('BRIDGE', 'BridgeAccessor 初始化失败:', e.message);
                }
            }

            // 非 WebView2 环境（网页版/无插件）：乐观建立本地桥代理，统一路由到 9527
            this._initLocalBridge();
        },

        /**
         * 【守卫修正·关键】WebView2 hostObject：
         *   - typeof 任意名都是 'function'，缺方法时异步抛 e is not a function / 0x80020006
         *   - 缓存的 host 引用上调方法会偶发投影失败；必须每次用 hostObjects.excelBridge 新鲜调用
         *   - 用 async/await 接住拒绝（比 .then 更稳），失败返回结构化 JSON，不抛未处理拒绝
         */
        _wrapHostBridge: function(obj) {
            if (!obj) return obj;
            if (this._wrappedBridge) return this._wrappedBridge;
            if (typeof Proxy === 'undefined') return obj;
            var self = this;
            var warned = {};
            var wrapped = new Proxy(obj, {
                get: function(_target, prop) {
                    if (typeof prop !== 'string' || prop === 'then' || prop === 'catch' || prop === 'finally') {
                        var host0 = self._rawHostBridge() || _target;
                        try { return host0[prop]; } catch (_r) { return undefined; }
                    }
                    // 一律当作可调用：WebView2 对任意名 typeof 都是 function
                    return function() {
                        var args = Array.prototype.slice.call(arguments);
                        return self._invokeHostMethod(prop, args, warned);
                    };
                }
            });
            this._wrappedSource = obj;
            this._wrappedBridge = wrapped;
            return wrapped;
        },

        _invokeHostMethod: function(prop, args, warned) {
            var self = this;
            warned = warned || {};
            return (async function() {
                self._ensureAsyncHostObjects();
                var host = self._rawHostBridge();
                if (!host) {
                    if (!warned[prop]) { warned[prop] = 1; self._warnBridgeFailure(prop, new Error('host bridge missing')); }
                    return self._bridgeFailure(prop, new Error('host bridge missing'));
                }
                try {
                    // 必须 host[prop](...args)；拆成 fn=host[prop]; fn() 会丢 receiver → e is not a function
                    var out = host[prop](...args);
                    if (out != null && typeof out.then === 'function') {
                        try {
                            return await out;
                        } catch (err) {
                            if (!warned[prop]) { warned[prop] = 1; self._warnBridgeFailure(prop, err); }
                            return self._bridgeFailure(prop, err);
                        }
                    }
                    return out;
                } catch (e) {
                    if (!warned[prop]) { warned[prop] = 1; self._warnBridgeFailure(prop, e); }
                    return self._bridgeFailure(prop, e);
                }
            })();
        },

        _bridgeFailure: function(methodName, err) {
            var msg = (err && err.message) ? err.message : String(err === undefined ? 'unknown' : err);
            return JSON.stringify({
                success: false,
                error: 'bridge-call-failed',
                method: methodName,
                detail: msg
            });
        },

        _warnBridgeFailure: function(methodName, err) {
            var msg = (err && err.message) ? err.message : String(err === undefined ? 'unknown' : err);
            var text = '[BridgeAccessor] 桥方法调用失败（该宿主可能未实现；已按结构化失败返回，不再抛未处理拒绝）: '
                + methodName + ' — ' + msg;
            var missing = /0x80020006|未知名称|0x80070490|找不到元素/i.test(msg);
            if (typeof ChatLog !== 'undefined' && ChatLog.warn) {
                if (missing && ChatLog.debug) ChatLog.debug('BRIDGE', text);
                else ChatLog.warn('BRIDGE', text);
            } else if (!missing) {
                console.warn(text);
            }
        },

        /**
         * 获取异步 Bridge 对象（用于 await 调用）
         * 未初始化时即时探测（避免 init 在桥就绪前被调用后 _bridge 长期为 null）
         */
        getBridge: function() {
            if (this._initialized && this._bridge) return this._bridge;
            this.init();
            if (!this._bridge) this._probe();
            return this._bridge;
        },

        /**
         * 探测当前环境是否已可用桥（window.excelBridge / App.excelBridge 已就绪但本访问器未初始化的场景）
         */
        _probe: function() {
            var live = (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge)
                || window.excelBridge || (typeof App !== 'undefined' && App.excelBridge);
            if (live) {
                // 本地 HTTP 桥自身已有 catch，不再二次包装；只包 WebView2 宿主桥。
                if (this._usingLocalBridge) {
                    this._bridge = live;
                    if (typeof App !== 'undefined') App.excelBridge = live;
                    window.excelBridge = live;
                    this._initialized = true;
                } else {
                    this.adoptHostBridge(live);
                }
            }
        },

        /**
         * 安全获取 Bridge，返回 { bridge, ok }，避免调用方各自判断桥可用性
         */
        getBridgeSafe: function() {
            var bridge = this.getBridge();
            return bridge ? { bridge: bridge, ok: true } : { bridge: null, ok: false };
        },

        /**
         * 获取同步 Bridge 对象（用于不需要 await 的同步调用）
         */
        getSyncBridge: function() {
            if (!this._initialized) this.init();
            if (this._syncBridge) return this._syncBridge;
            var hosts = window.chrome && window.chrome.webview && window.chrome.webview.hostObjects;
            if (hosts && hosts.sync) {
                this._syncBridge = hosts.sync.excelBridge;
            }
            return this._syncBridge;
        },

        /**
         * 把 callAsync/callSync 的剩余参数归一成扁平数组。
         * 兼容两种历史调用模式：
         *   callAsync('GetMcpStatus')              → args=[]              → 调用 bridge.GetMcpStatus()
         *   callAsync('GetOpenWorkbooks', [])      → args=[[]]（嵌套数组）→ 调用 bridge.GetOpenWorkbooks()
         *   callAsync('CallLocalTool', name, args) → args=[name,args]     → 调用 bridge.CallLocalTool(name, args)
         *   callAsync('X', ['a','b'])              → args=[['a','b']]     → 调用 bridge.X('a','b')
         * 规则：当且仅当"只有一个剩余参数且该参数本身是数组"时，视为调用方传了完整参数列表，展开之。
         */
        _normalizeArgs: function(slicedArgs) {
            if (slicedArgs.length === 1 && slicedArgs[0] instanceof Array) {
                return slicedArgs[0];
            }
            return slicedArgs;
        },

        /**
         * 调用 Bridge 方法（统一错误处理）
         */
        callAsync: async function(methodName) {
            this._ensureAsyncHostObjects();
            if (!this._rawHostBridge() && !this.getBridge()) {
                return JSON.stringify({ success: false, error: 'Bridge 未初始化' });
            }
            try {
                var args = this._normalizeArgs(Array.prototype.slice.call(arguments, 1));
                return await this._invokeHostMethod(methodName, args, {});
            } catch (e) {
                return this._bridgeFailure(methodName, e);
            }
        },

        /**
         * 调用 Bridge 同步方法
         */
        callSync: function(methodName) {
            var bridge = this.getSyncBridge() || this.getBridge();
            if (!bridge) {
                return JSON.stringify({ success: false, error: 'Bridge 未初始化' });
            }
            try {
                var args = this._normalizeArgs(Array.prototype.slice.call(arguments, 1));
                var method = bridge[methodName];
                if (typeof method === 'function') {
                    return method(...args);
                } else {
                    return JSON.stringify({ success: false, error: '方法不存在: ' + methodName });
                }
            } catch (e) {
                return JSON.stringify({ success: false, error: e.message });
            }
        },

        /**
         * 本地桥（网页版无插件）连接能力。
         * 当不存在 WebView2 宿主对象时，把所有桥调用统一路由到本地 HTTP 服务
         * （默认 http://localhost:9527/api/excel，由 py-excel-bridge.py 提供），
         * 使硬编码的 window.chrome.webview.hostObjects.excelBridge.xxx() 调用
         * 在浏览器环境下也能工作。
         */
        _bridgeUrl: function() {
            if (window.ExcelBridgeConfig && window.ExcelBridgeConfig.bridgeUrl) {
                return window.ExcelBridgeConfig.bridgeUrl;
            }
            return 'http://localhost:9527/api/excel';
        },

        _initLocalBridge: function() {
            if (this._localBridgeInited) return;
            this._localBridgeInited = true;
            this._usingLocalBridge = true; // 显式标记：当前走本地 HTTP 桥（网页版），非 WebView2 宿主
            var baseUrl = this._bridgeUrl();
            var proxy = this._createLocalBridgeProxy(baseUrl);
            // 乐观赋值：立即可用，调用时再实际请求 9527
            this._bridge = proxy;
            if (typeof App !== 'undefined') { App.excelBridge = proxy; }
            window.excelBridge = proxy;
            this._initialized = true;
            this._installHostObjectShim(proxy);
            this._probeLocalBridge(baseUrl);
            if (typeof ChatLog !== 'undefined') {
                ChatLog.info('BRIDGE', '🌉 本地桥代理已就绪 (目标 ' + baseUrl + ')');
            }
        },

        _createLocalBridgeProxy: function(baseUrl) {
            var self = this;
            return new Proxy({}, {
                get: function(target, prop) {
                    // 避免被当成 thenable；仅字符串方法名走代理
                    if (prop === 'then' || typeof prop !== 'string') return undefined;
                    return function() {
                        var args = Array.prototype.slice.call(arguments);
                        return self._callLocalBridge(baseUrl, prop, args);
                    };
                }
            });
        },

        _callLocalBridge: function(baseUrl, methodName, args) {
            // 契约对齐 VSTO：桥方法返回 JSON 字符串，前端统一 JSON.parse。
            // 因此这里也返回字符串，保证硬编码调用契约一致。
            return fetch(baseUrl + '/invoke', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ method: methodName, arguments: args })
            }).then(function(r) { return r.json();             }).then(function(resp) {
                if (resp && resp.success) {
                    var d = (resp.data !== undefined) ? resp.data : resp;
                    // VSTO 部分方法（如表名/引擎名）直接返回裸字符串，需原样透传；
                    // 其余返回对象的方法才 JSON.stringify，供前端 JSON.parse。
                    if (typeof d === 'string') return d;
                    return JSON.stringify(d);
                }
                return JSON.stringify({ success: false, error: (resp && resp.error) ? resp.error : '本地桥返回异常' });
            }).catch(function(err) {
                return JSON.stringify({ success: false, error: '本地桥连接失败: ' + (err && err.message ? err.message : err) });
            });
        },

        _installHostObjectShim: function(proxy) {
            try {
                var hosts = window.chrome && window.chrome.webview && window.chrome.webview.hostObjects;
                if (hosts && hosts.excelBridge) return; // WebView2 真环境不覆盖
                window.chrome = window.chrome || {};
                window.chrome.webview = window.chrome.webview || {};
                window.chrome.webview.hostObjects = window.chrome.webview.hostObjects || {};
                window.chrome.webview.hostObjects.excelBridge = proxy;
            } catch (e) { (window.ChatLog || console).warn?.('BRIDGE', '[bridge-accessor] 注入本地桥 mock 失败:', e); }
        },

        _probeLocalBridge: function(baseUrl) {
            var self = this;
            fetch(baseUrl + '/status', { method: 'GET', cache: 'no-store' })
                .then(function(r) { return r.json(); })
                .then(function(d) {
                    self._localBridgeOnline = !!(d && d.status === 'ok');
                    if (typeof ChatLog !== 'undefined') {
                        ChatLog.info('BRIDGE', '[本地桥] ' + (self._localBridgeOnline ? '在线' : '离线') +
                            (d && d.pywin32Available ? ' (pywin32 真实 Excel)' : (d && d.mockMode ? ' (模拟模式)' : '')));
                    }
                })
                .catch(function() { self._localBridgeOnline = false; });
        }

    };

    window.BridgeAccessor = BridgeAccessor;
    if (typeof App !== 'undefined') {
        App.BridgeAccessor = BridgeAccessor;
    }

})();