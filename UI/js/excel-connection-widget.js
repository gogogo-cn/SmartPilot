/**
 * ============================================================
 * excel-connection-widget.js — Excel Connection Widget
 * ============================================================
 * 功能：Excel 连接状态组件（VSTO 内嵌 / 本地 9527 桥）。
 *
 * 对外：App.ExcelConnection.{connect,disconnect,loadWorkbooks,switchWorkbook,...}
 * 宿主差异：
 *   - VSTO：始终内嵌，隐藏「连接/断开」，切换工作簿直接作用于当前 Excel
 *   - 本地桥：需手动连接；断开只清前端状态
 * ============================================================
 */

ChatLog.info('CONNECTOR', '[ExcelConnectionWidget] 📦 脚本开始加载...');

(function() {
    'use strict';

    function tt(key, vars) {
        try {
            if (window.App && App.i18n && typeof App.i18n.t === 'function') return App.i18n.t(key, vars);
            if (typeof window.t === 'function') return window.t(key, vars);
        } catch (e) { /* ignore */ }
        return key;
    }

    
    function getBridgeUrl() {
        return (window.ExcelBridgeConfig && window.ExcelBridgeConfig.defaultBridgeUrl)
            || (App.Endpoints ? App.Endpoints.excelBridge.baseUrl : 'http://localhost:9527/api/excel');
    }

    // 是否在真实 VSTO/WebView2 宿主内（Excel 由 hostObjects 提供，不走本地 9527 桥）。
    // 注意：网页版 bridge-accessor 会安装 hostObjects 垫片，因此不能只判 hostObjects 存在；
    // 真实 WebView2 必有 chrome.webview.postMessage，垫片没有；再结合 BridgeAccessor 的本地桥标记兜底。
    function isHostEnv() {
        if (window.BridgeAccessor && window.BridgeAccessor._usingLocalBridge) return false;
        var wv = window.chrome && window.chrome.webview;
        return !!(wv && wv.hostObjects && typeof wv.postMessage === 'function');
    }

    // 统一调用桥方法：优先 BridgeAccessor（已按环境选择 hostObjects / 本地 9527 桥），
    // 失败再回退到直接 HTTP（网页版本地桥）。返回解析后的对象或 { success:false, error }
    function invokeBridge(method, args) {
        if (window.BridgeAccessor) {
            return window.BridgeAccessor.callAsync(method, (args || []));
        }
        return fetch(getBridgeUrl() + '/invoke', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ method: method, arguments: (args || []) })
        }).then(function (r) { return r.json(); });
    }

    // 检测桥/Excel 是否在线
    function fetchStatus() {
        if (isHostEnv()) {
            // WebView2 宿主：hostObjects 存在即视为 Excel 可用
            return Promise.resolve({ excelConnected: true, status: 'ok' });
        }
        return fetch(getBridgeUrl() + '/status', { method: 'GET', cache: 'no-store', signal: AbortSignal.timeout(5000) })
            .then(function (r) { return r.json(); })
            .catch(function () { return { excelConnected: false }; });
    }

    // 将桥调用返回值归一化为 { success, data, error }
    // VSTO：GetActiveSheetName 返回裸字符串 "销售数据"；ActivateWorkbook 字段在顶层；
    // 本地桥：常包一层 {success,data:{...}}。二者都要兼容。
    function parseBridgeResult(res) {
        if (res == null || res === '') return { success: false, error: '空响应' };
        if (typeof res === 'object') {
            if ('success' in res) return res;
            return { success: true, data: res };
        }
        if (typeof res === 'string') {
            var trimmed = res.trim();
            if (trimmed.charAt(0) === '{' || trimmed.charAt(0) === '[') {
                try {
                    var obj = JSON.parse(trimmed);
                    if (obj && typeof obj === 'object' && 'success' in obj) return obj;
                    return { success: true, data: obj };
                } catch (e) {
                    return { success: false, error: trimmed };
                }
            }
            // 裸字符串（工作表名等）
            return { success: true, data: trimmed };
        }
        return { success: true, data: res };
    }

    function bridgeErr(data, fallback) {
        if (!data) return fallback || '操作失败';
        return String(data.error || data.message || data.detail || fallback || '操作失败');
    }

    /** 从 ActivateWorkbook 等响应里取出 activeSheet / sheetCount / sheets（兼容 VSTO 顶层与 Python data 层） */
    function pickActivateInfo(data) {
        var nested = (data && data.data && typeof data.data === 'object') ? data.data : null;
        var activeSheet = (nested && nested.activeSheet) || (data && data.activeSheet) || '';
        var sheets = null;
        if (nested && Array.isArray(nested.sheets)) sheets = nested.sheets.slice();
        else if (data && Array.isArray(data.sheets)) sheets = data.sheets.slice();
        var sheetCount = null;
        if (sheets) sheetCount = sheets.length;
        else if (nested && nested.sheetCount != null) sheetCount = nested.sheetCount;
        else if (data && data.sheetCount != null) sheetCount = data.sheetCount;
        return {
            activeSheet: activeSheet || '',
            sheetCount: sheetCount == null ? 0 : sheetCount,
            sheets: sheets
        };
    }

    /** 把 GetSheets / ActivateWorkbook 的各种返回收敛成 string[] */
    function coerceSheetList(parsed) {
        if (!parsed) return [];
        if (Array.isArray(parsed)) return parsed.filter(function (n) { return typeof n === 'string' && n; });
        if (parsed.success === false) return [];
        var list = parsed.data;
        if (typeof list === 'string') {
            try { list = JSON.parse(list); } catch (_e) { list = null; }
        }
        if (!Array.isArray(list) && Array.isArray(parsed.sheets)) list = parsed.sheets;
        if (!Array.isArray(list)) return [];
        return list.filter(function (n) { return typeof n === 'string' && n; });
    }

    function toast(msg, type) {
        if (App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function') {
            var color = type === 'error' ? '#c62828' : (type === 'warning' ? '#ef6c00' : '#999');
            App.fixedStatusBar.setExtraStatus(' <span style="color:' + color + ';font-size:11px;margin-left:8px;">' + msg + '</span>', true);
        } else {
            showToast(msg, type || 'info');
        }
    }

    function initialize() {
        ChatLog.info('CONNECTOR', '[ExcelConnectionWidget] 🚀 开始初始化...');
        
        if (typeof window.App === 'undefined') {
            ChatLog.warn('CONNECTOR', '[ExcelConnectionWidget] ⏳ App 未就绪，等待...');
            setTimeout(initialize, 200);
            return;
        }
        
        var container = document.getElementById('connection-status');
        if (!container) {
            ChatLog.error('CONNECTOR', '[ExcelConnectionWidget] ❌ 找不到 #connection-status！');
            return;
        }
        
        ChatLog.info('CONNECTOR', '[ExcelConnectionWidget] ✅ 找到容器');

        // Word/PPT：本控件无意义，直接隐藏（避免 host-context 延迟收口前闪现「已连接+断开」）
        if (window.SmartPilotHost && !window.SmartPilotHost.isExcel) {
            container.style.display = 'none';
            ChatLog.info('CONNECTOR', '[ExcelConnectionWidget] 非 Excel 宿主，已隐藏连接控件');
            return;
        }

        createWidget(container);
        container.title=tt('status.detecting');

        // WebView2/VSTO 宿主：Excel 由 hostObjects 直接提供，标记为已连接并拉取工作簿
        if (isHostEnv()) {
            var ex = window.App && window.App.ExcelConnection;
            if (ex) {
                ex.isConnected = true;
                ex.connectionTime = new Date();
                ex.updateStatus('connected', tt('conn.connected'));
                ex.applyModeCard('vsto');
                ex.loadWorkbooks();
                if (typeof startMonitoring === 'function') startMonitoring();
            }
        } else {
            // 网页版（本地 9527 桥）：加载时静默探测一次真实状态，不假装已连接
            fetchStatus().then(function (data) {
                var ex = window.App && window.App.ExcelConnection;
                if (!ex) return;
                if (data && data.excelConnected) {
                    ex.isConnected = true;
                    ex.connectionTime = new Date();
                    ex.updateStatus('connected', tt('conn.connected'));
                    ex.applyModeCard('bridge');
                    ex.loadWorkbooks();
                    if (typeof startMonitoring === 'function') startMonitoring();
                } else {
                    // 区分：桥没起 / Excel 没开 / Excel 在运行但 COM 连不上（权限不一致）
                    var label = tt('conn.serviceDown');
                    if (data && data.reason) {
                        label = data.reason.length > 24 ? tt('conn.cannotConnect') : data.reason;
                    } else if (data && data.excelRunning) {
                        label = tt('conn.cannotConnect');
                    }
                    ex.updateStatus('disconnected', label);
                    ex.hideWorkbookSelector();
                }
            }).catch(function () { /* 保持"检测中"，由用户手动点连接 */ });
        }
        ChatLog.info('CONNECTOR', '[ExcelConnectionWidget] ✅ 初始化完成！');
    }
    
    function createWidget(container) {
        container.innerHTML = `
            <div id="excel-connection-widget" style="display:flex;align-items:center;gap:8px;position:relative;">
                <div onclick="toggleMenu()" style="cursor:pointer;display:flex;align-items:center;gap:6px;padding:4px 10px;border-radius:20px;background:rgba(255,255,255,0.15);">
                    <span id="status-dot" style="width:8px;height:8px;border-radius:50%;background:#4caf50;box-shadow:0 0 8px rgba(76,175,80,0.6);animation:pulse 2s infinite;"></span>
                    <span id="status-text" style="font-size:13px;font-weight:500;color:white;" data-i18n="status.detecting">检测中</span>
                    <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"></polyline></svg>
                </div>
                
                <div id="connection-menu" style="display:none;position:absolute;top:-6px;left:calc(100% + 10px);right:auto;margin-top:0;background:white;border-radius:12px;box-shadow:0 8px 30px rgba(0,0,0,0.15);padding:16px;min-width:380px;z-index:10000;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;">
                    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:14px;padding-bottom:12px;border-bottom:2px solid #f0f0f0;">
                        <div style="font-weight:600;font-size:15px;color:#333;" data-i18n="conn.title">📡 Excel 连接状态</div>
                        <button onclick="closeMenu()" style="background:none;border:none;font-size:20px;cursor:pointer;color:#999;">&times;</button>
                    </div>
                    
                    <div id="status-card" style="padding:14px;border-radius:10px;margin-bottom:14px;background:linear-gradient(135deg,#e8f5e9 0%,#c8e6c9 100%);border-left:4px solid #4caf50;">
                        <div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;">
                            <span id="menu-dot" style="width:12px;height:12px;border-radius:50%;background:#4caf50;box-shadow:0 0 10px rgba(76,175,80,0.5);"></span>
                            <span id="menu-text" style="font-weight:600;font-size:14px;color:#2e7d32;" data-i18n="conn.checking">正在检测...</span>
                        </div>
                        <div id="status-card-meta" style="font-size:12px;color:#555;line-height:1.6;">
                            <div>🖥️ <span data-i18n="conn.mode">模式:</span> <strong id="conn-mode-label">--</strong></div>
                            <div>⏱️ <span data-i18n="conn.time">时间:</span> <strong id="conn-time">--</strong></div>
                        </div>
                        <div id="conn-host-hint" style="display:none;margin-top:8px;font-size:11px;color:#2e7d32;line-height:1.45;" data-i18n="conn.hostHint">已内嵌于 Excel，工作簿切换会直接作用于当前窗口</div>
                    </div>
                    
                    <!-- 工作簿选择器 -->
                    <div id="workbook-selector" style="margin-bottom:14px;display:none;">
                        <div style="font-weight:600;font-size:13px;color:#333;margin-bottom:8px;" data-i18n="conn.pickWorkbook">📚 选择Excel工作簿</div>
                        <select id="workbook-select" onchange="switchWorkbook(this.value)" style="width:100%;padding:10px;border:2px solid #e0e0e0;border-radius:8px;font-size:13px;background:white;cursor:pointer;transition:all 0.2s;">
                            <option value="" data-i18n="conn.loading">加载中</option>
                        </select>
                        <div id="workbook-info" style="margin-top:8px;padding:8px 10px;background:#f5f5f5;border-radius:6px;font-size:12px;color:#666;display:none;">
                            <div>📚 <span data-i18n="conn.workbook">工作簿:</span> <strong id="info-workbook">--</strong></div>
                            <div style="margin-top:4px;">📄 <span data-i18n="conn.activeSheet">活动工作表:</span> <strong id="active-sheet">--</strong></div>
                            <div style="margin-top:4px;">📋 <span data-i18n="conn.sheetCount">工作表数量:</span> <strong id="sheet-count">--</strong></div>
                        </div>
                    </div>

                    <!-- 工作表选择器 -->
                    <div id="sheet-selector" style="margin-bottom:14px;display:none;">
                        <div style="font-weight:600;font-size:13px;color:#333;margin-bottom:8px;" data-i18n="conn.pickSheet">📄 选择工作表</div>
                        <select id="sheet-select" onchange="switchSheet(this.value)" style="width:100%;padding:10px;border:2px solid #e0e0e0;border-radius:8px;font-size:13px;background:white;cursor:pointer;transition:all 0.2s;">
                            <option value="" data-i18n="conn.loading">加载中</option>
                        </select>
                    </div>
                    
                    <div style="display:flex;flex-direction:column;gap:8px;">
                        <button id="btn-connect" onclick="connect()" style="padding:11px;background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:white;border:none;border-radius:8px;cursor:pointer;font-size:14px;font-weight:600;display:none;" data-i18n="conn.connect">🔗 连接到 Excel</button>
                        <button id="btn-refresh-workbooks" onclick="refreshWorkbooks()" style="padding:11px;background:#fff;color:#667eea;border:2px solid #667eea;border-radius:8px;cursor:pointer;font-size:14px;display:none;" data-i18n="conn.refreshList">📚 刷新工作簿列表</button>
                        <button id="btn-disconnect" onclick="disconnect()" style="padding:11px;background:#fff;color:#e74c3c;border:2px solid #e74c3c;border-radius:8px;cursor:pointer;font-size:14px;" data-i18n="conn.disconnect">🔌 断开连接</button>
                        <button onclick="refreshStatus()" style="padding:10px;background:#f8f9fa;color:#666;border:1px solid #ddd;border-radius:8px;cursor:pointer;font-size:13px;" data-i18n="conn.refresh">🔄 刷新状态</button>
                    </div>
                </div>
            </div>
        `;
        
        injectStyles();
        try { if (window.App && App.i18n) App.i18n.apply(container); } catch (e) {}
        bindEvents();
        
        window.App.ExcelConnection = {
            isConnected:false,
            connectionTime:null,
            currentWorkbook:null,
            workbooks:[],
            
            toggleMenu:function(){
                var m=document.getElementById('connection-menu');
                if(m)m.style.display=m.style.display==='block'?'none':'block';
            },
            
            closeMenu:function(){
                var m=document.getElementById('connection-menu');
                if(m)m.style.display='none';
            },
            
            applyModeCard:function(mode){
                var label=document.getElementById('conn-mode-label');
                var hint=document.getElementById('conn-host-hint');
                if(label) label.textContent = mode==='vsto' ? tt('conn.modeVsto') : tt('conn.modeBridge');
                if(hint) hint.style.display = mode==='vsto' ? 'block' : 'none';
                var timeEl=document.getElementById('conn-time');
                if(timeEl && !timeEl.textContent) timeEl.textContent=tt('conn.justNow');
            },

            connect:async function(){
                if(isHostEnv()){
                    // VSTO 内嵌无需手动连接
                    this.isConnected=true;
                    this.connectionTime=new Date();
                    this.updateStatus('connected',tt('conn.connected'));
                    this.applyModeCard('vsto');
                    this.loadWorkbooks();
                    return;
                }
                ChatLog.info('CONNECTOR', '[ExcelConnectionWidget]  用户点击连接按钮');
                this.updateStatus('connecting',tt('conn.connecting'));
                try{
                    var data=await fetchStatus();
                    if(data.excelConnected){
                        this.connectionTime=new Date();
                        this.updateStatus('connected',tt('conn.connected'));
                        this.applyModeCard('bridge');
                        toast('✅ '+tt('conn.connectOk'),'success');
                        this.loadWorkbooks();
                        if(typeof startMonitoring === 'function') startMonitoring();
                    }else{
                        ChatLog.warn('CONNECTOR', '[ExcelConnectionWidget] ⚠️ Excel未打开');
                        this.updateStatus('disconnected',tt('conn.excelNotOpen'));
                        toast('⚠️ '+tt('conn.needExcel'),'warning');
                        showConnectDialog();
                    }
                }catch(e){
                    ChatLog.error('CONNECTOR', '[ExcelConnectionWidget] ❌ 连接失败:',e);
                    this.updateStatus('disconnected',tt('conn.serviceDown'));
                    showConnectDialog();
                }
            },
            
            disconnect:function(){
                if(isHostEnv()){
                    // 内嵌宿主无法真正断开；仅提示
                    toast(tt('conn.hostHint'),'info');
                    return;
                }
                if(confirm(tt('conn.confirmDisconnect'))){
                    this.updateStatus('disconnected',tt('conn.disconnected'));
                    this.connectionTime=null;
                    this.currentWorkbook=null;
                    this.workbooks=[];
                    this.hideWorkbookSelector();
                }
            },
            
            refreshStatus:async function(){
                await checkConn();
                if(this.isConnected){
                    this.loadWorkbooks();
                }
                toast('✅ '+tt('conn.refreshed'),'info');
            },
            
            loadWorkbooks:async function(){
                if (window.SmartPilotHost && !window.SmartPilotHost.isExcel) {
                    ChatLog.info('CONNECTOR', '[ExcelConnectionWidget] 非 Excel 宿主，跳过工作簿加载');
                    return;
                }
                ChatLog.info('CONNECTOR', '[ExcelConnectionWidget]  开始加载工作簿列表...');
                            
                if(!this.isConnected){
                    ChatLog.warn('CONNECTOR', '[ExcelConnectionWidget] ⚠️ 未连接，无法加载工作簿');
                    toast('⚠️ '+tt('conn.needConnect'),'warning');
                    return;
                }
                            
                try{
                    var raw=await invokeBridge('GetOpenWorkbooks', []);
                    var data=parseBridgeResult(raw);
                    ChatLog.info('CONNECTOR', '[ExcelConnectionWidget]  后端返回数据:',data);
                                
                    if(data.success !== false){
                        // 兼容两种返回：Python 桥 {data:{workbooks}} 与 VSTO 桥 {workbooks}
                        this.workbooks=(data.data && data.data.workbooks) || data.workbooks || [];
                        ChatLog.info('CONNECTOR', '[ExcelConnectionWidget] ✅ 获取到'+this.workbooks.length+'个工作簿');
                        this.renderWorkbookSelector();
                        this.loadSheets();
                    }else{
                        var errMsg=bridgeErr(data,'获取工作簿失败');
                        // 桥投影失败（旧 callAsync 拆方法 / 缺方法）：勿再 throw 成红条，避免刷屏
                        if(/bridge-call-failed|e is not a function|t is not a function|0x80020006|0x80070490|未知名称|找不到元素/i.test(errMsg)){
                            ChatLog.warn('CONNECTOR', '[ExcelConnectionWidget] 桥暂不可用，跳过工作簿列表:', errMsg, data.method||'GetOpenWorkbooks');
                            this.workbooks=[];
                            this.renderWorkbookSelector();
                            return;
                        }
                        throw new Error(errMsg);
                    }
                }catch(e){
                    var em=(e&&e.message)?e.message:String(e);
                    if(/e is not a function|t is not a function|0x80020006|0x80070490|未知名称|bridge-call-failed|找不到元素/i.test(em)){
                        ChatLog.warn('CONNECTOR', '[ExcelConnectionWidget] 桥调用失败（已降噪）:', em);
                    }else{
                        ChatLog.error('CONNECTOR', '[ExcelConnectionWidget] ❌ 加载工作簿失败:',e);
                        toast('❌ '+tt('conn.loadWbFail',{msg:em}),'error');
                    }
                    this.renderWorkbookSelector();
                }
            },
            
            updateWorkbookInfo:function(wbName, sheetName, sheetCount){
                var info=document.getElementById('workbook-info');
                var infoWb=document.getElementById('info-workbook');
                var activeSheet=document.getElementById('active-sheet');
                var sheetCountEl=document.getElementById('sheet-count');
                if(!info) return;
                info.style.display='block';
                if(infoWb) infoWb.textContent = wbName || '--';
                if(activeSheet) activeSheet.textContent = sheetName || '--';
                if(sheetCountEl) sheetCountEl.textContent = tt('conn.sheetUnit',{n: sheetCount == null ? 0 : sheetCount});
            },

            renderWorkbookSelector:function(){
                var selector=document.getElementById('workbook-selector');
                var select=document.getElementById('workbook-select');
                var refreshBtn=document.getElementById('btn-refresh-workbooks');
                var self=this;
                
                if(!selector||!select){
                    ChatLog.error('CONNECTOR', '[ExcelConnectionWidget] ❌ 找不到工作簿选择器元素');
                    return;
                }
                
                ChatLog.info('CONNECTOR', '[ExcelConnectionWidget]  渲染工作簿选择器，工作簿数量:',this.workbooks.length);
                
                selector.style.display='block';
                if(refreshBtn) refreshBtn.style.display='block';
                
                select.innerHTML='';
                
                if(this.workbooks.length===0){
                    select.innerHTML='<option value="">'+tt('conn.noWorkbooks')+'</option>';
                    var info0=document.getElementById('workbook-info');
                    if(info0)info0.style.display='none';
                    return;
                }
                
                this.workbooks.forEach(function(wb){
                    var option=document.createElement('option');
                    option.value=wb.name||wb.fullName;
                    option.textContent=wb.name||wb.fullName;
                    select.appendChild(option);
                });
                
                var firstWb=null;
                if(this.currentWorkbook){
                    select.value=this.currentWorkbook;
                    firstWb=this.workbooks.find(function(w){return (w.name||w.fullName)===self.currentWorkbook;})||null;
                }
                if(!firstWb){
                    firstWb=this.workbooks.find(function(w){return w.isActive;})||this.workbooks[0];
                    select.value=firstWb.name||firstWb.fullName;
                    this.currentWorkbook=select.value;
                }
                var nSheets=firstWb.sheets ? firstWb.sheets.length : (firstWb.sheetCount||0);
                // 活动表名可能尚未拿到：先填 sheetCount，表名由 loadSheets → GetActiveSheetName 补齐
                this.updateWorkbookInfo(
                    firstWb.name||firstWb.fullName,
                    firstWb.activeSheet || '',
                    nSheets
                );
            },
            
            loadSheets:async function(){
                // 序号守卫：快速连切工作簿时，丢弃过期的 GetSheets 响应，避免旧表覆盖新表
                var seq = (this._sheetLoadSeq = (this._sheetLoadSeq || 0) + 1);
                try{
                    var list = [];
                    for (var attempt = 0; attempt < 2; attempt++) {
                        if (seq !== this._sheetLoadSeq) return;
                        var raw = await invokeBridge('GetSheets', []);
                        if (seq !== this._sheetLoadSeq) return;
                        list = coerceSheetList(parseBridgeResult(raw));
                        if (list.length > 0) break;
                        // Activate 刚完成时偶发读到空列表，短暂重试一次
                        if (attempt === 0) {
                            await new Promise(function (r) { setTimeout(r, 120); });
                        }
                    }
                    if (seq !== this._sheetLoadSeq) return;
                    ChatLog.info('CONNECTOR', '[ExcelConnectionWidget] GetSheets →', list.length, '张表');
                    this.renderSheetSelector(list);
                }catch(e){
                    if (seq !== this._sheetLoadSeq) return;
                    ChatLog.warn('CONNECTOR', '[ExcelConnectionWidget] 加载工作表失败:',e);
                    var sel=document.getElementById('sheet-selector');
                    if(sel)sel.style.display='none';
                }
            },

            renderSheetSelector:function(sheetNames){
                var wrap=document.getElementById('sheet-selector');
                var sel=document.getElementById('sheet-select');
                if(!wrap||!sel)return;
                if(!sheetNames || sheetNames.length===0){
                    wrap.style.display='none';
                    return;
                }
                wrap.style.display='block';
                sel.innerHTML='';
                var self=this;
                sheetNames.forEach(function(n){
                    var opt=document.createElement('option');
                    opt.value=n; opt.textContent=n;
                    sel.appendChild(opt);
                });
                // 默认选中当前活动工作表（VSTO 返回裸字符串，由 parseBridgeResult 归一）
                invokeBridge('GetActiveSheetName', []).then(function(raw){
                    var r=parseBridgeResult(raw);
                    var cur=(r && r.success !== false && r.data != null) ? r.data : '';
                    if(typeof cur==='string' && cur){
                        sel.value=cur;
                        if(self.onSheetSelected) self.onSheetSelected(cur);
                    } else if(sheetNames.length){
                        // 回退：选第一项并写到信息区
                        sel.value=sheetNames[0];
                        if(self.onSheetSelected) self.onSheetSelected(sheetNames[0]);
                    }
                }).catch(function(_e){ (window.ChatLog || console).warn?.('UI', '[excel-connection] 读取活动表名失败:', _e); });
            },

            switchSheet:async function(sheetName){
                if(!sheetName)return;
                try{
                    var raw=await invokeBridge('ActivateSheet', [sheetName]);
                    var data=parseBridgeResult(raw);
                    if(data.success !== false){
                        this.onSheetSelected(sheetName);
                        toast('✅ '+tt('conn.switchedSheet',{name:sheetName}),'success');
                    }else{
                        throw new Error(bridgeErr(data,'切换工作表失败'));
                    }
                }catch(e){
                    ChatLog.error('CONNECTOR', '[ExcelConnectionWidget] 切换工作表失败:',e);
                    toast('❌ '+tt('conn.switchSheetFail',{msg:e.message||e}),'error');
                }
            },

            onSheetSelected:function(sheetName){
                ChatLog.info('CONNECTOR', '[ExcelConnectionWidget] 当前工作表:',sheetName);
                var el=document.getElementById('active-sheet');
                if(el)el.textContent=sheetName;
                var infoWb=document.getElementById('info-workbook');
                if(infoWb && this.currentWorkbook) infoWb.textContent=this.currentWorkbook;
            },

            hideWorkbookSelector:function(){
                var selector=document.getElementById('workbook-selector');
                var refreshBtn=document.getElementById('btn-refresh-workbooks');
                var info=document.getElementById('workbook-info');
                var sheetWrap=document.getElementById('sheet-selector');

                if(selector)selector.style.display='none';
                if(refreshBtn)refreshBtn.style.display='none';
                if(info)info.style.display='none';
                if(sheetWrap)sheetWrap.style.display='none';
            },
            
            switchWorkbook:async function(workbookName){
                if(!workbookName){
                    toast('⚠️ '+tt('conn.pickWbWarn'),'warning');
                    return;
                }
                
                try{
                    var raw=await invokeBridge('ActivateWorkbook', [workbookName]);
                    var data=parseBridgeResult(raw);
                    if(data.success !== false){
                        this.currentWorkbook=workbookName;
                        var info=pickActivateInfo(data);
                        // 若桥没回 sheetCount，用列表里缓存的
                        if(!info.sheetCount){
                            var cached=this.workbooks.find(function(w){return (w.name||w.fullName)===workbookName;});
                            if(cached) info.sheetCount = cached.sheets ? cached.sheets.length : (cached.sheetCount||0);
                        }
                        this.updateWorkbookInfo(workbookName, info.activeSheet, info.sheetCount);
                        toast('✅ '+tt('conn.switchedWb',{name:workbookName}),'success');

                        // 优先用 Activate 一并返回的 sheets（VSTO 新契约 / Python data.sheets），
                        // 避免再打 GetSheets 时 ActiveWorkbook 尚未切稳读到空列表。
                        if (info.sheets && info.sheets.length) {
                            this._sheetLoadSeq = (this._sheetLoadSeq || 0) + 1;
                            this.renderSheetSelector(info.sheets);
                            if (info.activeSheet) this.onSheetSelected(info.activeSheet);
                        } else {
                            await this.loadSheets();
                        }
                    }else{
                        throw new Error(bridgeErr(data,'切换失败'));
                    }
                }catch(e){
                    ChatLog.error('CONNECTOR', '[ExcelConnectionWidget] 切换工作簿失败:',e);
                    toast('❌ '+tt('conn.switchWbFail',{msg:e.message||e}),'error');
                }
            },
            
            updateStatus:function(status,text){
                this.isConnected=(status==='connected');
                var c={connected:{d:'#4caf50',t:'#2e7d32'},disconnected:{d:'#f44336',t:'#c62828'},connecting:{d:'#ff9800',t:'#ef6c00'}};
                var cfg=c[status]||c.disconnected;
                
                ['status-dot','menu-dot'].forEach(function(id){
                    var el=document.getElementById(id);
                    if(el){el.style.background=cfg.d;el.style.boxShadow='0 0 10px'+cfg.d+'80';}
                });
                
                var st=document.getElementById('status-text');
                var mt=document.getElementById('menu-text');
                var card=document.getElementById('status-card');
                if(st) st.textContent=text;
                if(mt) mt.textContent=text;
                if(card) card.style.borderLeftColor=cfg.d;
                
                var cs=document.getElementById('connection-status');
                if(cs) cs.title=text;
                
                var btnConn=document.getElementById('btn-connect');
                var btnDisc=document.getElementById('btn-disconnect');
                if(isHostEnv()){
                    // VSTO 内嵌：连接/断开无意义，全部隐藏
                    if(btnConn) btnConn.style.display='none';
                    if(btnDisc) btnDisc.style.display='none';
                }else{
                    if(btnConn) btnConn.style.display=this.isConnected?'none':'block';
                    if(btnDisc) btnDisc.style.display=this.isConnected?'block':'none';
                }
                
                if(this.isConnected&&this.connectionTime){
                    var el=document.getElementById('conn-time');
                    if(el){
                        var m=Math.floor((Date.now()-this.connectionTime.getTime())/60000);
                        el.textContent=m<1?tt('conn.justNow'):m<60?tt('conn.minsAgo',{n:m}):tt('conn.hoursAgo',{n:Math.floor(m/60)});
                    }
                }
            }
        };

        // 供单测 / 调试：归一化逻辑
        window.App.ExcelConnection._parseBridgeResult = parseBridgeResult;
        window.App.ExcelConnection._pickActivateInfo = pickActivateInfo;
        window.App.ExcelConnection._coerceSheetList = coerceSheetList;
        window.App.ExcelConnection._bridgeErr = bridgeErr;
        
        // 全局函数
        window.toggleMenu=function(){if(window.App.ExcelConnection)window.App.ExcelConnection.toggleMenu();};
        window.closeMenu=function(){if(window.App.ExcelConnection)window.App.ExcelConnection.closeMenu();};
        window.connect=function(){if(window.App.ExcelConnection)window.App.ExcelConnection.connect();};
        window.disconnect=function(){if(window.App.ExcelConnection)window.App.ExcelConnection.disconnect();};
        window.refreshStatus=function(){if(window.App.ExcelConnection)window.App.ExcelConnection.refreshStatus();};
        window.switchWorkbook=function(name){if(window.App.ExcelConnection)window.App.ExcelConnection.switchWorkbook(name);};
        window.switchSheet=function(name){if(window.App.ExcelConnection)window.App.ExcelConnection.switchSheet(name);};
        window.refreshWorkbooks=function(){if(window.App.ExcelConnection)window.App.ExcelConnection.loadWorkbooks();};
        
        setTimeout(checkConn,500);
    }
    
    function injectStyles(){
        if(document.getElementById('ecw-styles'))return;
        var s=document.createElement('style');
        s.id='ecw-styles';
        s.textContent='@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.5}} .connection-menu{animation:slideDown 0.2s ease-out}@keyframes slideDown{from{opacity:0;transform:translateY(-10px)}to{opacity:1;transform:translateY(0)}}';
        document.head.appendChild(s);
    }
    
    function bindEvents(){
        document.addEventListener('click',function(e){
            var w=document.getElementById('excel-connection-widget');
            var m=document.getElementById('connection-menu');
            if(w&&m&&!w.contains(e.target))m.style.display='none';
        });
        document.addEventListener('keydown',function(e){
            if(e.key==='Escape'){
                var m=document.getElementById('connection-menu');
                if(m&&m.style.display==='block')m.style.display='none';
            }
        });
    }
    
    function startMonitoring(){
        checkConn();
        setInterval(checkConn,30000);
    }
    
    async function checkConn(){
        // 只在用户已主动连接后才自动检测
        if(!window.App.ExcelConnection || !window.App.ExcelConnection.isConnected){
            return;
        }
        
        try{
            var data=await fetchStatus();
            // 检查Excel是否仍然连接
            if(data.excelConnected){
                    // 保持已连接状态，刷新连接时间
                    window.App.ExcelConnection.connectionTime=new Date();
                    window.App.ExcelConnection.updateStatus('connected',tt('conn.connected'));
                }else{
                    // Excel被关闭，断开连接
                    window.App.ExcelConnection.updateStatus('disconnected',tt('conn.excelClosed'));
                    window.App.ExcelConnection.hideWorkbookSelector();
                    toast('⚠️ '+tt('conn.excelClosedToast'),'warning');
                }
        }catch(e){
            // 连接失败，更新为断开状态
            if(window.App.ExcelConnection && window.App.ExcelConnection.isConnected){
                window.App.ExcelConnection.updateStatus('disconnected',tt('conn.linkLost'));
                window.App.ExcelConnection.hideWorkbookSelector();
            }
        }
    }
    
    function showConnectDialog(){
        var d=document.createElement('div');
        d.style.cssText='position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);display:flex;align-items:center;justify-content:center;z-index:10001;';
        d.innerHTML='<div style="background:white;border-radius:16px;padding:30px;max-width:480px;width:90%;text-align:center;"><div style="font-size:48px;margin-bottom:12px;">🔗</div><h2 style="margin:10px 0;color:#333;">连接到 Microsoft Excel</h2><p style="color:#666;margin:16px 0;">请先启动本地 Excel 桥服务（端口 9527）</p><div style="background:#f0f4ff;padding:16px;border-radius:10px;text-align:left;margin-bottom:20px;"><ol style="margin:0;padding-left:20px;color:#555;line-height:1.8;"><li>打开 Excel（至少新建/打开一个工作簿）</li><li>进入 Tools 文件夹，运行 <code>py-excel-bridge.py</code>（或对应的启动脚本）</li><li>看到本地桥监听 <code>http://localhost:9527</code> 后，回到此处点「🔗 连接到 Excel」</li></ol></div><button onclick="this.closest(\'div[style*=fixed]\').remove()" style="padding:11px 30px;background:#667eea;color:white;border:none;border-radius:8px;cursor:pointer;font-size:14px;font-weight:600;">知道了</button></div>';
        d.onclick=function(e){if(e.target===d)d.remove();};
        document.body.appendChild(d);
    }
    
    function showToast(msg,type){
        var t=document.createElement('div');
        var c={success:'#4caf50',warning:'#ff9800',error:'#f44336',info:'#2196f3'};
        t.style.cssText='position:fixed;top:20px;right:20px;background:'+(c[type]||c.info)+';color:white;padding:14px 24px;border-radius:10px;box-shadow:0 4px 20px rgba(0,0,0,0.2);z-index:10002;font-size:14px;';
        t.textContent=msg;
        document.body.appendChild(t);
        setTimeout(function(){t.remove()},3000);
    }
    
    if(document.readyState==='loading'){
        document.addEventListener('DOMContentLoaded',initialize);
    }else{
        initialize();
    }
})();