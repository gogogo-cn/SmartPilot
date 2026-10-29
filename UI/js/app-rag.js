/**
 * ============================================================
 * app-rag.js — RAG Retrieval
 * ============================================================
 * 功能：RAG 检索增强模块，提供本地知识库索引、向量检索与上下文注入能力。
 *
 * 核心类/函数：
 *   - buildIndex() — 构建索引
 *   - searchIndex() — 检索索引
 *   - addDocument() — 添加文档
 *   - getRelevantContext() — 获取相关上下文
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 *   - app-utils.js
 * ============================================================
 */

(function() {
    'use strict';

    window.App.rag = {
        currentState: {
            knowledgeBases: [],
            currentKbId: null,
            documents: [],
            operationLog: [],
            config: null,
            currentSubTab: 'kb-list',
            docAutoRefreshTimer: null,
            searchText: '',
            filterStatus: 'all'
        },

        // ========== 初始化 ==========
        init: async function() {
            if (window.SmartPilotHost && !window.SmartPilotHost.isExcel) {
                ChatLog.info('RAG', '[RAG] 非 Excel 宿主，跳过知识库初始化');
                return;
            }
            if (!App.excelBridge) {
                ChatLog.warn('RAG', '[RAG] Excel Bridge 未就绪，跳过初始化');
                return;
            }
            ChatLog.info('RAG', '[RAG] 初始化知识库模块...');
            try {
                await this.loadKnowledgeBases();
                await this.loadConfig();
                this.bindEvents();
                // 异步探测本地 RAG 环境（Ollama/模型），不阻塞初始化
                this.checkEnvironment(true);
                ChatLog.info('RAG', '[RAG] 初始化完成');
            } catch (ex) {
                ChatLog.error('RAG', '[RAG] 初始化失败:', ex);
            }
        },

        // ========== 子 Tab 切换 ==========
        switchSubTab: function(tabName) {
            // 隐藏所有子面板
            var subtabs = document.querySelectorAll('.subtab-content');
            subtabs.forEach(function(tab) {
                tab.style.display = 'none';
            });

            // 显示目标面板
            var targetPanel = document.getElementById('subtab-' + tabName);
            if (targetPanel) {
                targetPanel.style.display = 'block';
            }

            // 更新子导航 active 状态（与扩展面板一致）
            var tabs = document.querySelectorAll('.knowledge-sub-tabs .sub-tab');
            tabs.forEach(function(tab) {
                tab.classList.remove('active');
            });
            var activeTab = document.querySelector('.knowledge-sub-tabs .sub-tab[data-subtab="' + tabName + '"]');
            if (activeTab) {
                activeTab.classList.add('active');
            }

            var prevTab = this.currentState.currentSubTab;
            this.currentState.currentSubTab = tabName;

            // 离开文档管理时停止文档状态轮询，避免后台空转
            if (prevTab === 'doc-manage' && tabName !== 'doc-manage') {
                this.stopDocAutoRefresh();
            }

            // 切换到文档管理或检索测试时，更新下拉框
            if (tabName === 'doc-manage' || tabName === 'search-test') {
                this.updateKbSelects();
            }
            // 进入设置页时刷新默认知识库下拉（知识库可能刚变）
            if (tabName === 'rag-settings') {
                this.renderDefaultKbSelect();
            }
            // 进入知识库列表/文档管理时，若上次检测环境未就绪则复检（节流 10s）
            if (tabName === 'kb-list' || tabName === 'doc-manage') {
                var last = this.envGuideState.lastStatus;
                if (!last || !last.ready) {
                    this.checkEnvironment(false);
                }
            }
        },

        // ========== RAG 环境引导（Ollama / embedding 模型） ==========

        envGuideState: {
            lastStatus: null,     // 最近一次检测结果
            lastCheckAt: 0,       // 上次检测时间戳（节流）
            checking: false,
            pullTimer: null,      // 拉取进度轮询定时器
            pulling: false
        },

        /** 取翻译（带回退，与 app-history.js 一致） */
        _t: function(key, vars, fallback) {
            try {
                if (window.App && App.i18n && typeof App.i18n.t === 'function') {
                    var v = App.i18n.t(key, vars);
                    if (v && v !== key) return v;
                }
            } catch (_) { /* ignore */ }
            return fallback || key;
        },

        /** 轻提示（沿用项目惯例：优先 fixedStatusBar，回退 showToast） */
        _toast: function(msg) {
            if (App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function') {
                App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">' + this.escapeHtml(msg) + '</span>', true);
            } else if (typeof App.showToast === 'function') {
                App.showToast(msg);
            }
        },

        /**
         * 检测 RAG 本地环境并渲染引导卡片。
         * 节流 10s（force=true 跳过），避免切换 tab 时反复探测。
         */
        checkEnvironment: async function(force) {
            var st = this.envGuideState;
            if (st.checking) return;
            var now = Date.now();
            if (!force && st.lastCheckAt && (now - st.lastCheckAt) < 10000) {
                if (st.lastStatus) this.renderEnvGuide(st.lastStatus);
                return;
            }
            if (!App.excelBridge || typeof App.excelBridge.CheckRagEnvironmentAsync !== 'function') return;
            st.checking = true;
            this.renderEnvGuide({ nextStep: 'checking' });
            try {
                var raw = await App.excelBridge.CheckRagEnvironmentAsync();
                var status = this.safeParse(raw);
                if (!status || status.error) {
                    this.renderEnvGuide({ nextStep: 'error', error: status && status.error });
                } else {
                    st.lastStatus = status;
                    st.lastCheckAt = Date.now();
                    this.renderEnvGuide(status);
                    if (status.ready) this.stopPullPolling();
                }
            } catch (ex) {
                ChatLog.error('RAG', '[RAG] 环境检测失败:', ex);
                this.renderEnvGuide({ nextStep: 'error', error: ex && ex.message });
            } finally {
                st.checking = false;
            }
        },

        /**
         * 渲染环境引导卡片到两个挂载槽（kb-list / doc-manage）。
         * nextStep: checking / install_ollama / start_ollama / pull_model / ready / error
         */
        renderEnvGuide: function(status) {
            var slots = document.querySelectorAll('.rag-env-guide-slot');
            if (!slots || !slots.length) return;
            var html = '';
            var step = (status && status.nextStep) || 'ready';

            if (step === 'checking') {
                html = '<div class="rag-env-guide" style="background:#f0f7ff;border:1px solid #90caf9;border-radius:8px;padding:12px 16px;font-size:13px;color:#1a5a9a;">'
                    + '⏳ ' + this.escapeHtml(this._t('kb.envChecking', null, '正在检测本地 RAG 环境…'))
                    + '</div>';
            } else if (step === 'ready') {
                html = '';   // 就绪不占位
            } else {
                var model = (status && status.modelName) || 'bge-m3';
                var title = '', desc = '', buttons = '';
                if (step === 'install_ollama') {
                    title = this._t('kb.envInstallTitle', null, '未检测到 Ollama');
                    desc = this._t('kb.envInstallDesc', { model: model }, 'RAG 向量化依赖本地 Ollama 服务（免费开源，约 700MB）。请先下载安装，完成后点「重新检测」。Ollama 与模型不会随插件分发，需单独安装一次。');
                    buttons = '<button onclick="App.rag.openOllamaDownload()" style="padding:6px 14px;background:#0078d4;color:white;border:none;border-radius:4px;cursor:pointer;font-size:12px;">'
                        + this.escapeHtml(this._t('kb.envDownload', null, '⬇ 下载 Ollama'))
                        + '</button>'
                        + '<button onclick="App.rag.checkEnvironment(true)" style="padding:6px 14px;background:white;color:#333;border:1px solid #ccc;border-radius:4px;cursor:pointer;font-size:12px;">'
                        + this.escapeHtml(this._t('kb.envRecheck', null, '重新检测'))
                        + '</button>';
                } else if (step === 'start_ollama') {
                    title = this._t('kb.envStartTitle', null, 'Ollama 服务未运行');
                    desc = this._t('kb.envStartDesc', null, '检测到 Ollama 已安装但服务未启动。请打开 Ollama 应用（或在终端运行 ollama serve），然后点「重新检测」。');
                    buttons = '<button onclick="App.rag.checkEnvironment(true)" style="padding:6px 14px;background:#0078d4;color:white;border:none;border-radius:4px;cursor:pointer;font-size:12px;">'
                        + this.escapeHtml(this._t('kb.envRecheck', null, '重新检测'))
                        + '</button>';
                } else if (step === 'pull_model') {
                    var pulling = this.envGuideState.pulling;
                    title = this._t('kb.envPullTitle', { model: model }, '缺少向量模型 ' + model);
                    desc = this._t('kb.envPullDesc', { model: model }, 'RAG 需要嵌入模型 ' + model + '（约 1.2GB，仅首次）。点「一键拉起」自动下载，或复制命令到终端执行。');
                    if (pulling) {
                        buttons = '<button onclick="App.rag.stopModelPull()" style="padding:6px 14px;background:white;color:#333;border:1px solid #ccc;border-radius:4px;cursor:pointer;font-size:12px;">'
                            + this.escapeHtml(this._t('kb.envCancel', null, '取消'))
                            + '</button>';
                    } else {
                        buttons = '<button onclick="App.rag.startModelPull()" style="padding:6px 14px;background:#f5a623;color:white;border:none;border-radius:4px;cursor:pointer;font-size:12px;font-weight:600;">'
                            + this.escapeHtml(this._t('kb.envPullBtn', null, '⬇ 一键拉起 ' + model))
                            + '</button>'
                            + '<button onclick="App.rag.copyPullCommand()" style="padding:6px 14px;background:white;color:#333;border:1px solid #ccc;border-radius:4px;cursor:pointer;font-size:12px;">'
                            + this.escapeHtml(this._t('kb.envCopyCmd', null, '复制命令'))
                            + '</button>';
                    }
                } else {
                    title = this._t('kb.envErrorTitle', null, '环境检测失败');
                    desc = this.escapeHtml((status && status.error) || this._t('kb.envErrorDesc', null, '未知错误'));
                    buttons = '<button onclick="App.rag.checkEnvironment(true)" style="padding:6px 14px;background:#0078d4;color:white;border:none;border-radius:4px;cursor:pointer;font-size:12px;">'
                        + this.escapeHtml(this._t('kb.envRecheck', null, '重新检测'))
                        + '</button>';
                }

                html = '<div class="rag-env-guide" style="background:#fff8e6;border:1px solid #ffe0a3;border-radius:8px;padding:12px 16px;font-size:13px;">'
                    + '<div style="display:flex;gap:10px;align-items:flex-start;">'
                    + '  <div style="font-size:16px;line-height:1.4;">⚠️</div>'
                    + '  <div style="flex:1;min-width:0;">'
                    + '    <div style="font-weight:600;color:#8a6d1a;margin-bottom:4px;">' + this.escapeHtml(title) + '</div>'
                    + '    <div style="color:#6b5a2a;line-height:1.6;">' + desc + '</div>'
                    + '    <div style="margin-top:10px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;">' + buttons + '</div>'
                    + (this.envGuideState.pulling ? this._envProgressHtml() : '')
                    + '  </div>'
                    + '</div>'
                    + '</div>';
            }
            for (var i = 0; i < slots.length; i++) {
                slots[i].innerHTML = html;
            }
        },

        /** 拉取进度条 HTML（pull_model 且 pulling 时显示） */
        _envProgressHtml: function() {
            return '<div class="rag-env-progress" style="margin-top:10px;">'
                + '<div style="display:flex;justify-content:space-between;font-size:12px;color:#6b5a2a;margin-bottom:4px;">'
                + '  <span class="rag-env-pull-line">' + this.escapeHtml(this._t('kb.envPulling', null, '正在拉取…')) + '</span>'
                + '  <span class="rag-env-pull-pct">0%</span>'
                + '</div>'
                + '<div style="height:8px;background:#f0e6cc;border-radius:4px;overflow:hidden;">'
                + '  <div class="rag-env-fill" style="height:100%;width:0%;background:#f5a623;transition:width .3s;"></div>'
                + '</div>'
                + '</div>';
        },

        /** 更新进度条（两个挂载槽同步） */
        updatePullProgress: function(percent, line) {
            var pct = Math.max(0, Math.min(100, parseInt(percent, 10) || 0));
            var fills = document.querySelectorAll('.rag-env-fill');
            for (var i = 0; i < fills.length; i++) fills[i].style.width = pct + '%';
            var pcts = document.querySelectorAll('.rag-env-pull-pct');
            for (var j = 0; j < pcts.length; j++) pcts[j].textContent = pct + '%';
            if (line) {
                var lines = document.querySelectorAll('.rag-env-pull-line');
                for (var k = 0; k < lines.length; k++) lines[k].textContent = line;
            }
        },

        /** 一键拉起模型 */
        startModelPull: async function() {
            if (this.envGuideState.pulling) return;
            if (!App.excelBridge) return;
            try {
                var raw = await App.excelBridge.StartModelPull();
                var r = this.safeParse(raw);
                if (r && r.started) {
                    this.envGuideState.pulling = true;
                    this.updatePullProgress(0, this._t('kb.envPulling', null, '正在拉取…'));
                    this._startPullPolling();
                } else {
                    this._toast((r && r.message) || this._t('kb.envPullStartFail', null, '启动拉取失败'));
                }
            } catch (ex) {
                ChatLog.error('RAG', '[RAG] 启动模型拉取异常:', ex);
                this._toast(this._t('kb.envPullStartFail', null, '启动拉取失败'));
            }
        },

        _startPullPolling: function() {
            var self = this;
            this.stopPullPolling();
            this.envGuideState.pullTimer = setInterval(async function() {
                try {
                    var raw = await App.excelBridge.GetModelPullProgress();
                    var st = self.safeParse(raw);
                    if (!st) return;
                    self.updatePullProgress(st.percent, st.lastLine);
                    if (!st.running) {
                        self.stopPullPolling();
                        self.envGuideState.pulling = false;
                        if (st.error) {
                            self._toast(self._t('kb.envPullFail', null, '模型拉取失败') + ': ' + st.error);
                            self.renderEnvGuide({ nextStep: 'pull_model', modelName: self.envGuideState.lastStatus && self.envGuideState.lastStatus.modelName });
                        } else {
                            self._toast(self._t('kb.envPullDone', null, '模型拉取完成'));
                            self.checkEnvironment(true);
                        }
                    }
                } catch (_) { /* 轮询失败静默，下次继续 */ }
            }, 1500);
        },

        stopPullPolling: function() {
            if (this.envGuideState.pullTimer) {
                clearInterval(this.envGuideState.pullTimer);
                this.envGuideState.pullTimer = null;
            }
        },

        /** 取消拉取 */
        stopModelPull: async function() {
            this.stopPullPolling();
            this.envGuideState.pulling = false;
            try {
                if (App.excelBridge && typeof App.excelBridge.StopModelPull === 'function') {
                    await App.excelBridge.StopModelPull();
                }
            } catch (_) { /* ignore */ }
            this._toast(this._t('kb.envPullCancelled', null, '已取消拉取'));
            this.checkEnvironment(true);
        },

        /** 打开 Ollama 下载页 */
        openOllamaDownload: async function() {
            if (!App.excelBridge || typeof App.excelBridge.OpenExternalUrl !== 'function') {
                this._toast(this._t('kb.envNoBridge', null, '当前宿主不支持打开外部链接'));
                return;
            }
            try {
                var raw = await App.excelBridge.OpenExternalUrl('https://ollama.com/download');
                var r = this.safeParse(raw);
                if (r && r.success === false) {
                    this._toast(r.error || this._t('kb.envOpenFail', null, '打开链接失败'));
                }
            } catch (ex) {
                this._toast(this._t('kb.envOpenFail', null, '打开链接失败'));
            }
        },

        /** 复制手动拉取命令 */
        copyPullCommand: function() {
            var model = (this.envGuideState.lastStatus && this.envGuideState.lastStatus.modelName) || 'bge-m3';
            var cmd = 'ollama pull ' + model;
            if (typeof App.copyToClipboard === 'function') {
                App.copyToClipboard(cmd);
                this._toast(this._t('kb.envCopied', { cmd: cmd }, '已复制: ' + cmd));
            } else {
                this._toast(cmd);
            }
        },

        // ========== 知识库 CRUD ==========

        /**
         * 切换新建知识库表单显示/隐藏
         */
        toggleCreateForm: function() {
            var form = document.getElementById('kb-create-form');
            if (!form) return;
            if (form.style.display === 'none') {
                form.style.display = 'block';
                document.getElementById('kb-name-input').focus();
            } else {
                form.style.display = 'none';
                document.getElementById('kb-name-input').value = '';
                document.getElementById('kb-desc-input').value = '';
            }
        },

        /**
         * 渲染搜索和筛选工具栏
         */
        renderKbToolbar: function() {
            var html = '';
            html += '<div style="display:flex;gap:8px;margin-bottom:12px;align-items:center;flex-wrap:wrap;">';
            
            // 搜索框
            html += '  <div style="flex:1;min-width:200px;">';
            html += '    <input id="kb-search-input" type="text" placeholder="🔍 搜索知识库名称或描述..." value="' + this.escapeHtml(this.currentState.searchText || '') + '" ';
            html += '           style="width:100%;padding:6px 10px;border:1px solid #ddd;border-radius:4px;font-size:13px;" ';
            html += '           onkeyup="if(event.keyCode===13)App.rag.onSearchKnowledgeBases()" />' ;
            html += '  </div>';
            
            // 状态筛选器
            html += '  <div>';
            html += '    <select id="kb-filter-status" onchange="App.rag.onFilterKnowledgeBases(this.value)" style="padding:6px 10px;border:1px solid #ddd;border-radius:4px;font-size:13px;background:white;">';
            var filters = [
                { value: 'all', text: '全部状态' },
                { value: 'active', text: '✓ 活跃' },
                { value: 'pending', text: '⏳ 待同步' },
                { value: 'archived', text: '📦 已归档' }
            ];
            filters.forEach(function(f) {
                var selected = f.value === (this.currentState.filterStatus || 'all') ? ' selected' : '';
                html += '<option value="' + f.value + '"' + selected + '>' + f.text + '</option>';
            }.bind(this));
            html += '    </select>';
            html += '  </div>';
            
            // 清除筛选按钮
            if (this.currentState.searchText || (this.currentState.filterStatus && this.currentState.filterStatus !== 'all')) {
                html += '  <button onclick="App.rag.clearKbFilters()" style="padding:6px 12px;border:1px solid #ddd;background:white;border-radius:4px;cursor:pointer;font-size:12px;color:#666;">✕ 清除</button>';
            }
            
            html += '</div>';
            
            return html;
        },

        /**
         * 搜索知识库（回车触发）
         */
        onSearchKnowledgeBases: function() {
            var input = document.getElementById('kb-search-input');
            if (input) {
                this.currentState.searchText = input.value.trim();
                this.renderKbList();
                this.logOperation('search', '搜索知识库: "' + this.currentState.searchText + '"');
            }
        },

        /**
         * 筛选知识库状态
         */
        onFilterKnowledgeBases: function(status) {
            this.currentState.filterStatus = status;
            this.renderKbList();
            if (status !== 'all') {
                this.logOperation('filter', '筛选知识库状态: ' + status);
            }
        },

        /**
         * 清除所有筛选条件
         */
        clearKbFilters: function() {
            this.currentState.searchText = '';
            this.currentState.filterStatus = 'all';
            
            var searchInput = document.getElementById('kb-search-input');
            if (searchInput) searchInput.value = '';
            
            var filterSelect = document.getElementById('kb-filter-status');
            if (filterSelect) filterSelect.value = 'all';
            
            this.renderKbList();
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">已清除筛选条件</span>', true) : App.showToast('已清除筛选条件');
        },

        /**
         * 记录操作日志（用于审计）
         */
        logOperation: function(action, detail) {
            var logEntry = {
                timestamp: new Date().toISOString(),
                action: action,
                detail: detail,
                user: 'current_user'  // 可扩展为实际用户名
            };
            
            this.currentState.operationLog.unshift(logEntry);
            
            // 只保留最近100条日志
            if (this.currentState.operationLog.length > 100) {
                this.currentState.operationLog = this.currentState.operationLog.slice(0, 100);
            }
            
            ChatLog.info('RAG', '[RAG Operation]', logEntry.timestamp, '-', action, '-', detail);
        },

        /**
         * 创建新知识库（从 UI 表单读取输入）
         */
        createKnowledgeBase: async function() {
            if (!App.excelBridge) { App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">Excel 桥接未连接</span>', true) : App.showToast('Excel 桥接未连接'); return; }

            var nameInput = document.getElementById('kb-name-input');
            var descInput = document.getElementById('kb-desc-input');

            var name = (nameInput ? nameInput.value.trim() : '');
            if (!name) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">请输入知识库名称</span>', true) : App.showToast('请输入知识库名称');
                if (nameInput) nameInput.focus();
                return;
            }

            var description = descInput ? descInput.value.trim() : '';

            try {
                // ★ 状态已迁移到顶部固定状态栏，不再显示底部Toast
                // App.showToast('正在创建知识库...');

                // 同步更新固定状态栏
                if (window.App && App.fixedStatusBar && App.fixedStatusBar.isInitialized) {
                    App.fixedStatusBar.setProcessing('正在创建知识库...');
                }

                // 诊断日志
                ChatLog.info('RAG', '[RAG] App.excelBridge 类型:', typeof App.excelBridge);
                ChatLog.info('RAG', '[RAG] App.excelBridge 是否有 CreateKnowledgeBaseAsync:',
                    !!(App.excelBridge && App.excelBridge.CreateKnowledgeBaseAsync));
                ChatLog.info('RAG', '[RAG] App.excelBridge keys:', Object.keys(App.excelBridge || {}).slice(0, 20));

                var result = await App.excelBridge.CreateKnowledgeBaseAsync(name.trim(), description);
                ChatLog.info('RAG', '[RAG] CreateKnowledgeBaseAsync 原始返回:', typeof result, JSON.stringify(result));
                var data = this.safeParse(result);
                ChatLog.info('RAG', '[RAG] CreateKnowledgeBaseAsync 解析后:', data);

                if (data.success) {
                    var kbName = data.knowledgeBase ? (data.knowledgeBase.name || name.trim()) : name.trim();
                    var kbStatus = data.knowledgeBase ? (data.knowledgeBase.status || 'unknown') : 'unknown';
                    var msg = '知识库 "' + kbName + '" 创建成功';
                    if (kbStatus === 'pending') {
                        msg += '（Qdrant 未连接，状态: 待同步）';
                    }
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">' + msg + '</span>', true) : App.showToast(msg);
                    this.toggleCreateForm();  // 隐藏表单
                    await this.loadKnowledgeBases();  // 刷新列表
                    this.updateKbSelects();       // 更新下拉框
                } else {
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">创建失败: ' + (data.error || '未知错误') + '</span>', true) : App.showToast('创建失败: ' + (data.error || '未知错误'));
                }
            } catch (ex) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">创建异常: ' + ex.message + '</span>', true) : App.showToast('创建异常: ' + ex.message);
                ChatLog.error('RAG', '[RAG] createKnowledgeBase 异常:', ex);
            }
        },

        /**
         * 删除知识库
         */
        deleteKnowledgeBase: async function(kbId) {
            if (!App.excelBridge) { App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">Excel 桥接未连接</span>', true) : App.showToast('Excel 桥接未连接'); return; }
            if (!await App.showCustomConfirm('确定要删除这个知识库吗？\n\n注意：这将同时删除所有关联的文档和向量数据！')) {
                return;
            }

            try {
                // ★ 状态已迁移到顶部固定状态栏，不再显示底部Toast
                // App.showToast('正在删除...');

                // 同步更新固定状态栏
                if (window.App && App.fixedStatusBar && App.fixedStatusBar.isInitialized) {
                    App.fixedStatusBar.setProcessing('正在删除知识库...');
                }
                var result = await App.excelBridge.DeleteKnowledgeBaseAsync(kbId);
                var data = this.safeParse(result);

                if (data.success) {
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">' + (data.message || '知识库已删除') + '</span>', true) : App.showToast(data.message || '知识库已删除');
                    await this.loadKnowledgeBases();
                    this.updateKbSelects();

                    // 如果当前选中的是被删除的 KB，清空
                    if (this.currentState.currentKbId === kbId) {
                        this.currentState.currentKbId = null;
                        this.currentState.documents = [];
                    }
                } else {
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">删除失败: ' + (data.error || '未知错误') + '</span>', true) : App.showToast('删除失败: ' + (data.error || '未知错误'));
                }
            } catch (ex) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">删除异常: ' + ex.message + '</span>', true) : App.showToast('删除异常: ' + ex.message);
            }
        },

        /**
         * 加载所有知识库
         */
        loadKnowledgeBases: async function() {
            if (!App.excelBridge) return;
            try {
                var result = await App.excelBridge.GetKnowledgeBasesAsync();
                var data = this.safeParse(result);

                if (data.success) {
                    this.currentState.knowledgeBases = data.knowledgeBases || [];
                    this.renderKbList();
                    this.updateKbSelects();
                    this.renderDefaultKbSelect();

                    // 存在"待同步"知识库时，静默尝试同步一次（Qdrant 可能已就绪）
                    var self = this;
                    var hasPending = this.currentState.knowledgeBases.some(function(kb) { return kb.status === 'pending'; });
                    if (hasPending && App.excelBridge && typeof App.excelBridge.SyncKnowledgeBaseStatusAsync === 'function') {
                        try {
                            var syncRes = this.safeParse(await App.excelBridge.SyncKnowledgeBaseStatusAsync());
                            if (syncRes.success && syncRes.synced > 0) {
                                var res2 = await App.excelBridge.GetKnowledgeBasesAsync();
                                var d2 = this.safeParse(res2);
                                if (d2.success) {
                                    this.currentState.knowledgeBases = d2.knowledgeBases || [];
                                    this.renderKbList();
                                    this.updateKbSelects();
                                    this.renderDefaultKbSelect();
                                }
                            }
                        } catch (syncErr) {
                            ChatLog.warn('RAG', '[RAG] 同步待同步知识库失败（不影响列表）:', syncErr);
                        }
                    }
                } else {
                    ChatLog.error('RAG', '[RAG] 加载知识库失败:', data.error);
                }
            } catch (ex) {
                ChatLog.error('RAG', '[RAG] loadKnowledgeBases 异常:', ex);
            }
        },

        /**
         * 渲染知识库列表
         */
        renderKbList: function() {
            var container = document.getElementById('kb-list-container');
            if (!container) return;

            var self = this;
            var kbs = this.currentState.knowledgeBases;

            if (!kbs || kbs.length === 0) {
                var noDataMsg = this.currentState.searchText || (this.currentState.filterStatus && this.currentState.filterStatus !== 'all') 
                    ? '<div style="text-align: center; color: #999; padding: 40px 20px; font-size: 13px;">未找到匹配的知识库<br><br>尝试调整搜索条件或筛选器</div>'
                    : '<div style="text-align: center; color: #999; padding: 40px 20px; font-size: 13px;">暂无知识库<br><br>点击右上角 "+ 新建知识库" 开始使用</div>';
                
                // 搜索/筛选工具栏
                var toolbarHtml = this.renderKbToolbar();
                container.innerHTML = toolbarHtml + noDataMsg;
                return;
            }

            var html = this.renderKbToolbar();
            
            // 搜索结果统计
            var totalCount = this.currentState.knowledgeBases.length;
            if (kbs.length !== totalCount) {
                html += '<div style="padding: 8px 12px; background: #e3f2fd; border-radius: 4px; margin-bottom: 8px; font-size: 12px; color: #1976d2;">';
                html += '  显示 ' + kbs.length + ' / ' + totalCount + ' 个知识库';
                if (this.currentState.searchText) html += ' (搜索: "' + this.escapeHtml(this.currentState.searchText) + '")';
                if (this.currentState.filterStatus && this.currentState.filterStatus !== 'all') html += ' (状态: ' + this.currentState.filterStatus + ')';
                html += '</div>';
            }
            kbs.forEach(function(kb) {
                var statusColor = '#999';
                var statusText = '未知';

                switch (kb.status) {
                    case 'active':
                        statusColor = '#4caf50';
                        statusText = '活跃';
                        break;
                    case 'pending':
                        statusColor = '#ff9800';
                        statusText = '待同步';
                        break;
                    case 'archived':
                        statusColor = '#999';
                        statusText = '已归档';
                        break;
                }

                html += '<div style="display: flex; justify-content: space-between; align-items: center; padding: 12px 16px; margin-bottom: 8px; background: white; border-radius: 6px; border: 1px solid #e0e0e0;">';
                html += '  <div style="flex: 1;">';
                html += '    <div style="font-weight: 500; font-size: 14px; color: #333; margin-bottom: 4px;">' + self.escapeHtml(kb.name) + '</div>';
                html += '    <div style="font-size: 12px; color: #666; margin-bottom: 2px;">' + self.escapeHtml(kb.description || '无描述') + '</div>';
                html += '    <div style="display: flex; gap: 12px; font-size: 11px; color: #888;">';
                html += '      <span>📄 文档: ' + (kb.documentCount || 0) + '</span>';
                html += '      <span>🔷 块: ' + (kb.chunkCount || 0) + '</span>';
                html += '      <span style="color: ' + statusColor + ';">● ' + statusText + '</span>';
                html += '    </div>';
                html += '  </div>';
                html += '  <div style="display: flex; gap: 6px;">';
                html += '    <button onclick="App.rag.viewDocuments(\'' + kb.id + '\')" style="padding: 4px 10px; border: 1px solid #ddd; background: white; border-radius: 4px; cursor: pointer; font-size: 11px;">查看</button>';
                html += '    <button onclick="App.rag.deleteKnowledgeBase(\'' + kb.id + '\')" style="padding: 4px 10px; border: 1px solid #f44336; color: #f44336; background: white; border-radius: 4px; cursor: pointer; font-size: 11px;">删除</button>';
                html += '  </div>';
                html += '</div>';
            });

            container.innerHTML = html;
        },

        /**
         * 查看某个知识库的文档（切换到文档管理 Tab）
         */
        viewDocuments: async function(kbId) {
            this.stopDocAutoRefresh();
            this.currentState.currentKbId = kbId;
            this.switchSubTab('doc-manage');
            await this.loadDocuments(kbId);
        },

        // ========== 文档管理 ==========

        /**
         * 判断输入是否为本地文件路径（而非 URL）
         * 只有带 scheme 的（http/https/ftp/oss 等）按远端处理；其余（盘符路径、UNC、含分隔符、裸文件名）都按本地路径处理，
         * 这样裸文件名会得到"本地文件不存在"的明确错误，而不是走进 OSS 下载逻辑报"文件下载失败"。
         */
        isLocalPath: function(input) {
            if (!input) return false;
            // URL: 带 scheme 的按远端处理
            if (/^[a-z][a-z0-9+.-]*:\/\//i.test(input)) return false;
            return true;
        },

        /**
         * 从路径/URL 提取小写扩展名（去掉 query/fragment），无扩展名返回 ''
         */
        getExtension: function(input) {
            if (!input) return '';
            var pathPart = String(input).split('?')[0].split('#')[0];
            var name = pathPart.split(/[\\\/]/).pop() || '';
            var m = /\.([A-Za-z0-9]+)$/.exec(name);
            return m ? '.' + m[1].toLowerCase() : '';
        },

        // 支持的导入格式（与后端 RagBridge.BrowseLocalFile 过滤器一致）
        SUPPORTED_EXTENSIONS: { '.pdf': 1, '.doc': 1, '.docx': 1, '.xls': 1, '.xlsx': 1, '.csv': 1, '.txt': 1, '.md': 1 },

        /**
         * 格式化文件大小
         */
        formatFileSize: function(bytes) {
            if (!bytes || bytes <= 0) return '';
            if (bytes < 1024) return bytes + ' B';
            if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
            return (bytes / 1024 / 1024).toFixed(2) + ' MB';
        },

        /**
         * 打开本地文件选择对话框
         */
        browseLocalFile: async function() {
            if (!App.excelBridge) { App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">Excel 桥接未连接</span>', true) : App.showToast('Excel 桥接未连接'); return; }
            try {
                var filePath = await App.excelBridge.BrowseLocalFile();
                if (filePath) {
                    document.getElementById('oss-url-input').value = filePath;
                }
            } catch (ex) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">打开文件选择框失败: ' + ex.message + '</span>', true) : App.showToast('打开文件选择框失败: ' + ex.message);
                ChatLog.error('RAG', '[RAG] browseLocalFile 异常:', ex);
            }
        },

        /**
         * 导入文档（支持本地路径 和 OSS URL）
         */
        importDocument: async function() {
            if (!App.excelBridge) { App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">Excel 桥接未连接</span>', true) : App.showToast('Excel 桥接未连接'); return; }
            var input = document.getElementById('oss-url-input').value.trim();
            var kbId = document.getElementById('target-kb-select').value;

            if (!input) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">请输入文件路径或 OSS URL</span>', true) : App.showToast('请输入文件路径或 OSS URL');
                return;
            }

            if (!kbId) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">请选择目标知识库</span>', true) : App.showToast('请选择目标知识库');
                return;
            }

            // 前置校验扩展名，避免走到后端才失败
            var ext = this.getExtension(input);
            if (!this.SUPPORTED_EXTENSIONS[ext]) {
                App.showToast('不支持的文件格式: ' + (ext || '未知') + '（支持 PDF/Word/Excel/CSV/TXT/MD）');
                return;
            }

            var isLocal = this.isLocalPath(input);

            // 防重复提交
            var btn = document.querySelector('button[onclick="App.rag.importDocument()"]');
            var btnText = btn ? btn.textContent : '';
            if (btn) { btn.disabled = true; btn.textContent = '导入中...'; }

            try {
                // ★ 状态已迁移到顶部固定状态栏，不再显示底部Toast
                // App.showToast('正在导入文档...');

                // 同步更新固定状态栏
                if (window.App && App.fixedStatusBar && App.fixedStatusBar.isInitialized) {
                    App.fixedStatusBar.setProcessing('正在导入文档...');
                }
                var result;
                if (isLocal) {
                    // 本地文件路径，调用本地导入方法
                    result = await App.excelBridge.ImportLocalDocumentAsync(input, kbId);
                } else {
                    // OSS URL
                    result = await App.excelBridge.ImportDocumentAsync(input, kbId);
                }
                var data = this.safeParse(result);

                if (data.success) {
                    var docTitle = data.document ? (data.document.title || data.document.id) : input;
                    App.showToast('文档导入成功：' + docTitle);
                    document.getElementById('oss-url-input').value = '';  // 清空输入框

                    // 如果当前查看的是该知识库的文档，刷新列表
                    if (this.currentState.currentKbId === kbId) {
                        await this.loadDocuments(kbId);
                    }

                    // 更新知识库统计
                    await this.loadKnowledgeBases();

                    // 后台正在处理：启动状态轮询，完成后自动停
                    this.startDocAutoRefresh(kbId);
                } else {
                    App.showToast('导入失败: ' + (data.error || '未知错误'));
                }
            } catch (ex) {
                App.showToast('导入异常: ' + ex.message);
                ChatLog.error('RAG', '[RAG] importDocument 异常:', ex);
            } finally {
                if (btn) { btn.disabled = false; btn.textContent = btnText; }
            }
        },

        /**
         * 刷新当前知识库的文档列表（手动刷新按钮）
         */
        refreshDocuments: async function() {
            var kbId = document.getElementById('target-kb-select').value;
            if (!kbId) {
                App.showToast('请先选择知识库');
                return;
            }
            await this.loadDocuments(kbId);
            await this.loadKnowledgeBases();
        },

        /**
         * 启动文档状态自动轮询（导入/重试后调用）
         * 每 2 秒拉一次列表，全部文档脱离 pending/processing 后自动停止（上限 90 秒）
         */
        startDocAutoRefresh: function(kbId) {
            this.stopDocAutoRefresh();
            var self = this;
            var ticks = 0;
            this.currentState.docAutoRefreshTimer = setInterval(async function() {
                ticks++;
                if (ticks > 45) { self.stopDocAutoRefresh(); return; }
                try {
                    await self.loadDocuments(kbId);
                    var docs = self.currentState.documents || [];
                    var processing = docs.some(function(d) { return d.status === 'pending' || d.status === 'processing'; });
                    if (!processing) {
                        await self.loadKnowledgeBases();
                        self.stopDocAutoRefresh();
                    }
                } catch (e) {
                    // 轮询失败不打断用户，下一轮再试
                }
            }, 2000);
        },

        /**
         * 停止文档状态轮询
         */
        stopDocAutoRefresh: function() {
            if (this.currentState.docAutoRefreshTimer) {
                clearInterval(this.currentState.docAutoRefreshTimer);
                this.currentState.docAutoRefreshTimer = null;
            }
        },

        /**
         * 加载指定知识库的文档列表
         */
        loadDocuments: async function(kbId) {
            if (!App.excelBridge) return;
            try {
                var result = await App.excelBridge.GetDocumentsAsync(kbId);
                var data = this.safeParse(result);

                if (data.success) {
                    this.currentState.documents = data.documents || [];
                    this.renderDocList();
                } else {
                    ChatLog.error('RAG', '[RAG] 加载文档失败:', data.error);
                }
            } catch (ex) {
                ChatLog.error('RAG', '[RAG] loadDocuments 异常:', ex);
            }
        },

        /**
         * 渲染文档列表
         */
        renderDocList: function() {
            var container = document.getElementById('doc-list-container');
            if (!container) return;

            var self = this;
            var docs = this.currentState.documents;

            if (!docs || docs.length === 0) {
                container.innerHTML = '<div style="text-align: center; color: #999; padding: 40px 20px; font-size: 13px;">暂无文档<br><br>在上方输入本地文件路径或 OSS URL 导入文档</div>';
                return;
            }

            var html = '';
            docs.forEach(function(doc) {
                var docId = doc.id || '';
                var statusIcon = '';
                var statusText = '';
                var statusColor = '';

                switch (doc.status) {
                    case 'completed':
                        statusIcon = '✅';
                        statusText = '已完成';
                        statusColor = '#4caf50';
                        break;
                    case 'processing':
                        statusIcon = '⏳';
                        statusText = '处理中';
                        statusColor = '#ff9800';
                        break;
                    case 'failed':
                        statusIcon = '❌';
                        statusText = '失败';
                        statusColor = '#f44336';
                        break;
                    default:
                        statusIcon = '⏸️';
                        statusText = '等待中';
                        statusColor = '#999';
                }

                html += '<div style="display: flex; justify-content: space-between; align-items: center; padding: 10px 16px; margin-bottom: 6px; background: white; border-radius: 4px; border: 1px solid #eee;">';
                html += '  <div style="flex: 1; min-width: 0;">';
                html += '    <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 4px;">';
                html += '      <span style="font-weight: 500; font-size: 13px; color: #333; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">' + self.escapeHtml(doc.title || (docId ? docId.substring(0, 8) : '未命名文档')) + '</span>';
                html += '      <span style="font-size: 11px; color: ' + statusColor + '; white-space: nowrap;">' + statusIcon + ' ' + statusText + '</span>';
                html += '    </div>';
                var metaParts = ['块数: ' + (doc.chunkCount || 0), '类型: ' + (doc.fileType || '?')];
                var sizeText = self.formatFileSize(doc.fileSize);
                if (sizeText) metaParts.push('大小: ' + sizeText);
                html += '    <div style="font-size: 11px; color: #888;">' + metaParts.join(' | ') + '</div>';
                if (doc.status === 'failed' && doc.errorMessage) {
                    html += '    <div style="font-size: 11px; color: #f44336; margin-top: 2px;">错误: ' + self.escapeHtml(doc.errorMessage) + '</div>';
                }
                html += '  </div>';
                html += '  <div style="display: flex; gap: 6px; white-space: nowrap;">';
                if (doc.status === 'failed') {
                    html += '    <button onclick="App.rag.retryDocument(\'' + docId + '\')" style="padding: 4px 10px; border: 1px solid #ff9800; color: #e65100; background: white; border-radius: 4px; cursor: pointer; font-size: 11px;">重试</button>';
                }
                html += '    <button onclick="App.rag.deleteDocument(\'' + docId + '\')" style="padding: 4px 10px; border: 1px solid #f44336; color: #f44336; background: white; border-radius: 4px; cursor: pointer; font-size: 11px;">删除</button>';
                html += '  </div>';
                html += '</div>';
            });

            container.innerHTML = html;
        },

        /**
         * 重试失败的文档（后端重置状态并重新后台处理，不产生新记录）
         */
        retryDocument: async function(docId) {
            if (!App.excelBridge) { App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">Excel 桥接未连接</span>', true) : App.showToast('Excel 桥接未连接'); return; }
            if (!docId) return;
            try {
                var result = await App.excelBridge.RetryDocumentAsync(docId);
                var data = this.safeParse(result);
                if (data.success) {
                    App.showToast(data.message || '已重新开始处理');
                    var kbId = this.currentState.currentKbId;
                    if (kbId) {
                        await this.loadDocuments(kbId);
                        this.startDocAutoRefresh(kbId);
                    }
                } else {
                    App.showToast('重试失败: ' + (data.error || '未知错误'));
                }
            } catch (ex) {
                App.showToast('重试异常: ' + ex.message);
                ChatLog.error('RAG', '[RAG] retryDocument 异常:', ex);
            }
        },

        /**
         * 删除文档
         */
        deleteDocument: async function(docId) {
            if (!App.excelBridge) { App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">Excel 桥接未连接</span>', true) : App.showToast('Excel 桥接未连接'); return; }
            if (!await App.showCustomConfirm('确定要删除这个文档吗？')) return;

            try {
                var result = await App.excelBridge.DeleteDocumentAsync(docId);
                var data = this.safeParse(result);

                if (data.success) {
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">文档已删除</span>', true) : App.showToast('文档已删除');

                    // 刷新当前知识库的文档列表
                    if (this.currentState.currentKbId) {
                        await this.loadDocuments(this.currentState.currentKbId);
                    }

                    // 更新知识库统计
                    await this.loadKnowledgeBases();
                    
                    this.logOperation('delete_doc', '删除文档 (ID: ' + docId + ')');
                } else {
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">删除失败: ' + (data.error || '未知错误') + '</span>', true) : App.showToast('删除失败: ' + (data.error || '未知错误'));
                }
            } catch (ex) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">删除异常: ' + ex.message + '</span>', true) : App.showToast('删除异常: ' + ex.message);
            }
        },

        // ========== 检索测试 ==========

        /**
         * 执行 RAG 检索测试
         */
        testQuery: async function() {
            if (!App.excelBridge) { App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">Excel 桥接未连接</span>', true) : App.showToast('Excel 桥接未连接'); return; }
            var question = document.getElementById('rag-test-question').value.trim();
            var kbId = document.getElementById('test-kb-select').value;

            if (!question) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">请输入测试问题</span>', true) : App.showToast('请输入测试问题');
                return;
            }

            if (!kbId) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">请选择知识库</span>', true) : App.showToast('请选择知识库');
                return;
            }

            try {
                // ★ 状态已迁移到顶部固定状态栏，不再显示底部Toast
                // App.showToast('正在检索...');

                // 同步更新固定状态栏
                if (window.App && App.fixedStatusBar && App.fixedStatusBar.isInitialized) {
                    App.fixedStatusBar.setProcessing('正在检索知识库...');
                }
                var result = await App.excelBridge.TestRagQueryAsync(question, kbId);
                var data = this.safeParse(result);

                this.renderTestResults(data);
            } catch (ex) {
                App.showToast('检索异常: ' + ex.message);
                ChatLog.error('RAG', '[RAG] testQuery 异常:', ex);
            }
        },

        /**
         * 渲染测试结果
         */
        renderTestResults: function(data) {
            var container = document.getElementById('rag-test-results');
            if (!container) return;

            var self = this;
            var html = '';

            if (!data.success) {
                html = '<div style="padding: 20px; background: #fff3f3; border: 1px solid #ffcdd2; border-radius: 6px; color: #c00; font-size: 13px;">';
                html += '  <strong>检索失败:</strong> ' + (data.error || '未知错误');
                html += '</div>';
                container.innerHTML = html;
                return;
            }

            // 检索结果摘要
            html += '<div style="margin-bottom: 16px; padding: 12px; background: #e3f2fd; border: 1px solid #90caf9; border-radius: 6px; font-size: 13px;">';
            html += '  <strong>检索完成</strong> | 检索到 <strong>' + (data.totalRetrieved || 0) + '</strong> 个相关片段 | 耗时 ' + ((data.processingTimeMs || 0) / 1000).toFixed(2) + ' 秒';
            html += '</div>';

            // 结果列表
            if (data.results && data.results.length > 0) {
                html += '<div style="margin-bottom: 16px;"><strong>检索结果：</strong></div>';

                data.results.forEach(function(item, index) {
                    var scorePercent = (item.score * 100).toFixed(1);
                    html += '<div style="padding: 12px; margin-bottom: 8px; background: white; border-left: 3px solid #2196f3; border-radius: 0 6px 6px 0; font-size: 13px;">';
                    html += '  <div style="display: flex; justify-content: space-between; margin-bottom: 6px;">';
                    html += '    <span style="font-weight: 500;">#' + (index + 1) + ' 相关度: <span style="color: #1976d2;">' + scorePercent + '%</span></span>';
                    html += '    <span style="font-size: 11px; color: #888;">来源: ' + self.escapeHtml(item.docTitle || '未知') + '</span>';
                    html += '  </div>';
                    html += '  <div style="color: #333; line-height: 1.6; max-height: 150px; overflow-y: auto; font-size: 12px;">' + self.escapeHtml(item.text || '') + '</div>';
                    html += '</div>';
                });
            } else {
                // 0 命中排障指引（耗时 0 秒通常意味着向量库未连接，用户无从判断）
                var elapsedSec = ((data.processingTimeMs || 0) / 1000).toFixed(2);
                html += '<div style="margin-bottom: 16px; padding: 12px; background: #fff8e1; border: 1px solid #ffcc80; border-radius: 6px; font-size: 12px; color: #8d6e00; line-height: 1.9;">';
                html += '  <strong>未检索到相关内容，请按以下顺序排查：</strong><br>';
                if (parseFloat(elapsedSec) < 0.1) {
                    html += '  1. <strong>Qdrant 大概率未连接</strong>（耗时 ' + elapsedSec + ' 秒说明请求未真正发出）——请到「RAG 设置」启动 Qdrant（默认端口 6333）<br>';
                } else {
                    html += '  1. <strong>Qdrant 是否已启动</strong>（RAG 设置页可一键启动，默认端口 6333）<br>';
                }
                html += '  2. <strong>文档是否真的入库</strong>——在文档列表点「重试」重新处理；若 Ollama 未安装向量模型（bge-m3），导入会失败<br>';
                html += '  3. <strong>相似度阈值是否过高</strong>（RAG 设置 → 检索配置，默认 0.7 偏高，可降到 0.3 再试）<br>';
                html += '  4. <strong>换更自然的关键词</strong>——尽量使用文档中出现过的原文词汇';
                html += '</div>';
            }

            // 增强后的上下文预览
            if (data.enhancedAnswer) {
                html += '<div style="margin-top: 16px;"><strong>增强上下文预览：</strong></div>';
                html += '<div style="padding: 12px; background: #f5f5f5; border: 1px solid #ddd; border-radius: 6px; font-size: 12px; line-height: 1.6; max-height: 300px; overflow-y: auto; white-space: pre-wrap;">';
                html += self.escapeHtml(data.enhancedAnswer);
                html += '</div>';
            }

            container.innerHTML = html;
        },

        // ========== 配置管理 ==========

        /**
         * 加载配置并填充表单
         */
        loadConfig: async function() {
            if (!App.excelBridge) return;
            try {
                var result = await App.excelBridge.GetRagConfigAsync();
                var data = this.safeParse(result);

                if (data.success && data.config) {
                    this.currentState.config = data.config;
                    this.renderConfigForm();
                }
            } catch (ex) {
                ChatLog.error('RAG', '[RAG] loadConfig 异常:', ex);
            }
        },

        /**
         * 将配置渲染到表单
         */
        renderConfigForm: function() {
            var config = this.currentState.config;
            if (!config) return;

            var setVal = function(id, val) {
                var el = document.getElementById(id);
                if (el && val !== undefined && val !== null && val !== '') el.value = val;
            };
            var setChecked = function(id, val) {
                var el = document.getElementById(id);
                if (el) el.checked = !!val;
            };

            // Embedding 设置
            setVal('rag-embedding-provider', config.defaultEmbeddingProvider || 'local_bge');
            setVal('rag-local-endpoint', config.localBgeEndpoint || 'http://localhost:11434');
            setVal('rag-local-model', config.localBgeModel || 'bge-m3');
            setVal('rag-cloud-endpoint', config.cloudEndpoint || '');
            setVal('rag-cloud-model', config.cloudModel || 'text-embedding-v3');
            // 密钥字段：有值只显示掩码，不回显明文
            setVal('rag-cloud-apikey', config.cloudApiKey ? '***' : '');
            setVal('rag-qdrant-apikey', config.qdrantApiKey ? '***' : '');

            // Qdrant 设置
            setVal('rag-qdrant-endpoint', config.qdrantEndpoint || 'http://localhost:6333');

            // 对话增强设置
            setChecked('rag-enable-chat', config.enableRagForChat !== false);
            this.renderDefaultKbSelect(config.defaultKnowledgeBaseId || '');
            setVal('rag-topk', config.defaultTopK || 5);
            var threshold = parseFloat(config.scoreThreshold);
            setVal('rag-threshold', isNaN(threshold) ? 0.7 : threshold);
            setVal('rag-max-chunks', config.maxContextChunks || 3);
        },

        /**
         * 渲染"默认知识库"下拉框（RAG 设置 → 对话增强）
         */
        renderDefaultKbSelect: function(selectedId) {
            var select = document.getElementById('rag-default-kb-select');
            if (!select) return;
            if (selectedId === undefined || selectedId === null) {
                selectedId = (this.currentState.config && this.currentState.config.defaultKnowledgeBaseId) || '';
            }

            var currentValue = select.value || selectedId || '';
            select.innerHTML = '';

            var defaultOption = document.createElement('option');
            defaultOption.value = '';
            defaultOption.textContent = '-- 不使用（关闭对话增强）--';
            select.appendChild(defaultOption);

            var self = this;
            (this.currentState.knowledgeBases || []).forEach(function(kb) {
                var option = document.createElement('option');
                option.value = kb.id;
                option.textContent = kb.name + ' (' + (kb.chunkCount || 0) + ' 块)';
                select.appendChild(option);
            });

            // 恢复选中值（配置里的默认 KB 可能已被删除，此时回落空值）
            if (currentValue && select.querySelector('option[value="' + currentValue + '"]')) {
                select.value = currentValue;
            }
        },

        /**
         * 从表单收集配置并保存
         */
        saveSettings: async function() {
            var saveBtn = document.getElementById('rag-save-btn');
            var saveBtnText = saveBtn ? saveBtn.textContent : '保存配置';
            var restoreSaveBtn = function(text, bg) {
                if (!saveBtn) return;
                saveBtn.disabled = false;
                saveBtn.textContent = text || saveBtnText;
                saveBtn.style.background = bg || '#0078d4';
                saveBtn.style.transform = '';
                saveBtn.style.filter = '';
            };
            if (!App.excelBridge) {
                this._flash('Excel 桥接未连接', 'error');
                return;
            }
            if (saveBtn) {
                saveBtn.disabled = true;
                saveBtn.textContent = '保存中…';
                saveBtn.style.transform = '';
                saveBtn.style.filter = '';
            }
            try {
                var getVal = function(id, dft) {
                    var el = document.getElementById(id);
                    return el ? el.value : (dft || '');
                };
                var getChecked = function(id, dft) {
                    var el = document.getElementById(id);
                    return el ? el.checked : dft;
                };

                var topK = parseInt(getVal('rag-topk', '5'), 10);
                if (isNaN(topK) || topK <= 0) topK = 5;
                var threshold = parseFloat(getVal('rag-threshold', '0.7'));
                if (isNaN(threshold) || threshold < 0) threshold = 0.7;
                var maxChunks = parseInt(getVal('rag-max-chunks', '3'), 10);
                if (isNaN(maxChunks) || maxChunks <= 0) maxChunks = 3;

                var config = {
                    defaultEmbeddingProvider: getVal('rag-embedding-provider', 'local_bge'),
                    localBgeEndpoint: getVal('rag-local-endpoint', 'http://localhost:11434'),
                    localBgeModel: getVal('rag-local-model', 'bge-m3'),
                    cloudEndpoint: getVal('rag-cloud-endpoint', ''),
                    cloudApiKey: getVal('rag-cloud-apikey', ''),
                    cloudModel: getVal('rag-cloud-model', 'text-embedding-v3'),
                    qdrantEndpoint: getVal('rag-qdrant-endpoint', 'http://localhost:6333'),
                    qdrantApiKey: getVal('rag-qdrant-apikey', ''),
                    defaultTopK: topK,
                    scoreThreshold: threshold,
                    enableRagForChat: getChecked('rag-enable-chat', true),
                    maxContextChunks: maxChunks,
                    defaultKnowledgeBaseId: getVal('rag-default-kb-select', '')
                };

                // 密钥掩码语义：原配置有密钥且输入框为空/仍是掩码 → 传 "***" 让后端保留原值
                var prev = this.currentState.config || {};
                if (prev.cloudApiKey && (!config.cloudApiKey || config.cloudApiKey === '***')) {
                    config.cloudApiKey = '***';
                }
                if (prev.qdrantApiKey && (!config.qdrantApiKey || config.qdrantApiKey === '***')) {
                    config.qdrantApiKey = '***';
                }

                // ★ 状态已迁移到顶部固定状态栏，不再显示底部Toast
                // App.showToast('正在保存配置...');
                var result = await App.excelBridge.SaveRagConfigAsync(JSON.stringify(config));
                var data = this.safeParse(result);

                if (data.success) {
                    this.currentState.config = config;
                    if (saveBtn) {
                        saveBtn.textContent = '已保存';
                        saveBtn.style.background = '#2e7d32';
                    }
                    this._flash('配置已保存', 'success');
                    setTimeout(function() { restoreSaveBtn(); }, 1200);
                } else {
                    var failMsg = '保存失败: ' + (data.error || '未知错误');
                    if (saveBtn) {
                        saveBtn.textContent = '保存失败';
                        saveBtn.style.background = '#c62828';
                    }
                    this._flash(failMsg, 'error');
                    setTimeout(function() { restoreSaveBtn(); }, 1600);
                }
            } catch (ex) {
                if (saveBtn) {
                    saveBtn.textContent = '保存失败';
                    saveBtn.style.background = '#c62828';
                }
                this._flash('保存异常: ' + ex.message, 'error');
                setTimeout(function() { restoreSaveBtn(); }, 1600);
                ChatLog.error('RAG', '[RAG] saveSettings 异常:', ex);
            }
        },

        /** 可见反馈。状态条的 setExtraStatus 已被清空，不能再拿它当提示。 */
        _flash: function(message, type) {
            try {
                if (App.utils && typeof App.utils.showToast === 'function') {
                    App.utils.showToast(message, type || 'info');
                } else {
                    var toast = document.getElementById('toast');
                    if (toast) {
                        toast.textContent = message;
                        toast.style.background = type === 'error' ? '#ef4444' : '#10b981';
                        toast.classList.add('show');
                        setTimeout(function() { toast.classList.remove('show'); }, 2500);
                    }
                }
            } catch (e) { /* 提示失败不影响保存 */ }
            var fsb = App.fixedStatusBar;
            if (!fsb) return;
            if (type === 'error' && typeof fsb.setError === 'function') fsb.setError(message);
            else if (typeof fsb.setSuccess === 'function') fsb.setSuccess(message);
        },

        /**
         * RAG 对话增强（供 text-pipeline 调用）。
         * 返回可注入 LLM 的上下文字符串；未启用/未配默认库/未命中/异常一律返回 null（聊天绝不被 RAG 拖垮）。
         * 带 5 秒超时保护：Embedding 服务挂起时聊天照常进行。
         */
        enhanceForChat: function(userMessage) {
            var self = this;
            return new Promise(function(resolve) {
                try {
                    if (!userMessage || !window.App || !App.excelBridge ||
                        typeof App.excelBridge.EnhanceWithRagAsync !== 'function') {
                        resolve(null);
                        return;
                    }
                    var cfg = self.currentState.config || {};
                    if (cfg.enableRagForChat === false) { resolve(null); return; }
                    var kbId = cfg.defaultKnowledgeBaseId || '';
                    if (!kbId) { resolve(null); return; }

                    var timeoutGuard = new Promise(function(res) {
                        setTimeout(function() { res(null); }, 5000);
                    });
                    var call = Promise.resolve()
                        .then(function() { return App.excelBridge.EnhanceWithRagAsync(userMessage, kbId); })
                        .then(function(raw) {
                            var data = self.safeParse(raw);
                            if (data && data.success && data.hasContext && data.contextText) {
                                ChatLog.info('RAG', '[RAG] 对话增强命中: ' + (data.retrievedChunks || 0) + ' 块 (KB=' + kbId + ')');
                                return data.contextText;
                            }
                            return null;
                        })
                        .catch(function(err) {
                            ChatLog.warn('RAG', '[RAG] 对话增强失败（不影响聊天）:', err && err.message);
                            return null;
                        });

                    Promise.race([call, timeoutGuard]).then(resolve);
                } catch (e) {
                    resolve(null);
                }
            });
        },

        /**
         * 测试 Qdrant 和 Embedding 连接
         */
        testConnections: async function() {
            if (!App.excelBridge) { App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">Excel 桥接未连接</span>', true) : App.showToast('Excel 桥接未连接'); return; }
            try {
                // 先检查 Qdrant 安装状态
                var installStatusResult = await App.excelBridge.GetQdrantInstallStatus();
                var installStatus = this.safeParse(installStatusResult);

                // 如果 Qdrant 未安装，显示安装向导
                if (!installStatus.success || !installStatus.installed) {
                    this.showQdrantInstallWizard();
                    return;
                }

                // 如果已安装但未运行，尝试自动启动
                if (installStatus.installed && !installStatus.running) {
                    var startResult = await App.excelBridge.StartQdrantAsync();
                    var startData = this.safeParse(startResult);
                    if (!startData.success) {
                        // 启动失败，显示安装向导让用户手动处理
                        this.showQdrantInstallWizard('Qdrant 启动失败，请尝试重新安装或手动启动');
                        return;
                    }
                }

                // 测试连接
                var [qdrantResult, embeddingResult] = await Promise.all([
                    App.excelBridge.TestQdrantConnectionAsync(),
                    App.excelBridge.TestEmbeddingConnectionAsync('local_bge')
                ]);

                var qdrantData = this.safeParse(qdrantResult);
                var embeddingData = this.safeParse(embeddingResult);

                var messages = [];

                if (qdrantData.success) {
                    messages.push('✅ Qdrant: ' + (qdrantData.message || qdrantData.Message || '连接成功') + ' (' + (qdrantData.endpoint || qdrantData.Endpoint || '未知') + ')');
                } else {
                    messages.push('❌ Qdrant: ' + (qdrantData.message || qdrantData.Message || qdrantData.error || qdrantData.Error || '连接失败'));
                }

                if (embeddingData.success) {
                    messages.push('✅ Embedding: ' + (embeddingData.message || embeddingData.Message || '连接成功') + ' (' + (embeddingData.latencyMs || embeddingData.LatencyMs || 0) + 'ms)');
                } else {
                    messages.push('❌ Embedding: ' + (embeddingData.message || embeddingData.Message || embeddingData.error || embeddingData.Error || '连接失败'));
                }

                // 显示测试结果
                var resultHtml = messages.map(function(msg) {
                    var isError = msg.startsWith('❌');
                    return '<div style="padding:8px;margin:4px 0;border-radius:4px;background:' + (isError ? '#ffebee' : '#e8f5e9') + ';color:' + (isError ? '#c62828' : '#2e7d32') + ';">' + msg + '</div>';
                }).join('');

                var overlay = document.createElement('div');
                overlay.id = 'test-result-overlay';
                overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:100000;display:flex;align-items:center;justify-content:center;';
                overlay.innerHTML = '<div id="test-result-dialog" style="background:white;border-radius:8px;padding:24px;max-width:500px;width:90%;box-shadow:0 4px 20px rgba(0,0,0,0.3);"><div style="font-size:16px;font-weight:500;margin-bottom:16px;color:#333;">连接测试结果</div>' + resultHtml + '<div style="display:flex;justify-content:flex-end;margin-top:16px;"><button type="button" id="test-result-close" style="padding:8px 16px;border:none;background:#1976d2;color:white;border-radius:4px;cursor:pointer;font-size:14px;">关闭</button></div></div>';
                document.body.appendChild(overlay);

                var closeOverlay = function() {
                    if (overlay && overlay.parentNode) overlay.parentNode.removeChild(overlay);
                    document.removeEventListener('keydown', escHandler);
                };
                var escHandler = function(e) {
                    if (e.key === 'Escape') closeOverlay();
                };
                document.getElementById('test-result-close').addEventListener('click', closeOverlay);
                var _downX = 0, _downY = 0;
                overlay.addEventListener('pointerdown', function(e) {
                    _downX = e.clientX; _downY = e.clientY;
                });
                overlay.addEventListener('click', function(e) {
                    if (e.target !== overlay) return;
                    if (window.getSelection && window.getSelection().toString().trim()) return;
                    var moved = Math.hypot(e.clientX - _downX, e.clientY - _downY);
                    if (moved > 8) return;
                    closeOverlay();
                });
                document.addEventListener('keydown', escHandler);
                document.getElementById('test-result-close').focus();

            } catch (ex) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">连接测试异常: ' + ex.message + '</span>', true) : App.showToast('连接测试异常: ' + ex.message);
            }
        },

        /**
         * 显示 Qdrant 安装向导
         */
        showQdrantInstallWizard: function(hintMessage) {
            var self = this;
            var overlay = document.createElement('div');
            overlay.id = 'qdrant-install-overlay';
            overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:100000;display:flex;align-items:center;justify-content:center;';
            
            var hintHtml = hintMessage ? '<div style="padding:10px;margin-bottom:16px;background:#fff3cd;border-radius:4px;color:#856404;font-size:13px;">' + hintMessage + '</div>' : '';
            
            overlay.innerHTML = '<div style="background:white;border-radius:8px;padding:24px;max-width:500px;width:90%;box-shadow:0 4px 20px rgba(0,0,0,0.3);">' +
                '<div style="font-size:18px;font-weight:500;margin-bottom:8px;color:#333;">📦 安装 Qdrant 向量数据库</div>' +
                '<div style="font-size:13px;color:#666;margin-bottom:16px;">RAG 功能需要 Qdrant 向量数据库支持。选择以下方式安装：</div>' +
                hintHtml +
                '<div style="margin-bottom:16px;padding:12px;background:#f5f5f5;border-radius:4px;">' +
                    '<div style="font-size:13px;color:#333;margin-bottom:8px;"><strong>方式一：一键安装</strong>（自动下载并安装）</div>' +
                    '<div style="font-size:12px;color:#666;">适合网络畅通的用户，自动从 GitHub 下载</div>' +
                '</div>' +
                '<div style="margin-bottom:16px;padding:12px;background:#f5f5f5;border-radius:4px;">' +
                    '<div style="font-size:13px;color:#333;margin-bottom:8px;"><strong>方式二：手动安装</strong>（指定本地安装包）</div>' +
                    '<div style="font-size:12px;color:#666;margin-bottom:8px;">1. 下载 Qdrant（如浏览器下载慢，可用迅雷/IDM）：</div>' +
                    '<div style="font-size:12px;margin-bottom:4px;"><a href="https://ghfast.top/https://github.com/qdrant/qdrant/releases/download/v1.13.2/qdrant-x86_64-pc-windows-msvc.zip" target="_blank" style="color:#1976d2;">国内镜像下载</a> | <a href="https://github.com/qdrant/qdrant/releases/download/v1.13.2/qdrant-x86_64-pc-windows-msvc.zip" target="_blank" style="color:#1976d2;">GitHub 原站</a></div>' +
                    '<div style="font-size:12px;color:#666;">2. 点击下方按钮选择下载的 ZIP 文件</div>' +
                '</div>' +
                '<div id="qdrant-install-progress" style="display:none;margin-bottom:16px;">' +
                    '<div style="background:#f5f5f5;border-radius:4px;height:20px;overflow:hidden;">' +
                        '<div id="qdrant-progress-bar" style="background:#1976d2;height:100%;width:0%;transition:width 0.3s;"></div>' +
                    '</div>' +
                    '<div id="qdrant-progress-text" style="font-size:12px;color:#666;margin-top:8px;">准备中</div>' +
                '</div>' +
                '<div style="display:flex;gap:10px;justify-content:flex-end;">' +
                    '<button id="qdrant-install-cancel" style="padding:8px 16px;border:1px solid #ddd;background:white;border-radius:4px;cursor:pointer;font-size:14px;color:#666;">取消</button>' +
                    '<button id="qdrant-install-manual" style="padding:8px 16px;border:1px solid #1976d2;background:white;color:#1976d2;border-radius:4px;cursor:pointer;font-size:14px;">📁 选择安装包</button>' +
                    '<button id="qdrant-install-btn" style="padding:8px 16px;border:none;background:#1976d2;color:white;border-radius:4px;cursor:pointer;font-size:14px;">🚀 一键安装</button>' +
                '</div>' +
            '</div>';
            
            document.body.appendChild(overlay);
            
            document.getElementById('qdrant-install-cancel').onclick = function() {
                if (overlay.parentNode) overlay.parentNode.removeChild(overlay);
            };
            
            // 手动安装按钮
            document.getElementById('qdrant-install-manual').onclick = function() {
                // 检测是否为 VSTO 模式
                var isVstoMode = App.excelBridge && typeof App.excelBridge.ShowOpenFileDialog === 'function';
                
                // 显示路径输入框
                var manualDiv = document.createElement('div');
                manualDiv.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:100001;display:flex;align-items:center;justify-content:center;';
                
                // 根据模式生成不同的 HTML
                var inputAreaHtml = '';
                if (isVstoMode) {
                    // VSTO 模式：显示浏览按钮
                    inputAreaHtml = '<div style="display:flex;gap:8px;margin-bottom:16px;">' +
                        '<input type="text" id="qdrant-file-path" placeholder="文件路径..." style="flex:1;padding:10px;border:1px solid #ddd;border-radius:4px;font-size:14px;box-sizing:border-box;font-family:monospace;" />' +
                        '<button id="qdrant-browse-btn" style="padding:10px 16px;border:1px solid #1976d2;background:white;color:#1976d2;border-radius:4px;cursor:pointer;font-size:14px;white-space:nowrap;">浏览...</button>' +
                    '</div>';
                } else {
                    // Web 模式：只显示输入框 + 警告提示
                    inputAreaHtml = '<div style="margin-bottom:16px;">' +
                        '<input type="text" id="qdrant-file-path" placeholder="请粘贴文件完整路径..." style="width:100%;padding:10px;border:1px solid #ddd;border-radius:4px;font-size:14px;box-sizing:border-box;font-family:monospace;" />' +
                    '</div>' +
                    '<div style="font-size:12px;color:#ff9800;margin-bottom:16px;padding:8px;background:#fff3e0;border-radius:4px;border-left:3px solid #ff9800;">' +
                        '⚠️ Web 模式不支持文件浏览，请手动输入路径<br/>' +
                        '提示：在文件资源管理器中按住 Shift 右键文件 → "复制为路径" → 粘贴到上方' +
                    '</div>';
                }
                
                manualDiv.innerHTML = '<div style="background:white;border-radius:8px;padding:24px;max-width:600px;width:90%;box-shadow:0 4px 20px rgba(0,0,0,0.3);">' +
                    '<div style="font-size:16px;font-weight:500;margin-bottom:16px;color:#333;"> 指定 Qdrant 安装包路径</div>' +
                    '<div style="font-size:13px;color:#666;margin-bottom:16px;">请选择或输入 ZIP 文件的完整路径：</div>' +
                    inputAreaHtml +
                    '<div style="font-size:12px;color:#999;margin-bottom:16px;padding:8px;background:#f5f5f5;border-radius:4px;">' +
                        '示例：C:\\Users\\Administrator\\Downloads\\qdrant-x86_64-pc-windows-msvc.zip' +
                    '</div>' +
                    '<div style="display:flex;gap:10px;justify-content:flex-end;">' +
                        '<button id="qdrant-manual-cancel" style="padding:8px 16px;border:1px solid #ddd;background:white;border-radius:4px;cursor:pointer;font-size:14px;color:#666;">取消</button>' +
                        '<button id="qdrant-manual-confirm" style="padding:8px 16px;border:none;background:#1976d2;color:white;border-radius:4px;cursor:pointer;font-size:14px;">开始安装</button>' +
                    '</div>' +
                '</div>';
                document.body.appendChild(manualDiv);
                
                document.getElementById('qdrant-manual-cancel').onclick = function() {
                    if (manualDiv.parentNode) manualDiv.parentNode.removeChild(manualDiv);
                };
                
                // 浏览按钮 - 仅在 VSTO 模式下绑定
                if (isVstoMode) {
                    document.getElementById('qdrant-browse-btn').onclick = async function() {
                        try {
                            var result = await App.excelBridge.ShowOpenFileDialog();
                            var data = self.safeParse(result);
                            if (data.success && data.filePath) {
                                document.getElementById('qdrant-file-path').value = data.filePath;
                            }
                        } catch (ex) {
                            ChatLog.error('RAG', '[Qdrant] 打开文件对话框失败:', ex);
                        }
                    };
                }
                
                document.getElementById('qdrant-manual-confirm').onclick = async function() {
                    var filePath = document.getElementById('qdrant-file-path').value.trim();
                    if (!filePath) {
                        // WebView2 中 alert 不工作，使用自定义提示
                        var alertOverlay = document.createElement('div');
                        alertOverlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:100002;display:flex;align-items:center;justify-content:center;';
                        alertOverlay.innerHTML = '<div style="background:white;border-radius:8px;padding:24px;max-width:400px;width:90%;"><div style="font-size:16px;font-weight:500;margin-bottom:16px;">提示</div><div style="font-size:14px;color:#666;margin-bottom:20px;">请先输入或选择文件路径</div><div style="text-align:right;"><button onclick="this.closest(\'div[style*=fixed]\').remove()" style="padding:8px 16px;background:#1976d2;color:white;border:none;border-radius:4px;cursor:pointer;">确定</button></div></div>';
                        document.body.appendChild(alertOverlay);
                        return;
                    }
                    
                    // 移除路径输入框
                    if (manualDiv.parentNode) manualDiv.parentNode.removeChild(manualDiv);
                    
                    var btn = document.getElementById('qdrant-install-manual');
                    var progressDiv = document.getElementById('qdrant-install-progress');
                    var progressBar = document.getElementById('qdrant-progress-bar');
                    var progressText = document.getElementById('qdrant-progress-text');
                    
                    btn.disabled = true;
                    btn.textContent = '安装中';
                    progressDiv.style.display = 'block';
                    progressText.textContent = '正在解压...';
                    
                    try {
                        var result = await App.excelBridge.InstallQdrantFromLocalFileAsync(filePath);
                        var data = self.safeParse(result);
                        
                        if (data.success) {
                            progressBar.style.width = '100%';
                            progressBar.style.background = '#4caf50';
                            progressText.textContent = data.message || '安装完成';
                            progressText.style.color = '#4caf50';
                            
                            btn.textContent = '✅ 安装成功';
                            btn.style.background = '#4caf50';
                            btn.style.color = 'white';
                            
                            setTimeout(function() {
                                if (overlay.parentNode) overlay.parentNode.removeChild(overlay);
                                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">Qdrant 安装完成，请重新测试连接</span>', true) : App.showToast('Qdrant 安装完成，请重新测试连接');
                            }, 2000);
                        } else {
                            progressBar.style.background = '#f44336';
                            progressText.textContent = data.message || '安装失败';
                            progressText.style.color = '#f44336';
                            
                            btn.disabled = false;
                            btn.textContent = ' 选择安装包';
                        }
                    } catch (ex) {
                        progressText.textContent = '安装异常: ' + ex.message;
                        progressText.style.color = '#f44336';
                        
                        btn.disabled = false;
                        btn.textContent = '📁 选择安装包';
                    }
                };
            };
            
            // 一键安装按钮
            document.getElementById('qdrant-install-btn').onclick = async function() {
                var btn = this;
                var progressDiv = document.getElementById('qdrant-install-progress');
                var progressBar = document.getElementById('qdrant-progress-bar');
                var progressText = document.getElementById('qdrant-progress-text');
                
                btn.disabled = true;
                btn.textContent = '安装中';
                progressDiv.style.display = 'block';
                
                try {
                    var result = await App.excelBridge.InstallQdrantAsync();
                    var data = self.safeParse(result);
                    
                    if (data.success) {
                        progressBar.style.width = '100%';
                        progressBar.style.background = '#4caf50';
                        progressText.textContent = data.message || '安装完成';
                        progressText.style.color = '#4caf50';
                        
                        btn.textContent = '✅ 安装成功';
                        btn.style.background = '#4caf50';
                        
                        // 2秒后关闭并刷新页面
                        setTimeout(function() {
                            if (overlay.parentNode) overlay.parentNode.removeChild(overlay);
                            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">Qdrant 安装完成，请重新测试连接</span>', true) : App.showToast('Qdrant 安装完成，请重新测试连接');
                        }, 2000);
                    } else {
                        progressBar.style.background = '#f44336';
                        progressText.textContent = data.message || '安装失败';
                        progressText.style.color = '#f44336';
                        
                        btn.disabled = false;
                        btn.textContent = '🔄 重试安装';
                        btn.style.background = '#f44336';
                    }
                } catch (ex) {
                    progressText.textContent = '安装异常: ' + ex.message;
                    progressText.style.color = '#f44336';
                    
                    btn.disabled = false;
                    btn.textContent = '🔄 重试安装';
                }
            };
        },

        // ========== 工具方法 ==========

        /**
         * 更新所有知识库下拉框
         */
        updateKbSelects: function() {
            var self = this;
            var selects = [
                document.getElementById('target-kb-select'),
                document.getElementById('test-kb-select')
            ];

            selects.forEach(function(select) {
                if (!select) return;

                var currentValue = select.value;
                select.innerHTML = '';

                // 添加默认选项
                var defaultOption = document.createElement('option');
                defaultOption.value = '';
                defaultOption.textContent = '-- 请选择知识库 --';
                select.appendChild(defaultOption);

                self.currentState.knowledgeBases.forEach(function(kb) {
                    var option = document.createElement('option');
                    option.value = kb.id;
                    option.textContent = kb.name + ' (' + (kb.chunkCount || 0) + ' 块)';
                    select.appendChild(option);
                });

                // 尝试恢复选中值
                if (currentValue && select.querySelector('option[value="' + currentValue + '"]')) {
                    select.value = currentValue;
                }
            });
        },

        /**
         * 安全解析 COM 桥接返回的 JSON
         */
        safeParse: function(result) {
            if (!result) return { success: false, error: '空返回值' };
            if (typeof result === 'object') return result;  // 已经是对象（某些桥接模式）
            try {
                return JSON.parse(result);
            } catch (e) {
                ChatLog.error('RAG', '[RAG] JSON 解析失败, 原始值:', typeof result, result);
                return { success: false, error: 'JSON 解析失败: ' + e.message };
            }
        },

        /**
         * HTML 转义
         */
        escapeHtml: function(text) {
            if (!text) return '';
            var div = document.createElement('div');
            div.textContent = text;
            return div.innerHTML;
        },

        /**
         * 绑定事件
         */
        bindEvents: function() {
            // Tab 切换事件监听（如果主框架支持的话）
            // 这里主要依赖 HTML 中的 inline onclick
        }
    };

    // 不在此处自动调用 init()
    // 由 app-model.js 在 Excel Bridge 初始化完成后调用 App.rag.init()
    // 这样确保 App.excelBridge 已就绪

})();