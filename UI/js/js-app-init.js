/**
 * ============================================================
 * js-app-init.js — JS Application Initializer
 * ============================================================
 * 功能：JS 应用初始化模块，负责将各 IIFE 模块函数暴露到 window、加载脚本、注册全局回调。
 *
 * 核心类/函数：
 *   - exposeFunctions() — 暴露函数到全局
 *   - initializeApp() — 初始化应用
 *   - functionsToExpose — 全局暴露函数清单
 *
 * 依赖：
 *   - app-bootstrap.js
 *   - app.js
 * ============================================================
 */

(function() {
    'use strict';

    // ==================== 需要暴露的函数清单 ====================
    var functionsToExpose = [
        'sendMessage', 'toggleEditor', 'showCodeEditor', 'executeCode',
        'saveCodeSnippet', 'confirmSaveSnippet', 'closeSnippetDialog',
        'showCodeLibrary', 'closeCodeLibrary', 'clearEditor', 'undoLastAction', 'redoLastAction',
        'clearAllCache', 'clearAllHistory', 'deleteHistoryItem',
        'initTabs', 'onProviderChange', 'saveSettings', 'saveOcrSettings', 'updateContextStats',
        'testAIConfig', 'testCloudOcrConfig',
        'detectLocalModels', 'selectModel', 'refreshModelsList', 'onChatModelChange',
        'installOcrEngine', 'installPaddleOcr', 'runPaddleOCRInstaller',
        'showAddSkillDialog', 'closeSkillDialog', 'switchSkillTab', 'toggleAdvancedSection',
        'formatSkillJson', 'validateSkillJson', 'editSkill', 'saveSkill', 'loadCustomSkills',
        'switchSettingsTab', 'switchExtensionTab', 'clearAttachment', 'collapseEditor',
        'startOrStepInto', 'stepInto', 'stepOver', 'continueExecution',
        'stopDebug', 'refreshWatchWindow',
        'deleteSnippetFromLibrary', 'loadSnippetFromLibrary', 'editSnippetFromLibrary',
        'executeCacheClear', 'closeCacheClearDialog',
        'copyCode', 'toEditor', 'previewHtml',
        'insertTableToExcel', 'copyTableToClipboard', 'downloadTableAsCsv', 'insertLastTableToExcel',
        'copyLastTableToClipboard', 'editLastTableData',
        'copyOcrResultToClipboard', 'splitOcrResultByLine', 'extractOcrKeywords',
        'copyLinesToClipboard', 'insertLinesToExcel', 'copyKeywordsToClipboard', 'insertKeywordsToExcel',
        'fillRecognitionDataToExcel', 'copyRecognitionResult', 'downloadRecognitionResult',
        'processImageGeneric', 'parseRecognitionData',
        'sendQuickPrompt', 'triggerQuickSkill', 'bindQuickSkill', 'clearQuickSkillBinding', 'renderQuickSkillChip',
        'refreshMcpStatus', 'startMcpServer', 'stopMcpServer',
        'showAddMcpClientDialog', 'editMcpClient', 'closeMcpClientDialog', 'onMcpTransportChange',
        'saveMcpClient', 'testMcpClient', 'testMcpClientFromDialog',
        'removeMcpClient', 'toggleMcpClient', 'autoReconnectMcpClients', 'onOcrEngineChange',
        'showToast', 'addMessage', 'renderMessageContent',
        'handlePaste', 'autoResizeTextarea', 'newChat', 'loadHistory', 'initBridge',
        'showWelcomePanel', 'addWelcomeMessage',
        'saveUIState', 'restoreUIState',
        'switchEditorLanguage', 'executePythonCode', 'executeJSCode',
        'openDebugWindow'
    ];

    // ==================== API 暴露函数（支持多次调用，处理 defer 模块延迟加载）====================
    function exposeApi(isReload) {
        // 来源列表：按优先级排序（每次调用都重新构建，以捕获 defer 加载后新增的命名空间）
        var sources = [
            App,                          // App 顶层（直接挂载的函数）
            window.AppHistory,            // app-history.js 导出
            window.App && App.Editor,     // app-editor.js 导出
            window.App && App.McpModule,  // app-mcp.js 导出
            window.App && App.utils,      // app-utils.js 导出
            window.App && App.ImageTable, // app-image-table.js 导出
            window.App && App.recognition // app-recognition.js 导出
        ];

        function findFunction(name) {
            // 先检查 window 直接挂载
            if (typeof window[name] === 'function') return window[name];
            // 再检查 App 顶层
            if (App && typeof App[name] === 'function') return App[name];
            // 最后遍历所有子命名空间
            for (var i = 0; i < sources.length; i++) {
                var src = sources[i];
                if (src && typeof src[name] === 'function') return src[name];
            }
            return undefined;
        }

        // 扁平化已移除：子命名空间函数不再被提升到 App 顶层
        // 请通过 App.Editor.xxx、App.McpModule.xxx 等命名空间访问
        // 或使用 functionsToExpose 清单中暴露到 window 的全局函数
        if (!isReload) {
            ChatLog.info('INIT', '🔧 子命名空间保持独立，不再扁平化到 App 顶层');
        }

        // stopDebug 别名
        if (!App.stopDebug && typeof App.stopExecution === 'function') {
            App.stopDebug = App.stopExecution;
        }

        // 统一暴露到 window
        var exposedCount = 0;
        var failedList = [];

        functionsToExpose.forEach(function(fnName) {
            var fn = findFunction(fnName);
            if (typeof fn === 'function') {
                window[fnName] = fn;
                exposedCount++;
            } else {
                failedList.push(fnName);
            }
        });

        // 暴露关键变量
        if (App.config) window.config = App.config;
        if (App.state) window.chatHistory = App.state.chatHistory || [];
        if (App.excelBridge) window.excelBridge = App.excelBridge;

        // 更新延迟全局变量
        if (typeof App.callAIUnified === 'function') { window.callAIUnified = App.callAIUnified; }
        if (typeof App.exactMatchSkill === 'function') { window.exactMatchSkill = App.exactMatchSkill; }
        if (typeof App.matchSkillByDescription === 'function') { window.matchSkillByDescription = App.matchSkillByDescription; }
        if (typeof App.showToast === 'function') { window.showToast = App.showToast; }

        // 输出结果
        if (isReload) {
            ChatLog.info('INIT', '📤 API 二次暴露（defer 模块已加载）: 成功', exposedCount, '/', functionsToExpose.length, '个函数');
        } else {
            if (failedList.length > 0) {
                ChatLog.info('INIT', '⏳ 部分函数将在 defer 模块加载后二次暴露，当前缺失:', failedList.length, '个');
            }
            ChatLog.info('INIT', '='.repeat(60));
            ChatLog.info('INIT', '✅ SmartPilot 智航 启动完成');
            ChatLog.info('INIT', '📊 已加载模块: bootstrap/utils/ocr/history/mcp/config/model/editor/skill/recognition/image-table/file-batch/chat-ui/ai-core');
            ChatLog.info('INIT', '📤 API 暴露: 成功', exposedCount, '/', functionsToExpose.length, '个函数');
            ChatLog.info('INIT', '='.repeat(60));
        }

        return { exposedCount, failedList };
    }

    // 立即暴露一次（同步模块已可用）
    exposeApi(false);

    // 所有资源（含 defer 脚本）加载完成后，再次暴露延迟模块的函数
    if (typeof window !== 'undefined') {
        window.addEventListener('load', function() {
            ChatLog.info('INIT', '[Init] 页面资源全部加载完成，执行 API 二次暴露...');
            var result = exposeApi(true);
            if (result.failedList.length > 0) {
                ChatLog.warn('INIT', '⚠️ 以下函数在 defer 加载后仍未找到 (' + result.failedList.length + '个):');
                result.failedList.forEach(function(f) { ChatLog.warn('INIT', '  -', f); });
            }
        });
    }

    // ==================== DOMContentLoaded 初始化（来自原app.js:L1687-L2031）====================
    // 这是拆分后最关键的初始化入口：绑定事件、调用各模块初始化函数
    // 注意：所有函数通过 App.* 或 window.* 调用（已由上方扁平化逻辑保证可用）
    document.addEventListener('DOMContentLoaded', async function() {
        ChatLog.info('INIT', '[Init] DOMContentLoaded 触发，开始应用初始化...');

        // --- 启动进度条初始化 ---
        var startupSteps = [
            { id: 'load-config', name: '加载配置...' },
            { id: 'load-recognition-types', name: '加载识别类型...' },
            { id: 'load-user-settings', name: '加载用户设置...' },
            { id: 'load-history', name: '加载历史记录...' },
            { id: 'init-ui', name: '初始化界面组件...' },
            { id: 'ready', name: (window.App && App.i18n && App.i18n.t) ? App.i18n.t('status.ready') : '系统就绪' }
        ];
        if (window.startupProgress) {
            window.startupProgress.init(startupSteps);
        }

        // --- 固定状态栏初始化（已移至 defer 组，做安全保护）---
        try {
            if (window.App && App.fixedStatusBar && typeof App.fixedStatusBar.init === 'function') {
                App.fixedStatusBar.init();
                App.fixedStatusBar.setLoading((window.App && App.i18n && App.i18n.t) ? App.i18n.t('status.detecting') : '系统准备中', 0);
            }
        } catch(e) { ChatLog.warn('INIT', '[Init] 固定状态栏初始化失败:', e); }

        // ============================================================
        // 关键路径：串行执行，完成后立即显示欢迎面板
        // ============================================================

        // --- 阶段1：加载配置 ---
        if (window.startupProgress) window.startupProgress.startStep('load-config');
        try { if (typeof App.loadPromptConfig === 'function') await App.loadPromptConfig(); } catch(e) { ChatLog.warn('INIT', '[Init] loadPromptConfig 失败:', e); }
        if (window.startupProgress) window.startupProgress.completeStep('load-config');
        if (App.config && !App.config.systemPrompt) {
            try { App.config.systemPrompt = App.getPrompt ? App.getPrompt('default') : ''; } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('INIT', 'loadPromptConfig 失败:', e && e.message); }
        }
        try { if (typeof App.waitForCardTypesConfig === 'function') await App.waitForCardTypesConfig(); } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('INIT', 'waitForCardTypesConfig 失败:', e && e.message); }
        if (typeof window.initRecognitionTypesConfig === 'function') {
            window.initRecognitionTypesConfig();
        }

        // --- 阶段2：加载用户设置 ---
        if (window.startupProgress) window.startupProgress.startStep('load-user-settings');
        try { if (typeof App.loadConfig === 'function') await App.loadConfig(); } catch(e) { ChatLog.warn('INIT', '[Init] loadConfig 失败:', e); }
        if (window.startupProgress) window.startupProgress.completeStep('load-user-settings');

        // --- 阶段3：预初始化 IndexedDB 存储层 ---
        if (typeof StorageDB !== 'undefined' && StorageDB.isAvailable()) {
            try { await StorageDB.init(); } catch(e) { ChatLog.warn('INIT', '[Init] StorageDB 初始化失败:', e); }
        }

        // --- 阶段4：加载历史记录 ---
        if (window.startupProgress) window.startupProgress.startStep('load-history');
        try { if (typeof App.loadHistory === 'function') await App.loadHistory(); else if (typeof AppHistory === 'object' && typeof AppHistory.loadHistory === 'function') await AppHistory.loadHistory(); } catch(e) { ChatLog.warn('INIT', '[Init] loadHistory 失败:', e); }
        if (window.startupProgress) window.startupProgress.completeStep('load-history');

        // --- 阶段5：初始化界面组件（导航切换依赖此步骤）---
        ChatLog.info('INIT', '[Init] 初始化标签页切换...');
        if (window.startupProgress) window.startupProgress.startStep('init-ui');
        try { if (typeof App.initTabs === 'function') App.initTabs(); } catch(e) { ChatLog.error('INIT', '[Init] initTabs 失败:', e); }
        try { if (typeof App.renderHistory === 'function') App.renderHistory(); else if (typeof AppHistory === 'object' && typeof AppHistory.renderHistory === 'function') AppHistory.renderHistory(); } catch(e) { ChatLog.warn('INIT', '[Init] renderHistory 失败:', e); }
        try { if (typeof App.onProviderChange === 'function') App.onProviderChange(); } catch(e) { ChatLog.warn('INIT', '[Init] onProviderChange 失败:', e); }
        if (window.startupProgress) window.startupProgress.completeStep('init-ui');

        // --- 关键路径完成：立即显示欢迎面板 ---
        ChatLog.info('INIT', '[Init] 关键路径完成，准备显示欢迎面板...');
        try {
            if (typeof window.showWelcomePanel === 'function') {
                window.showWelcomePanel();
            } else if (typeof App.addWelcomeMessage === 'function') {
                App.addWelcomeMessage();
            }
        } catch(e) {
            ChatLog.warn('INIT', '[Init] showWelcomePanel 失败:', e);
        }

        // --- 强制进入「对话」视图并展示像素欢迎页 ---
        // 目的：覆盖 index.html 静态/被外部进程重置的初始面板，保证每次刷新/启动都落在对话+欢迎页。
        try {
            var chatTab = document.querySelector('.tab-btn[data-tab="chat"]');
            if (chatTab) chatTab.click();
        } catch(e) {
            ChatLog.warn('INIT', '[Init] 强制进入对话视图失败:', e);
        }
        // 默认折叠代码编辑器，避免刷新后编辑器铺满界面
        try {
            var leftPanel = document.getElementById('left-panel');
            if (leftPanel && leftPanel.classList.contains('expanded') && typeof window.toggleEditor === 'function') {
                window.toggleEditor();
            }
        } catch(e) {
            ChatLog.warn('INIT', '[Init] 折叠代码编辑器失败:', e);
        }
        try {
            if (typeof window.ensurePixelWelcome === 'function') {
                window.ensurePixelWelcome();
            } else if (typeof window.addWelcomeMessage === 'function') {
                window.addWelcomeMessage();
            }
        } catch(e) {
            ChatLog.warn('INIT', '[Init] 注入像素欢迎页失败:', e);
        }

        // --- 隐藏启动进度条（欢迎面板已显示，无需再等）---
        if (window.startupProgress) {
            try { await window.startupProgress.hide(); } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('INIT', '隐藏启动进度条失败:', e && e.message); }
        }

        // ============================================================
        // 非关键任务：后台并行执行，不阻塞首屏
        // ============================================================
        ChatLog.info('INIT', '[Init] 启动后台非关键任务...');

        Promise.all([
            // 预加载 Skills
            (function() {
                return (typeof App.preloadAllSkills === 'function')
                    ? App.preloadAllSkills().catch(function(e) { ChatLog.warn('INIT', '[Init] preloadAllSkills 失败:', e); })
                    : Promise.resolve();
            })(),
            // 共享 Skill 加载
            (function() {
                return (window.SharedSkillLoader && typeof window.SharedSkillLoader.load === 'function')
                    ? window.SharedSkillLoader.load().catch(function(e) { ChatLog.warn('INIT', '[Init] 加载共享Skill失败:', e); })
                    : Promise.resolve();
            })(),
            // 初始化通信桥接
            (function() {
                return (typeof App.initBridge === 'function')
                    ? Promise.resolve(App.initBridge()).catch(function(e) { ChatLog.warn('INIT', '[Init] initBridge 失败:', e); })
                    : Promise.resolve();
            })(),
            // 检测本地模型
            (function() {
                return (typeof App.detectLocalModels === 'function')
                    ? App.detectLocalModels().catch(function(e) { ChatLog.warn('INIT', '[Init] detectLocalModels 失败:', e); })
                    : Promise.resolve();
            })(),
            // 更新模型选择器
            (function() {
                return (typeof App.updateChatModelSelector === 'function')
                    ? Promise.resolve(App.updateChatModelSelector()).catch(function(e) { ChatLog.warn('INIT', '[Init] updateChatModelSelector 失败:', e); })
                    : Promise.resolve();
            })(),
            // 恢复上次界面状态
            (function() {
                return (typeof App.restoreUIState === 'function')
                    ? Promise.resolve(App.restoreUIState()).catch(function(e) { ChatLog.warn('INIT', '[Init] restoreUIState 失败:', e); })
                    : Promise.resolve();
            })()
        ]).then(async function() {
            ChatLog.info('INIT', '[Init] 后台非关键任务完成');

            // 【修复】等 initBridge 完成后再初始化 MCP，避免桥接未就绪导致状态/工具加载失败
            try {
                let autoStart = false;
                try {
                    if (App.excelBridge && typeof App.excelBridge.GetMcpConfig === 'function') {
                        const configResult = await App.excelBridge.GetMcpConfig();
                        const configData = (typeof configResult === 'string')
                            ? JSON.parse(configResult || '{}')
                            : (configResult || {});
                        autoStart = configData && configData.success && configData.config && configData.config.autoStartClients;
                    } else {
                        try {
                            const bridgeBase = App.Endpoints ? App.Endpoints.excelBridge.baseUrl : (window.ExcelBridgeConfig ? window.ExcelBridgeConfig.baseUrl : 'http://localhost:9527');
                            const response = await fetch(`${bridgeBase}/api/mcp/config`);
                            const configData = await response.json();
                            autoStart = configData && configData.success && configData.config && configData.config.autoStartClients;
                            ChatLog.info('INIT', '[Init] Web模式MCP配置:', autoStart ? '自动启动' : '手动启动');
                        } catch (e) {
                            ChatLog.warn('INIT', '[Init] Web模式读取MCP配置失败:', e);
                        }
                    }
                } catch (e) {
                    const msg = (e && e.message) ? e.message : String(e || '');
                    if (!/e is not a function|0x80020006|未知名称|bridge-call-failed/i.test(msg)) {
                        ChatLog.warn('INIT', '[Init] 读取 MCP 配置失败:', msg);
                    }
                }

                var refreshFn = window.refreshMcpStatus || (App.McpModule && App.McpModule.refreshMcpStatus);
                if (typeof refreshFn === 'function') {
                    await refreshFn().catch(function(e) { ChatLog.warn('INIT', '[Init] 刷新MCP状态失败:', e); });
                }
                var loadToolsFn = window.loadMcpToolsForAI || (App.McpModule && App.McpModule.loadMcpToolsForAI);
                if (typeof loadToolsFn === 'function') {
                    await loadToolsFn().catch(function(e) { ChatLog.warn('INIT', '[Init] 加载MCP工具失败:', e); });
                }
                if (autoStart && typeof App.autoReconnectMcpClients === 'function') {
                    App.autoReconnectMcpClients().catch(function(e) { ChatLog.warn('INIT', '[Init] MCP 自动重连失败:', e); });
                }
            } catch (e) {
                ChatLog.warn('INIT', '[Init] MCP 初始化失败:', e);
            }
        }).catch(function(e) {
            ChatLog.warn('INIT', '[Init] 后台非关键任务异常:', e);
        });

        // --- 初始化扩展面板默认标签（后台） ---
        setTimeout(function() {
            try {
                var switchFn = window.switchExtensionTab || (App && App.switchExtensionTab);
                if (typeof switchFn === 'function') {
                    switchFn('skills');
                }
            } catch(e) {
                ChatLog.warn('INIT', '[Init] 初始化扩展面板失败:', e);
            }
        }, 0);

        // --- OCR 设置延迟加载 ---
        setTimeout(function() {
            try { if (window.SmartPilotHost && !window.SmartPilotHost.isExcel) return; if (typeof App.loadOcrSettings === 'function') App.loadOcrSettings(); } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('INIT', 'loadOcrSettings 失败:', e && e.message); }
        }, 500);

        // --- OCR 引擎初始化 ---
        setTimeout(async function() {
            if (window.SmartPilotHost && !window.SmartPilotHost.isExcel) return;
            try {
                if (App.excelBridge && typeof App.excelBridge.GetCurrentOcrEngine === 'function') {
                    const currentEngine = await App.excelBridge.GetCurrentOcrEngine();
                    window.__currentOcrEngine = currentEngine || '';
                    ChatLog.info('OCR', `启动时 OCR 引擎：${currentEngine}`);
                    if (currentEngine === 'cloud') {
                        const cloudConfig = await App.excelBridge.GetCloudOcrConfig();
                        const config = cloudConfig ? JSON.parse(cloudConfig) : {};
                        if (!config.apiKey && !config.secretKey) {
                            ChatLog.warn('OCR', '云端 OCR 未配置 API Key，强制切换到 PaddleOCR');
                            await App.excelBridge.SetOcrEngine('paddleocr');
                        }
                    }
                }
            } catch(e) {
                ChatLog.warn('OCR', 'OCR 引擎初始化失败:', e.message);
            }
        }, 1000);

        // --- 同步配置到后端 ---
        setTimeout(function() {
            try { if (typeof App.syncConfigToBackend === 'function') App.syncConfigToBackend(); } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('INIT', 'syncConfigToBackend 失败:', e && e.message); }
        }, 2000);

        // ==================== 事件绑定 ====================
        var messageInput = document.getElementById('message-input');
        if (messageInput) {
            // =================== Debug: 监测输入框被程序修改的情况 ===================
            // 在开发环境或调试时临时启用，用以定位谁在通过 script 设置 input.value
            try {
                (function(inputEl) {
                    var proto = Object.getPrototypeOf(inputEl);
                    var desc = Object.getOwnPropertyDescriptor(proto, 'value');
                    if (desc && typeof desc.set === 'function') {
                        var origSetter = desc.set.bind(inputEl);
                        var origGetter = desc.get ? desc.get.bind(inputEl) : function() { return this.getAttribute('value') || ''; };
                        Object.defineProperty(inputEl, 'value', {
                            get: function() { return origGetter(); },
                            set: function(v) {
                                var stack = '';
                                try {
                                    // 标准方式获取堆栈
                                    stack = (new Error()).stack || '';
                                } catch (e) {
                                    stack = '无法获取堆栈: ' + (e && e.message);
                                }
                                try {
                                    // 保存到全局变量，便于复制粘贴
                                    window.__lastProgrammaticValueSetStack = stack;
                                } catch (e) { /* ignore */ }

                                try {
                                    console.groupCollapsed('%c[DEBUG] message-input.value 被设置', 'color:#0b79ff; font-weight:600');
                                    ChatLog.info('INIT', '新值:', v);
                                    // 也把堆栈作为普通可复制日志输出
                                    ChatLog.info('INIT', '堆栈(复制此行以粘贴):\n' + stack);
                                    console.groupEnd();
                                } catch (e) { ChatLog.warn('INIT', '[DEBUG] 打印堆栈失败', e); }

                                // 调用原始 setter
                                try { origSetter(v); } catch (e) { /* ignore setter errors */ }

                                // 触发自定义事件，便于监听
                                try { inputEl.dispatchEvent(new Event('programmatic-value-set', { bubbles: true })); } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('INIT', 'dispatchEvent 失败:', e && e.message); }
                            },
                            configurable: true
                        });
                        // 监听自定义事件，方便外部捕获
                        inputEl.addEventListener('programmatic-value-set', function() {
                            
                        });
                    }
                })(messageInput);
            } catch (e) {
                ChatLog.warn('INIT', '[DEBUG] 输入框修改监测初始化失败:', e);
            }

            // 回车发送
            messageInput.addEventListener('keypress', function(e) {
                if (e.key === 'Enter' && !e.shiftKey) {
                    e.preventDefault();
                    if (typeof window.sendMessage === 'function') window.sendMessage();
                }
            });

            // 输入框点击焦点处理（WebView2兼容）
            messageInput.addEventListener('click', function(e) {
                e.stopPropagation();
                if (document.activeElement !== messageInput) {
                    var savedStart = messageInput.selectionStart;
                    var savedEnd = messageInput.selectionEnd;
                    messageInput.focus();
                    if (savedStart > 0 || savedEnd > savedStart) {
                        messageInput.setSelectionRange(savedStart, savedEnd);
                    }
                }
            });

            // 自动调整输入框高度（就地实现。原先借道 App.autoResizeTextarea 兼容存根，
            // 每敲一个键就打一条"调用了已删除的函数"警告，且存根语义上是空实现兜底）
            messageInput.addEventListener('input', function() {
                var el = this;
                if (!el) return;
                el.style.height = 'auto';
                el.style.height = Math.min(el.scrollHeight, 200) + 'px';
            });

            // 粘贴图片
            messageInput.addEventListener('paste', function(e) {
                if (typeof window.handlePaste === 'function') window.handlePaste(e);
            });
        }

        // 快捷指令
        try { if (typeof App.initQuickCommands === 'function') App.initQuickCommands(); } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('INIT', 'initQuickCommands 失败:', e && e.message); }

        // 引用区域管理器（依赖 Excel 选区接口，非 Excel 宿主不初始化，避免 700ms 空轮询报错）
        try {
            var _isExcelHost = !(window.SmartPilotHost && window.SmartPilotHost.isExcel === false);
            if (_isExcelHost && App.ReferenceAreaManager && typeof App.ReferenceAreaManager.init === 'function') {
                App.ReferenceAreaManager.init();
                // 定期刷新当前选区显示 / 添加模式自动捕获选区
                setInterval(function() {
                    if (App.ReferenceAreaManager && typeof App.ReferenceAreaManager.refresh === 'function') {
                        App.ReferenceAreaManager.refresh();
                    }
                }, 700);
            }
        } catch(e) { ChatLog.warn('INIT', '[Init] ReferenceAreaManager 初始化失败:', e); }

        // 新建对话按钮
        var newChatBtn = document.getElementById('new-chat-btn');
        if (newChatBtn) {
            newChatBtn.addEventListener('click', function() {
                if (typeof window.newChat === 'function') window.newChat();
            });
        }

        // 清空历史记录按钮
        var clearAllBtn = document.getElementById('clear-all-btn');
        if (clearAllBtn) {
            clearAllBtn.addEventListener('click', function(e) {
                e.preventDefault();
                e.stopPropagation();
                if (typeof window.clearAllHistory === 'function') window.clearAllHistory();
            });
        }

        // Ollama / llama.cpp 本地模型自动检测
        if (App.config && (App.config.apiType === 'ollama' || App.config.apiType === 'llamacpp')) {
            try { if (typeof App.detectLocalModels === 'function') App.detectLocalModels(); } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('INIT', 'detectLocalModels 失败:', e && e.message); }
        }

        // 代码片段列表初始化
        try { if (typeof App.loadCodeSnippets === 'function') App.loadCodeSnippets().then(function() {}).catch(function(_e){ if (typeof ChatLog !== 'undefined') ChatLog.debug('INIT', 'loadCodeSnippets 内部失败:', _e && _e.message); }); } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('INIT', 'loadCodeSnippets 失败:', e && e.message); }

        // 常驻快捷图标栏初始化（模型选择器上方）
        try { if (typeof App.renderQuickShortcutBar === 'function') App.renderQuickShortcutBar().catch(function(_e){ if (typeof ChatLog !== 'undefined') ChatLog.debug('INIT', 'renderQuickShortcutBar 内部失败:', _e && _e.message); }); } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('INIT', 'renderQuickShortcutBar 失败:', e && e.message); }

        // 欢迎面板快捷提示词网格初始化（从后端恢复用户标记的快捷入口）
        try { if (typeof App.renderQuickPrompts === 'function') App.renderQuickPrompts().catch(function(_e){ if (typeof ChatLog !== 'undefined') ChatLog.debug('INIT', 'renderQuickPrompts 内部失败:', _e && _e.message); }); } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('INIT', 'renderQuickPrompts 失败:', e && e.message); }

        // 图片缩略图点击放大功能
        try {
            const chatMessages = document.getElementById('chat-messages');
            if (chatMessages) {
                // 使用事件委托监听聊天消息中的图片点击
                chatMessages.addEventListener('click', function(e) {
                    // 查找被点击的 img 元素
                    const img = e.target.closest('.message img, .image-container img');
                    if (img) {
                        // 【修复】优先使用 data-base64 属性（blob URL revoke 后仍有效）
                        const imgSrc = img.getAttribute('data-base64') || img.src;
                        if (imgSrc && typeof App.openZoom === 'function') {
                            e.preventDefault();
                            e.stopPropagation();
                            App.openZoom(imgSrc);
                        }
                    }
                });
            }

            // 点击遮罩层关闭放大视图
            const zoomOverlay = document.getElementById('image-zoom-overlay');
            if (zoomOverlay) {
                zoomOverlay.addEventListener('click', function() {
                    if (typeof App.closeZoom === 'function') {
                        App.closeZoom();
                    }
                });
            }

            // ESC 键关闭放大视图
            document.addEventListener('keydown', function(e) {
                if (e.key === 'Escape' || e.keyCode === 27) {
                    const overlay = document.querySelector('.image-zoom-overlay');
                    if (overlay && overlay.classList.contains('active')) {
                        if (typeof App.closeZoom === 'function') {
                            App.closeZoom();
                        }
                    }
                }
            });

            ChatLog.info('INIT', '[Init] ✅ 图片缩略图点击放大功能已初始化');
        } catch(e) {
            ChatLog.warn('INIT', '[Init] ⚠️ 图片缩略图点击放大功能初始化失败:', e.message);
        }

        // 初始化 AI Pipeline（延迟到所有模块加载完成后）
        try {
            if (typeof App.AIPipeline === 'function') {
                App.aiPipeline = new App.AIPipeline();
                ChatLog.info('INIT', '[Init] ✅ AI Pipeline 初始化完成');
            }
        } catch(e) {
            ChatLog.warn('INIT', '[Init] ⚠️ AI Pipeline 初始化失败:', e.message);
        }

        ChatLog.info('INIT', '[Init] ✅ DOMContentLoaded 初始化完成');
    });

})();