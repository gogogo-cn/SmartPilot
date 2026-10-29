/**
 * ============================================================
 * app-skill-mgmt.code-library.js — Skill 代码库管理子模块
 * ============================================================
 * 从 app-skill-mgmt.js L1363-2313 抽取的「代码库/片段管理」职责域。
 * 与原文件同步维护同一组 App.* 与 window.* 导出契约。
 * 依赖：window.App、App.codeSnippets、App.excelBridge、window.ExcelBridgeConfig。
 * ============================================================
 */

(function() {
    'use strict';

    // ---- 子模块自带辅助（与主文件 app-skill-mgmt.js 同源，保证加载顺序自足） ----
    function isRealWebView2() {
        return !!(window.chrome && window.chrome.webview && !window.chrome.webview._mockMode &&
                  window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge);
    }

    function getBridge() {
        return App.excelBridge ||
            (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge);
    }

    async function callBridgeMethod(method, args) {
        if (isRealWebView2()) {
            const bridge = getBridge();
            if (bridge && typeof bridge[method] === 'function') {
                return await bridge[method](...args);
            }
            throw new Error(`WebView2 中不存在方法: ${method}`);
        }
        const baseUrl = window.ExcelBridgeConfig ? window.ExcelBridgeConfig.baseUrl : 'http://localhost:9527';
        const response = await fetch(`${baseUrl}/api/excel/invoke`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ method: method, arguments: args })
        });
        const result = await response.json();
        return JSON.stringify(result.data || result);
    }

    function escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.appendChild(document.createTextNode(text));
        return div.innerHTML;
    }

    // ---- 全局操作锁（与 app-skill-mgmt.js 同源；2026-10-21 修复）----
    // deleteSnippetFromLibrary 从 app.js 抽取时漏带了本函数的定义，点击删除即报
    // ReferenceError: _tryAcquireActionLock is not defined。
    // 语义：防止重复绑定的监听器导致 confirm 弹多次（同一 actionKey 在 lockMs 内只放行一次）。
    // 用独立 Map（_snippetActionLock）而非与主文件共享——两个模块各自锁、互不干扰。
    const _snippetActionLock = new Map();
    function _tryAcquireActionLock(actionKey, lockMs = 1500) {
        const now = Date.now();
        const expires = _snippetActionLock.get(actionKey) || 0;
        if (now < expires) return false;
        _snippetActionLock.set(actionKey, now + lockMs);
        // 自动清理，防止 Map 无限增长
        setTimeout(() => _snippetActionLock.delete(actionKey), lockMs + 100);
        return true;
    }

    window.App = window.App || {};
    var CODE_FINGERPRINTS = {
        '清空工作表': {
            patterns: [/清空.*工作表|清除.*工作表|清空.*sheet|清空.*数据/i, /clear.*sheet|clear.*data/i],
            keywords: ['清空', '清除', 'clear']
        },
        '复制工作表': {
            patterns: [/复制.*工作表|拷贝.*工作表|复制.*sheet/i, /copy.*sheet|duplicate.*sheet/i],
            keywords: ['复制', '拷贝', 'copy']
        },
        '删除重复项': {
            patterns: [/删除.*重复|去重|重复项|删除重复/i, /remove.*duplicate/i],
            keywords: ['去重', '重复', 'duplicate']
        },
        '合并单元格': {
            patterns: [/合并.*单元格|单元格.*合并/i, /merge.*cell/i],
            keywords: ['合并', 'merge']
        },
        '设置格式': {
            patterns: [/设置.*格式|格式化|设置.*样式|设置.*颜色|设置.*字体/i, /format|style|color|font/i],
            keywords: ['格式', '样式', 'format']
        },
        '创建图表': {
            patterns: [/创建.*图表|生成.*图表|画图|绘制.*图/i, /create.*chart|generate.*chart/i],
            keywords: ['图表', 'chart', '图']
        },
        '排序': {
            patterns: [/排序|按.*排序|升序|降序/i, /sort|ascending|descending/i],
            keywords: ['排序', 'sort']
        },
        '筛选': {
            patterns: [/筛选|过滤|条件筛选/i, /filter/i],
            keywords: ['筛选', '过滤', 'filter']
        },
        '求和': {
            patterns: [/求和|合计|总计|sum/i, /sum|total/i],
            keywords: ['求和', '合计', 'sum']
        },
        '平均值': {
            patterns: [/平均|均值|average/i, /average|mean/i],
            keywords: ['平均', 'average']
        },
        '最大值': {
            patterns: [/最大|max|maximum/i, /max|maximum|largest/i],
            keywords: ['最大', 'max']
        },
        '最小值': {
            patterns: [/最小|min|minimum/i, /min|minimum|smallest/i],
            keywords: ['最小', 'min']
        },
        '查找替换': {
            patterns: [/查找.*替换|搜索.*替换|替换.*内容/i, /find.*replace|search.*replace/i],
            keywords: ['查找', '替换', 'find']
        },
        '插入行列': {
            patterns: [/插入.*行|插入.*列|添加.*行|添加.*列/i, /insert.*row|insert.*column/i],
            keywords: ['插入', 'insert']
        },
        '删除行列': {
            patterns: [/删除.*行|删除.*列|移除.*行|移除.*列/i, /delete.*row|delete.*column/i],
            keywords: ['删除行', '删除列']
        },
        '数据验证': {
            patterns: [/数据验证|验证.*输入|输入.*限制/i, /data.*validation|validation/i],
            keywords: ['验证', 'validation']
        },
        '条件格式': {
            patterns: [/条件格式|条件.*格式/i, /conditional.*format/i],
            keywords: ['条件格式', 'conditional']
        },
        '保护工作表': {
            patterns: [/保护.*工作表|锁定.*工作表|工作表.*保护/i, /protect.*sheet|lock.*sheet/i],
            keywords: ['保护', '锁定', 'protect']
        },
        '打印设置': {
            patterns: [/打印.*设置|设置.*打印|打印.*区域/i, /print.*setting|print.*area/i],
            keywords: ['打印', 'print']
        },
        '导出PDF': {
            patterns: [/导出.*PDF|保存.*PDF|转.*PDF/i, /export.*pdf|save.*pdf/i],
            keywords: ['PDF', '导出']
        }
    };

    // --- 来自 app.js:L23144-L23160 ---
    function generateCodeFingerprint(code) {
        if (!code) return null;
        const lowerMessage = code.toLowerCase();

        for (const [feature, config] of Object.entries(CODE_FINGERPRINTS)) {
            for (const pattern of config.patterns) {
                if (pattern.test(code)) {
                    return {
                        feature: feature,
                        keywords: config.keywords,
                        confidence: 'high'
                    };
                }
            }
        }
        return null;
    }

    // --- 来自 app.js:L23163-L23180 ---
    function calculateSimilarity(fp1, fp2) {
        if (!fp1 || !fp2) return 0;

        const s1 = fp1.toLowerCase();
        const s2 = fp2.toLowerCase();

        // 完全匹配
        if (s1 === s2) return 1;

        // 包含关系
        if (s1.includes(s2) || s2.includes(s1)) return 0.9;

        // 关键词匹配
        const words1 = s1.split(/\s+/);
        const words2 = s2.split(/\s+/);
        const common = words1.filter(w => words2.includes(w));
        return common.length / Math.max(words1.length, words2.length);
    }

    // --- 来自 app.js:L23183-L23224 ---
    function findSimilarCodeInLibrary(code) {
        if (!App.codeSnippets || App.codeSnippets.length === 0) return null;

        const lowerMessage = code.toLowerCase();
        let bestMatch = null;
        let bestScore = 0;

        for (const snippet of App.codeSnippets) {
            let score = 0;
            const snippetText = (snippet.name + ' ' + (snippet.description || '')).toLowerCase();

            // 1. 精确匹配名称或描述
            if (snippetText.includes(lowerMessage) || lowerMessage.includes(snippet.name.toLowerCase())) {
                score += 1.0;
            }

            // 2. 功能指纹匹配
            const fingerprint = generateCodeFingerprint(code);
            if (fingerprint) {
                const snippetFingerprint = generateCodeFingerprint(snippet.name + ' ' + (snippet.description || ''));
                if (snippetFingerprint && snippetFingerprint.feature === fingerprint.feature) {
                    score += 0.8;
                }

                // 3. 关键词匹配
                for (const keyword of fingerprint.keywords) {
                    if (snippetText.includes(keyword.toLowerCase())) {
                        score += 0.3;
                    }
                }
            }

            // 4. 相似度计算
            score += calculateSimilarity(code, snippet.name) * 0.5;

            if (score > bestScore && score >= 0.5) { // 阈值0.5
                bestScore = score;
                bestMatch = { ...snippet, matchScore: score };
            }
        }

        return bestMatch;
    }

    // --- 来自 app.js:L23227-L23258 ---
    function findAllSimilarSnippets(code) {
        if (!App.codeSnippets || App.codeSnippets.length === 0) return [];

        const matches = [];
        const lowerMessage = code.toLowerCase();
        const fingerprint = generateCodeFingerprint(code);

        for (const snippet of App.codeSnippets) {
            let score = 0;
            const snippetText = (snippet.name + ' ' + (snippet.description || '')).toLowerCase();

            if (fingerprint) {
                const snippetFingerprint = generateCodeFingerprint(snippet.name + ' ' + (snippet.description || ''));
                if (snippetFingerprint && snippetFingerprint.feature === fingerprint.feature) {
                    score += 0.8;
                }

                for (const keyword of fingerprint.keywords) {
                    if (snippetText.includes(keyword.toLowerCase())) {
                        score += 0.3;
                    }
                }
            }

            score += calculateSimilarity(code, snippet.name) * 0.5;

            if (score >= 0.4) {
                matches.push({ ...snippet, matchScore: score });
            }
        }

        return matches.sort((a, b) => b.matchScore - a.matchScore);
    }

    // --- 来自 app.js:L23261-L23387 ---
    function addCacheHitMessage(snippetName) {
        const snippet = App.codeSnippets.find(s => s.name === snippetName);
        if (!snippet) return;

        const container = document.getElementById('chat-messages');
        if (!container) return;

        // 移除欢迎消息
        const welcomeMessage = container.querySelector('.welcome-message');
        if (welcomeMessage) {
            welcomeMessage.remove();
        }

        // 创建消息包装器
        const wrapper = document.createElement('div');
        wrapper.className = 'message-wrapper ai animate-in';

        // 创建头像
        const avatar = document.createElement('div');
        avatar.className = 'message-avatar ai';
        avatar.innerHTML = '<img src="images/chat.png" alt="SmartPilot" style="width:28px;height:28px;border-radius:50%;object-fit:cover;">';

        const contentDiv = document.createElement('div');
        contentDiv.className = 'message-content';

        // 创建消息头部
        const header = document.createElement('div');
        header.className = 'message-header';
        const providerNames = {
            'ollama': 'Ollama',
            'openai': 'OpenAI',
            'azure': 'Azure',
            'wenxin': '文心一言',
            'qianwen': '通义千问',
            'zhipu': '智谱 AI',
            'xunfei': '讯飞星火',
            'moonshot': 'Moonshot',
            'deepseek': '深度求索 (DeepSeek)',
            'volcengine': '火山引擎 (豆包)',
            'custom': '自定义'
        };
        const _rawModel = App.currentChatModel || App.config.model || '';
        const _providerConfig = (typeof App !== 'undefined' && App.PROVIDER_CONFIGS) ? App.PROVIDER_CONFIGS[App.config.apiType] : null;
        const _actualModel = (_rawModel === '默认模型' || _rawModel === '当前模型' || _rawModel === 'AUTO' || !_rawModel)
            ? (_providerConfig ? (_providerConfig.model || '') : '')
            : _rawModel;
        const modelName = _actualModel || '默认模型';
        header.innerHTML = `
            <span>${providerNames[App.config.apiType] || 'AI'}</span>
            <span class="model-badge">${modelName}</span>
        `;

        // 创建消息正文
        const messageDiv = document.createElement('div');
        messageDiv.className = 'message ai';

        // 构建内容：提示 + 代码块
        const escapedCode = escapeHtml(snippet.code);
        messageDiv.innerHTML = `
            <div style="color: #666; margin-bottom: 10px; font-size: 13px;">💡 从您的代码库中找到匹配代码（${escapeHtml(snippet.name)}）</div>
            <div class="code-block" data-original-code="${escapedCode.replace(/"/g, '&quot;')}">${escapedCode}</div>
        `;

        // 创建按钮容器
        const buttonWrapper = document.createElement('div');
        buttonWrapper.className = 'cache-action-buttons';
        buttonWrapper.style.cssText = 'display: flex; gap: 10px; margin-top: 10px;';

        // 直接执行按钮
        const executeBtn = document.createElement('button');
        executeBtn.innerHTML = '▶️ 执行';
        executeBtn.style.cssText = 'padding: 6px 14px; border: none; border-radius: 6px; background: linear-gradient(135deg, #4CAF50, #45a049); color: white; cursor: pointer; font-size: 13px; font-weight: 500; box-shadow: 0 2px 4px rgba(76,175,80,0.3); transition: all 0.2s;';
        executeBtn.onmouseenter = () => executeBtn.style.transform = 'translateY(-1px)';
        executeBtn.onmouseleave = () => executeBtn.style.transform = 'translateY(0)';
        executeBtn.onclick = () => {
            loadSnippetToEditorAndExecute(snippet.name);
        };

        // 修改按钮
        const modifyBtn = document.createElement('button');
        modifyBtn.innerHTML = '✏️ 修改';
        modifyBtn.style.cssText = 'padding: 6px 14px; border: 1px solid #e0e0e0; border-radius: 6px; background: white; color: #666; cursor: pointer; font-size: 13px; font-weight: 500; transition: all 0.2s;';
        modifyBtn.onmouseenter = () => { modifyBtn.style.background = '#f5f5f5'; modifyBtn.style.borderColor = '#ccc'; };
        modifyBtn.onmouseleave = () => { modifyBtn.style.background = 'white'; modifyBtn.style.borderColor = '#e0e0e0'; };

        // 其他按钮
        const othersBtn = document.createElement('button');
        othersBtn.innerHTML = '📚 其他';
        othersBtn.style.cssText = 'padding: 6px 14px; border: 1px solid #e0e0e0; border-radius: 6px; background: white; color: #666; cursor: pointer; font-size: 13px; font-weight: 500; transition: all 0.2s;';
        othersBtn.onmouseenter = () => { othersBtn.style.background = '#f5f5f5'; othersBtn.style.borderColor = '#ccc'; };
        othersBtn.onmouseleave = () => { othersBtn.style.background = 'white'; othersBtn.style.borderColor = '#e0e0e0'; };
        othersBtn.onclick = () => {
            const similarSnippets = findAllSimilarSnippets(snippet.code);
            showSimilarSnippetsDialog(similarSnippets);
        };

        buttonWrapper.appendChild(executeBtn);
        buttonWrapper.appendChild(modifyBtn);
        buttonWrapper.appendChild(othersBtn);
        messageDiv.appendChild(buttonWrapper);

        // 组装消息
        contentDiv.appendChild(header);
        contentDiv.appendChild(messageDiv);
        wrapper.appendChild(avatar);
        wrapper.appendChild(contentDiv);
        container.appendChild(wrapper);

        // 滚动到底部
        container.scrollTop = container.scrollHeight;

        // 同时提取代码到编辑器
        const editor = document.getElementById('code-editor');
        if (editor) {
            editor.value = snippet.code;
            editor.style.display = 'block';
            document.getElementById('editor-actions').style.display = 'flex';
            if (typeof App.updateLineNumbers === 'function') {
                App.updateLineNumbers();
            }
        }
    }

    // --- 来自 app.js:L23390-L23405 ---
    function loadSnippetToEditorAndExecute(snippetName) {
        const snippet = App.codeSnippets.find(s => s.name === snippetName);
        if (!snippet) return;

        // 加载到编辑器
        const editor = document.getElementById('code-editor');
        if (editor) {
            editor.value = snippet.code;
            if (typeof App.updateLineNumbers === 'function') {
                App.updateLineNumbers();
            }
            App.showToast('代码已加载到编辑器');
        }

        // 显示编辑器
        if (typeof App.showCodeEditor === 'function') {
            App.showCodeEditor();
        }

        // 自动执行
        setTimeout(() => {
            if (typeof App.executeCode === 'function') {
                App.executeCode();
            }
        }, 500);
    }

    // --- 来自 app.js:L23408-L23458 ---
    function showSimilarSnippetsDialog(similarList) {
        if (!similarList || similarList.length === 0) {
            App.showToast('没有其他相似代码');
            return;
        }

        // 创建对话框
        const dialog = document.createElement('div');
        dialog.className = 'similar-snippets-dialog';
        dialog.style.cssText = 'position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.5); display: flex; align-items: center; justify-content: center; z-index: 10000;';

        const content = document.createElement('div');
        content.style.cssText = 'background: white; border-radius: 12px; width: 90%; max-width: 600px; max-height: 80vh; overflow: hidden; display: flex; flex-direction: column;';

        // 头部
        const header = document.createElement('div');
        header.style.cssText = 'padding: 16px 20px; border-bottom: 1px solid #eee; display: flex; justify-content: space-between; align-items: center;';
        header.innerHTML = `<h3 style="margin: 0; font-size: 16px;">找到 ${similarList.length} 个相似代码</h3>`;

        const closeBtn = document.createElement('button');
        closeBtn.innerHTML = '×';
        closeBtn.style.cssText = 'background: none; border: none; font-size: 24px; cursor: pointer; color: #999;';
        closeBtn.onclick = () => dialog.remove();
        header.appendChild(closeBtn);

        // 列表
        const list = document.createElement('div');
        list.style.cssText = 'overflow-y: auto; padding: 16px; max-height: 60vh;';

        similarList.forEach((snippet, index) => {
            const item = document.createElement('div');
            item.style.cssText = 'padding: 12px; border: 1px solid #eee; border-radius: 8px; margin-bottom: 12px; cursor: pointer; transition: background 0.2s;';
            item.innerHTML = `
                <div style="font-weight: bold; margin-bottom: 4px;">${index + 1}. ${escapeHtml(snippet.name)}</div>
                <div style="font-size: 12px; color: #666; margin-bottom: 8px;">${escapeHtml(snippet.description || '无描述')}</div>
                <div style="font-size: 11px; color: #4CAF50;">匹配度: ${Math.round(snippet.matchScore * 100)}%</div>
            `;
            item.onmouseenter = () => item.style.background = '#f5f5f5';
            item.onmouseleave = () => item.style.background = 'white';
            item.onclick = () => {
                loadSnippetToEditorAndExecute(snippet.name);
                dialog.remove();
            };
            list.appendChild(item);
        });

        content.appendChild(header);
        content.appendChild(list);
        dialog.appendChild(content);
        document.body.appendChild(dialog);
    }

    // --- 来自 app.js:L23467-L23545 ---
    async function loadCodeSnippets() {
        try {
            // 1. 先从 localStorage 加载用户手动保存的片段
            const saved = localStorage.getItem(App.SNIPPETS_KEY || 'code_snippets');
            if (saved) {
                App.codeSnippets = JSON.parse(saved);
            } else {
                App.codeSnippets = [];
            }

            // 2. 从后端 CustomSkills 文件夹加载
            const bridge = getBridge();
            if (bridge) {
                try {
                    const allSkillsJson = await bridge.GetAllSkills();
                    if (allSkillsJson) {
                        const allSkills = typeof allSkillsJson === 'string'
                            ? JSON.parse(allSkillsJson)
                            : allSkillsJson;

                        const allSkillsList = Array.isArray(allSkills.skills) ? allSkills.skills : Array.isArray(allSkills) ? allSkills : [];

                        // 只加载 CustomSkills 文件夹中的文件(名称以 Snippet_ 开头)
                        const customSkillFiles = allSkillsList.filter(s =>
                            s.templateCode &&
                            s.templateCode.trim() !== '' &&
                            s.name &&
                            s.name.startsWith('Snippet_')
                        );

                        ChatLog.info('CODELIB', '[CodeLibrary] CustomSkills 文件夹中有', customSkillFiles.length, '个文件');

                        // 同步到 codeSnippets(如果不存在则添加)
                        customSkillFiles.forEach(skill => {
                            const existingIndex = App.codeSnippets.findIndex(s =>
                                s.skillFileName === skill.name ||
                                s.id === skill.name
                            );

                            if (existingIndex === -1) {
                                // 添加新片段
                                App.codeSnippets.push({
                                    id: skill.name,
                                    name: skill.displayName || skill.name,
                                    description: skill.description || '',
                                    code: skill.templateCode,
                                    savedAsSkill: true,
                                    skillId: skill.name,
                                    skillFileName: skill.name,
                                    timestamp: new Date(skill.lastModified || skill.createdDate).getTime() || Date.now()
                                });
                                ChatLog.info('CODELIB', `[CodeLibrary] 已添加: ${skill.displayName || skill.name}`);
                            } else {
                                // 更新现有片段
                                App.codeSnippets[existingIndex].code = skill.templateCode;
                                App.codeSnippets[existingIndex].name = skill.displayName || skill.name;
                                App.codeSnippets[existingIndex].description = skill.description || '';
                                App.codeSnippets[existingIndex].savedAsSkill = true;
                                ChatLog.info('CODELIB', `[CodeLibrary] 已更新: ${skill.displayName || skill.name}`);
                            }
                        });

                        // 保存到 localStorage
                        if (customSkillFiles.length > 0) {
                            saveCodeSnippetsToStorage();
                        }
                    }
                } catch (syncError) {
                    const msg = syncError && syncError.message ? syncError.message : String(syncError || '');
                    if (!/e is not a function|0x80020006|未知名称|bridge-call-failed/i.test(msg)) {
                        ChatLog.warn('CODELIB', '[CodeLibrary] 从 CustomSkills 加载失败:', msg);
                    }
                }
            }

            ChatLog.info('CODELIB', '[CodeLibrary] 最终加载了', App.codeSnippets.length, '个代码片段');
            App.codeSnippetsLoaded = true;
        } catch (error) {
            ChatLog.error('CODELIB', '加载代码片段失败:', error);
            App.codeSnippets = [];
            App.codeSnippetsLoaded = true;
        }
    }

    // --- 来自 app.js:L23548-L23555 ---
    function saveCodeSnippetsToStorage() {
        try {
            localStorage.setItem(App.SNIPPETS_KEY || 'code_snippets', JSON.stringify(App.codeSnippets));
        } catch (error) {
            ChatLog.error('CODELIB', '保存代码片段失败:', error);
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">保存失败: ' + error.message + '</span>', true) : App.showToast('保存失败: ' + error.message);
        }
    }

    // --- 来自 app.js:L23558-L23597 ---
    function renderCodeLibrary() {
        const container = document.getElementById('code-library-list');
        if (!container) {
            ChatLog.error('CODELIB', '[CodeLibrary] 错误: 找不到 code-library-list 容器');
            return;
        }

        ChatLog.info('CODELIB', `[CodeLibrary] 开始渲染, codeSnippets 数量: ${App.codeSnippets.length}`);
        ChatLog.info('CODELIB', '[CodeLibrary] codeSnippets 数据:', App.codeSnippets);

        if (App.codeSnippets.length === 0) {
            container.innerHTML = '<div class="snippet-empty">暂无保存的代码片段<br>在编辑器中编写代码后点击"保存代码"添加</div>';
            return;
        }

        // 按时间倒序排列
        const sortedSnippets = [...App.codeSnippets].sort((a, b) => b.timestamp - a.timestamp);
        ChatLog.info('CODELIB', '[CodeLibrary] 排序后 snippets:', sortedSnippets);

        const loadIcon = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>`;
        const editIcon = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>`;
        const deleteIcon = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/></svg>`;

        const html = sortedSnippets.map(snippet => `
            <div class="snippet-item" data-snippet-id="${snippet.id}">
                <div class="snippet-item-main">
                    <div class="snippet-item-title-row">
                        <span class="snippet-item-name">${escapeHtml(snippet.name || '(无名称)')}</span>
                        ${snippet.savedAsSkill ? '<span class="snippet-skill-badge" title="已在快捷指令中可用">快捷指令</span>' : ''}
                    </div>
                    <div class="snippet-item-desc">${escapeHtml(snippet.description || '无描述')}</div>
                </div>
                <div class="snippet-item-actions">
                    <button class="snippet-action-btn" onclick="App.loadSnippetFromLibrary('${snippet.id}')" title="加载到编辑器">${loadIcon}</button>
                    <button class="snippet-action-btn" onclick="App.editSnippetFromLibrary('${snippet.id}')" title="编辑">${editIcon}</button>
                    <button class="snippet-action-btn delete" onclick="App.deleteSnippetFromLibrary('${snippet.id}')" title="删除">${deleteIcon}</button>
                </div>
            </div>
        `).join('');

        ChatLog.info('CODELIB', '[CodeLibrary] 生成的 HTML 长度:', html.length);
        ChatLog.info('CODELIB', '[CodeLibrary] snippet-item 数量:', (html.match(/snippet-item/g) || []).length);

        container.innerHTML = html;
        ChatLog.info('CODELIB', '[CodeLibrary] DOM 中实际渲染的 snippet-item 数量:', container.querySelectorAll('.snippet-item').length);
    }

    // --- 来自 app.js:L23600-L23606 ---
    async function showCodeLibrary() {
        // 强制重新加载，确保获取最新数据
        ChatLog.info('CODELIB', '[CodeLibrary] 打开代码库，重新加载数据...');
        await loadCodeSnippets();
        renderCodeLibrary();
        document.getElementById('code-library-dialog').style.display = 'flex';
    }

    // --- 来自 app.js:L23609-L23611 ---
    function closeCodeLibrary() {
        document.getElementById('code-library-dialog').style.display = 'none';
    }

    // --- 来自 app.js:L23614-L23633 ---
    function loadSnippetFromLibrary(name) {
        const snippet = App.codeSnippets.find(s => s.id === name || s.name === name);
        if (!snippet) return;

        // 如果编辑器中有代码，询问是否覆盖
        const currentCode = document.getElementById('code-editor').value.trim();
        if (currentCode && currentCode !== snippet.code) {
            if (!confirm('编辑器中已有代码，是否覆盖？')) {
                return;
            }
        }

        // 自动格式化代码缩进
        let codeToLoad = snippet.code;
        if (typeof App.autoFormatCode === 'function') {
            codeToLoad = App.autoFormatCode(snippet.code, 'vba');
        }

        document.getElementById('code-editor').value = codeToLoad;
        // 更新行号、高亮、折叠按钮
        if (typeof App.updateLineNumbers === 'function') {
            setTimeout(() => {
                App.updateLineNumbers();
                if (typeof App.updateSyntaxHighlight === 'function') {
                    App.updateSyntaxHighlight();
                }
                if (typeof App.updateFoldButtons === 'function') {
                    App.updateFoldButtons();
                }
            }, 10);
        }
        App.showToast('代码已加载到编辑器');
        closeCodeLibrary();
    }

    // --- 来自 app.js:L23636-L23654 ---
    function editSnippetFromLibrary(name) {
        const snippet = App.codeSnippets.find(s => s.id === name || s.name === name);
        if (!snippet) return;

        closeCodeLibrary();

        document.getElementById('snippet-edit-id').value = snippet.id;
        document.getElementById('snippet-input-name').value = snippet.name;
        document.getElementById('snippet-input-desc').value = snippet.description || '';
        document.getElementById('snippet-input-code').value = snippet.code;

        // 恢复"保存为 Skill"勾选状态
        const saveAsSkillCheckbox = document.getElementById('snippet-save-as-skill');
        if (saveAsSkillCheckbox) {
            saveAsSkillCheckbox.checked = snippet.savedAsSkill || false;
        }

        document.getElementById('snippet-save-dialog').style.display = 'flex';
    }

    // --- 来自 app.js:L23657-L23693 ---
    async function deleteSnippetFromLibrary(name) {
        const snippet = App.codeSnippets.find(s => s.id === name || s.name === name);
        if (!snippet) return;

        const lockKey = 'deleteSnippet:' + (snippet.id || snippet.name);
        if (!_tryAcquireActionLock(lockKey, 1500)) return;

        if (!confirm(`确定要删除代码片段 "${snippet.name}" 吗？\n\n${snippet.savedAsSkill ? '⚠️ 同时会从快捷指令中移除' : ''}`)) {
            return;
        }

        // 如果该片段之前保存为 Skill，同时删除 CustomSkill
        if (snippet.savedAsSkill && getBridge()) {
            try {
                const skillNameToDelete = 'Snippet_' + sanitizeSkillFileName(snippet.name);
                ChatLog.info('SKILL', '[SKILL] 🗑️ 删除片段，同步删除 CustomSkill:', skillNameToDelete);

                const result = await callBridgeMethod('DeleteSkill', [skillNameToDelete]);
                const resultObj = JSON.parse(result);
                ChatLog.info('SKILL', '[SKILL] DeleteSkill 返回:', result);

                if (resultObj.success) {
                    ChatLog.info('SKILL', '[SKILL] ✅ 已同步删除 CustomSkill:', skillNameToDelete);
                } else {
                    ChatLog.warn('SKILL', '[SKILL] ⚠️ DeleteSkill 返回失败:', resultObj.error);
                }

                if (typeof App.loadSkillsAsCommands === 'function') {
                    await App.loadSkillsAsCommands(true);
                }
            } catch (e) {
                ChatLog.warn('SKILL', '删除关联 CustomSkill 失败（非致命）:', e.message);
            }

            if (typeof App.loadSkillsAsCommands === 'function') {
                await App.loadSkillsAsCommands(true);
            }
        }

        App.codeSnippets = App.codeSnippets.filter(s => s.id !== snippet.id);
        saveCodeSnippetsToStorage();
        renderCodeLibrary();
        App.showToast('代码片段已删除' + (snippet.savedAsSkill ? '（已从快捷指令移除）' : ''));
    }

    // --- 来自 app.js:L23703-L23745 ---
    function extractCodeTags(code) {
        const tags = [];
        const codeLower = code.toLowerCase();

        const tagPatterns = {
            '清空': /clear|清空|清除|cells\.clear/i,
            '复制': /copy|复制|拷贝/i,
            '删除': /delete|remove|清除|清空/i,
            '插入': /insert|插入|add/i,
            '排序': /sort|排序/i,
            '筛选': /filter|筛选|autofilter/i,
            '格式化': /format|样式|颜色|字体|interior|font|border/i,
            '图表': /chart|图表|graph|series/i,
            '数据验证': /validation|验证|validation/i,
            '条件格式': /conditional|条件格式|formatconditions/i,
            '保护工作表': /protect|保护|锁定|lock/i,
            '打印': /print|打印|printout/i,
            '导出': /export|导出|saveas/i,
            '求和': /sum|求和|合计|总计/i,
            '平均': /average|avg|平均/i,
            '最大': /max|最大/i,
            '最小': /min|最小/i,
            '查找': /find|查找|search/i,
            '替换': /replace|替换/i,
            '合并': /merge|合并/i,
            '拆分': /split|拆分/i,
            '循环': /for|each|while|do|loop/i,
            '条件': /if|then|else|elseif|select/i,
            '函数': /function|sub|end sub|end function/i,
            '事件': /worksheet_|workbook_|selectionchange|change/i,
            '对话框': /msgbox|inputbox|userform/i,
            '文件操作': /workbooks\.open|workbooks\.save|files\.open/i,
            '错误处理': /on error|err\.|resume/i
        };

        for (const [tag, pattern] of Object.entries(tagPatterns)) {
            if (pattern.test(codeLower)) {
                tags.push(tag);
            }
        }

        return tags;
    }

    // --- 来自 app.js:L23748-L23774 ---
    function generateSnippetNameSuggestion(code, description) {
        const codeLower = code.toLowerCase();

        // 尝试从Sub/Function名称提取
        const subMatch = code.match(/sub\s+(\w+)/i);
        if (subMatch) {
            return subMatch[1];
        }

        const funcMatch = code.match(/function\s+(\w+)/i);
        if (funcMatch) {
            return funcMatch[1];
        }

        // 根据功能生成建议名称
        if (/clear|清空|清除/i.test(codeLower)) return '清空数据';
        if (/copy|复制|拷贝/i.test(codeLower)) return '复制数据';
        if (/delete|删除/i.test(codeLower)) return '删除数据';
        if (/sort|排序/i.test(codeLower)) return '排序数据';
        if (/filter|筛选/i.test(codeLower)) return '筛选数据';
        if (/format|格式/i.test(codeLower)) return '设置格式';
        if (/chart|图表/i.test(codeLower)) return '创建图表';
        if (/sum|求和/i.test(codeLower)) return '求和计算';
        if (/print|打印/i.test(codeLower)) return '打印设置';

        return '';
    }

    // 生成 Skill 文件名：保留中文/英文/数字，仅将 Windows 非法文件名字符替换为下划线。
    // 旧逻辑用 [^a-zA-Z0-9] 白名单，会把中文全部替换成下划线，导致文件名变成 Snippet_____（不完整）。
    function sanitizeSkillFileName(name) {
        const s = String(name || '').trim();
        if (!s) return 'Snippet';
        return s
            .replace(/[\\/:*?"<>|]/g, '_')   // Windows 非法文件名字符
            .replace(/\s+/g, '_')            // 空格 → 下划线
            .replace(/\.+$/, '')             // 去除末尾点(否则会被当成无扩展名)
            .slice(0, 80);                   // 限制长度，避免文件名过长
    }

    // --- 来自 app.js:L23777-L23798 ---
    function saveCodeSnippet() {
        const code = document.getElementById('code-editor').value.trim();
        if (!code) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">代码编辑器为空，请先输入代码</span>', true) : App.showToast('代码编辑器为空，请先输入代码');
            return;
        }

        // 自动生成建议名称
        const nameSuggestion = generateSnippetNameSuggestion(code);

        // 描述字段默认为空，由用户自行填写
        const autoDescription = '';

        // 清空对话框
        document.getElementById('snippet-edit-id').value = '';
        document.getElementById('snippet-input-name').value = nameSuggestion;
        document.getElementById('snippet-input-desc').value = autoDescription;
        document.getElementById('snippet-input-code').value = code;

        // 显示对话框
        document.getElementById('snippet-save-dialog').style.display = 'flex';
    }

    // --- 来自 app.js:L23801-L23803 ---
    function closeSnippetDialog() {
        document.getElementById('snippet-save-dialog').style.display = 'none';
    }

    // --- 来自 app.js:L23806-L23895 ---
    async function confirmSaveSnippet() {
        const id = document.getElementById('snippet-edit-id').value;
        const name = document.getElementById('snippet-input-name').value.trim();
        const description = document.getElementById('snippet-input-desc').value.trim();
        const code = document.getElementById('snippet-input-code').value.trim();
        const saveAsSkill = document.getElementById('snippet-save-as-skill').checked;

        if (!name) {
            App.showToast('请输入代码片段名称');
            return;
        }

        if (!code) {
            App.showToast('代码不能为空');
            return;
        }

        // 处理"保存为 Skill"的勾选变化
        const oldSnippet = id ? App.codeSnippets.find(s => s.id === id) : null;
        const wasSavedAsSkill = oldSnippet?.savedAsSkill;
        ChatLog.info('SKILL', '[SKILL] 📝 保存片段: name=', name, 'saveAsSkill=', saveAsSkill, 'wasSavedAsSkill=', wasSavedAsSkill, 'oldSnippet.skillFileName=', oldSnippet?.skillFileName);

        if (id) {
            // 编辑现有代码片段
            const index = App.codeSnippets.findIndex(s => s.id === id);
            if (index !== -1) {
                App.codeSnippets[index] = {
                    ...App.codeSnippets[index],
                    name,
                    description,
                    code,
                    savedAsSkill: saveAsSkill,  // 保存勾选状态
                    timestamp: Date.now()
                };
                App.showToast('代码片段已更新');
            }
        } else {
            // 添加新代码片段
            const newSnippet = {
                id: 'snippet_' + Date.now(),
                name,
                description,
                code,
                savedAsSkill: saveAsSkill,  // 记录是否保存为 Skill
                timestamp: Date.now()
            };
            App.codeSnippets.push(newSnippet);
            App.showToast('代码片段已保存');
        }

        // 处理 CustomSkill 同步逻辑
        if (saveAsSkill) {
            // 勾选了"保存为 Skill" → 创建/更新 CustomSkill
            try {
                await saveSnippetAsSkill(name, description, code);
                App.showToast('已同时保存为 Skill（可在快捷指令中调用）');

                if (typeof App.loadSkillsAsCommands === 'function') {
                    await App.loadSkillsAsCommands(true);
                }
            } catch (error) {
                ChatLog.error('SKILL', '保存为 Skill 失败:', error);
                App.showToast('保存为 Skill 失败: ' + error.message);
            }
        } else if (wasSavedAsSkill && !saveAsSkill) {
            try {
                if (getBridge()) {
                    const skillNameToDelete = 'Snippet_' + sanitizeSkillFileName(name);
                    ChatLog.info('SKILL', '[SKILL] 🗑️ 取消勾选，直接删除 CustomSkill:', skillNameToDelete);

                    const result = await callBridgeMethod('DeleteSkill', [skillNameToDelete]);
                    const resultObj = JSON.parse(result);
                    ChatLog.info('SKILL', '[SKILL] DeleteSkill 返回:', result);

                    if (resultObj.success) {
                        ChatLog.info('SKILL', '[SKILL] ✅ 已移除 CustomSkill:', skillNameToDelete);
                        App.showToast('已从快捷指令中移除');

                        if (typeof App.loadSkillsAsCommands === 'function') {
                            await App.loadSkillsAsCommands(true);
                        }
                    } else {
                        ChatLog.warn('SKILL', '[SKILL] ⚠️ DeleteSkill 返回失败:', resultObj.error);
                    }
                }
            } catch (e) {
                ChatLog.warn('SKILL', '移除 CustomSkill 失败（非致命）:', e.message);
            }
        }

        saveCodeSnippetsToStorage();
        renderCodeLibrary();
        closeSnippetDialog();
    }

    // --- 来自 app.js:L23898-L23941 ---
    async function saveSnippetAsSkill(name, description, code) {
        // 从描述中提取关键词
        const keywords = extractKeywordsFromDescription(name, description);

        // 构建 Skill 对象
        const skill = {
            name: 'Snippet_' + sanitizeSkillFileName(name),
            displayName: name,
            description: description || name,
            keywords: keywords,
            templateCode: code,
            examples: [name],
            isGeneric: false,
            needsConfig: false,
            configFields: [],
            processingMode: 0,
            supportsImageInput: false,
            supportsFileInput: false,
            outputFormat: 'text',
            version: 1,
            isCustom: true,
            createdDate: new Date().toISOString(),
            lastModifiedDate: new Date().toISOString(),
            createdBy: 'User',
            usageCount: 0,
            tags: ['代码片段', '自定义']
        };

        // 调用后端保存 Skill（Web 版通过 Local Bridge HTTP API）
        const result = await callBridgeMethod('SaveSkill', [JSON.stringify(skill)]);
        const resultObj = JSON.parse(result);
        if (!resultObj.success) {
            throw new Error(resultObj.error || '保存失败');
        }

        if (typeof App.loadSkillsAsCommands === 'function') {
            await App.loadSkillsAsCommands(true);
        }
        if (typeof ChatLog !== 'undefined' && ChatLog.info) {
            ChatLog.info('SKILL', `✅ 自定义 Skill 已保存: ${skill.name}`);
        }
    }

    // --- 来自 app.js:L23944-L23973 ---
    function extractKeywordsFromDescription(description) {
        const keywords = [];
        const text = description.toLowerCase();

        // 常见操作关键词
        const commonKeywords = {
            '拆分': ['拆分', '分割', '分开', '导出'],
            '合并': ['合并', '汇总', '整合', '汇总'],
            '复制': ['复制', '拷贝', '克隆'],
            '删除': ['删除', '移除', '清除', '清空'],
            '插入': ['插入', '添加', '新建'],
            '格式化': ['格式', '样式', '美化'],
            '计算': ['计算', '求和', '统计', '汇总'],
            '排序': ['排序', '排列', '顺序'],
            '筛选': ['筛选', '过滤', '查找']
        };

        for (const [key, variants] of Object.entries(commonKeywords)) {
            if (variants.some(v => text.includes(v))) {
                keywords.push(key);
            }
        }

        // 如果没有提取到关键词，使用描述中的主要词汇
        if (keywords.length === 0) {
            const words = text.split(/\s+/).filter(w => w.length > 1);
            if (words.length > 0) {
                keywords.push(words[0]);
            }
        }

        return keywords;
    }

// ==================== 代码库模块导出 ====================

    App.generateCodeFingerprint = generateCodeFingerprint;
    App.calculateSimilarity = calculateSimilarity;
    App.findSimilarCodeInLibrary = findSimilarCodeInLibrary;
    App.findAllSimilarSnippets = findAllSimilarSnippets;
    App.addCacheHitMessage = addCacheHitMessage;
    App.loadSnippetToEditorAndExecute = loadSnippetToEditorAndExecute;
    App.showSimilarSnippetsDialog = showSimilarSnippetsDialog;
    App.loadCodeSnippets = loadCodeSnippets;
    App.saveCodeSnippetsToStorage = saveCodeSnippetsToStorage;
    App.renderCodeLibrary = renderCodeLibrary;
    App.showCodeLibrary = showCodeLibrary;
    App.closeCodeLibrary = closeCodeLibrary;
    App.loadSnippetFromLibrary = loadSnippetFromLibrary;
    App.editSnippetFromLibrary = editSnippetFromLibrary;
    App.deleteSnippetFromLibrary = deleteSnippetFromLibrary;
    App.extractCodeTags = extractCodeTags;
    App.generateSnippetNameSuggestion = generateSnippetNameSuggestion;
    App.saveCodeSnippet = saveCodeSnippet;
    App.closeSnippetDialog = closeSnippetDialog;
    App.confirmSaveSnippet = confirmSaveSnippet;
    App.saveSnippetAsSkill = saveSnippetAsSkill;
    App.extractKeywordsFromDescription = extractKeywordsFromDescription;

    // 兼容 index.html 中 onclick 直接调用
    window.saveCodeSnippet = saveCodeSnippet;
    window.closeCodeLibrary = closeCodeLibrary;

    App.CODE_FINGERPRINTS = CODE_FINGERPRINTS;

})();
