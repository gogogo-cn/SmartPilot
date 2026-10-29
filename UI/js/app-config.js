/**
 * ============================================================
 * app-config.js — Configuration System
 * ============================================================
 * 功能：配置系统模块，管理用户设置、提示词模板、卡片类型与持久化存储。
 *
 * 核心类/函数：
 *   - loadConfig() — 加载配置
 *   - saveConfig() — 保存配置
 *   - getPromptConfig() — 获取提示词配置
 *   - resetConfig() — 重置配置
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 *   - app-utils.js
 * ============================================================
 */

(function() {
    'use strict';

    // HTML 转义（防止 XSS）
    function escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.appendChild(document.createTextNode(text));
        return div.innerHTML;
    }

    // --- 来自 app.js:L199-L239 ---
    /**
     * 加载 Prompt 配置文件（【v=56重构】支持Markdown格式）
     * @returns {Promise<boolean>} 是否成功加载
     */
    async function loadPromptConfig() {
        try {
            ChatLog.info('CONFIG', '[PromptConfig] 正在加载 Prompt 配置 (prompts.md中');
            const response = await fetch('js/prompts.md');
            if (!response.ok) {
                throw new Error(`HTTP ${response.status}: ${response.statusText}`);
            }

            const markdownText = await response.text();

            App.promptConfig = parsePromptsFromMarkdown(markdownText);

            // 非 Excel 宿主（Word / PowerPoint）：把 Excel 专属措辞/工具名改写为当前宿主。
            // Excel 宿主下 App.adaptPromptsForHost 为 no-op，零影响。
            try {
                if (typeof App.adaptPromptsForHost === 'function') {
                    App.adaptPromptsForHost(App.promptConfig.prompts);
                }
            } catch (adaptErr) {
                ChatLog.warn('CONFIG', '[PromptConfig] 宿主提示词适配失败:', adaptErr && adaptErr.message);
            }

            if (App.promptConfig && Object.keys(App.promptConfig.prompts).length > 0) {
                App.promptConfigLoaded = true;
                ChatLog.info('CONFIG', '[PromptConfig] ✅ Prompts.md 配置加载成功，版本:', App.promptConfig.version);
                ChatLog.info('CONFIG', '[PromptConfig] 可用Prompt列表:', Object.keys(App.promptConfig.prompts));
                return true;
            } else {
                throw new Error('解析Markdown后无有效Prompt');
            }
        } catch (error) {
            ChatLog.warn('CONFIG', '[PromptConfig] ⚠️ Prompts.md 加载失败，使用默认配置:', error.message);
            // 【降级】如果MD文件加载失败，尝试JSON格式
            try {
                const jsonResponse = await fetch('js/prompts.json');
                if (jsonResponse.ok) {
                    App.promptConfig = await jsonResponse.json();
                    try {
                        if (typeof App.adaptPromptsForHost === 'function') {
                            App.adaptPromptsForHost(App.promptConfig.prompts);
                        }
                    } catch (adaptErr) {
                        ChatLog.warn('CONFIG', '[PromptConfig] 宿主提示词适配失败:', adaptErr && adaptErr.message);
                    }
                    App.promptConfigLoaded = true;
                    ChatLog.info('CONFIG', '[PromptConfig] ✅ 降级使用 prompts.json 成功');
                    return true;
                }
            } catch (jsonError) {
                ChatLog.warn('CONFIG', '[PromptConfig] ⚠️ prompts.json 也加载失败');
            }

            // 最终降级：使用硬编码默认值
            App.promptConfig = { prompts: App.DEFAULT_PROMPTS, skillInfo: { template: App.DEFAULT_SKILL_INFO }, version: 'default' };
            App.promptConfigLoaded = false;
            return false;
        }
    }

    // --- 来自 app.js:L246-L283 ---
    /**
     * 【v=56新增】从Markdown文本解析Prompt配置
     * @param {string} markdown - Markdown文本
     * @returns {Object} 解析后的配置对象
     */
    function parsePromptsFromMarkdown(markdown) {
        const config = {
            version: '2.1.0',
            prompts: {},
            skillInfo: {
                template: `- 智能分析(AI): 使用 AI 对图片/文件进行智能分析和总结\n- VBA 助手: 根据自然语言描述生成 Excel VBA 代码`
            }
        };

        // 使用正则提取每个 ### 标题下的代码块内容
        const promptRegex = /### \d+\.\s*(.+?)\s*[^\n]*\n\n\*\*ID\*\*:\s*`(.+?)`/g;
        let match;

        while ((match = promptRegex.exec(markdown)) !== null) {
            const title = match[1].trim();
            const id = match[2].trim();

            // 找到对应的 System Prompt Content 代码块
            const contentStartRegex = new RegExp(`### \\d+\\.\\s*${escapeRegExp(title)}[\\s\\S]*?#### System Prompt Content:\\n\\n\`\`\`\n([\\s\\S]*?)\`\`\``);
            const contentMatch = markdown.match(contentStartRegex);

            if (contentMatch) {
                config.prompts[id] = {
                    id: id,
                    name: title,
                    content: contentMatch[1].trim()
                };
            }
        }

        // 提取版本号
        const versionMatch = markdown.match(/\*\*版本\*\*:([\d.]+)/);
        if (versionMatch) {
            config.version = versionMatch[1];
        }

        return config;
    }

    // --- 来自 app.js:L288-L290 ---
    /**
     * 转义正则表达式特殊字符
     */
    function escapeRegExp(string) {
        return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

    // --- 来自 app.js:L298-L308 ---
    /**
     * 获取指定 Prompt
     * @param {string} promptId - Prompt ID
     * @param {Object} variables - 要替换的变量
     * @returns {string} Prompt 内容
     */
    function getPrompt(promptId, variables = {}) {
        const prompts = App.promptConfig?.prompts || App.DEFAULT_PROMPTS;
        let content = prompts[promptId]?.content || prompts.default?.content || App.DEFAULT_PROMPTS.default.content;

        // 替换变量
        for (const [key, value] of Object.entries(variables)) {
            content = content.replace(new RegExp(`{${key}}`, 'g'), value);
        }

        return content;
    }

    // --- 来自 app.js:L314-L316 ---
    /**
     * 获取技能信息（优先使用缓存的真实技能数据）
     * @returns {string} 技能信息文本
     */
    // 后台实时拉取开关：缓存为空时只触发一次，避免每轮 sendMessage 都去轰击桥接
    let _skillInfoLiveFetchStarted = false;

    function getSkillInfo() {
        const cached = App.cachedSkills || window.cachedSkills;
        if (cached && cached.length > 0) {
            const seen = new Set();
            const uniqueSkills = cached.filter(s => {
                const key = (s.displayName || s.name || '').toLowerCase().trim();
                if (seen.has(key)) return false;
                seen.add(key);
                return true;
            });
            const systemSkills = uniqueSkills.filter(s => !(s.isCustom || s.IsCustom));
            const customSkills = uniqueSkills.filter(s => s.isCustom || s.IsCustom);
            const summary = `【技能统计】共 ${uniqueSkills.length} 个技能（系统内置 ${systemSkills.length} 个，自定义 ${customSkills.length} 个）\n\n`;
            const lines = uniqueSkills.map(s => {
                const execName = s.name || s.Name || '未命名';
                const displayName = s.displayName || s.DisplayName || execName;
                const desc = s.description || s.Description || '';
                const type = (s.isCustom || s.IsCustom) ? '自定义' : '系统';
                const keywords = (s.keywords || s.Keywords || []).join(', ');
                return `- 名称: ${displayName} | 类型: ${type} | 描述: ${desc || '无描述'} | 关键词: ${keywords || '无'}`;
            });
            if (cached.length !== uniqueSkills.length) {
                ChatLog.warn('CONFIG', `[SkillInfo] ⚠️ 检测到重复技能，已去重: ${cached.length} -> ${uniqueSkills.length}`);
            }
            return summary + lines.join('\n');
        }
        const fallback = App.promptConfig?.skillInfo?.template || App.DEFAULT_SKILL_INFO;
        ChatLog.warn('CONFIG', '[SkillInfo] ⚠️ 缓存为空，触发后台异步加载（本轮用 fallback）');
        // 【修复】曾经在此同步调用 hostObjects.sync.excelBridge.GetAvailableSkills()，而访问 .sync
        // 代理会触发 WebView2 注入的投影脚本抛出未捕获的 "e is not a function" Promise 拒绝
        // （实测每轮刷屏、GlobalErrorBoundary 限流丢弃）。
        // 现改为：立即返回 fallback；仅在尚未触发过时，后台异步调用已自带 try/catch 的
        // loadSkillsForLLM 拉取真实技能，下一次调用缓存命中即返回权威数据。绝不触碰 .sync 代理。
        if (!_skillInfoLiveFetchStarted && typeof loadSkillsForLLM === 'function') {
            _skillInfoLiveFetchStarted = true;
            loadSkillsForLLM(true)
                .then(skills => {
                    if (Array.isArray(skills) && skills.length > 0) {
                        ChatLog.info('CONFIG', `[SkillInfo] ✅ 后台加载成功，共 ${skills.length} 个技能`);
                    } else {
                        ChatLog.warn('CONFIG', '[SkillInfo] ⚠️ 后台加载未返回技能（桥接可能无 GetAvailableSkills）');
                    }
                })
                .catch(err => {
                    ChatLog.warn('CONFIG', '[SkillInfo] ⚠️ 后台加载失败:', err && err.message);
                });
        }
        return fallback;
    }

    // --- 来自 app.js:L322-L326 ---
    /**
     * 重新加载 Prompt 配置
     * @returns {Promise<boolean>}
     */
    async function reloadPromptConfig() {
        ChatLog.info('CONFIG', '[PromptConfig] 重新加载配置...');
        App.promptConfigLoaded = false;
        return await loadPromptConfig();
    }

    // --- 来自 app.js:L332-L338 ---
    /**
     * 获取 Prompt 配置状态
     * @returns {Object}
     */
    function getPromptConfigStatus() {
        return {
            loaded: App.promptConfigLoaded,
            version: App.promptConfig?.version || 'default',
            availablePrompts: Object.keys(App.promptConfig?.prompts || App.DEFAULT_PROMPTS)
        };
    }

    // --- 来自 app.js:L340-L390 ---
    // 解析模型识别响应（容错）并安全渲染识别结果
    function parseModelRecognitionResponse(text) {
        if (!text) return null;
        // 1) 直接尝试 JSON
        try {
            return JSON.parse(text);
        } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('CONFIG', 'JSON.parse 直接解析失败:', e && e.message); }

        // 2) 提取第一个 JSON 对象块
        const jsonMatch = text.match(/\{[\s\S]*\}/);
        if (jsonMatch) {
            try { return JSON.parse(jsonMatch[0]); } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('CONFIG', 'JSON 块解析失败:', e && e.message); }
        }

        // 3) 把文本当 HTML 解析，查找 data-json 或表格
        try {
            const parser = new DOMParser();
            const doc = parser.parseFromString(text, 'text/html');

            // 优先查找带 data-json 的元素
            const el = doc.querySelector('[data-json]');
            if (el) {
                const attr = el.getAttribute('data-json');
                if (attr) {
                    try { return JSON.parse(attr); } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('CONFIG', 'data-json 解析失败:', e && e.message); }
                    const firstBrace = attr.indexOf('{');
                    const lastBrace = attr.lastIndexOf('}');
                    if (firstBrace >= 0 && lastBrace > firstBrace) {
                        try { return JSON.parse(attr.slice(firstBrace, lastBrace + 1)); } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('CONFIG', '切片 JSON 解析失败:', e && e.message); }
                    }
                }
            }

            // 查找带有识别表格的节点
            const table = doc.querySelector('.recognition-table') || doc.querySelector('table');
            if (table) {
                const obj = {};
                table.querySelectorAll('tr').forEach(tr => {
                    const th = tr.querySelector('th');
                    const td = tr.querySelector('td');
                    if (th && td) {
                        const key = th.textContent.trim();
                        const val = td.textContent.trim();
                        if (key) obj[key] = val;
                    }
                });
                if (Object.keys(obj).length) return obj;
            }
        } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('CONFIG', 'HTML 解析识别结果失败:', e && e.message); }

        return null;
    }

    // --- 来自 app.js:L392-L460 ---
    function buildRecognitionCardFromObject(data) {
        const container = document.createElement('div');
        container.className = 'recognition-result-container';

        const h3 = document.createElement('h3');
        h3.textContent = data?.title || '识别结果';
        container.appendChild(h3);

        const card = document.createElement('div');
        card.className = 'recognition-result-card';

        const header = document.createElement('div');
        header.className = 'recognition-header';
        const titleWrap = document.createElement('div');
        titleWrap.className = 'recognition-title';
        const icon = document.createElement('span'); icon.className = 'recognition-title-icon'; icon.textContent = '🔍';
        const titleText = document.createElement('span'); titleText.textContent = data?.type || '表格';
        titleWrap.appendChild(icon);
        titleWrap.appendChild(titleText);

        const actions = document.createElement('div'); actions.className = 'recognition-actions';
        // 填入 Excel：仅 Excel 宿主显示（Word/PPT 无填表能力）
        const _isExcelHost = !(window.SmartPilotHost) || !!window.SmartPilotHost.isExcel;
        if (_isExcelHost && data && Object.keys(data).length) {
            const btn = document.createElement('button');
            btn.className = 'recognition-action-btn';
            btn.title = '填入 Excel';
            btn.style.background = '#4caf50'; btn.style.color = 'white';
            btn.textContent = '📊';
            try { btn.setAttribute('data-json', JSON.stringify(data)); } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('CONFIG', 'setAttribute data-json 失败:', e && e.message); }
            btn.setAttribute('data-type', data.type || '识别');
            btn.onclick = function() { if (typeof fillRecognitionDataToExcel === 'function') fillRecognitionDataToExcel(this); };
            actions.appendChild(btn);
        }

        // 复制按钮
        const copyBtn = document.createElement('button'); copyBtn.className = 'recognition-action-btn'; copyBtn.textContent='📋'; copyBtn.title='复制数据'; copyBtn.onclick = function(){ copyRecognitionResult(this); };
        const dlBtn = document.createElement('button'); dlBtn.className = 'recognition-action-btn'; dlBtn.textContent='⬇️'; dlBtn.title='下载数据'; dlBtn.onclick = function(){ downloadRecognitionResult(this); };
        actions.appendChild(copyBtn); actions.appendChild(dlBtn);

        header.appendChild(titleWrap);
        header.appendChild(actions);
        card.appendChild(header);

        const table = document.createElement('table'); table.className = 'recognition-table';
        const tbody = document.createElement('tbody');

        // 按键值顺序渲染字段
        for (const key of Object.keys(data)) {
            const tr = document.createElement('tr');
            const th = document.createElement('th'); th.textContent = key;
            const td = document.createElement('td');
            // 对于对象或数组，保持 JSON 字符串
            const val = (typeof data[key] === 'object') ? JSON.stringify(data[key], null, 2) : String(data[key]);
            td.textContent = val;
            tr.appendChild(th); tr.appendChild(td);
            tbody.appendChild(tr);
        }

        table.appendChild(tbody);
        card.appendChild(table);

        const footer = document.createElement('div'); footer.className = 'recognition-footer';
        const p = document.createElement('p'); p.textContent = '说明：该结果为系统识别，供参考，请核对原始信息';
        footer.appendChild(p);
        card.appendChild(footer);

        container.appendChild(card);
        return container;
    }

    // --- 来自 app.js:L462-L507 ---
    function processElementForRecognition(el) {
        if (!el || el.nodeType !== Node.ELEMENT_NODE) return;

        // 如果元素本身带 data-json 属性，尝试修正并返回
        if (el.hasAttribute && el.hasAttribute('data-json')) {
            const raw = el.getAttribute('data-json');
            const parsed = parseModelRecognitionResponse(raw);
            if (parsed) {
                try { el.setAttribute('data-json', JSON.stringify(parsed)); } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('CONFIG', 'setAttribute data-json 失败:', e && e.message); }
            }
        }

        // 如果是识别结果容器或包含识别表格，尝试替换为安全渲染
        if (el.classList && el.classList.contains('recognition-result-container')) {
            const parsed = parseModelRecognitionResponse(el.innerHTML) || parseModelRecognitionResponse(el.textContent);
            if (parsed) {
                const safe = buildRecognitionCardFromObject(parsed);
                el.parentNode.replaceChild(safe, el);
                return;
            }
        }

        // 查找内部可能包含 data-json 或 table 的子节点
        const dataAttrEl = el.querySelector && el.querySelector('[data-json]');
        if (dataAttrEl) {
            const raw = dataAttrEl.getAttribute('data-json');
            const parsed = parseModelRecognitionResponse(raw);
            if (parsed) {
                const safe = buildRecognitionCardFromObject(parsed);
                dataAttrEl.parentNode.replaceChild(safe, dataAttrEl);
                return;
            }
        }

        const tableEl = el.querySelector && el.querySelector('table.recognition-table');
        if (tableEl) {
            const parsed = parseModelRecognitionResponse(el.innerHTML) || parseModelRecognitionResponse(el.textContent);
            if (parsed) {
                const safe = buildRecognitionCardFromObject(parsed);
                // replace the closest ancestor container if exists
                const container = el.closest('.recognition-result-container') || tableEl.parentNode;
                container.parentNode.replaceChild(safe, container);
                return;
            }
        }
    }

    // ==================== 卡片类型配置函数 ====================

    // --- 来自 app.js:L791-L838 ---
    /**
     * 从后端 API 加载卡片类型配置（从 Skill 定义中读取）
     * @returns {Promise<boolean>} 是否成功加载
     */
    function _bridgeErrorMessage(err) {
        if (!err) return 'unknown';
        if (typeof err === 'string') return err;
        if (err.message) return err.message;
        // WebView2 hostObjects 拒绝形态：{ parameters: { error, kind }, callId, ... }
        try {
            const p = err.parameters;
            if (p && p.error) return String(p.error);
        } catch (_) { /* ignore */ }
        try { return JSON.stringify(err); } catch (_) { return String(err); }
    }

    async function loadCardTypesConfig() {
        try {
            ChatLog.info('CONFIG', '[Config] 正在从后端加载卡片类型配置...');

            // 优先 BridgeAccessor：吞掉 WebView2 的 e/t is not a function / 0x80070490，避免未处理拒绝
            let result;
            if (window.BridgeAccessor && typeof window.BridgeAccessor.callAsync === 'function') {
                result = await window.BridgeAccessor.callAsync('GetRecognitionTypes');
            } else {
                const bridge = window.excelBridge ||
                               (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge);
                if (!bridge) {
                    throw new Error('excelBridge 未初始化，稍后重试');
                }
                result = await bridge.GetRecognitionTypes();
            }
            if (result && typeof result === 'object') {
                // callAsync 偶发直接返回对象；统一成可 parse 的字符串
                result = JSON.stringify(result);
            }
            const response = JSON.parse(result);

            if (!response.success) {
                throw new Error(response.error || '后端返回错误');
            }

            if (response.recognitionTypes && response.recognitionTypes.length > 0) {
                // 将后端返回的配置转换为前端格式
                const backendConfig = {};
                for (const typeInfo of response.recognitionTypes) {
                    backendConfig[typeInfo.id] = convertSkillConfigToImageTypeConfig(typeInfo);
                }

                // 合并后端配置和默认配置（后端配置优先）
                App.IMAGE_TYPE_CONFIG = { ...App.DEFAULT_IMAGE_TYPE_CONFIG, ...backendConfig };
                App.cardTypesConfigLoaded = true;
                ChatLog.info('CONFIG', '[Config] 卡片类型配置加载成功:', Object.keys(App.IMAGE_TYPE_CONFIG));
                ChatLog.info('CONFIG', '[Config] 从后端加载了', response.recognitionTypes.length, '种识别类型');
                return true;
            } else {
                // 后端返回成功但没有数据，使用默认通用配置（已简化，不再需要后端返回识别类型）
                ChatLog.info('CONFIG', '[Config] 使用默认通用图片分析配置');
                App.IMAGE_TYPE_CONFIG = { ...App.DEFAULT_IMAGE_TYPE_CONFIG };
                App.cardTypesConfigLoaded = true;
                return true;
            }
        } catch (error) {
            ChatLog.warn('CONFIG', '[Config] 从后端加载配置失败，使用默认通用配置:', _bridgeErrorMessage(error));
            // 使用默认配置
            App.IMAGE_TYPE_CONFIG = { ...App.DEFAULT_IMAGE_TYPE_CONFIG };
            App.cardTypesConfigLoaded = true;
            return true;
        }
    }

    // --- 来自 app.js:L843-L876 ---
    /**
     * 将后端 Skill 配置转换为前端 IMAGE_TYPE_CONFIG 格式
     */
    function convertSkillConfigToImageTypeConfig(typeInfo) {
        const config = {
            keywords: typeInfo.keywords || [],
            icon: typeInfo.icon || '📄',
            displayName: typeInfo.displayName || '识别' + typeInfo.id + '信息',
            backendType: typeInfo.id,
            skillName: typeInfo.name,
            extractFields: typeInfo.extractFields || [],
            defaultAction: (!(window.SmartPilotHost) || window.SmartPilotHost.isExcel) ? '填入 Excel' : '复制结果',
            defaultWorksheet: typeInfo.excelConfig?.defaultWorksheet || typeInfo.id + '信息',
            primaryAction: (!(window.SmartPilotHost) || window.SmartPilotHost.isExcel)
                ? { id: 'fill-excel', label: '📊 填入 Excel', action: 'fillExcelAuto' }
                : { id: 'copy-result', label: '📋 复制结果', action: 'copyResult' },
            secondaryActions: [],
            fieldList: typeInfo.extractFields || []
        };

        // 添加字段映射
        if (typeInfo.excelConfig?.fieldMapping) {
            config.fieldMapping = {};
            for (const [fieldName, mapping] of Object.entries(typeInfo.excelConfig.fieldMapping)) {
                config.fieldMapping[fieldName] = mapping.column;
            }
        }

        // 添加字段名映射
        if (typeInfo.excelConfig?.fieldNameMapping) {
            config.fieldNameMapping = typeInfo.excelConfig.fieldNameMapping;
        }

        return config;
    }

    // --- 来自 app.js:L882-L886 ---
    /**
     * 重新加载卡片类型配置（用于配置更新后）
     * @returns {Promise<boolean>}
     */
    async function reloadCardTypesConfig() {
        ChatLog.info('CONFIG', '[Config] 重新从后端加载配置...');
        App.cardTypesConfigLoaded = false;
        return await loadCardTypesConfig();
    }

    // --- 来自 app.js:L892-L898 ---
    /**
     * 获取当前配置状态
     * @returns {Object}
     */
    function getCardTypesConfigStatus() {
        return {
            loaded: App.cardTypesConfigLoaded,
            types: Object.keys(App.IMAGE_TYPE_CONFIG),
            config: App.IMAGE_TYPE_CONFIG
        };
    }

    // --- 来自 app.js:L905-L907 ---
    /**
     * 获取指定类型的配置
     * @param {string} type - 卡片类型
     * @returns {Object|null}
     */
    function getCardTypeConfig(type) {
        return App.IMAGE_TYPE_CONFIG[type] || null;
    }

    // --- 来自 app.js:L913-L915 ---
    /**
     * 获取所有可用的卡片类型
     * @returns {string[]}
     */
    function getAvailableCardTypes() {
        return Object.keys(App.IMAGE_TYPE_CONFIG);
    }

    // 页面加载时自动加载配置
    let cardTypesConfigPromise = null;

    (async function initCardTypesConfig() {
        cardTypesConfigPromise = loadCardTypesConfig();
        await cardTypesConfigPromise;
    })().catch(function (e) {
        // 双保险：避免任何未捕获拒绝冒泡到 bootstrap 的 unhandledrejection
        try {
            ChatLog.warn('CONFIG', '[Config] initCardTypesConfig 异常（已忽略）:', _bridgeErrorMessage(e));
        } catch (_) { /* ignore */ }
        App.IMAGE_TYPE_CONFIG = { ...App.DEFAULT_IMAGE_TYPE_CONFIG };
        App.cardTypesConfigLoaded = true;
    });

    // --- 来自 app.js:L929-L933 ---
    /**
     * 等待卡片类型配置加载完成
     * @returns {Promise<void>}
     */
    async function waitForCardTypesConfig() {
        if (cardTypesConfigPromise) {
            await cardTypesConfigPromise;
        }
    }

    // --- 来自 app.js:L936-L942 ---
    // 获取字段名映射（英文名 -> 中文名）
    function getFieldNameMapping(imageType) {
        const config = App.IMAGE_TYPE_CONFIG[imageType];
        if (config && config.fieldNameMapping) {
            return config.fieldNameMapping;
        }
        return null;
    }

    // --- 来自 app.js:L945-L951 ---
    // 获取字段列表（英文名）
    function getFieldList(imageType) {
        const config = App.IMAGE_TYPE_CONFIG[imageType];
        if (config && config.fieldList) {
            return config.fieldList;
        }
        return [];
    }

    // --- 来自 app.js:L954-L980 ---
    // 判断两个中文名是否表示同一个字段
    function isSameField(name1, name2) {
        // 去除空格和特殊字符进行比较
        const normalize = (name) => name.replace(/[\s\-\_\*]/g, '');

        const n1 = normalize(name1);
        const n2 = normalize(name2);

        // 完全相同
        if (n1 === n2) return true;

        // 身份证号码变体
        const idNumberPatterns = ['身份证号码', '公民身份号码', '身份证号', '号码', 'idnumber', 'idnum'];
        if (idNumberPatterns.includes(n1) && idNumberPatterns.includes(n2)) return true;

        // 有效期限变体
        const validDatePatterns = ['有效期限', '有效期', '有效日期', 'validdate', 'valid'];
        if (validDatePatterns.includes(n1) && validDatePatterns.includes(n2)) return true;

        // 签发机关变体
        const issuePatterns = ['签发机关', '签发单位', 'issueauthority'];
        if (issuePatterns.includes(n1) && issuePatterns.includes(n2)) return true;

        // 一个名字包含另一个（处理类似"性别"和"性"的情况）
        if (n1.includes(n2) || n2.includes(n1)) return true;

        return false;
    }

    // --- 来自 app.js:L983-L1048 ---
    // 转换字段名（英文名 → 中文名）
    function transformFieldNames(structuredData, imageType) {
        const mapping = getFieldNameMapping(imageType);
        const fieldList = getFieldList(imageType);
        const result = {};

        // 创建反向映射：中文名 -> 英文名
        const reverseMapping = {};
        if (mapping) {
            for (const [enName, zhName] of Object.entries(mapping)) {
                reverseMapping[zhName] = enName;
            }
        }

        // 检查结果中是否已经有相似的字段
        const hasSimilarInResult = (newKey) => {
            for (const existingKey of Object.keys(result)) {
                if (isSameField(newKey, existingKey)) {
                    return true;
                }
            }
            return false;
        };

        if (mapping) {
            for (const [key, value] of Object.entries(structuredData)) {
                if (fieldList.includes(key)) {
                    // 英文名，需要映射到中文名
                    const chineseName = mapping[key] || key;
                    // 检查是否已经有相似的字段
                    if (!hasSimilarInResult(chineseName)) {
                        result[chineseName] = value;
                    } else {
                        ChatLog.info('CONFIG', `跳过重复字段: ${key} -> ${chineseName} (已有相似字段)`);
                    }
                } else if (Object.values(mapping).includes(key)) {
                    // 已经是中文名，且在映射中存在
                    // 检查是否已经有相似的字段
                    if (!hasSimilarInResult(key)) {
                        result[key] = value;
                    } else {
                        ChatLog.info('CONFIG', `跳过重复字段: ${key} (已有相似字段)`);
                    }
                } else {
                    // 中文名或其他字段
                    // 检查这个中文名是否已经被对应的英文名映射过
                    const correspondingEnName = reverseMapping[key];
                    if (correspondingEnName && structuredData.hasOwnProperty(correspondingEnName)) {
                        // 如果已经有对应的英文名数据，跳过这个中文名，避免重复
                        ChatLog.info('CONFIG', `跳过重复字段: ${key} (已有对应英文名: ${correspondingEnName})`);
                    } else {
                        // 检查是否已经有相似的字段名
                        if (!hasSimilarInResult(key)) {
                            // 没有对应的英文名，也没有相似字段，保留中文名
                            result[key] = value;
                        } else {
                            ChatLog.info('CONFIG', `跳过重复字段: ${key} (已有相似字段)`);
                        }
                    }
                }
            }
        } else {
            Object.assign(result, structuredData);
        }

        return result;
    }

    // --- 来自 app.js:L1051-L1057 ---
    // 获取表头顺序
    function getHeaderOrder(imageType) {
        const config = App.IMAGE_TYPE_CONFIG[imageType];
        if (config && config.headerOrder) {
            return config.headerOrder;
        }
        return [];
    }

    // --- 来自 app.js:L1065-L1194 ---
    /**
     * 格式化识别结果为卡片表格HTML
     * @param {Object} data - 识别结果数据
     * @param {string} imageType - 图片类型（身份证、名片等）
     * @returns {string} HTML字符串
     */
    function formatRecognitionResult(data, imageType) {
        ChatLog.info('CONFIG', '🔍 formatRecognitionResult 被调用:', { imageType, dataKeys: Object.keys(data) });
        if (!data || typeof data !== 'object') {
            ChatLog.warn('CONFIG', '⚠️ 数据无效:', data);
            return '<div style="color: #999; padding: 16px;">暂无识别数据</div>';
        }

        // 转换字段名（英文名转中文名）
        const transformedData = transformFieldNames(data, imageType);
        ChatLog.info('CONFIG', '🔍 转换后的数据:', transformedData);

        // 获取配置的字段顺序
        const headerOrder = getHeaderOrder(imageType);
        ChatLog.info('CONFIG', '🔍 表头顺序:', headerOrder);

        // 按顺序排列字段
        const orderedData = {};
        headerOrder.forEach(field => {
            if (transformedData.hasOwnProperty(field)) {
                orderedData[field] = transformedData[field];
                delete transformedData[field];
            }
        });
        // 添加剩余未排序的字段
        Object.assign(orderedData, transformedData);

        // 生成表格行
        let rowsHtml = '';
        Object.entries(orderedData).forEach(([key, value]) => {
            // 跳过空值和元数据字段
            if (value === undefined || value === null || value === '' ||
                ['imageType', 'image_type', 'description', 'structuredData',
                 'suggestedAction', 'suggested_action', 'extractedText',
                 'extracted_text', 'type', 'success', 'error', 'executionLog',
                 'totalExecutionTime'].includes(key)) {
                return;
            }

            // 对值进行处理
            let processedValue = String(value);

            // 处理 rawText 等特殊字段
            if (key === 'rawText' || key === 'extractedText' || key === 'description') {
                // 检查是否已经包含HTML标签（来自前端的格式化结果）
                if (processedValue.includes('<div') || processedValue.includes('<br')) {
                    // 已经包含HTML标签，直接显示（不再转义）
                    // 但需要确保其他潜在的危险标签被清理
                    processedValue = processedValue
                        .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
                        .replace(/javascript:/gi, '');
                } else {
                    // 纯文本，将换行符转换为 <br>
                    processedValue = processedValue
                        .replace(/\\r\\n/g, '<br>')
                        .replace(/\\n/g, '<br>')
                        .replace(/\r\n/g, '<br>')
                        .replace(/\r/g, '<br>')
                        .replace(/\n/g, '<br>');
                }
            }

            // 对于包含HTML的rawText，不再使用escapeHtml
            const escapedValue = (key === 'rawText' && (String(value).includes('<div') || String(value).includes('<br')))
                ? processedValue
                : escapeHtml(processedValue);
            rowsHtml += `
                <tr>
                    <th>${escapeHtml(key)}</th>
                    <td>${escapedValue}</td>
                </tr>
            `;
        });

        // 获取图片类型对应的图标
        const typeIcons = {
            '身份证': '',
            '名片': '📇',
            '发票': '🧾',
            '表格': '📊',
            '文档': '📝',
            '图片': '🖼️'
        };
        const icon = typeIcons[imageType] || '🔍';

        // 获取类型显示名称
        const typeNames = {
            '身份证': '身份证',
            '名片': '名片',
            '发票': '发票',
            '表格': '表格数据',
            '文档': '文档内容',
            '图片': '图片分析'
        };
        const typeName = typeNames[imageType] || '识别结果';

        // 生成完整的卡片HTML
        // 使用 data-json 属性存储 JSON，避免在 onclick 中直接嵌入 JSON 字符串
        // 这样可以避免 HTML/JS 解析器对转义字符的不同处理导致 JSON 损坏
        const dataJson = JSON.stringify(data);
        const html = `
            <div class="recognition-result-container">
                <h3>${typeName}提取</h3>
                <div class="recognition-result-card">
                    <div class="recognition-header">
                        <div class="recognition-title">
                            <span class="recognition-title-icon">${icon}</span>
                            <span>表格</span>
                        </div>
                        <div class="recognition-actions">
                            ${ ((!(window.SmartPilotHost) || window.SmartPilotHost.isExcel) && (imageType === '身份证' || imageType === '名片' || imageType === '发票')) ?
                                `<button class="recognition-action-btn" data-json="${escapeHtml(dataJson)}" data-type="${escapeHtml(imageType)}" onclick="fillRecognitionDataToExcel(this)" title="填入 Excel" style="background: #4caf50; color: white;">📊</button>` : '' }
                                <button class="recognition-action-btn" onclick="copyRecognitionResult(this)" title="复制数据">📋</button>
                                <button class="recognition-action-btn" onclick="downloadRecognitionResult(this)" title="下载数据">⬇️</button>
                        </div>
                    </div>
                    <table class="recognition-table">
                        <tbody>
                            ${rowsHtml}
                        </tbody>
                    </table>
                    <div class="recognition-footer">
                        <p>说明：该${imageType === '身份证' ? '身份证' : '识别结果'}为系统自动识别提取，仅供参考，请核对原始信息</p>
                    </div>
                </div>
            </div>
        `;
        ChatLog.info('CONFIG', '🔍 生成的HTML:', html.substring(0, 200) + '...');

        return html;
    }

    // ==================== 技能加载函数 ====================

    // --- 来自 app.js:L1624-L1647 ---
    // 为 LLM 加载 Skill 列表（全局函数，供多处调用）
    // 【修复】Web 版没有 WebView2 时，回退到 Local Bridge HTTP API 加载共享 Skill
    async function loadSkillsForLLM(forceRefresh = false) {
        // 修复：空数组 [] 是 truthy，会导致缓存命中但实际无数据
        // 只有当 cachedSkills 是非空数组时才使用缓存
        if (!forceRefresh && Array.isArray(App.cachedSkills) && App.cachedSkills.length > 0) return App.cachedSkills;

        try {
            const hasHostBridge = !!(window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge);
            const callBridge = async (method) => {
                // 统一走 BridgeAccessor：避免裸 host 拆方法 → e is not a function / 未处理拒绝
                if (window.BridgeAccessor && typeof window.BridgeAccessor.callAsync === 'function') {
                    return window.BridgeAccessor.callAsync(method);
                }
                if (hasHostBridge) {
                    return window.chrome.webview.hostObjects.excelBridge[method]();
                }
                return null;
            };
            // 真实 WebView2 或已挂本地桥代理时走桥；纯浏览器无 shim 时走下方 HTTP
            if (hasHostBridge) {
                // 【强制刷新】先通知后端重新加载 Markdown Skills
                if (forceRefresh) {
                    try {
                        ChatLog.info('SKILL', '[Skill加载] forceRefresh=true, 正在通知后端重新加载 Skills...');
                        const reloadResult = await callBridge('ReloadAllSkills');
                        const reloadParsed = typeof reloadResult === 'string' ? JSON.parse(reloadResult) : reloadResult;
                        ChatLog.info('SKILL', '[Skill加载] 后端刷新结果:', reloadParsed);
                    } catch (reloadErr) {
                        ChatLog.warn('SKILL', '[Skill加载] 后端刷新失败(可能接口不存在):', reloadErr && reloadErr.message);
                    }
                }

                const skillsJson = await callBridge('GetAvailableSkills');
                if (skillsJson == null) {
                    ChatLog.warn('SKILL', '[Skill加载] 桥未返回 GetAvailableSkills');
                } else {
                    const parsed = typeof skillsJson === 'string' ? JSON.parse(skillsJson) : skillsJson;
                    if (parsed && parsed.error === 'bridge-call-failed') {
                        ChatLog.warn('SKILL', '[Skill加载] 桥调用失败:', parsed.method, parsed.detail);
                    } else if (parsed && parsed.success && Array.isArray(parsed.skills)) {
                        App.cachedSkills = parsed.skills;
                        ChatLog.info('SKILL', `[Skill加载] 成功加载 ${parsed.skills.length} 个 Skills`);
                        if (typeof notifyCachedSkillsLoaded === 'function') {
                            notifyCachedSkillsLoaded();
                        }
                        return App.cachedSkills;
                    }
                }
            } else {
                // Web 版（浏览器环境）：通过 Local Bridge HTTP API 加载共享 Skill
                ChatLog.info('SKILL', '[Skill加载] Web 版未检测到 WebView2，回退到 HTTP API 加载共享 Skill...');
                const skillsUrl = window.ExcelBridgeConfig
                    ? window.ExcelBridgeConfig.withCacheBuster(window.ExcelBridgeConfig.skillsUrl)
                    : `http://localhost:9527/api/skills/shared?t=${Date.now()}`;
                const response = await fetch(skillsUrl);
                const parsed = await response.json();
                if (parsed && parsed.success && Array.isArray(parsed.skills)) {
                    // 统一字段：后端返回 name/displayName/description，VSTO 返回 Name/DisplayName/Description
                    // 这里保留原始字段同时补充 Pascal 命名字段，确保后续代码兼容
                    App.cachedSkills = parsed.skills.map(skill => ({
                        ...skill,
                        Name: skill.Name || skill.name,
                        DisplayName: skill.DisplayName || skill.displayName,
                        Description: skill.Description || skill.description,
                        Keywords: skill.Keywords || skill.keywords,
                        Examples: skill.Examples || skill.examples,
                        TemplateCode: skill.TemplateCode || skill.templateCode,
                        ProcessingMode: skill.ProcessingMode || skill.processingMode,
                        Pipeline: skill.Pipeline || skill.pipeline,
                        IsCustom: true,
                        source: 'shared'
                    }));
                    ChatLog.info('SKILL', `[Skill加载] Web 版成功加载 ${App.cachedSkills.length} 个共享 Skills`);
                    if (typeof notifyCachedSkillsLoaded === 'function') {
                        notifyCachedSkillsLoaded();
                    }
                    return App.cachedSkills;
                }
            }
        } catch (e) {
            ChatLog.error('SKILL', '加载 Skill 列表失败', e);
        }
        return [];
    }

    // --- 来自 app.js:L1650-L1678 ---
    // 【两阶段匹配优化】启动时预加载所有Skills到全局缓存
    async function preloadAllSkills() {
        ChatLog.info('SKILL', '[Skill预加载] 开始加载所有Skills到缓存...');

        try {
            const skills = await loadSkillsForLLM(true);  // 强制刷新

            if (skills && skills.length > 0) {
                window.cachedSkills = skills;  // 存入全局变量（供sendMessage使用）
                App.cachedSkills = skills;     // 同步到 App 对象（供 exactMatchSkill 使用）
                window.cachedSkillsLoaded = true;  // 标记已加载
                ChatLog.info('SKILL', `[Skill预加载] ✅ 完成，共${skills.length}个Skills`);

                // 输出所有Skill名称（方便调试）
                skills.forEach((skill, idx) => {
                    ChatLog.info('SKILL', `  [${idx + 1}] ${skill.name}: ${skill.description?.substring(0, 50)}...`);
                });
            } else {
                ChatLog.info('SKILL', '[Skill预加载] 当前没有可用的 Skills');
                window.cachedSkills = [];
                App.cachedSkills = [];       // 同步到App对象
                window.cachedSkillsLoaded = false;
            }

            return skills;
        } catch (error) {
            ChatLog.error('SKILL', '[Skill预加载] ❌ 失败:', error);
            window.cachedSkills = [];
            App.cachedSkills = [];           // 同步到App对象
            window.cachedSkillsLoaded = false;
            return [];
        } finally {
            // 【全量清单·启动即加载】GetAvailableSkills 只返回启用项，无法区分"用户主动停用"
            // 还是"后端压根没这个 Skill"——而这个区别决定了两件事：
            //   ① 停用感知 isDomainDisabledByUser：避免被停用的领域被 DESCRIPTORS 空壳复活；
            //   ② 背书包袱 isDomainBacked：避免给"后端压根没有"的领域注入空壳抢路由
            //      （实测：幽灵 formatter(kw=32,score=5) 抢走已启用的 SmartTableBeautifier(score=1)，
            //        用户反馈"我明明打开了 beautifier，却没匹配到"）。
            // 这份全量清单此前只在打开技能管理面板时才拉取，用户不打开面板就整体退化成旧行为
            // （路由结果随"是否点开过技能面板"而变，极难排查）。这里在启动预加载时顺手拉一次，
            // preloadAllSkills 本身已是非阻塞后台任务。
            try {
                let allRaw = null;
                if (window.BridgeAccessor && typeof window.BridgeAccessor.callAsync === 'function') {
                    allRaw = await window.BridgeAccessor.callAsync('GetAllSkills');
                } else {
                    const _b = (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects
                        && window.chrome.webview.hostObjects.excelBridge) || null;
                    if (_b) allRaw = await _b.GetAllSkills();
                }
                if (allRaw != null) {
                    const allParsed = typeof allRaw === 'string' ? JSON.parse(allRaw) : allRaw;
                    if (allParsed && allParsed.error === 'bridge-call-failed') {
                        ChatLog.warn('SKILL', '[Skill预加载] GetAllSkills 桥失败:', allParsed.detail);
                    } else if (allParsed && allParsed.success && Array.isArray(allParsed.skills)) {
                        App.allSkillsFull = allParsed.skills;
                        window.allSkillsFull = allParsed.skills;
                        App.disabledSkillNames = allParsed.skills
                            .filter(s => s && (s.isEnabled === false || s.IsEnabled === false))
                            .map(s => s.name || s.Name)
                            .filter(Boolean);
                        window.disabledSkillNames = App.disabledSkillNames;
                        ChatLog.info('SKILL', `[Skill预加载] 全量清单 ${allParsed.skills.length} 个`
                            + `（停用 ${App.disabledSkillNames.length} 个）→ 供停用感知/背书校验使用`);
                    }
                }
            } catch (e) {
                const msg = (e && e.message) ? e.message : String(e);
                if (!/e is not a function|0x80020006|未知名称/i.test(msg)) {
                    ChatLog.warn('SKILL', '[Skill预加载] 全量清单拉取失败（停用感知/背书校验将退化为保守行为）:', msg);
                }
            }
        }
    }

    // ==================== 导出公共接口 ====================
    // 将函数挂载到 App 对象上
    App.loadPromptConfig = loadPromptConfig;
    App.parsePromptsFromMarkdown = parsePromptsFromMarkdown;
    App.getPrompt = getPrompt;
    App.getSkillInfo = getSkillInfo;
    App.reloadPromptConfig = reloadPromptConfig;
    App.getPromptConfigStatus = getPromptConfigStatus;
    App.parseModelRecognitionResponse = parseModelRecognitionResponse;
    App.buildRecognitionCardFromObject = buildRecognitionCardFromObject;
    App.processElementForRecognition = processElementForRecognition;
    App.loadCardTypesConfig = loadCardTypesConfig;
    App.convertSkillConfigToImageTypeConfig = convertSkillConfigToImageTypeConfig;
    App.reloadCardTypesConfig = reloadCardTypesConfig;
    App.getCardTypesConfigStatus = getCardTypesConfigStatus;
    App.getCardTypeConfig = getCardTypeConfig;
    App.getAvailableCardTypes = getAvailableCardTypes;
    App.waitForCardTypesConfig = waitForCardTypesConfig;
    App.getFieldNameMapping = getFieldNameMapping;
    App.getFieldList = getFieldList;
    App.isSameField = isSameField;
    App.transformFieldNames = transformFieldNames;
    App.getHeaderOrder = getHeaderOrder;
    App.formatRecognitionResult = formatRecognitionResult;
    App.loadSkillsForLLM = loadSkillsForLLM;
    App.preloadAllSkills = preloadAllSkills;

})();