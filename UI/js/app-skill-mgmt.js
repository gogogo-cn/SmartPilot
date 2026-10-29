/**
 * ============================================================
 * app-skill-mgmt.js — Skill Management UI
 * ============================================================
 * 功能：Skill 管理界面模块，提供 Skill 的增删改查、代码库管理与模板维护。
 *
 * 核心类/函数：
 *   - openSkillManager() — 打开 Skill 管理器
 *   - renderSkillList() — 渲染 Skill 列表
 *   - saveSkill() — 保存 Skill
 *   - deleteSkill() — 删除 Skill
 *
 * 依赖：
 *   - app.js
 *   - app-ai-skill.js
 *   - app-editor.js
 * ============================================================
 */

(function() {
    'use strict';

    // ==================== 辅助函数 ====================

    // 获取 ExcelBridge 实例（兼容不同初始化方式）
    function getBridge() {
        return App.excelBridge ||
            (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge);
    }

    // 检测是否为真实 WebView2（非 Mock）
    function isRealWebView2() {
        return !!(window.chrome && window.chrome.webview && !window.chrome.webview._mockMode &&
                  window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge);
    }

    // 【修复】Web 版（Mock 模式）下通过 Local Bridge HTTP API 调用后端方法
    async function callBridgeMethod(method, args) {
        if (isRealWebView2()) {
            const bridge = getBridge();
            if (bridge && typeof bridge[method] === 'function') {
                return await bridge[method](...args);
            }
            throw new Error(`WebView2 中不存在方法: ${method}`);
        }

        // 非 WebView2 环境：优先使用 Local Bridge HTTP API
        const baseUrl = App.Endpoints ? App.Endpoints.excelBridge.baseUrl : (window.ExcelBridgeConfig ? window.ExcelBridgeConfig.baseUrl : 'http://localhost:9527');
        const response = await fetch(`${baseUrl}/api/excel/invoke`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ method: method, arguments: args })
        });
        const result = await response.json();
        // 兼容 Local Bridge 的返回结构：外层 { success, data }，里层是实际执行结果
        return JSON.stringify(result.data || result);
    }

    // HTML 转义（防止 XSS）
    function escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.appendChild(document.createTextNode(text));
        return div.innerHTML;
    }

    // ==================== 技能管理 UI 函数 ====================

    // --- 来自 app.js:L22559-L22678 ---
    function switchSkillTab(tabId) {
        // 更新标签按钮状态
        document.querySelectorAll('.skill-tab').forEach(tab => {
            tab.classList.remove('active');
            if (tab.dataset.tab === tabId) {
                tab.classList.add('active');
            }
        });

        // 更新内容区域显示
        document.querySelectorAll('.skill-tab-content').forEach(content => {
            content.classList.remove('active');
        });
        document.getElementById('skill-tab-' + tabId).classList.add('active');

        // 如果切换到 JSON 标签页，同步当前表单数据到 JSON
        if (tabId === 'json') {
            syncFormToJson();
        }
    }

    // --- 来自 app.js:L22681-L22692 ---
    function toggleAdvancedSection(headerEl) {
        // 支持直接传 header 元素，或从全局找（兼容 onclick="toggleAdvancedSection()" 旧调用）
        const header = (typeof headerEl === 'string') ?
            document.querySelector(headerEl) : headerEl ||
            document.querySelector('.advanced-section .section-title.collapsible');
        if (!header) return;

        // 从 header 向上找 form-section，再找其直接子 .advanced-content
        const parent = header.closest('.form-section');
        const content = parent ? parent.querySelector('.advanced-content') : null;
        const icon = header.querySelector('.toggle-icon');
        if (!content) return;

        const isHidden = content.style.display === 'none';
        content.style.display = isHidden ? '' : 'none';
        if (icon) icon.textContent = isHidden ? '▼' : '▶';
    }

    // --- 来自 app.js:L22695-L22706 ---
    function onProcessingModeChange(mode) {
        const modeValue = parseInt(document.getElementById('skill-input-processing-mode').value);
        const vbaCodeGroup = document.getElementById('skill-code-group');
        if (!vbaCodeGroup) return;

        // LLM 模式(1) 和本地OCR模式(3) 隐藏代码编辑器
        // VBA 模式(0) 和混合模式(2) 显示代码编辑器
        if (modeValue === 1 || modeValue === 3) {
            vbaCodeGroup.style.display = 'none';
        } else {
            vbaCodeGroup.style.display = 'block';
        }
    }

    /**
     * 快捷入口复选框切换：显示/隐藏快捷入口配置区域
     */
    function onQuickShortcutToggle() {
        const checked = document.getElementById('skill-input-quick-shortcut').checked;
        const configDiv = document.getElementById('skill-quick-shortcut-config');
        if (configDiv) {
            configDiv.style.display = checked ? '' : 'none';
        }
        // 同步更新 JSON 编辑器
        syncFormToJson();
    }

    // 绑定快捷入口复选框事件（DOM 就绪后执行）
    (function bindQuickShortcutEvent() {
        const checkbox = document.getElementById('skill-input-quick-shortcut');
        if (checkbox) {
            checkbox.addEventListener('change', onQuickShortcutToggle);
        }
    })();

    // --- 来自 app.js:L22709-L22712 ---
    function syncFormToJson() {
        const skillData = collectSkillDataFromForm(App.currentEditingSkill);
        document.getElementById('skill-input-json').value = JSON.stringify(skillData, null, 2);
    }

    /**
     * 表单 <-> JSON 双向同步绑定（只执行一次）
     * - 表单修改 → 实时同步到 JSON tab（debounce 150ms）
     * - JSON 手工修改 → 暂不反向覆盖表单（避免互相打架）
     */
    let _formJsonSyncBound = false;
    let _formJsonSyncTimer = null;
    function bindFormJsonSync() {
        if (_formJsonSyncBound) return;
        const body = document.querySelector('#skill-edit-dialog .dialog-body');
        if (!body) return;
        const scheduleSync = function(e) {
            // 排除 JSON tab 的手动编辑：只同步表单 tab → JSON，不反向
            if (e.target && e.target.id === 'skill-input-json') return;
            if (_formJsonSyncTimer) clearTimeout(_formJsonSyncTimer);
            _formJsonSyncTimer = setTimeout(syncFormToJson, 150);
        };
        body.addEventListener('input', scheduleSync);
        body.addEventListener('change', scheduleSync);
        _formJsonSyncBound = true;
    }

    /**
     * 渲染 Skill 声明的 specFields 到编辑器表单
     * 支持 select / boolean / string / number 四类字段，仅允许编辑 defaultValue。
     * @param {Object} skill - 当前编辑的 Skill（含 specFields 或 SpecFields）
     */
    function renderSpecFields(skill) {
        const section = document.getElementById('skill-spec-fields-section');
        const container = document.getElementById('skill-spec-fields-container');
        if (!section || !container) return;

        const fields = getSkillProp(skill, 'specFields', 'SpecFields', []);
        if (!Array.isArray(fields) || fields.length === 0) {
            section.style.display = 'none';
            container.innerHTML = '';
            return;
        }

        container.innerHTML = fields.map((f, idx) => {
            const fName = f.name || f.Name || '';
            const fLabel = f.label || f.Label || fName;
            const fType = String(f.type || f.Type || 'string').toLowerCase();
            const fOptions = Array.isArray(f.options) ? f.options : (Array.isArray(f.Options) ? f.Options : []);
            const fDefault = (f.defaultValue !== undefined && f.defaultValue !== null) ? f.defaultValue : (f.DefaultValue !== undefined ? f.DefaultValue : '');
            const inputId = 'skill-spec-field-' + idx;

            let control;
            if (fType === 'select') {
                const opts = fOptions.map(opt => {
                    const optVal = (opt && opt.value !== undefined && opt.value !== null) ? opt.value : opt;
                    const optLabel = (opt && opt.label) ? opt.label : optVal;
                    const sel = String(optVal) === String(fDefault) ? ' selected' : '';
                    return `<option value="${escapeHtml(String(optVal))}"${sel}>${escapeHtml(String(optLabel))}</option>`;
                }).join('');
                control = `<select id="${inputId}" class="spec-field-select" data-name="${escapeHtml(fName)}" data-type="select">
                    <option value="">（不指定）</option>${opts}</select>`;
            } else if (fType === 'boolean') {
                const checked = String(fDefault) === 'true' ? ' checked' : '';
                control = `<input type="checkbox" id="${inputId}" class="spec-field-boolean" data-name="${escapeHtml(fName)}" data-type="boolean"${checked}>`;
            } else if (fType === 'number') {
                control = `<input type="number" id="${inputId}" class="spec-field-number" data-name="${escapeHtml(fName)}" data-type="number" value="${escapeHtml(String(fDefault))}" placeholder="默认值">`;
            } else {
                control = `<input type="text" id="${inputId}" class="spec-field-string" data-name="${escapeHtml(fName)}" data-type="string" value="${escapeHtml(String(fDefault))}" placeholder="默认值">`;
            }

            return `<div class="spec-field-row" style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
                <label style="min-width:110px;font-size:13px;color:#555;" title="${escapeHtml(String(fType))}">${escapeHtml(fLabel)}</label>
                <div style="flex:1;">${control}</div>
            </div>`;
        }).join('');

        section.style.display = 'block';
    }

    /**
     * 从表单收集 specFields 的 defaultValue 修改
     * @returns {Array} specFields 数组（含最新 defaultValue）
     */
    function collectSpecFieldsFromForm(originalSkill) {
        const orig = originalSkill || {};
        const fields = getSkillProp(orig, 'specFields', 'SpecFields', []);
        if (!Array.isArray(fields) || fields.length === 0) return fields;

        return fields.map((f, idx) => {
            const el = document.getElementById('skill-spec-field-' + idx);
            const result = {
                name: f.name || f.Name || '',
                label: f.label || f.Label || '',
                type: f.type || f.Type || 'string'
            };
            if (Array.isArray(f.options)) result.options = f.options;
            else if (Array.isArray(f.Options)) result.options = f.Options;
            if (el) {
                if (el.type === 'checkbox') {
                    result.defaultValue = el.checked ? 'true' : 'false';
                } else {
                    result.defaultValue = el.value || '';
                }
            } else if (f.defaultValue !== undefined) {
                result.defaultValue = f.defaultValue;
            }
            return result;
        });
    }

    // --- 来自 app.js:L22715-L22758 ---
    function collectSkillDataFromForm(originalSkill) {
        const name = document.getElementById('skill-input-name').value.trim();
        const displayName = document.getElementById('skill-input-display').value.trim();
        const description = document.getElementById('skill-input-desc').value.trim();
        const keywords = document.getElementById('skill-input-keywords').value.split(',').map(k => k.trim()).filter(k => k);
        const code = document.getElementById('skill-input-code').value.trim();
        const examples = document.getElementById('skill-input-examples').value.split(',').map(e => e.trim()).filter(e => e);
        const processingMode = parseInt(document.getElementById('skill-input-processing-mode').value);
        const supportsImage = document.getElementById('skill-input-supports-image').checked;

        // 高级模式字段
        const llmPrompt = document.getElementById('skill-input-llm-prompt').value.trim();
        const llmUserTemplate = document.getElementById('skill-input-llm-user-template').value.trim();
        const outputFormat = document.getElementById('skill-input-output-format').value;
        const supportsFile = document.getElementById('skill-input-supports-file').checked;
        const version = parseInt(document.getElementById('skill-input-version').value) || 1;
        const tags = document.getElementById('skill-input-tags').value.split(',').map(t => t.trim()).filter(t => t);

        // 快捷入口字段
        const isQuickShortcut = document.getElementById('skill-input-quick-shortcut').checked;
        const quickShortcutIcon = document.getElementById('skill-input-quick-icon').value.trim() || '⚡';
        const quickShortcutPrompt = document.getElementById('skill-input-quick-prompt').value.trim();
        
        // UI按钮配置
        const uiActions = {
            ShowCopy: document.getElementById('skill-ui-show-copy').checked,
            ShowSaveMd: document.getElementById('skill-ui-show-save-md').checked,
            ShowRegenerate: document.getElementById('skill-ui-show-regenerate').checked
        };

        // 从原始 Skill 中保留系统级字段（不允许表单修改的）
        const orig = originalSkill || {};
        const isGeneric = getSkillProp(orig, 'isGeneric', 'IsGeneric', false);
        const needsConfig = getSkillProp(orig, 'needsConfig', 'NeedsConfig', false);
        const isCustom = getSkillProp(orig, 'isCustom', 'IsCustom', false);
        const isComposite = getSkillProp(orig, 'isComposite', 'IsComposite', false);

        return {
            Name: name,
            DisplayName: displayName,
            Description: description,
            Keywords: keywords,
            TemplateCode: code,
            Examples: examples,
            ProcessingMode: processingMode,
            LLMSystemPrompt: llmPrompt,
            LLMUserPromptTemplate: llmUserTemplate || null,
            SupportsImageInput: supportsImage,
            SupportsFileInput: supportsFile,
            OutputFormat: outputFormat,
            Version: version,
            Tags: tags,
            IsQuickShortcut: isQuickShortcut,
            QuickShortcutIcon: quickShortcutIcon,
            QuickShortcutPrompt: quickShortcutPrompt,
            IsGeneric: isGeneric,
            NeedsConfig: needsConfig,
            IsCustom: isCustom,
            IsComposite: isComposite,
            WorkflowSteps: orig.workflowSteps || orig.WorkflowSteps || [],
            CustomWorkflowSteps: orig.customWorkflowSteps || orig.CustomWorkflowSteps || [],
            InputParameters: orig.inputParameters || orig.InputParameters || [],
            OutputParameters: orig.outputParameters || orig.OutputParameters || [],
            FieldMappings: orig.fieldMappings || orig.FieldMappings || [],
            // 保留并收集 specFields（含表单中的 defaultValue 修改）
            SpecFields: collectSpecFieldsFromForm(originalSkill),
            // 保留 quickPrompts
            QuickPrompts: getSkillProp(orig, 'quickPrompts', 'QuickPrompts', []),
            UIActions: uiActions
        };
    }

    // --- 来自 app.js:L22761-L22771 ---
    function formatSkillJson(skillData) {
        try {
            const jsonText = document.getElementById('skill-input-json').value;
            const parsed = JSON.parse(jsonText);
            document.getElementById('skill-input-json').value = JSON.stringify(parsed, null, 2);
            document.getElementById('skill-json-error').style.display = 'none';
        } catch (error) {
            document.getElementById('skill-json-error').textContent = 'JSON 格式错误: ' + error.message;
            document.getElementById('skill-json-error').style.display = 'block';
        }
    }

    // --- 来自 app.js:L22774-L22789 ---
    function validateSkillJson(jsonStr) {
        try {
            const jsonText = document.getElementById('skill-input-json').value;
            const parsed = JSON.parse(jsonText);

            // 验证必需字段
            if (!parsed.Name) throw new Error('缺少必需字段: Name');
            if (!parsed.DisplayName) throw new Error('缺少必需字段: DisplayName');

            document.getElementById('skill-json-error').innerHTML = '<span style="color: #4caf50;">✓ JSON 格式正确，必需字段完整</span>';
            document.getElementById('skill-json-error').style.display = 'block';
        } catch (error) {
            document.getElementById('skill-json-error').textContent = '验证失败: ' + error.message;
            document.getElementById('skill-json-error').style.display = 'block';
        }
    }

    // --- 来自 app.js:L22791-L22821 ---
    function showAddSkillDialog() {
        document.getElementById('skill-dialog-title').textContent = '添加 Skill';
        document.getElementById('skill-edit-name').value = '';
        document.getElementById('skill-input-name').value = '';
        document.getElementById('skill-input-display').value = '';
        document.getElementById('skill-input-desc').value = '';
        document.getElementById('skill-input-keywords').value = '';
        document.getElementById('skill-input-code').value = '';
        document.getElementById('skill-input-examples').value = '';
        document.getElementById('skill-input-name').disabled = false;

        // 新增字段默认值
        document.getElementById('skill-input-processing-mode').value = '0';
        document.getElementById('skill-input-supports-image').checked = false;
        document.getElementById('skill-input-llm-prompt').value = '';
        document.getElementById('skill-input-llm-user-template').value = '';
        document.getElementById('skill-input-output-format').value = 'text';
        document.getElementById('skill-input-supports-file').checked = false;
        document.getElementById('skill-input-version').value = '1';
        document.getElementById('skill-input-tags').value = '';
        document.getElementById('skill-input-json').value = '';
        document.getElementById('skill-json-error').style.display = 'none';

        // 快捷入口字段重置
        document.getElementById('skill-input-quick-shortcut').checked = false;
        document.getElementById('skill-input-quick-icon').value = '⚡';
        document.getElementById('skill-input-quick-prompt').value = '';
        document.getElementById('skill-quick-shortcut-config').style.display = 'none';

        // UI 按钮重置（默认全部开）
        document.getElementById('skill-ui-show-copy').checked = true;
        document.getElementById('skill-ui-show-save-md').checked = false;
        document.getElementById('skill-ui-show-regenerate').checked = true;

        // 隐藏 specFields 区块（新 Skill 默认没有）
        const specSection = document.getElementById('skill-spec-fields-section');
        const specContainer = document.getElementById('skill-spec-fields-container');
        if (specSection) specSection.style.display = 'none';
        if (specContainer) specContainer.innerHTML = '';

        // 重置标签页到表单模式，折叠高级选项（HTML 默认已折叠，这里兜底）
        switchSkillTab('form');
        onProcessingModeChange();
        document.getElementById('skill-advanced-content').style.display = 'none';
        document.querySelector('#skill-edit-dialog .section-title.collapsible .toggle-icon').textContent = '▶';

        bindFormJsonSync();
        document.getElementById('skill-edit-dialog').style.display = 'flex';
    }

    // 【新增】解析Markdown为Skill对象
    function parseMarkdownToSkill(content, fileName) {
        const yamlMatch = content.match(/^---[\r\n]+([\s\S]*?)[\r\n]+---/);
        const mdContent = content.replace(/^---[\r\n]+[\s\S]*?[\r\n]+---/, '').trim();
        
        const skill = {
            name: fileName,
            displayName: fileName,
            description: '',
            keywords: [],
            templateCode: mdContent,
            examples: [],
            processingMode: 0,
            llmSystemPrompt: '',
            supportsImageInput: false,
            supportsFileInput: false,
            outputFormat: 'text',
            version: 1,
            tags: []
        };
        
        if (yamlMatch) {
            const yaml = yamlMatch[1];
            const lines = yaml.split('\n');
            
            lines.forEach(line => {
                const trimmed = line.trim();
                if (trimmed.startsWith('displayName:')) {
                    skill.displayName = trimmed.substring('displayName:'.length).trim();
                } else if (trimmed.startsWith('description:')) {
                    skill.description = trimmed.substring('description:'.length).trim();
                } else if (trimmed.startsWith('processingMode:')) {
                    const mode = trimmed.substring('processingMode:'.length).trim();
                    skill.processingMode = mode === 'Hybrid' ? 2 : mode === 'LLM' ? 1 : 0;
                } else if (trimmed.startsWith('supportsImageInput:')) {
                    skill.supportsImageInput = trimmed.includes('true');
                } else if (trimmed.startsWith('supportsFileInput:')) {
                    skill.supportsFileInput = trimmed.includes('true');
                } else if (trimmed.startsWith('outputFormat:')) {
                    skill.outputFormat = trimmed.substring('outputFormat:'.length).trim();
                } else if (trimmed.startsWith('version:')) {
                    skill.version = parseInt(trimmed.substring('version:'.length).trim()) || 1;
                }
            });
        }
        
        return skill;
    }
    
    /**
     * 从后端返回的 skill（PascalCase / camelCase 混杂）中一次性提取常用字段，
     * 返回一个全 camelCase + 默认值的对象。避免 fillEditForm / renderSpecFields 等处
     * 反复写 getSkillProp(skill, 'x', 'X', default)。
     */
    function normalizeSkill(skill) {
        const s = skill || {};
        const g = function(k1, k2, def) { return getSkillProp(s, k1, k2, def); };
        return {
            name:               g('name', 'Name', '') || '',
            displayName:        g('displayName', 'DisplayName', '') || '',
            description:        g('description', 'Description', '') || '',
            keywords:           g('keywords', 'Keywords', []),
            templateCode:       g('templateCode', 'TemplateCode', '') || '',
            examples:           g('examples', 'Examples', []),
            processingMode:     g('processingMode', 'ProcessingMode', 0),
            llmSystemPrompt:    g('llmSystemPrompt', 'LLMSystemPrompt', '') || '',
            llmUserPromptTemplate: g('llmUserPromptTemplate', 'LLMUserPromptTemplate', '') || '',
            supportsImageInput: g('supportsImageInput', 'SupportsImageInput', false),
            supportsFileInput:  g('supportsFileInput', 'SupportsFileInput', false),
            outputFormat:       g('outputFormat', 'OutputFormat') || 'text',
            version:            g('version', 'Version', 1),
            tags:               g('tags', 'Tags', []),
            isQuickShortcut:    g('isQuickShortcut', 'IsQuickShortcut', false),
            quickShortcutIcon:  g('quickShortcutIcon', 'QuickShortcutIcon') || '⚡',
            quickShortcutPrompt:g('quickShortcutPrompt', 'QuickShortcutPrompt') || '',
            isGeneric:          g('isGeneric', 'IsGeneric', false),
            needsConfig:        g('needsConfig', 'NeedsConfig', false),
            isCustom:           g('isCustom', 'IsCustom', false),
            isComposite:        g('isComposite', 'IsComposite', false),
            uiActions:          g('uiActions', 'UIActions', {}),
            specFields:         g('specFields', 'SpecFields', []),
        };
    }

    // 【新增】填充编辑表单（从VSTO版复制）
    function fillEditForm(skill) {
        const n = normalizeSkill(skill);

        // 渲染 Spec 参数配置区块
        renderSpecFields(skill);

        document.getElementById('skill-dialog-title').textContent = '编辑 Skill';
        document.getElementById('skill-edit-name').value = n.name;
        document.getElementById('skill-input-name').value = n.name;
        document.getElementById('skill-input-name').disabled = true;
        document.getElementById('skill-input-display').value = n.displayName;
        document.getElementById('skill-input-desc').value = n.description;
        document.getElementById('skill-input-keywords').value = Array.isArray(n.keywords) ? n.keywords.join(',') : '';
        document.getElementById('skill-input-code').value = n.templateCode;
        document.getElementById('skill-input-examples').value = Array.isArray(n.examples) ? n.examples.join(',') : '';
        document.getElementById('skill-input-processing-mode').value = String(n.processingMode);
        document.getElementById('skill-input-supports-image').checked = n.supportsImageInput;
        document.getElementById('skill-input-llm-prompt').value = n.llmSystemPrompt;
        document.getElementById('skill-input-llm-user-template').value = n.llmUserPromptTemplate;
        document.getElementById('skill-input-output-format').value = n.outputFormat;
        document.getElementById('skill-input-supports-file').checked = n.supportsFileInput;
        document.getElementById('skill-input-version').value = String(n.version);
        document.getElementById('skill-input-tags').value = Array.isArray(n.tags) ? n.tags.join(',') : '';
        document.getElementById('skill-input-quick-shortcut').checked = n.isQuickShortcut;
        document.getElementById('skill-input-quick-icon').value = n.quickShortcutIcon;
        document.getElementById('skill-input-quick-prompt').value = n.quickShortcutPrompt;
        document.getElementById('skill-quick-shortcut-config').style.display = n.isQuickShortcut ? '' : 'none';

        // UI按钮配置（兼容两种命名）
        const ua = n.uiActions;
        document.getElementById('skill-ui-show-copy').checked      = getSkillProp(ua, 'showCopy',      'ShowCopy',      true);
        document.getElementById('skill-ui-show-save-md').checked  = getSkillProp(ua, 'showSaveMd',    'ShowSaveMd',    false);
        document.getElementById('skill-ui-show-regenerate').checked= getSkillProp(ua, 'showRegenerate','ShowRegenerate',true);

        // 同步到 JSON 编辑器 — 确保包含系统级字段
        const jsonSkill = collectSkillDataFromForm(skill);
        document.getElementById('skill-input-json').value = JSON.stringify(jsonSkill, null, 2);
        document.getElementById('skill-json-error').style.display = 'none';

        // 重置标签页到表单模式，折叠高级选项
        switchSkillTab('form');
        onProcessingModeChange();
        document.getElementById('skill-advanced-content').style.display = 'none';
        document.querySelector('#skill-edit-dialog .section-title.collapsible .toggle-icon').textContent = '▶';
        
        bindFormJsonSync();
        document.getElementById('skill-edit-dialog').style.display = 'flex';
    }
    
    // 【暴露到App对象】
    window.editSkill = editSkill;
    if (window.App) {
        window.App.editSkill = editSkill;
    }
    
    function isVstoWebView2() {
        return !!(window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge);
    }

    // --- 来自 app.js:L22824-L22892 ---
    async function editSkill(skillName) {
        try {
            if (!skillName) {
                ChatLog.warn('SKILL', '[editSkill] skillName 为空或 undefined');
                return;
            }

            ChatLog.info('SKILL', '[editSkill] 开始加载 Skill:', skillName);
            
            let result;

            // VSTO WebView2 环境优先使用 COM 桥接
            if (isVstoWebView2()) {
                ChatLog.info('SKILL', '[editSkill] VSTO WebView2 环境，使用 COM 桥接');
                const bridge = getBridge();
                if (!bridge) {
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">无法获取 Skill：Excel 桥接不可用</span>', true) : App.showToast('无法获取 Skill：Excel 桥接不可用');
                    return;
                }
                const skillJson = await bridge.GetSkillByName(skillName);
                result = JSON.parse(skillJson);
            } else {
                // 非 VSTO 环境使用 Web API
                const skillUrl = window.ExcelBridgeConfig ? window.ExcelBridgeConfig.skillUrl(skillName) : `http://localhost:9527/api/skills/shared/${encodeURIComponent(skillName)}`;
                ChatLog.info('SKILL', '[editSkill] 请求URL:', skillUrl);
                
                try {
                    const response = await fetch(skillUrl);
                    ChatLog.info('SKILL', '[editSkill] HTTP状态:', response.status);
                    result = await response.json();
                    ChatLog.info('SKILL', '[editSkill] API返回:', JSON.stringify(result).substring(0, 200));
                } catch (fetchErr) {
                    ChatLog.warn('SKILL', '[editSkill] Web API失败:', fetchErr.message);
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">无法获取 Skill：Web API 不可用</span>', true) : App.showToast('无法获取 Skill：Web API 不可用');
                    return;
                }
            }
            
            if (!result.success) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">获取 Skill 失败: ' + (result.error || '未知错误') + '</span>', true) : App.showToast('获取 Skill 失败: ' + (result.error || '未知错误'));
                return;
            }
            
            const skill = result.skill;
            if (!skill) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">获取 Skill 失败：返回数据为空</span>', true) : App.showToast('获取 Skill 失败：返回数据为空');
                return;
            }
            
            ChatLog.info('SKILL', '[editSkill] Skill数据:', skill.name, skill.displayName);
            App.currentEditingSkill = skill;
            
            // 保存当前编辑的 Skill 名称，用于打开源文件
            _currentEditSkillName = skill.name || skill.Name || skillName;
            
            // 填充表单
            fillEditForm(skill);
        } catch (error) {
            ChatLog.error('SKILL', '[editSkill] 异常:', error);
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">加载 Skill 失败: ' + error.message + '</span>', true) : App.showToast('加载 Skill 失败: ' + error.message);
        }
    }

    // --- 来自 app.js:L22895-L22897 ---
    function closeSkillDialog() {
        document.getElementById('skill-edit-dialog').style.display = 'none';
    }

    // 【新增】打开 Skill 的 MD 源文件
    let _currentEditSkillName = null;
    
    async function openSkillSourceFile() {
        if (!_currentEditSkillName) {
            App.showToast('无法获取 Skill 名称');
            return;
        }
        
        try {
            const bridge = getBridge();
            if (!bridge) {
                App.showToast('Excel 桥接不可用');
                return;
            }
            
            const resultJson = await bridge.OpenSkillSourceFile(_currentEditSkillName);
            const result = JSON.parse(resultJson);
            
            if (!result.success) {
                App.showToast('打开失败: ' + (result.error || '未知错误'));
                return;
            }
            
            // 显示文件路径提示
            const typeLabel = result.isCustom ? '自定义' : '系统';
            App.showToast(`已在默认程序打开${typeLabel}Skill源文件: ${result.path}`);
        } catch (err) {
            console.error('[openSkillSourceFile] 错误:', err);
            App.showToast('打开源文件失败');
        }
    }
    
    // 暴露到全局
    window.openSkillSourceFile = openSkillSourceFile;

    // --- 来自 app.js:L22900-L22979 ---
    async function saveSkill() {
        try {
            if (!getBridge()) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">Excel 未连接</span>', true) : App.showToast('Excel 未连接');
                return;
            }

            let skillData;

            // 检查当前激活的标签页
            const activeTab = document.querySelector('.skill-tab.active');
            const currentTab = activeTab ? activeTab.dataset.tab : 'simple';

            if (currentTab === 'json') {
                // JSON 模式：直接从 JSON 编辑器获取
                try {
                    const jsonText = document.getElementById('skill-input-json').value;
                    skillData = JSON.parse(jsonText);

                    // 确保必需字段
                    if (!skillData.Name || !skillData.DisplayName) {
                        App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">JSON 中缺少必需字段: Name 和 DisplayName</span>', true) : App.showToast('JSON 中缺少必需字段: Name 和 DisplayName');
                        return;
                    }
                    
                    // 确保 UIActions 字段存在（如果没有则添加默认值）
                    if (!skillData.UIActions) {
                        skillData.UIActions = {
                            ShowCopy: true,
                            ShowSaveMd: false,
                            ShowRegenerate: true
                        };
                    }
                } catch (error) {
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">JSON 格式错误: ' + error.message + '</span>', true) : App.showToast('JSON 格式错误: ' + error.message);
                    return;
                }
            } else {
                // 表单模式：从表单收集数据，传递原始 Skill 以保留系统级字段
                skillData = collectSkillDataFromForm(App.currentEditingSkill);

                if (!skillData.Name || !skillData.DisplayName) {
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">请填写 Skill 名称和显示名称</span>', true) : App.showToast('请填写 Skill 名称和显示名称');
                    return;
                }

                // 验证名称格式（只能包含字母、数字和下划线）
                if (!/^[a-zA-Z][a-zA-Z0-9_]*$/.test(skillData.Name)) {
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">Skill 名称必须以字母开头，只能包含字母、数字和下划线</span>', true) : App.showToast('Skill 名称必须以字母开头，只能包含字母、数字和下划线');
                    return;
                }

                // 合并原始 Skill 的其他字段（防止丢失未在表单中显示的数据）
                if (App.currentEditingSkill) {
                    const orig = App.currentEditingSkill;
                    // 保留原始 Skill 中的工作流、参数等复杂字段（兼容双命名）
                    skillData.WorkflowSteps = getSkillProp(orig, 'workflowSteps', 'WorkflowSteps', []);
                    skillData.CustomWorkflowSteps = getSkillProp(orig, 'customWorkflowSteps', 'CustomWorkflowSteps', []);
                    skillData.InputParameters = getSkillProp(orig, 'inputParameters', 'InputParameters', []);
                    skillData.OutputParameters = getSkillProp(orig, 'outputParameters', 'OutputParameters', []);
                    skillData.FieldMappings = getSkillProp(orig, 'fieldMappings', 'FieldMappings', []);
                    skillData.ConfigFields = getSkillProp(orig, 'configFields', 'ConfigFields', []);
                    
                    // 保留 UIActions 配置
                    const origUIActions = getSkillProp(orig, 'uiActions', 'UIActions', null);
                    if (origUIActions) {
                        skillData.UIActions = origUIActions;
                    } else if (!skillData.UIActions) {
                        skillData.UIActions = {
                            ShowCopy: true,
                            ShowSaveMd: false,
                            ShowRegenerate: true
                        };
                    }

                    // 保留时间戳和统计信息
                    const createdDate = getSkillProp(orig, 'createdDate', 'CreatedDate', null);
                    if (createdDate) {
                        skillData.CreatedDate = createdDate;
                    }
                    const usageCount = getSkillProp(orig, 'usageCount', 'UsageCount', undefined);
                    if (usageCount !== undefined) {
                        skillData.UsageCount = usageCount || 0;
                    }
                } else if (!skillData.UIActions) {
                    // 新Skill，添加默认UIActions
                    skillData.UIActions = {
                        ShowCopy: true,
                        ShowSaveMd: false,
                        ShowRegenerate: true
                    };
                }
            }

            // 入库前清洗 + 校验：与 SkillCreator 卡片保存同一道防线
            sanitizeSkillExamples(skillData);
            const preStore = await validateStepsBeforeSave(skillData);
            if (!preStore.ok) {
                const msg = '❌ ' + preStore.message;
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">' + msg + '</span>', true) : App.showToast(msg);
                return;
            }

            console.log('[Skill保存] 开始保存 Skill:', skillData.name || skillData.Name);
            const result = await callBridgeMethod('SaveSkill', [JSON.stringify(skillData)]);
            console.log('[Skill保存] SaveSkill 返回:', result);
            const resultObj = JSON.parse(result);

            if (resultObj.success) {
                console.log('[Skill保存] ✅ 保存成功，开始刷新列表...');
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">Skill 保存成功</span>', true) : App.showToast('Skill 保存成功');
                closeSkillDialog();

                if (typeof App.loadCustomSkills === 'function') {
                    console.log('[Skill保存] 调用 App.loadCustomSkills()...');
                    App.loadCustomSkills();
                }

                if (typeof App.loadSkillsAsCommands === 'function') {
                    console.log('[Skill保存] 调用 App.loadSkillsAsCommands(true)...');
                    await App.loadSkillsAsCommands(true);
                }
                // 技能管理面板必须重绘：上面只刷新了路由用缓存 App.cachedSkills，
                // 而面板数据来自 App.allSkillsFull + #skill-list-container，不重绘就是
                // "提示保存成功，但列表里看不到"（实测 2026-10-02 22:09，md 已落盘）。
                if (typeof App.loadSkillsForManagement === 'function') {
                    try { await App.loadSkillsForManagement(); } catch (e) { /* 面板未打开时忽略 */ }
                }
                console.log('[Skill保存] ✅ 刷新完成');
                if (typeof ChatLog !== 'undefined' && ChatLog.info) {
                    ChatLog.info('SKILL', '🔄 Skill 缓存已刷新，新 Skill 已生效');
                }
            } else {
                console.error('[Skill保存] ❌ 保存失败:', resultObj.error);
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">保存失败: ' + resultObj.error + '</span>', true) : App.showToast('保存失败: ' + resultObj.error);
            }
        } catch (error) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">保存失败: ' + error.message + '</span>', true) : App.showToast('保存失败: ' + error.message);
        }
    }

    /**
     * 保存 SkillCreator 生成的 Skill
     * @param {HTMLElement} btn - 被点击的按钮元素，包含 data-skill-json 属性
     */
    /**
     * 智能字段映射：将 camelCase 键映射到 JSON 中可能存在的 PascalCase/其他变体
     * 后端 SerializeSkillCompact 使用 PascalCase（如 LLMSystemPrompt, TemplateCode）
     * 此函数处理常见的大写前缀（LLM, API, URL, ID, HTTP, VBA 等）
     */
    function resolveJsonField(source, camelKey) {
        if (source.hasOwnProperty(camelKey)) return source[camelKey];

        // 常见大写前缀映射表
        const prefixMap = {
            llm: 'LLM', api: 'API', url: 'URL', id: 'ID', http: 'HTTP',
            vba: 'VBA', ocr: 'OCR', ai: 'AI', excel: 'Excel', xml: 'XML',
            json: 'JSON', pdf: 'PDF', vba: 'VBA'
        };

        // 尝试 PascalCase（首字母大写）
        let pascalKey = camelKey.charAt(0).toUpperCase() + camelKey.slice(1);
        if (source.hasOwnProperty(pascalKey)) return source[pascalKey];

        // 尝试常见前缀替换
        for (const [prefix, upperPrefix] of Object.entries(prefixMap)) {
            if (camelKey.startsWith(prefix) && camelKey.length > prefix.length) {
                const rest = camelKey.charAt(prefix.length).toUpperCase() + camelKey.slice(prefix.length + 1);
                const candidate = upperPrefix + rest;
                if (source.hasOwnProperty(candidate)) return source[candidate];
            }
        }

        // 尝试全小写匹配
        const lowerMap = {};
        for (const k of Object.keys(source)) {
            lowerMap[k.toLowerCase()] = source[k];
        }
        if (lowerMap[camelKey.toLowerCase()] !== undefined) return lowerMap[camelKey.toLowerCase()];

        return undefined;
    }

    /**
     * 保存前校验声明式 steps（PreStore 防线）：
     *   1) STEPS_CONTRACT.validateSteps —— direct 步必须有 tool、code 步必须有完整 VBA
     *   2) STEPS_CONTRACT.validateTools —— tool 必须在 toolRegistry 中真实注册（skill.X 委托同理）
     * 无 steps 的技能（多模态/Python/spec 等形态）直接放行。
     * @param {Object} skillData
     * @returns {Promise<{ok:boolean, message:string}>}
     */
    /**
     * 清洗生成技能的 examples：剔除"元指令"示例。
     * examples 的语义是"使用该技能时用户会说的话"，不是"创建这个技能的请求"。
     * 模型常照抄 SkillCreator 自身的示例（"生成一个批量发送邮件的技能"）当 examples，
     * 这类文本进入技能匹配库后会让该技能被"创建技能"类请求误触发。
     * 这里按语义硬性剔除（并打日志），而不是拦下保存——脏示例不该入库，但不值得让用户重来一次。
     * @returns {number} 剔除条数
     */
    function sanitizeSkillExamples(skillData) {
        try {
            if (!skillData || typeof skillData !== 'object') return 0;
            const key = Array.isArray(skillData.examples) ? 'examples'
                : (Array.isArray(skillData.Examples) ? 'Examples' : null);
            if (!key) return 0;

            const META_WORD_RE = /(\bskill\b|技能)/i;
            const META_VERB_RE = /^(帮我|请|帮忙)?\s*(生成|创建|新建|做一个|做个|弄一个|搞一个)/;
            const META_NOUN_RE = /(工具|功能|器|程序|脚本|插件)$/;

            const arr = skillData[key];
            const dropped = [];
            const kept = arr.filter(function (ex) {
                const s = String(ex == null ? '' : ex).trim();
                if (!s) return false;
                const isMeta = META_WORD_RE.test(s) || (META_VERB_RE.test(s) && META_NOUN_RE.test(s));
                if (isMeta) { dropped.push(s); return false; }
                return true;
            });

            if (dropped.length > 0) {
                skillData[key] = kept;
                console.log('[Skill保存] 已剔除 ' + dropped.length + ' 条元指令 examples（examples 应为"使用该技能时用户会说的话"，且总条数需保留 ≥1）:', dropped);
            }
            return dropped.length;
        } catch (e) {
            return 0;
        }
    }

    async function validateStepsBeforeSave(skillData) {
        const raw = skillData && (skillData.steps ?? skillData.Steps);
        if (!Array.isArray(raw) || raw.length === 0) return { ok: true, message: '' };

        const contract = (window.App && App.ChatV2 && App.ChatV2.skillDomain && App.ChatV2.skillDomain.STEPS_CONTRACT) || null;
        if (!contract || typeof contract.validateSteps !== 'function') return { ok: true, message: '' };

        const v = contract.validateSteps(raw);
        if (!v.ok) {
            return { ok: false, message: 'steps 契约校验失败：' + (v.errors || []).join('；') };
        }

        const reg = App.ChatV2 && App.ChatV2.toolRegistry;
        // MCP 工具为异步注册：先 init 再校验，避免把"尚未注册"误判为"工具不存在"
        if (reg && typeof reg.init === 'function') {
            try { await reg.init(); } catch (e) { /* init 失败则走 registry 现状 */ }
        }
        if (typeof contract.validateTools !== 'function') return { ok: true, message: '' };

        const t = contract.validateTools(raw, reg);
        if (t.skipped || t.ok) return { ok: true, message: '' };

        const detail = (t.unknown || []).map(u =>
            `步骤[${u.index}] "${u.tool}"${u.suggestion ? ' → 是否应为 "' + u.suggestion + '"' : ''}`
        ).join('；');
        return { ok: false, message: 'steps 引用了未注册的工具：' + detail };
    }

    async function saveGeneratedSkill(btn) {
        const skillJson = btn.getAttribute('data-skill-json');
        if (!skillJson) {
            if (App.showToast) App.showToast('❌ 找不到 Skill 数据', 'error');
            else alert('找不到 Skill 数据');
            return;
        }

        btn.disabled = true;
        btn.innerHTML = '⏳ 保存中...';

        try {
            let parsed;
            try {
                parsed = JSON.parse(skillJson);
            } catch (e) {
                throw new Error('Skill 数据 JSON 解析失败: ' + e.message);
            }

            // 智能清洗：保留所有字段，仅做类型修正和必填项补全
            const cleaned = {};
            const validFields = {
                name: 'string', displayName: 'string', description: 'string',
                keywords: 'array', examples: 'array', isGeneric: 'boolean',
                needsConfig: 'boolean', configFields: 'array', processingMode: 'string',
                supportsImageInput: 'boolean', supportsFileInput: 'boolean',
                outputFormat: 'string', llmSystemPrompt: 'string', templateCode: 'string',
                llmUserPromptTemplate: 'string',
                isCustom: 'boolean', tags: 'array', category: 'string',
                isComposite: 'boolean', fieldMappings: 'array',
                workflowSteps: 'array', customWorkflowSteps: 'array',
                inputParameters: 'array', outputParameters: 'array', version: 'number',
                securityLevel: 'string', ragTopK: 'number', ragScoreThreshold: 'number',
                ragFallbackToNormal: 'boolean',
                isQuickShortcut: 'boolean', quickShortcutIcon: 'string',
                createdDate: 'string', lastModifiedDate: 'string',
                createdBy: 'string', lastModifiedBy: 'string'
            };

            // 先处理已知字段（带类型修正）
            for (const key of Object.keys(validFields)) {
                const val = resolveJsonField(parsed, key);
                if (val === undefined || val === null) continue;

                const expectedType = validFields[key];
                if (expectedType === 'array' && !Array.isArray(val)) {
                    cleaned[key] = [];
                } else if (expectedType === 'boolean' && typeof val !== 'boolean') {
                    cleaned[key] = !!val;
                } else if (expectedType === 'string' && typeof val !== 'string') {
                    cleaned[key] = String(val ?? '');
                } else if (expectedType === 'number' && typeof val !== 'number') {
                    cleaned[key] = Number(val) || 1;
                } else {
                    cleaned[key] = val;
                }
            }

            // 透传所有其他字段（不丢弃任何后端字段）
            const knownKeys = Object.keys(validFields);
            for (const key of Object.keys(parsed)) {
                if (knownKeys.indexOf(key.toLowerCase()) === -1 && !cleaned.hasOwnProperty(key)) {
                    cleaned[key] = parsed[key];
                }
            }

            // 补全必填项
            cleaned.name = cleaned.name || resolveJsonField(parsed, 'name') || 'UnnamedSkill';
            cleaned.displayName = cleaned.displayName || cleaned.name;
            cleaned.isCustom = true;

            // 类型修正
            if (Array.isArray(cleaned.keywords) && cleaned.keywords.length === 0) delete cleaned.keywords;
            if (Array.isArray(cleaned.examples) && cleaned.examples.length === 0) delete cleaned.examples;

            // 确保数组字段存在
            ['configFields','tags','fieldMappings','workflowSteps','customWorkflowSteps','inputParameters','outputParameters'].forEach(k => {
                if (cleaned[k] !== undefined && !Array.isArray(cleaned[k])) cleaned[k] = [];
            });

            // 入库前清洗 + 校验：先剔除元指令 examples，再校验 steps 契约与工具名真实性
            sanitizeSkillExamples(cleaned);
            const preStore = await validateStepsBeforeSave(cleaned);
            if (!preStore.ok) throw new Error(preStore.message);

            const cleanJsonStr = JSON.stringify(cleaned);

            // 【修复】统一走 callBridgeMethod：兼容真实 WebView2 与 Web/Mock(Local Bridge HTTP) 两种环境，
            // 避免非 WebView2 环境抛出「Excel 未连接」导致 Skill 未被真实保存。
            const result = await callBridgeMethod('SaveSkill', [cleanJsonStr]);
            const resultObj = JSON.parse(result);
            if (!resultObj.success) {
                throw new Error(resultObj.error || '保存失败');
            }
            if (App.showToast) App.showToast('✅ 自定义 Skill 保存成功', 'success');
            btn.innerHTML = '✅ 已保存';
            btn.style.background = '#9e9e9e';

            // 重新加载 Skill 列表
            if (typeof App.loadCustomSkills === 'function') {
                App.loadCustomSkills();
            }
            if (typeof App.loadSkillsAsCommands === 'function') {
                await App.loadSkillsAsCommands(true);
            }
            // 同 saveSkill：重绘技能管理面板，否则新 Skill 只存在于缓存里、列表看不见。
            if (typeof App.loadSkillsForManagement === 'function') {
                try { await App.loadSkillsForManagement(); } catch (e) { /* 面板未打开时忽略 */ }
            }
        } catch (error) {
            console.error('[saveGeneratedSkill] 保存失败:', error);
            if (App.showToast) App.showToast('❌ 保存失败: ' + error.message, 'error');
            else alert('保存失败: ' + error.message);
            btn.disabled = false;
            btn.innerHTML = '💾 保存到自定义Skill';
        }
    }

    // --- Skill 卡片编辑功能 ---
    
    /**
     * 切换 Skill 卡片的编辑视图
     */
    function editSkillCard(name) {
        const viewEl = document.getElementById('skill-view-' + name);
        const editEl = document.getElementById('skill-edit-' + name);
        if (viewEl && editEl) {
            viewEl.style.display = 'none';
            editEl.style.display = 'block';
        }
    }

    /**
     * 取消编辑，恢复只读视图
     */
    function cancelSkillEdit(name) {
        const viewEl = document.getElementById('skill-view-' + name);
        const editEl = document.getElementById('skill-edit-' + name);
        if (viewEl && editEl) {
            viewEl.style.display = 'block';
            editEl.style.display = 'none';
        }
    }

    /**
     * 应用编辑修改，更新卡片显示
     */
    function applySkillEdit(name) {
        const displayNameInput = document.getElementById('edit-displayName-' + name);
        const categoryInput = document.getElementById('edit-category-' + name);
        const descriptionInput = document.getElementById('edit-description-' + name);

        if (!displayNameInput) {
            if (App.showToast) App.showToast('❌ 找不到编辑字段', 'error');
            return;
        }

        const newDisplayName = displayNameInput.value.trim();
        const newCategory = categoryInput ? categoryInput.value.trim() : '';
        const newDescription = descriptionInput ? descriptionInput.value.trim() : '';

        if (!newDisplayName) {
            if (App.showToast) App.showToast('❌ 名称不能为空', 'error');
            return;
        }

        // 更新卡片只读视图
        const viewEl = document.getElementById('skill-view-' + name);
        if (viewEl) {
            const card = viewEl.closest('.skill-card');
            const header = card ? card.querySelector('div[style*="background:#f8f9fa"]') : null;
            if (header) {
                // 更新头部：名称 + 分类
                header.innerHTML = '<span style="font-weight:600;">' + newDisplayName + '</span>';
                if (newCategory) header.innerHTML += '<span style="color:#999;font-size:12px;">' + newCategory + '</span>';
            }
            // 更新描述
            const descDiv = viewEl.querySelector('div[style*="margin-bottom:8px"]');
            if (descDiv) {
                if (newDescription) descDiv.textContent = newDescription;
                else descDiv.style.display = 'none';
            } else if (newDescription) {
                viewEl.insertBefore(document.createElement('div'), viewEl.firstChild);
                viewEl.firstChild.style.cssText = 'margin-bottom:8px;';
                viewEl.firstChild.textContent = newDescription;
            }
        }

        cancelSkillEdit(name);
        if (App.showToast) App.showToast('✅ 已更新', 'success');
    }

    /**
     * 复制 Skill JSON 到剪贴板
     * 从全局存储 window.__skillJsonStore[name] 读取 JSON
     */
    function copySkillJson(name) {
        const store = window.__skillJsonStore || {};
        const jsonStr = store[name];
        if (!jsonStr) {
            if (App.showToast) App.showToast('❌ 没有可复制的 Skill JSON', 'error');
            return;
        }
        try {
            navigator.clipboard.writeText(jsonStr).then(() => {
                if (App.showToast) App.showToast('✅ Skill JSON 已复制到剪贴板', 'success');
            }).catch(() => {
                // fallback
                const textarea = document.createElement('textarea');
                textarea.value = jsonStr;
                document.body.appendChild(textarea);
                textarea.select();
                document.execCommand('copy');
                document.body.removeChild(textarea);
                if (App.showToast) App.showToast('✅ Skill JSON 已复制', 'success');
            });
        } catch (e) {
            if (App.showToast) App.showToast('❌ 复制失败: ' + e.message, 'error');
        }
    }

    /**
     * 打开 JSON 编辑器（全编辑）
     * 从全局存储 window.__skillJsonStore[name] 读取 JSON，避免内联导致属性被破坏
     */
    function openJsonEditor(name) {
        const store = window.__skillJsonStore || {};
        const jsonStr = store[name];
        if (!jsonStr) {
            if (App.showToast) App.showToast('❌ 找不到 Skill JSON 数据', 'error');
            return;
        }
        try {
            const parsed = JSON.parse(jsonStr);
            const formatted = JSON.stringify(parsed, null, 2);
            
            // 创建模态对话框
            const modalId = 'json-editor-modal-' + name;
            
            // 如果已存在，先移除
            const existing = document.getElementById(modalId);
            if (existing) existing.remove();
            
            const modal = document.createElement('div');
            modal.id = modalId;
            modal.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:9999;display:flex;align-items:center;justify-content:center;';
            
            modal.innerHTML = `
                <div style="background:white;border-radius:8px;width:90%;max-width:800px;max-height:90vh;display:flex;flex-direction:column;box-shadow:0 4px 20px rgba(0,0,0,0.2);">
                    <div style="padding:12px 16px;border-bottom:1px solid #eee;display:flex;justify-content:space-between;align-items:center;">
                        <span style="font-size:14px;font-weight:600;">编辑 Skill JSON - ${name}</span>
                        <button onclick="document.getElementById('${modalId}').remove()" style="background:none;border:none;font-size:18px;cursor:pointer;color:#999;">✕</button>
                    </div>
                    <div style="padding:12px;16px;flex:1;overflow:auto;">
                        <textarea id="json-editor-textarea-${name}" style="width:100%;min-height:400px;padding:10px;border:1px solid #ddd;border-radius:4px;font-family:monospace;font-size:12px;resize:vertical;" spellcheck="false"></textarea>
                    </div>
                    <div style="padding:12px 16px;border-top:1px solid #eee;display:flex;justify-content:space-between;gap:8px;">
                        <div style="color:#999;font-size:12px;align-self:center;">支持编辑完整 JSON（包含 VBA 代码等所有字段）</div>
                        <div style="display:flex;gap:8px;">
                            <button onclick="App.formatJsonInEditor('${name}')" style="padding:6px 14px;background:#607d8b;color:white;border:none;border-radius:4px;cursor:pointer;font-size:12px;">格式化</button>
                            <button onclick="App.cancelJsonEditor('${name}')" style="padding:6px 14px;background:#ccc;color:#666;border:none;border-radius:4px;cursor:pointer;font-size:12px;">取消</button>
                            <button onclick="App.saveJsonEditor('${name}')" style="padding:6px 14px;background:#4CAF50;color:white;border:none;border-radius:4px;cursor:pointer;font-size:12px;">应用</button>
                        </div>
                    </div>
                </div>
            `;
            
            document.body.appendChild(modal);
            
            // 填充 JSON
            setTimeout(() => {
                const textarea = document.getElementById('json-editor-textarea-' + name);
                if (textarea) textarea.value = formatted;
            }, 50);
            
            // 点击背景关闭
            modal.addEventListener('click', (e) => {
                if (e.target === modal) modal.remove();
            });
            
        } catch (e) {
            if (App.showToast) App.showToast('❌ JSON 解析失败: ' + e.message, 'error');
        }
    }

    /**
     * 格式化编辑器中的 JSON
     */
    function formatJsonInEditor(name) {
        const textarea = document.getElementById('json-editor-textarea-' + name);
        if (!textarea) return;
        try {
            const parsed = JSON.parse(textarea.value);
            textarea.value = JSON.stringify(parsed, null, 2);
            if (App.showToast) App.showToast('✅ 已格式化', 'success');
        } catch (e) {
            if (App.showToast) App.showToast('❌ JSON 格式错误', 'error');
        }
    }

    /**
     * 取消 JSON 编辑
     */
    function cancelJsonEditor(name) {
        const modal = document.getElementById('json-editor-modal-' + name);
        if (modal) modal.remove();
    }

    /**
     * 保存 JSON 编辑并应用
     */
    async function saveJsonEditor(name) {
        const textarea = document.getElementById('json-editor-textarea-' + name);
        if (!textarea) return;
        
        let parsed;
        try {
            parsed = JSON.parse(textarea.value);
        } catch (e) {
            if (App.showToast) App.showToast('❌ JSON 格式错误: ' + e.message, 'error');
            return;
        }
        
        // 【修复】Skill 名称可能是 camelCase(name) 或 PascalCase(Name)，兼容两种命名
        const skillNameValue = parsed.name || parsed.Name;
        if (!skillNameValue) {
            if (App.showToast) App.showToast('❌ 缺少 name/Name 字段', 'error');
            return;
        }
        
        // 尝试保存到后端（【修复】统一走 callBridgeMethod，兼容 WebView2 与 Web/Mock）
        try {
            const result = await callBridgeMethod('SaveSkill', [JSON.stringify(parsed)]);
            const resultObj = JSON.parse(result);
            if (resultObj.success) {
                if (App.showToast) App.showToast('✅ Skill 保存成功', 'success');
            } else {
                if (App.showToast) App.showToast('❌ 保存失败: ' + (resultObj.error || '未知错误'), 'error');
            }
        } catch (e) {
            if (App.showToast) App.showToast('❌ 保存失败: ' + e.message, 'error');
        }
        
        // 更新前端卡片显示
        try {
            // 更新保存按钮的 data-skill-json
            const card = document.getElementById('skill-view-' + name);
            if (card) {
                const skillCard = card.closest('.skill-card');
                const saveBtn = skillCard ? skillCard.querySelector('button[onclick^="App.saveGeneratedSkill"]') : null;
                if (saveBtn) {
                    saveBtn.setAttribute('data-skill-json', JSON.stringify(parsed));
                }
                
                // 更新头部显示名称
                const header = skillCard ? skillCard.querySelector('div[style*="background:#f8f9fa"]') : null;
                if (header) {
                    const newName = parsed.displayName || parsed.name || name;
                    const newCategory = parsed.category || '';
                    header.innerHTML = '<span style="font-weight:600;">' + newName + '</span>';
                    if (newCategory) header.innerHTML += '<span style="color:#999;font-size:12px;">' + newCategory + '</span>';
                }
                
                // 更新描述
                const descDiv = card.querySelector('div[style*="margin-bottom:8px"]');
                if (descDiv) {
                    const newDesc = parsed.description || '';
                    if (newDesc) descDiv.textContent = newDesc;
                    else descDiv.style.display = 'none';
                }
            }
        } catch (e) {
            console.warn('更新卡片显示失败:', e);
        }
        
        // 关闭编辑器
        cancelJsonEditor(name);
        
        // 刷新 Skill 列表
        if (typeof App.loadCustomSkills === 'function') {
            App.loadCustomSkills();
        }
    }

    // 暴露到全局
    App.editSkillCard = editSkillCard;
    App.cancelSkillEdit = cancelSkillEdit;
    App.applySkillEdit = applySkillEdit;
    App.copySkillJson = copySkillJson;
    App.openJsonEditor = openJsonEditor;
    App.formatJsonInEditor = formatJsonInEditor;
    App.cancelJsonEditor = cancelJsonEditor;
    App.saveJsonEditor = saveJsonEditor;

    // --- 来自 app.js:L22982-L23008 ---
    // ★ 全局操作锁：防止重复绑定的监听器导致confirm弹多次（取消/确定都只弹一次）
    const _skillActionLock = new Map();
    function _tryAcquireActionLock(actionKey, lockMs = 1500) {
        const now = Date.now();
        const expires = _skillActionLock.get(actionKey) || 0;
        if (now < expires) return false;
        _skillActionLock.set(actionKey, now + lockMs);
        // 自动清理
        setTimeout(() => _skillActionLock.delete(actionKey), lockMs + 100);
        return true;
    }

    async function deleteSkill(skillName) {
        // ★ 防止重复绑定导致 confirm 弹多次：同一skill 1.5秒内只允许一次弹窗
        const lockKey = 'delete:' + skillName;
        if (!_tryAcquireActionLock(lockKey, 1500)) return;

        if (!confirm('确定要删除 Skill "' + skillName + '" 吗？')) {
            return;
        }

        try {
            if (!getBridge()) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">Excel 未连接</span>', true) : App.showToast('Excel 未连接');
                return;
            }

            const result = await callBridgeMethod('DeleteSkill', [skillName]);
            const resultObj = JSON.parse(result);

            if (resultObj.success) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">Skill 删除成功</span>', true) : App.showToast('Skill 删除成功');

                if (typeof App.loadCustomSkills === 'function') {
                    App.loadCustomSkills();
                }

                if (typeof App.loadSkillsAsCommands === 'function') {
                    await App.loadSkillsAsCommands(true);
                }
                if (typeof ChatLog !== 'undefined' && ChatLog.info) {
                    ChatLog.info('SKILL', '🔄 Skill 缓存已刷新');
                }
            } else {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">删除失败: ' + resultObj.error + '</span>', true) : App.showToast('删除失败: ' + resultObj.error);
            }
        } catch (error) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">删除失败: ' + error.message + '</span>', true) : App.showToast('删除失败: ' + error.message);
        }
    }

    // --- 从当前编辑器还原 Skill 为系统默认版本 ---
    async function restoreCurrentSkill() {
        if (!App.currentEditingSkill) {
            App.showToast('没有可恢复的 Skill');
            return;
        }
        const skillName = App.currentEditingSkill.name || App.currentEditingSkill.Name;
        if (!skillName) {
            App.showToast('无法识别 Skill 名称');
            return;
        }

        const lockKey = 'restoreCurrent:' + skillName;
        if (!_tryAcquireActionLock(lockKey, 1500)) return;

        if (!confirm('确定要将 Skill "' + skillName + '" 恢复为系统默认版本吗？\n\n这将删除您的所有自定义修改。')) {
            return;
        }

        try {
            if (!getBridge()) {
                App.showToast('Excel 未连接');
                return;
            }

            const result = await callBridgeMethod('DeleteSkill', [skillName]);
            const resultObj = JSON.parse(result);

            if (resultObj.success) {
                App.showToast('✓ 已恢复为系统默认版本');
                closeSkillDialog();
                if (typeof App.loadCustomSkills === 'function') {
                    App.loadCustomSkills();
                }
                if (typeof App.loadSkillsAsCommands === 'function') {
                    await App.loadSkillsAsCommands(true);
                }
                await editSkill(skillName);
            } else {
                App.showToast('恢复失败: ' + (resultObj.error || '未知错误'));
            }
        } catch (error) {
            App.showToast('恢复失败: ' + error.message);
        }
    }

    // --- 还原被覆盖的系统 Skill ---
    async function restoreSkill(skillName) {
        const lockKey = 'restore:' + skillName;
        if (!_tryAcquireActionLock(lockKey, 1500)) return;

        if (!confirm('确定要将 Skill "' + skillName + '" 还原为系统默认版本吗？\n\n这将删除您的自定义修改。')) {
            return;
        }

        try {
            if (!getBridge()) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">Excel 未连接</span>', true) : App.showToast('Excel 未连接');
                return;
            }

            const result = await callBridgeMethod('DeleteSkill', [skillName]);
            const resultObj = JSON.parse(result);

            if (resultObj.success) {
                App.showToast('✅ Skill 已还原为系统默认版本');

                if (typeof App.loadCustomSkills === 'function') {
                    App.loadCustomSkills();
                }
                if (typeof App.loadSkillsAsCommands === 'function') {
                    await App.loadSkillsAsCommands(true);
                }
            } else {
                App.showToast('还原失败: ' + resultObj.error);
            }
        } catch (error) {
            App.showToast('还原失败: ' + error.message);
        }
    }

    // --- 来自 app.js:L23011-L23038 ---
    async function resetSkill(skillName) {
        const lockKey = 'reset:' + skillName;
        if (!_tryAcquireActionLock(lockKey, 1500)) return;

        if (!confirm('确定要将 Skill "' + skillName + '" 还原为系统默认版本吗？\n\n这将删除您的自定义修改，恢复到初始状态。')) {
            return;
        }

        try {
            if (!getBridge()) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">Excel 未连接</span>', true) : App.showToast('Excel 未连接');
                return;
            }

            // 调用删除接口，删除自定义版本后系统会自动回退到系统版本
            const result = await callBridgeMethod('DeleteSkill', [skillName]);
            const resultObj = JSON.parse(result);

            if (resultObj.success) {
                App.showToast('Skill 已还原为系统默认版本');

                if (typeof App.loadCustomSkills === 'function') {
                    App.loadCustomSkills();
                }

                if (typeof App.loadSkillsAsCommands === 'function') {
                    await App.loadSkillsAsCommands(true);
                }
                if (typeof ChatLog !== 'undefined' && ChatLog.info) {
                    ChatLog.info('SKILL', '🔄 Skill 已还原，缓存已刷新');
                }
            } else {
                App.showToast('还原失败: ' + resultObj.error);
            }
        } catch (error) {
            App.showToast('还原失败: ' + error.message);
        }
    }

    // --- 来自 app.js:L23041-L23055 ---
    function toggleSection(sectionId) {
        const header = typeof sectionId === 'string' ?
            document.querySelector(`[data-section="${sectionId}"]`) :
            sectionId;

        if (!header) return;

        const content = header.nextElementSibling;
        const icon = header.querySelector('.toggle-icon');
        if (content.style.display === 'none') {
            content.style.display = 'block';
            icon.textContent = '▼';
            // 如果是 Skill 管理区域，加载列表
            if (header.textContent.includes('Skill')) {
                if (typeof App.loadCustomSkills === 'function') {
                    App.loadCustomSkills();
                }
            }
        } else {
            content.style.display = 'none';
            icon.textContent = '▶';
        }
    }

    // ==================== 代码库管理函数 ====================

    // 功能指纹映射表 - 用于识别用户需求类型
    // --- 来自 app.js:L23060-L23141 ---
    // ==================== 模块导出 ====================
    
    // 将所有函数挂载到 App 对象上
    window.App = window.App || {};
    
    // 技能管理 UI 函数
    App.switchSkillTab = switchSkillTab;
    App.toggleAdvancedSection = toggleAdvancedSection;
    App.onProcessingModeChange = onProcessingModeChange;
    App.syncFormToJson = syncFormToJson;
    App.collectSkillDataFromForm = collectSkillDataFromForm;
    App.formatSkillJson = formatSkillJson;
    App.validateSkillJson = validateSkillJson;
    App.showAddSkillDialog = showAddSkillDialog;
    App.editSkill = editSkill;
    App.closeSkillDialog = closeSkillDialog;
    App.saveSkill = saveSkill;
    App.deleteSkill = deleteSkill;
    App.saveGeneratedSkill = saveGeneratedSkill;
    App.validateStepsBeforeSave = validateStepsBeforeSave;
    App.sanitizeSkillExamples = sanitizeSkillExamples;
    App.resetSkill = resetSkill;
    App.restoreSkill = restoreSkill;
    App.restoreCurrentSkill = restoreCurrentSkill;
    App.toggleSection = toggleSection;

    // 【修复】Web 版中 index.html 的 renderSkillListDirect 通过 onclick 直接调用全局函数名
    // 同时将常用 Skill 管理函数挂载到 window，避免 HTML 内联事件报未定义
    window.showAddSkillDialog = showAddSkillDialog;
    window.editSkill = editSkill;
    window.closeSkillDialog = closeSkillDialog;
    window.saveSkill = saveSkill;
    window.deleteSkill = deleteSkill;
    window.saveGeneratedSkill = saveGeneratedSkill;
    window.resetSkill = resetSkill;
    window.restoreSkill = restoreSkill;
    window.restoreCurrentSkill = restoreCurrentSkill;

    // 快捷入口管理函数
    App.onQuickShortcutToggle = onQuickShortcutToggle;
    App.toggleQuickShortcut = toggleQuickShortcut;

    /**
     * 快捷入口按钮点击处理：乐观 UI 更新（立即切换外观）+ 后台同步后端
     * 注：快捷 Skill 只渲染在常驻快捷图标栏，不再混入「快捷提示词」区域
     */
    App.onQuickShortcutBtnClick = function(btn) {
        const skillName = btn.getAttribute('data-skill-name');
        if (!skillName) return;

        // 读取当前状态
        const isActive = btn.classList.contains('active');
        const newState = !isActive;   // 切换目标状态

        // 从已渲染的 Skill 卡片中读取图标配置
        let quickIcon = '⚡';
        const allCards = document.querySelectorAll('.skill-card');
        for (let i = 0; i < allCards.length; i++) {
            const qsBtn = allCards[i].querySelector('.skill-quick-btn[data-skill-name="' + skillName + '"]');
            if (qsBtn) {
                // 初始渲染时按钮内容就是图标
                const rawIcon = qsBtn.textContent.trim();
                if (rawIcon) quickIcon = rawIcon;
                break;
            }
        }

        // ====== 立即更新 UI（乐观更新）======
        if (newState) {
            btn.classList.add('active');
            btn.title = '已加入快捷入口（点击移除）';
            btn.innerHTML = quickIcon;
            // 同步更新：常驻快捷图标栏
            appendQuickShortcutBarItem(skillName);
        } else {
            btn.classList.remove('active');
            btn.title = '点击添加到快捷入口';
            btn.innerHTML = '☆';
            // 同步移除：常驻快捷图标栏
            removeQuickShortcutBarItem(skillName);
        }

        // ====== 后台异步同步到后端（失败时回滚）======
        toggleQuickShortcut(skillName, newState).catch(function() {
            // 回滚 UI
            if (newState) {
                btn.classList.remove('active');
                btn.title = '点击添加到快捷入口';
                btn.innerHTML = '☆';
                removeQuickShortcutBarItem(skillName);
            } else {
                btn.classList.add('active');
                btn.title = '已加入快捷入口（点击移除）';
                btn.innerHTML = quickIcon;
                appendQuickShortcutBarItem(skillName);
            }
            App.showToast('❌ 快捷入口同步失败，已回滚');
        });
    };

    /**
     * Skill 开关 toggle 处理：乐观更新卡片样式 + 后台同步后端
     */
    App.onSkillToggleEnabled = async function(skillName, isEnabled) {
        if (!skillName) return;

        // ====== 立即更新卡片 UI（乐观更新）======
        const allCards = document.querySelectorAll('.skill-card');
        let targetCard = null;
        for (let i = 0; i < allCards.length; i++) {
            const card = allCards[i];
            // 通过 meta 区域的 name 匹配
            const metaItems = card.querySelectorAll('.skill-meta-item');
            for (let j = 0; j < metaItems.length; j++) {
                if (metaItems[j].textContent.trim() === skillName) {
                    targetCard = card;
                    break;
                }
            }
            if (targetCard) break;
        }

        if (targetCard) {
            if (isEnabled) {
                targetCard.classList.remove('skill-card-disabled');
                // 移除"已停用"标签
                const disabledTag = targetCard.querySelector('.skill-tag-disabled');
                if (disabledTag) disabledTag.remove();
            } else {
                targetCard.classList.add('skill-card-disabled');
                // 添加"已停用"标签（如果没有）
                const tagsContainer = targetCard.querySelector('.skill-card-tags');
                if (tagsContainer && !tagsContainer.querySelector('.skill-tag-disabled')) {
                    const tag = document.createElement('span');
                    tag.className = 'skill-tag skill-tag-disabled';
                    tag.textContent = '已停用';
                    tagsContainer.appendChild(tag);
                }
            }
        }

        // ====== 后台异步同步到后端 ======
        try {
            const bridge = getBridge();
            if (!bridge || typeof bridge.ToggleSkillEnabled !== 'function') {
                throw new Error('Excel 桥接不可用');
            }
            const result = await bridge.ToggleSkillEnabled(skillName, !!isEnabled);
            const resultObj = JSON.parse(result);
            if (!resultObj.success) {
                throw new Error(resultObj.error || '未知错误');
            }

            // 同步更新管理面板全量缓存（保证切换面板后开关状态不回弹）
            if (Array.isArray(App.allSkillsFull)) {
                App.allSkillsFull.forEach(s => {
                    const n = s && (s.name || s.Name);
                    if (n === skillName) {
                        s.isEnabled = !!isEnabled;
                        if ('IsEnabled' in s) s.IsEnabled = !!isEnabled;
                    }
                });
            }
            // 【停用名单·会话级】显式维护一份"被用户关掉的 Skill 名"。
            // 为什么需要：DESCRIPTORS 是前端静态领域目录、不感知 isEnabled——停用后
            // GetAvailableSkills 不再返回该 Skill，反而让路由层"后端没返回 → 注入空壳"的
            // 分支把它复活（实测：SmartTableBeautifier(kw=27) → beautifier(kw=9) 空壳），
            // LLM 路由目录也照旧列出它。落一份名单供 skill-domain 判定"这是用户主动关的"。
            try {
                const _off = new Set(Array.isArray(App.disabledSkillNames) ? App.disabledSkillNames : []);
                if (isEnabled) _off.delete(skillName); else _off.add(skillName);
                if (Array.isArray(App.allSkillsFull)) {
                    App.allSkillsFull.forEach(s => {
                        const n = s && (s.name || s.Name);
                        if (!n) return;
                        if (s.isEnabled === false || s.IsEnabled === false) _off.add(n); else _off.delete(n);
                    });
                }
                App.disabledSkillNames = Array.from(_off);
                window.disabledSkillNames = App.disabledSkillNames;
            } catch (_e) { /* 名单维护失败不影响开关本身 */ }
            // 刷新路由缓存（确保停用的 Skill 不再参与路由匹配，仅含启用项）
            if (typeof App.preloadAllSkills === 'function') {
                try { await App.preloadAllSkills(); } catch(e) { (window.ChatLog || console).warn?.('SKILL', '[skill-mgmt] 刷新路由缓存失败:', e); }
            }
            // 刷新管理面板（完整列表，含停用项）
            if (typeof App.loadSkillsForManagement === 'function') {
                try { await App.loadSkillsForManagement(); } catch(e) { (window.ChatLog || console).warn?.('SKILL', '[skill-mgmt] 刷新管理面板失败:', e); }
            }
            App.showToast(`✅ Skill 已${isEnabled ? '启用' : '停用'}`);
        } catch (err) {
            console.error('[onSkillToggleEnabled] 失败:', err);
            App.showToast('❌ ' + (err.message || '同步失败'));
            // 回滚 UI
            if (targetCard) {
                if (isEnabled) {
                    targetCard.classList.add('skill-card-disabled');
                    const tagsContainer = targetCard.querySelector('.skill-card-tags');
                    if (tagsContainer && !tagsContainer.querySelector('.skill-tag-disabled')) {
                        const tag = document.createElement('span');
                        tag.className = 'skill-tag skill-tag-disabled';
                        tag.textContent = '已停用';
                        tagsContainer.appendChild(tag);
                    }
                } else {
                    targetCard.classList.remove('skill-card-disabled');
                    const disabledTag = targetCard.querySelector('.skill-tag-disabled');
                    if (disabledTag) disabledTag.remove();
                }
                // 回滚 toggle checkbox
                const toggle = targetCard.querySelector('.skill-toggle-switch input');
                if (toggle) toggle.checked = !isEnabled;
            }
        }
    };

    /**
     * 在常驻快捷图标栏（模型选择器上方）追加一个图标项（乐观更新用）
     */
    function appendQuickShortcutBarItem(skillName) {
        const barInner = document.getElementById('quick-shortcut-bar-inner');
        if (!barInner) return;
        // 避免重复
        if (barInner.querySelector('[data-qs-name="' + skillName + '"]')) return;
        // 移除空状态提示
        const emptyHint = barInner.querySelector('.qs-empty-hint');
        if (emptyHint) emptyHint.remove();

        // 从技能列表中获取信息
        let displayName = skillName, icon = '⚡', promptText = skillName;
        const allItems = document.querySelectorAll('.skill-item');
        for (let i = 0; i < allItems.length; i++) {
            const nameEl = allItems[i].querySelector('.skill-item-meta span:first-child');
            if (nameEl && nameEl.textContent === ('ID: ' + skillName)) {
                displayName = allItems[i].querySelector('.skill-item-name').childNodes[0].textContent.trim();
                const descEl = allItems[i].querySelector('.skill-item-description');
                promptText = descEl ? descEl.textContent.trim() : displayName;
                break;
            }
        }

        const item = document.createElement('span');
        item.className = 'qs-item';
        item.setAttribute('data-qs-name', skillName);
        item.title = promptText;
        item.style.animation = 'fadeIn 0.25s ease';
        item.innerHTML = '<span class="qs-icon">' + icon + '</span><span class="qs-label">' + displayName + '</span>';
        item.onclick = function() { App.triggerQuickSkill(skillName, promptText); };
        barInner.appendChild(item);
    }

    /**
     * 从常驻快捷图标栏移除指定图标项（乐观更新用）
     */
    function removeQuickShortcutBarItem(skillName) {
        const barInner = document.getElementById('quick-shortcut-bar-inner');
        if (!barInner) return;
        const existing = barInner.querySelector('[data-qs-name="' + skillName + '"]');
        if (existing) {
            existing.style.opacity = '0';
            existing.style.transform = 'scale(0.8)';
            existing.style.transition = 'all 0.2s ease';
            setTimeout(function() { existing.remove(); }, 200);
        }
        // 如果栏空了，显示空状态提示
        setTimeout(function() {
            if (barInner.children.length === 0 || !barInner.querySelector('.qs-item')) {
                barInner.innerHTML = '<span class="qs-empty-hint">在 Skill 管理中标记「⭐快捷入口」即可在此显示</span>';
            }
        }, 250);
    }

    /**
     * 快速切换 Skill 的快捷入口状态（从技能列表直接操作，无需打开编辑器）
     */
    async function toggleQuickShortcut(skillName, isQuick) {
        const bridge = getBridge();
        if (!bridge) {
            App.showToast('后端连接不可用');
            return;
        }
        try {
            const result = await bridge.ToggleQuickShortcut(skillName, isQuick);
            const data = JSON.parse(result);
            if (data.success) {
                App.showToast((isQuick ? '✅ 已添加到' : '❌ 已移出') + '快捷入口: ' + skillName);
                // 刷新常驻快捷图标栏
                if (window.App && window.App.renderQuickShortcutBar) {
                    await window.App.renderQuickShortcutBar();
                }
            } else {
                App.showToast('操作失败: ' + (data.error || '未知错误'));
            }
        } catch (e) {
            App.showToast('快捷入口切换异常: ' + e.message);
        }
    }

    // 代码库管理函数

    // ==================== Skill 通用工具函数 ====================
    // getSkillProp(skill, key1, key2, ..., fallbackValue)
    // 遍历所有 key 返回第一个有效值；最后一个非字符串参数视为 fallback
    // 示例: getSkillProp(skill, 'uiActions', 'UIActions', {}) 返回 uiActions 或 {}
    //       getSkillProp(skill, 'name', 'Name') 无默认值，找不到则返回 undefined
    function getSkillProp(skill) {
        if (skill === null || skill === undefined || typeof skill !== 'object') {
            for (var a = arguments.length - 1; a >= 1; a--) {
                if (typeof arguments[a] !== 'string') return arguments[a];
            }
            return undefined;
        }
        for (var i = 1; i < arguments.length; i++) {
            var k = arguments[i];
            if (typeof k !== 'string') continue;
            var v = skill[k];
            if (v !== undefined && v !== null) return v;
        }
        for (var j = arguments.length - 1; j >= 1; j--) {
            if (typeof arguments[j] !== 'string') return arguments[j];
        }
        return undefined;
    }

    function renderSkillListDirect(container, skills) {
        console.log('[renderSkillListDirect] called with', skills?.length, 'skills');
        if (!container) return;
        if (!skills || skills.length === 0) {
            console.log('[renderSkillListDirect] skills is empty, showing empty message');
            container.innerHTML = `
                <div class="skill-list-header">
                    <span>自定义 Skill</span>
                    <button class="btn-add-skill" onclick="showAddSkillDialog()">+ 添加 Skill</button>
                </div>
                <div style="text-align:center;color:#999;padding:20px;">暂无可用 Skills</div>`;
            return;
        }

        const categoryIcons = {
            '数据处理': '📊', '数据分析': '📈', '效率工具': '🛠️', '高级功能': '⚙️',
            '其他': '📦', '代码片段': '📝', '智能识别': '🤖', '自动化': '⚡'
        };
        const skillIcons = {
            '营业执照': '📜', '银行回单': '🏦', '名片': '👤', '驾驶证': '🪪', '快递单': '📦',
            '身份证': '🆔', '增值税发票': '🧾', '电子发票': '📄', '火车票': '🚄', '行驶证': '🚙',
            '合同': '📋', '表格': '📊', '图表': '📈', '分析': '🔍', '统计': '📉', '汇总': '📝',
            '转换': '🔄', '格式化': '✨', '拆分': '✂️', '合并': '🔗', '填充': '🖍️', '筛选': '🔎',
            '排序': '⬆️', '删除': '🗑️', '复制': '📋', '搜索': '🔍', '替换': '🔄', '对比': '⚖️',
            '计算': '🧮', '求和': '➕', '平均值': '📊', '最大值': '🔝', '最小值': '🔻', '计数': '🔢',
            '生成': '✨', '创建': '➕', '插入': '📥', '提取': '📤', '导出': '📤', '导入': '📥',
            '批量': '📦', '自动': '⚡', 'OCR': '👁️', '识别': '🔍', '截图': '📷', '图片': '🖼️',
            'PDF': '📕', 'Word': '📘', 'Excel': '📗', 'PowerPoint': '📙', '邮件': '📧', '报告': '📝',
            '配置': '⚙️', '设置': '⚙️', '帮助': '❓', '模板': '📋'
        };

        function getSkillIcon(skill) {
            const name = String(getSkillProp(skill, 'displayName', 'DisplayName', 'name', 'Name') || '');
            if (skillIcons[name]) return skillIcons[name];
            for (const kw in skillIcons) {
                if (skillIcons.hasOwnProperty(kw) && name.indexOf(kw) !== -1) return skillIcons[kw];
            }
            const cat = String(getSkillProp(skill, 'category', 'Category') || '其他');
            return categoryIcons[cat] || '📦';
        }

        const snippets = skills.filter(s =>
            getSkillProp(s, 'isCustom', 'IsCustom') &&
            /^snippet_/i.test(String(getSkillProp(s, 'name', 'Name') || ''))
        );
        const customSkills = skills.filter(s =>
            getSkillProp(s, 'isCustom', 'IsCustom') &&
            !/^snippet_/i.test(String(getSkillProp(s, 'name', 'Name') || ''))
        );
        const systemSkills = skills.filter(s =>
            !getSkillProp(s, 'isCustom', 'IsCustom') &&
            !/^snippet_/i.test(String(getSkillProp(s, 'name', 'Name') || ''))
        );

        // 【调试日志】输出分类结果
        console.log('[renderSkillListDirect] 分类结果: snippets=', snippets.length, 'customSkills=', customSkills.length, 'systemSkills=', systemSkills.length);
        customSkills.forEach(s => console.log('[renderSkillListDirect]   自定义:', getSkillProp(s, 'name', 'Name'), 'isCustom=', getSkillProp(s, 'isCustom', 'IsCustom')));
        skills.forEach(s => {
            const name = getSkillProp(s, 'name', 'Name');
            const isCustom = getSkillProp(s, 'isCustom', 'IsCustom');
            console.log(`[renderSkillListDirect]   skill: ${name}, isCustom=${isCustom}, raw keys=${Object.keys(s).join(',')}`);
        });

        function groupByCategory(list) {
            const groups = {};
            list.forEach(skill => {
                const cat = String(getSkillProp(skill, 'category', 'Category') || '其他');
                if (!groups[cat]) groups[cat] = [];
                groups[cat].push(skill);
            });
            return groups;
        }

        function renderSkillCard(skill) {
            const name = String(getSkillProp(skill, 'name', 'Name') || '未命名');
            const displayName = String(getSkillProp(skill, 'displayName', 'DisplayName') || name);
            const description = String(getSkillProp(skill, 'description', 'Description') || '无描述');
            const isCustom = !!getSkillProp(skill, 'isCustom', 'IsCustom');
            const isSnippet = /^snippet_/i.test(name);
            const isSystem = !isCustom && !isSnippet;
            const isOverridden = isCustom && systemSkills.some(ss => (getSkillProp(ss, 'name', 'Name') || '') === name);
            const keywords = getSkillProp(skill, 'keywords', 'Keywords') || [];
            const isQuickShortcut = !!getSkillProp(skill, 'isQuickShortcut', 'IsQuickShortcut');
            const quickIcon = String(getSkillProp(skill, 'quickShortcutIcon', 'QuickShortcutIcon') || '⚡');
            const skillIcon = getSkillIcon(skill);
            const isEnabled = getSkillProp(skill, 'isEnabled', 'IsEnabled', true);

            const actionButtons = `
                <button class="skill-action-btn skill-action-edit" data-action="edit" data-skill-name="${escapeHtml(name)}" title="编辑">✏️</button>
                ${isCustom ? `<button class="skill-action-btn skill-action-delete" data-action="delete" data-skill-name="${escapeHtml(name)}" title="删除">🗑️</button>` : ''}
                ${isOverridden ? `<button class="skill-action-btn skill-action-restore" data-action="restore" data-skill-name="${escapeHtml(name)}" title="还原">🔙</button>` : ''}
            `;
            const quickBtn = `
                <button class="skill-quick-btn ${isQuickShortcut ? 'active' : ''}"
                        data-skill-name="${escapeHtml(name)}"
                        onclick="App.onQuickShortcutBtnClick(this)"
                        title="${isQuickShortcut ? '已加入快捷入口（点击移除）' : '点击添加到快捷入口'}">
                    ${isQuickShortcut ? quickIcon : '☆'}
                </button>
            `;

            // Skill 开关 toggle
            const toggleId = 'skill-toggle-' + name.replace(/[^a-zA-Z0-9]/g, '_');
            const toggleBtn = `
                <label class="skill-toggle-switch" title="${isEnabled ? '已启用（点击停用，停用后不参与路由匹配）' : '已停用（点击启用）'}">
                    <input type="checkbox" id="${toggleId}"
                           ${isEnabled ? 'checked' : ''}
                           onchange="App.onSkillToggleEnabled('${escapeHtml(name)}', this.checked)">
                    <span class="skill-toggle-slider"></span>
                </label>
            `;

            return `
                <div class="skill-card${isEnabled ? '' : ' skill-card-disabled'}">
                    <div class="skill-card-content">
                        <div class="skill-card-tags">
                            ${isSnippet ? '<span class="skill-tag skill-tag-snippet">代码片段</span>' : ''}
                            ${isCustom && !isSnippet ? '<span class="skill-tag skill-tag-custom">自定义</span>' : ''}
                            ${isSystem ? '<span class="skill-tag skill-tag-system">系统</span>' : ''}
                            ${isOverridden ? '<span class="skill-tag skill-tag-overridden">已覆盖</span>' : ''}
                            ${!isEnabled ? '<span class="skill-tag skill-tag-disabled">已停用</span>' : ''}
                        </div>
                        <div class="skill-card-header">
                            <div class="skill-card-icon">${skillIcon}</div>
                            <div class="skill-card-name">${escapeHtml(displayName)}</div>
                            <div class="skill-card-header-actions">
                                ${actionButtons}
                                ${quickBtn}
                                ${toggleBtn}
                            </div>
                        </div>
                        <div class="skill-card-desc">${escapeHtml(description)}</div>
                        <div class="skill-card-meta">
                            <span class="skill-meta-item">${escapeHtml(name)}</span>
                            ${Array.isArray(keywords) && keywords.length > 0 ? `<span class="skill-meta-item">${escapeHtml(keywords.join(', '))}</span>` : ''}
                        </div>
                    </div>
                </div>
            `;
        }

        function renderCategoryGroup(categoryName, skillsInCategory) {
            const icon = categoryIcons[categoryName] || '📦';
            return `
                <div class="skill-category-group">
                    <div class="skill-category-title">
                        <span class="skill-category-icon">${icon}</span>
                        <span class="skill-category-name">${escapeHtml(categoryName)}</span>
                        <span class="skill-category-count">${skillsInCategory.length} 个</span>
                    </div>
                    <div class="skill-category-items">
                        ${skillsInCategory.map(s => renderSkillCard(s)).join('')}
                    </div>
                </div>
            `;
        }

        const categoryOrder = ['数据处理', '数据分析', '效率工具', '高级功能', '智能识别', '自动化', '代码片段', '其他'];
        function sortCategories(a, b) {
            const ai = categoryOrder.indexOf(a);
            const bi = categoryOrder.indexOf(b);
            if (ai === -1 && bi === -1) return a.localeCompare(b);
            if (ai === -1) return 1;
            if (bi === -1) return -1;
            return ai - bi;
        }

        let html = `
            <div class="skill-list-header">
                <span>Skill 管理</span>
                <button class="btn-add-skill" onclick="showAddSkillDialog()">+ 添加 Skill</button>
            </div>`;

        const systemGroups = groupByCategory(systemSkills);
        const systemCategories = Object.keys(systemGroups).sort(sortCategories);
        if (systemCategories.length > 0) {
            html += `<div class="skill-section-title">🔧 系统技能</div>`;
            systemCategories.forEach(cat => { html += renderCategoryGroup(cat, systemGroups[cat]); });
        }

        const customGroups = groupByCategory(customSkills);
        const customCategories = Object.keys(customGroups).sort(sortCategories);
        if (customCategories.length > 0) {
            html += `<div class="skill-section-title">✨ 自定义技能</div>`;
            customCategories.forEach(cat => { html += renderCategoryGroup(cat, customGroups[cat]); });
        }

        const snippetGroups = groupByCategory(snippets);
        const snippetCategories = Object.keys(snippetGroups).sort(sortCategories);
        if (snippetCategories.length > 0) {
            html += `<div class="skill-section-title">📝 代码片段</div>`;
            snippetCategories.forEach(cat => { html += renderCategoryGroup(cat, snippetGroups[cat]); });
        }

        container.innerHTML = html;

        // 事件委托：统一处理 Skill 卡片的编辑/删除/还原按钮点击（★ 只绑定一次，避免多次渲染导致confirm弹N次）
        if (!container.dataset.actionEventsBound) {
            container.addEventListener('click', function(e) {
                const btn = e.target.closest('.skill-action-btn');
                if (!btn) return;
                e.stopPropagation();
                const action = btn.getAttribute('data-action');
                const skillName = btn.getAttribute('data-skill-name');
                if (!action || !skillName) return;

                if (action === 'edit' && typeof window.editSkill === 'function') {
                    window.editSkill(skillName);
                } else if (action === 'delete' && typeof window.deleteSkill === 'function') {
                    window.deleteSkill(skillName);
                } else if (action === 'restore' && typeof window.restoreSkill === 'function') {
                    window.restoreSkill(skillName);
                }
            });
            container.dataset.actionEventsBound = '1';
        }

        // 事件委托：添加 Skill 按钮（★ 只绑定一次）
        if (!container.dataset.addEventsBound) {
            container.addEventListener('click', function(e) {
                const btn = e.target.closest('.btn-add-skill');
                if (!btn) return;
                e.stopPropagation();
                if (typeof window.showAddSkillDialog === 'function') {
                    window.showAddSkillDialog();
                }
            });
            container.dataset.addEventsBound = '1';
        }

        // ★ 【性能】preloadAllSkills 非阻塞：不要 await（避免 skill 页面首次渲染卡顿）
        if (typeof App.preloadAllSkills === 'function') {
            App.preloadAllSkills().catch(() => {});
        }
    }

    // 供 app-model.js / skill-shared-loader.js 直接调用
    window.renderSkillListDirect = renderSkillListDirect;
    App.renderSkillListDirect = renderSkillListDirect;

    // ★ Skill 管理面板专用加载器
    // 必须用 GetAllSkills（完整列表，含停用项 + isEnabled 字段）。
    // 不能用 App.cachedSkills：它来自 GetAvailableSkills，只含启用项，
    // 停用的 Skill 会在此处直接消失；若走 __sharedSkillsData 兜底（该数据无 isEnabled 字段），
    // 又会被当成"启用"，表现为"切换开关后换界面就回弹"。
    App.loadSkillsForManagement = async function() {
        const renderNow = (list) => {
            const c = document.getElementById('skill-list-container');
            if (c && Array.isArray(list) && list.length > 0 && typeof renderSkillListDirect === 'function') {
                renderSkillListDirect(c, list);
            }
        };

        // 1) 先用已有全量缓存秒出
        renderNow(App.allSkillsFull);

        // 2) 拉取后端完整列表（含停用项）
        try {
            const bridge = getBridge();
            if (bridge && typeof bridge.GetAllSkills === 'function') {
                const raw = await bridge.GetAllSkills();
                const parsed = typeof raw === 'string' ? JSON.parse(raw) : raw;
                if (parsed && parsed.error === 'bridge-call-failed') {
                    /* 桥投影失败：走下方 fallback，不打红字 */
                } else if (parsed && parsed.success && Array.isArray(parsed.skills)) {
                    App.allSkillsFull = parsed.skills;
                    window.allSkillsFull = parsed.skills;
                    // 同步"停用名单"（供 skill-domain.isDomainDisabledByUser / isSkillNameDisabledByUser，
                    // 避免被用户停用的领域被 DESCRIPTORS 空壳复活、或被 LLM 路由目录选中）
                    try {
                        App.disabledSkillNames = parsed.skills
                            .filter(s => s && (s.isEnabled === false || s.IsEnabled === false))
                            .map(s => s.name || s.Name)
                            .filter(Boolean);
                        window.disabledSkillNames = App.disabledSkillNames;
                    } catch (_e) { /* 名单同步失败不影响面板 */ }
                    renderNow(App.allSkillsFull);
                    return App.allSkillsFull;
                }
            }
        } catch (e) {
            const msg = (e && e.message) ? e.message : String(e || '');
            if (!/e is not a function|0x80020006|未知名称/i.test(msg)) {
                console.warn('[loadSkillsForManagement] GetAllSkills 失败:', msg);
            }
        }

        // 3) 回退：VSTO 用 cachedSkills（仅启用），浏览器用 __sharedSkillsData
        const fallback = (App.cachedSkills && App.cachedSkills.length > 0)
            ? App.cachedSkills
            : ((window.__sharedSkillsData && window.__sharedSkillsData.length > 0) ? window.__sharedSkillsData : []);
        renderNow(fallback);
        return fallback;
    };

    // 新版 loadCustomSkills 的入口：保存/删除后触发刷新
    App.skillMgmt = {
        refresh: async function(forceRefresh = false) {
            console.log('[skillMgmt.refresh] called, forceRefresh=', forceRefresh);
            const container = document.getElementById('skill-list-container');
            if (!container) {
                return;
            }
            // 立即用已有全量缓存渲染（UI 秒出）
            if (Array.isArray(App.allSkillsFull) && App.allSkillsFull.length > 0) {
                renderSkillListDirect(container, App.allSkillsFull);
            }
            // 后台拉取完整列表（含停用项），保证开关状态准确
            if (typeof App.loadSkillsForManagement === 'function') {
                App.loadSkillsForManagement().catch(() => {});
            }
            // 同步刷新路由用缓存（仅启用项）
            if (typeof App.preloadAllSkills === 'function') {
                App.preloadAllSkills().catch(() => {});
            }
        }
    };

    window.diagnoseCustomSkills = async function() {
        console.log('========== 自定义Skill诊断 ==========');
        console.log('1. App.cachedSkills 状态:', App.cachedSkills?.length, '个');
        if (App.cachedSkills && App.cachedSkills.length > 0) {
            const customInCache = App.cachedSkills.filter(s => s.isCustom || s.IsCustom);
            const systemInCache = App.cachedSkills.filter(s => !(s.isCustom || s.IsCustom));
            console.log('   缓存中: 自定义=', customInCache.length, '系统=', systemInCache.length);
            customInCache.forEach(s => console.log('   [自定义]', s.name||s.Name, 'isCustom=', s.isCustom, 'IsCustom=', s.IsCustom));
        }
        console.log('2. window.__sharedSkillsData 状态:', window.__sharedSkillsData?.length, '个');
        try {
            const bridge = window.chrome?.webview?.hostObjects?.excelBridge;
            if (bridge) {
                console.log('3. 调用后端 GetAllSkills...');
                const raw = await bridge.GetAllSkills();
                const data = JSON.parse(raw);
                console.log('   后端返回: success=', data.success, 'skills数量=', data.skills?.length);
                if (data.skills) {
                    const customFromBackend = data.skills.filter(s => s.isCustom || s.IsCustom);
                    console.log('   后端自定义skill:', customFromBackend.length, '个');
                    customFromBackend.forEach(s => console.log('   [后端自定义]', s.name||s.Name, 'isCustom=', s.isCustom));
                }
                console.log('4. 调用后端 GetCustomSkillDebugInfo...');
                try {
                    const debugRaw = await bridge.GetCustomSkillDebugInfo();
                    const debugData = JSON.parse(debugRaw);
                    console.log('   CustomSkillDebugInfo:', debugData);
                } catch (e) {
                    console.warn('   GetCustomSkillDebugInfo 失败:', e.message);
                }
            } else {
                console.warn('3. WebView2 bridge 不可用');
            }
        } catch (e) {
            console.error('后端调用失败:', e);
        }
        console.log('5. DOM 容器状态:');
        const container = document.getElementById('skill-list-container');
        console.log('   skill-list-container 存在:', !!container);
        if (container) {
            console.log('   innerHTML 长度:', container.innerHTML?.length);
            const customSection = container.querySelector('.skill-section-title');
            console.log('   自定义技能标题存在:', !!customSection);
        }
        console.log('========== 诊断完成 ==========');
    };

    // 兼容 index.html 中 onclick 直接调用

    // 常量

})();