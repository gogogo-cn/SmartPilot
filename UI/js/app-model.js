/**
 * ============================================================
 * app-model.js — AI Provider & Model Manager
 * ============================================================
 * 功能：AI Provider 与模型管理模块，管理各服务商配置、模型列表、视觉模型检测与自动选择逻辑。
 *
 * 核心类/函数：
 *   - getProviderConfig() — 获取 Provider 配置
 *   - resolveActualModel() — 解析实际模型
 *   - isVisionModel() — 检测视觉模型
 *   - autoSelectModel() — 自动选择模型
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 *   - app-config.js
 * ============================================================
 */

(function() {
    'use strict';

    function normalizeUiLocale(raw) {
        try {
            if (window.App && App.i18n && typeof App.i18n.normalizeLocale === 'function') {
                return App.i18n.normalizeLocale(raw);
            }
        } catch (_) { /* ignore */ }
        return raw === 'en' ? 'en' : 'zh';
    }

    // ==================== 模型能力关键词 ====================
    const VISION_KEYWORDS = [
        'vl', 'vision', 'visual',
        'gpt-4v', 'gpt-4o', 'gpt-4.1', 'gpt-5.4',
        'claude-3', 'claude-4',
        'qwen-vl', 'qwen2.5-vl', 'qwen3-vl', 'qwen3-vt',
        'llava', 'bakllava', 'minicpm-v', 'moondream',
        'gemma3', 'gemini', 'pixtral', 'grok-vision', 'internvl',
        'kimi-k3', 'kimi-k2',
        'glm-5v', 'glm-4.6v', 'glm-4.1v'
    ];

    const REASONING_KEYWORDS = [
        'r1', 'reasoning', 'thinking',
        'o1', 'o3', 'o4',
        'kimi-k3', 'kimi-k2',
        'qwen3', 'qwq', 'qwen2.5:32b', 'qwen2.5:72b'
    ];

    /**
     * 检测模型名称是否视觉模型
     * @param {string} modelName
     * @returns {boolean}
     */
    function isVisionModel(modelName) {
        if (!modelName) return false;
        const lower = String(modelName).toLowerCase();
        return VISION_KEYWORDS.some(keyword => lower.includes(keyword));
    }

    /**
     * 检测模型名称是否推理模型
     * @param {string} modelName
     * @returns {boolean}
     */
    function isReasoningModel(modelName) {
        if (!modelName) return false;
        const lower = String(modelName).toLowerCase();
        return REASONING_KEYWORDS.some(keyword => lower.includes(keyword));
    }

    // ==================== 提供商可用性检测 ====================

    // --- 来自 app.js:L2040-L2101 ---
    /**
     * 判断本地检测结果中是否存在指定来源的模型
     * @param {string} source - 'ollama' | 'llamacpp'
     */
    function hasDetectedLocalModels(source) {
        return detectedModels.some(m => (m.provider || 'ollama') === source);
    }

    function isProviderAvailable(providerKey) {
        const providerConfig = PROVIDER_CONFIGS[providerKey];
        if (!providerConfig) {
            ChatLog.debug('CONFIG', `    ├─ ${providerKey}: 配置不存在`);
            return false;
        }

        // 本地服务（Ollama / llama.cpp）：需要检测到对应来源的模型，或已配置 localhost URL
        if (providerKey === 'ollama' || providerKey === 'llamacpp') {
            if (hasDetectedLocalModels(providerKey)) {
                ChatLog.debug('CONFIG', `    ├─ ${providerKey}: ✅ 已检测到该来源的本地模型`);
                return true;
            }
            // 只按该提供商自身的 apiUrl 判断是否"可用但待检测"，
            // 绝不能复用全局 config.apiUrl（否则 llamacpp 会借用 ollama 的 URL 被误判）；
            // 且这里必须是纯判断，禁止异步触发 detectLocalModels——
            // 否则 updateChatModelSelector→getAvailableProviders 会反复排队检测，
            // 与 detectOllamaModels 成功后的 updateChatModelSelector 构成无限循环。
            const localUrl = providerConfig.apiUrl;
            if (localUrl && localUrl.includes('localhost')) {
                ChatLog.debug('CONFIG', `    ├─ ${providerKey}: ⚠️ 已配置URL但未检测到模型`);
                return true; // 暂时返回可用，让用户可以刷新
            }
            ChatLog.debug('CONFIG', `    ├─ ${providerKey}: ❌ 未配置URL且未检测到模型`);
            return false;
        }

        // 云端提供商：需要有有效的 API Key
        if (providerConfig.needsKey) {
            // 检查当前配置中是否有该提供商的 API Key
            const providerSettings = config.providers?.[providerKey];
            const savedApiKey = providerSettings?.apiKey;

            // 同时检查表单中当前的值（用户可能正在填写）
            const currentApiKey = document.getElementById('api-key')?.value || '';
            const currentProvider = document.getElementById('api-type')?.value || '';

            ChatLog.debug('CONFIG', `    ├─ ${providerKey}: 检查API Key...`);
            ChatLog.debug('CONFIG', `    │   ├─ 需要API Key: ${providerConfig.needsKey}`);
            ChatLog.debug('CONFIG', `    │   ├─ 已保存的Key: ${savedApiKey ? '有(' + savedApiKey.substring(0, 10) + '...)' : '无'}`);
            ChatLog.debug('CONFIG', `    │   ├─ 当前表单Key: ${currentApiKey ? '有(' + currentApiKey.substring(0, 10) + '...)' : '无'}`);
            ChatLog.debug('CONFIG', `    │   └─ 当前选中提供商: ${currentProvider}`);

            // 如果是当前选中的提供商，检查表单中的值
            if (currentProvider === providerKey && currentApiKey.trim()) {
                ChatLog.debug('CONFIG', `    ├─ ${providerKey}: ✅ 表单中已填写 API Key`);
                return true;
            }

            // 否则检查已保存的配置
            if (!savedApiKey || !savedApiKey.trim()) {
                ChatLog.debug('CONFIG', `    ├─ ${providerKey}: ❌ 没有配置 API Key`);
                return false;
            }

            ChatLog.debug('CONFIG', `    ├─ ${providerKey}: ✅ 已配置 API Key`);
            return true;
        }

        // 不需要 Key 的提供商（如某些本地服务）默认可用
        ChatLog.debug('CONFIG', `    ├─ ${providerKey}: ✅ 不需要API Key，默认可用`);
        return true;
    }

    // --- 来自 app.js:L2104-L2128 ---
    function getControlFlowHint(lineContent) {
        const hints = {
            'for': 'For循环需要配对Next，请使用"执行代码"运行完整代码块',
            'foreach': 'ForEach循环需要配对Next，请使用"执行代码"运行完整代码块',
            'while': 'While循环需要配对Wend，请使用"执行代码"运行完整代码块',
            'do': 'Do循环需要配对Loop，请使用"执行代码"运行完整代码块',
            'loop': 'Loop需要配对Do，请使用"执行代码"运行完整代码块',
            'if': 'If语句需要配对End If，请使用"执行代码"运行完整代码块',
            'else': 'Else需要配合If使用，请使用"执行代码"运行完整代码块',
            'select': 'Select Case需要配对End Select，请使用"执行代码"运行完整代码块',
            'case': 'Case需要在Select Case内使用，请使用"执行代码"运行完整代码块',
            'with': 'With需要配对End With，请使用"执行代码"运行完整代码块',
            'next': 'Next是循环结束标记，会自动跳过',
            'wend': 'Wend是While结束标记，会自动跳过',
            'end': 'End是结构结束标记，会自动跳过',
        };

        for (let [key, hint] of Object.entries(hints)) {
            if (lineContent.startsWith(key)) {
                return hint;
            }
        }

        return '此语句为控制流的一部分，建议使用"执行代码"按钮运行完整代码';
    }

    // --- 来自 app.js:L2134-L2143 ---
    function getAvailableProviders() {
        ChatLog.info('CONFIG', '🔍 检查可用提供商...');
        const providers = Object.keys(PROVIDER_CONFIGS).filter(providerKey => {
            const isAvailable = isProviderAvailable(providerKey);
            ChatLog.debug('CONFIG', `  ├─ ${providerKey}: ${isAvailable ? '✅ 可用' : '❌ 不可用'}`);
            return isAvailable;
        });
        ChatLog.info('CONFIG', `  └─ 可用提供商: [${providers.join(', ')}]`);
        return providers;
    }

    // ==================== 对话框模型选择 ====================

    /**
     * 提取模型家族 key：同一基础模型的不同 size/context/tag 版本归为一组
     * @param {string} modelName
     * @returns {string|null} 家族 key，无法聚合时返回 null
     */
    function getModelFamilyKey(modelName) {
        // Ollama: name:tag
        if (modelName.includes(':')) {
            return modelName.split(':')[0];
        }
        // 云端 context-size 变体：xxx-v1-8k / xxx-v1-128k-vision-preview
        const match = modelName.match(/^(.+)-(8k|32k|128k|256k)(-vision-preview|-vision)?$/i);
        if (match) {
            return match[1] + (match[3] || '');
        }
        return null;
    }

    /**
     * 解析模型的 size/context 信息，用于排序选择代表版本
     * @param {string} modelName
     * @returns {{size:string, context:string}}
     */
    function parseModelSizeContext(modelName) {
        if (modelName.includes(':')) {
            return { size: modelName.split(':')[1] || '', context: '' };
        }
        const match = modelName.match(/^(.+)-(8k|32k|128k|256k)(-vision-preview|-vision)?$/i);
        if (match) {
            return { size: '', context: match[2] };
        }
        return { size: '', context: '' };
    }

    /**
     * 按模型家族聚合模型列表：把同一基础模型的多个 size/context/tag 版本合并为一个代表项
     * @param {{value:string,label:string,provider:string,isVision:boolean,isReasoning?:boolean}[]} models
     * @returns {{value:string,label:string,provider:string,isVision:boolean,isReasoning?:boolean}[]}
     */
    function aggregateModelFamilies(models) {
        const families = {};
        const standalone = [];

        models.forEach(model => {
            const familyKey = getModelFamilyKey(model.value);
            if (familyKey) {
                if (!families[familyKey]) families[familyKey] = [];
                families[familyKey].push(model);
            } else {
                standalone.push(model);
            }
        });

        // 参数从小到大排序，帮助选择最轻量/经济的代表
        const sizeOrder = {
            '1.5b': 1, '4b': 2, '7b': 3, '8b': 4, '9b': 5,
            '12b': 6, '13b': 7, '14b': 8, '16b': 9,
            '27b': 10, '32b': 11, '70b': 12, '72b': 13
        };
        const contextOrder = { '8k': 1, '32k': 2, '128k': 3, '256k': 4 };

        const aggregated = Object.keys(families).map(familyKey => {
            const group = families[familyKey];
            let rep = group.find(m => m.value === config.model);

            if (!rep) {
                rep = group.slice().sort((a, b) => {
                    const aParts = parseModelSizeContext(a.value);
                    const bParts = parseModelSizeContext(b.value);
                    const aSize = sizeOrder[aParts.size] || 999;
                    const bSize = sizeOrder[bParts.size] || 999;
                    if (aSize !== bSize) return aSize - bSize;
                    return (contextOrder[aParts.context] || 999) - (contextOrder[bParts.context] || 999);
                })[0];
            }

            const displayName = familyKey.replace(/-preview$/i, '');
            // 本地/视觉/推理标识统一由 updateChatModelSelector 用图标添加，这里不再拼接文字标签

            return {
                value: rep.value,
                label: displayName,
                provider: rep.provider,
                isVision: group.some(m => m.isVision),
                isReasoning: group.some(m => m.isReasoning)
            };
        });

        return [...standalone, ...aggregated];
    }

    /**
     * 构建模型选项图标前缀
     * - 🖥️ 本地模型（Ollama）
     * - ☁️ 云端模型
     * - 👁️ 视觉模型
     * - 🧠 推理模型（thinking / reasoning）
     */
    function buildModelIcons(model) {
        const icons = [];
        // 本地/云端基础标识
        icons.push(model.provider === 'ollama' ? '🖥️' : '☁️');
        // 能力标识
        if (model.isVision) icons.push('👁️');
        if (model.isReasoning) icons.push('🧠');
        return icons.join('') + ' ';
    }

    // --- 来自 app.js:L2146-L2331 ---
    function updateChatModelSelector() {
        const selector = document.getElementById('chat-model-select');

        if (!selector) {
            // chat-model-select 元素尚未加载，忽略（初始化时序问题）
            return;
        }

        try {
            // 清空现有选项
            selector.innerHTML = '';

            // 添加 AUTO 自动选择选项（置顶，作为默认选中）
            const autoOption = document.createElement('option');
            autoOption.value = 'AUTO';
            autoOption.dataset.provider = 'auto';
            autoOption.textContent = '🤖 AUTO (智能自动选择)';
            autoOption.style.fontWeight = 'bold';
            autoOption.style.color = '#667eea';
            autoOption.style.backgroundColor = '#f0f4ff';

            // 如果当前是 AUTO 模式或未选择模型，默认选中 AUTO
            if (currentChatModel === 'AUTO' || config.model === 'AUTO' || !currentChatModel || !config.model) {
                autoOption.selected = true;
                if (!currentChatModel && !config.model) {
                    currentChatModel = 'AUTO';
                    config.model = 'AUTO';
                }
            }
            selector.appendChild(autoOption);

            /** @type {{value: string, label: string, provider: string, isVision: boolean, isReasoning?: boolean}[]} */
            let allModels = [];

            // 【诊断】仅 debug 开关打开时输出（禁止生产路径打印 apiKey 前缀）
            if (typeof window !== 'undefined' && window.__SP_DEBUG_MODEL__) {
                console.log('[MODEL][DIAG] config.providers:', JSON.stringify(config.providers));
                console.log('[MODEL][DIAG] config.apiType:', config.apiType);
                console.log('[MODEL][DIAG] config.apiKey set:', !!(config.apiKey && String(config.apiKey).trim()));
                console.log('[MODEL][DIAG] App.config === config:', App.config === config);
            }

            // 收集所有真正可用的提供商的模型
            // 使用 isProviderAvailable 进行严格检测
            const availableProviders = getAvailableProviders();
            ChatLog.info('MODEL', 'updateChatModelSelector: 可用提供商:', availableProviders);

            availableProviders.forEach(providerKey => {
                const providerConfig = PROVIDER_CONFIGS[providerKey];

                if (providerConfig && providerConfig.models && providerConfig.models.length > 0) {
                    providerConfig.models.forEach(model => {
                        // 兼容两种配置格式：字符串或 { name: '...', contextWindow: ... }
                        const modelName = typeof model === 'string' ? model : (model && model.name ? model.name : '');
                        if (!modelName) return;

                        allModels.push({
                            value: modelName,
                            label: modelName,
                            provider: providerKey,
                            isVision: false
                        });
                    });
                }
            });

            // 处理本地模型列表（Ollama + llama.cpp）
            // 移除预定义的本地模型（如果有检测到的模型，用检测到的替换）
            allModels = allModels.filter(m => m.provider !== 'ollama' && m.provider !== 'llamacpp');

            if (detectedModels.length > 0) {
                // 添加检测到的本地模型（按来源区分提供商：ollama / llamacpp）
                // value 用原始 name（请求 model 字段），label 用清理后的展示名
                detectedModels.forEach(m => {
                    const prov = m.provider || 'ollama';
                    allModels.push({
                        value: m.name,
                        label: getLocalModelDisplayName(m),
                        provider: prov,
                        isVision: false
                    });
                });
            }

            // Ollama 兜底占位：仅当可用但一个模型都没检测到时，才显示预定义列表。
            // llama.cpp 无预置模型，只显示真正检测到的。
            const defaultOllamaModels = ['llava:latest', 'llama3.2-vision:latest', 'minicpm-v:latest'];
            if (availableProviders.includes('ollama') && !hasDetectedLocalModels('ollama')) {
                defaultOllamaModels.forEach(modelName => {
                    allModels.push({
                        value: modelName,
                        label: modelName,
                        provider: 'ollama',
                        isVision: false
                    });
                });
                ChatLog.info('MODEL', 'updateChatModelSelector: Ollama 可用但未检测到模型，添加默认模型列表');
            }

            // 标记视觉模型和推理模型
            allModels.forEach(model => {
                model.isVision = isVisionModel(model.value);
                model.isReasoning = isReasoningModel(model.value);
            });

            // 聚合模型家族，避免同一模型的多个 size/context 版本刷屏
            allModels = aggregateModelFamilies(allModels);

            // 按提供商分组并添加到选择器（Ollama 放在前面）
            const providers = [...new Set(allModels.map(m => m.provider))];

            // Ollama 优先显示
            const sortedProviders = [];
            if (providers.includes('ollama')) {
                sortedProviders.push('ollama');
            }
            providers.forEach(p => {
                if (p !== 'ollama') sortedProviders.push(p);
            });

            sortedProviders.forEach(provider => {
                const providerModels = allModels.filter(m => m.provider === provider);

                // 添加分组标签
                const groupLabel = document.createElement('option');
                groupLabel.value = '';
                groupLabel.disabled = true;
                groupLabel.style.fontWeight = 'bold';
                groupLabel.style.backgroundColor = '#f5f5f5';
                groupLabel.textContent = `─── ${getProviderName(provider)} ───`;
                selector.appendChild(groupLabel);

                providerModels.forEach(model => {
                    try {
                        const option = document.createElement('option');
                        option.value = model.value;
                        option.dataset.provider = provider;

                        // 添加图标标识：本地/云端、视觉、推理
                        option.textContent = buildModelIcons(model) + model.label;
                        if (model.isVision) {
                            option.style.color = '#4a90d9';
                        }

                        // 如果是当前选择的模型，选中它
                        if (model.value === config.model) {
                            option.selected = true;
                            autoOption.selected = false;
                            currentChatModel = model.value;
                        }

                        selector.appendChild(option);
                    } catch (e) {
                        ChatLog.error('MODEL', '添加选项失败:', e, model);
                    }
                });
            });

            // 记录最终添加的模型数量
            ChatLog.info('CONFIG', `📊 模型选择器更新完成`);
            ChatLog.info('CONFIG', `  ├─ 总模型数: ${allModels.length}`);
            ChatLog.info('CONFIG', `  ├─ 提供商数: ${sortedProviders.length}`);
            sortedProviders.forEach(provider => {
                const count = allModels.filter(m => m.provider === provider).length;
                ChatLog.info('CONFIG', `  ├─ ${getProviderName(provider)}: ${count} 个模型`);
            });
            ChatLog.info('CONFIG', `  ├─ 视觉模型: ${allModels.filter(m => m.isVision).length} 个`);
            ChatLog.info('CONFIG', `  └─ 推理模型: ${allModels.filter(m => m.isReasoning).length} 个`);

            // 如果没有可用模型，显示明确的提示
            if (allModels.length === 0) {
                ChatLog.warn('CONFIG', '⚠️ 没有可用的模型');
                const noModelOption = document.createElement('option');
                noModelOption.value = '';
                noModelOption.disabled = true;

                // 根据情况显示不同的提示
                const hasAnyProviderConfig = Object.keys(PROVIDER_CONFIGS).some(key => {
                    if (key === 'ollama' || key === 'llamacpp') return hasDetectedLocalModels(key);
                    return config.providers?.[key]?.apiKey;
                });

                if (hasAnyProviderConfig) {
                    noModelOption.textContent = '配置已保存，请检查 API Key 是否有效';
                    ChatLog.warn('CONFIG', '  └─ 提示: 配置已保存，请检查 API Key 是否有效');
                } else {
                    noModelOption.textContent = '请先在设置中配置 API 密钥或启动 Ollama';
                    ChatLog.warn('CONFIG', '  └─ 提示: 请先在设置中配置 API 密钥或启动 Ollama');
                }

                noModelOption.style.color = '#999';
                selector.appendChild(noModelOption);
            }

            ChatLog.info('MODEL', 'updateChatModelSelector 完成，添加了', allModels.length, '个选项');
            try {
                if (typeof window.refreshWelcomeSetupCta === 'function') window.refreshWelcomeSetupCta();
            } catch (_) { /* ignore */ }
        } catch (error) {
            ChatLog.error('MODEL', 'updateChatModelSelector 错误:', error);
        }

        selector.style.display = 'block';
    }

    // ==================== AUTO 智能模型选择 ====================

    // --- 来自 app.js:L2334-L2350 ---
    const AUTO_MODEL_CONFIG = {
        vision: {
            keywords: ['图片', '图像', '照片', '截图', '识别', '视觉', '看图', '图中', '这张图', 'ocr', '表格识别', '视觉'],
            priority: ['qwen3-vt:4b', 'qwen3-vl:4b', 'qwen3-vl:8b', 'qwen2.5-vl:latest', 'llava:latest', 'minicpm-v:latest', 'gemma3:12b'],
            description: '视觉模型'
        },
        reasoning: {
            keywords: ['代码', '编程', 'debug', '算法', '数学', '推理', '逻辑', '证明', '优化', '复杂', '深度', '详细解释', 'step by step', '逐步', '深度分析', '逐步分析', '详细分析', '复杂计算'],
            priority: ['deepseek-r1:8b', 'gemma3:12b', 'qwen3:8b', 'qwen2.5:32b', 'qwen2.5:72b', 'qwen3:30b', 'qwen3:4b', 'gemma3:4b'],
            description: '推理模型'
        },
        simple: {
            keywords: [],
            priority: ['qwen2:1.8b', 'qwen2.5:7b', 'gemma3:4b', 'qwen3:4b', 'deepseek-r1:8b', 'qwen2.5:32b'],
            description: '轻量模型'
        }
    };

    // --- 来自 app.js:L2352-L2455 ---
    function autoSelectModel(userMessage, hasImage = false) {
        ChatLog.info('MODEL', '🤖 [AUTO] 开始智能选择模型...');
        ChatLog.info('MODEL', '🤖 [AUTO] 用户消息:', userMessage ? userMessage.substring(0, 100) + '...' : '(空)');
        ChatLog.info('MODEL', '🤖 [AUTO] 是否包含图片:', hasImage);

        const modelSelect = document.getElementById('chat-model-select');
        if (!modelSelect || modelSelect.options.length === 0) {
            ChatLog.warn('MODEL', '🤖 [AUTO] 模型选择器不可用，使用默认模型');
            const fallback = config.model || '';
            return fallback === 'AUTO' ? '' : fallback;
        }

        // 获取所有可用模型
        const availableModels = [];
        for (let i = 0; i < modelSelect.options.length; i++) {
            const option = modelSelect.options[i];
            if (option.value && option.value !== '' && option.value !== 'AUTO') {
                availableModels.push({
                    value: option.value,
                    provider: option.dataset.provider || ''
                });
            }
        }

        ChatLog.info('MODEL', '🤖 [AUTO] 可用模型数量:', availableModels.length);
        if (availableModels.length === 0) {
            ChatLog.warn('MODEL', '🤖 [AUTO] 没有可用模型');
            const fallback = config.model || '';
            return fallback === 'AUTO' ? '' : fallback;
        }

        // 任务类型检测
        let taskType = 'simple';
        const messageLower = (userMessage || '').toLowerCase();

        // 图片任务优先使用视觉模型
        if (hasImage) {
            taskType = 'vision';
            ChatLog.info('MODEL', `🤖 [AUTO] 检测到图片输入 → 使用视觉模型`);
        } else {
            // 检测是否为推理任务
            for (const keyword of AUTO_MODEL_CONFIG.reasoning.keywords) {
                if (messageLower.includes(keyword.toLowerCase())) {
                    taskType = 'reasoning';
                    ChatLog.info('MODEL', `🤖 [AUTO] 检测到关键词 "${keyword}" → 使用推理模型`);
                    break;
                }
            }
        }

        const taskConfig = AUTO_MODEL_CONFIG[taskType];
        ChatLog.info('MODEL', `🤖 [AUTO] 最终任务类型: ${taskType} (${taskConfig.description})`);
        ChatLog.info('MODEL', `🤖 [AUTO] 模型优先级列表:`, taskConfig.priority);

        // 根据优先级查找最佳匹配模型
        for (const preferredModel of taskConfig.priority) {
            const found = availableModels.find(m => {
                const modelLower = String(m.value || '').toLowerCase();
                const preferredLower = preferredModel.toLowerCase();
                return modelLower === preferredLower ||
                       modelLower.includes(preferredLower.replace(':', '').split('/')[0]) ||
                       preferredLower.includes(modelLower.replace(':', '').split('/')[0]);
            });

            if (found) {
                ChatLog.info('MODEL', `🤖 [AUTO] ✅ 选择模型: ${found.value} (${taskConfig.description})`);

                // 更新头部模型标识（仅显示信息，不改变下拉框）
                const headerModelBadge = document.querySelector('.model-badge');
                if (headerModelBadge) {
                    headerModelBadge.textContent = `🤖 AUTO → ${found.value}`;
                    headerModelBadge.style.color = '#667eea';
                }

                return found.value;
            }
        }

        // 如果没有找到匹配的模型，使用第一个可用模型
        const fallbackModel = availableModels[0].value;
        ChatLog.info('MODEL', `🤖 [AUTO] ⚠️ 未找到最佳匹配，使用备用模型: ${fallbackModel}`);

        // 更新头部模型标识
        const headerModelBadge = document.querySelector('.model-badge');
        if (headerModelBadge) {
            headerModelBadge.textContent = `🤖 AUTO → ${fallbackModel}`;
            headerModelBadge.style.color = '#667eea';
        }

        return fallbackModel;
    }

    // --- 来自 app.js:L2457-L2459 ---
    function isAutoMode() {
        return currentChatModel === 'AUTO' || config.model === 'AUTO';
    }

    // --- 来自 app.js:L2461-L2466 ---
    function getEffectiveModel(userMessage, hasImage = false) {
        if (!isAutoMode()) {
            return currentChatModel || config.model || '';
        }
        return autoSelectModel(userMessage, hasImage);
    }

    // --- 来自 app.js:L2469-L2482 ---
    function getProviderName(providerKey) {
        const names = {
            'ollama': 'Ollama 本地',
            'llamacpp': 'llama.cpp 本地',
            'openai': 'OpenAI',
            'azure': 'Azure OpenAI',
            'wenxin': '文心一言',
            'qianwen': '通义千问',
            'zhipu': '智谱 AI',
            'xunfei': '讯飞星火',
            'moonshot': 'Moonshot',
            'deepseek': '深度求索 (DeepSeek)',
            'volcengine': '火山引擎 (豆包)',
            'custom': '自定义'
        };
        return names[providerKey] || providerKey;
    }

    // --- 来自 app.js:L2485-L2497 ---
    async function refreshModelsList() {
        App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">正在刷新模型列表...</span>') : showToast('正在刷新模型列表...');

        // 【修复】重新从 localStorage 加载配置，确保获取最新的云端模型配置
        await loadConfig();

        // 总是检测 Ollama 本地模型，等待完成后再更新选择器
        await detectLocalModels();

        // 重新加载并显示模型列表
        updateChatModelSelector();

        setTimeout(() => {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">模型列表已刷新</span>', true) : showToast('模型列表已刷新');
        }, 500);
    }

    // --- 来自 app.js:L2499-L2566 ---
    function onChatModelChange() {
        const selector = document.getElementById('chat-model-select');
        // 【修复】selector可能为null
        if (!selector) {
            ChatLog.warn('MODEL', '[onChatModelChange] chat-model-select 元素不存在');
            return;
        }
        currentChatModel = selector.value;

        if (currentChatModel) {
            // 获取选中的选项元素
            const selectedOption = selector.options[selector.selectedIndex];
            const selectedProvider = selectedOption.dataset.provider;

            // 如果选择了 AUTO 模式
            if (currentChatModel === 'AUTO') {
                ChatLog.info('MODEL', '🤖 [AUTO] 用户选择了 AUTO 自动模式');
                config.model = 'AUTO';
                // 【修复】model 元素可能不存在
                const modelEl = document.getElementById('model');
                if (modelEl) modelEl.value = 'AUTO';

                // 更新头部模型标识
                const headerModelBadge = document.querySelector('.model-badge');
                if (headerModelBadge) {
                    headerModelBadge.textContent = '🤖 AUTO';
                    headerModelBadge.style.color = '#667eea';
                    headerModelBadge.style.fontWeight = 'bold';
                }

                saveSettings();
                return;
            }

            // 如果模型属于某个提供商，自动切换提供商
            if (selectedProvider && selectedProvider !== config.apiType && selectedProvider !== 'auto') {
                config.apiType = selectedProvider;
                // 【修复】api-type 元素只在设置面板存在，聊天面板中可能为 null
                const apiTypeEl = document.getElementById('api-type');
                if (apiTypeEl) apiTypeEl.value = selectedProvider;

                // 获取默认配置
                const providerConfig = PROVIDER_CONFIGS[selectedProvider];
                const defaultUrl = providerConfig ? providerConfig.apiUrl : '';

                // 从 providers 对象加载该提供商的配置
                if (config.providers && config.providers[selectedProvider]) {
                    const providerSettings = config.providers[selectedProvider];
                    config.apiKey = providerSettings.apiKey || '';
                    // 如果保存的 apiUrl 为空，使用默认配置
                    config.apiUrl = providerSettings.apiUrl || defaultUrl;
                    config.model = currentChatModel;

                    // 【修复】更新设置面板（元素可能不存在）
                    const apiKeyEl = document.getElementById('api-key');
                    if (apiKeyEl) apiKeyEl.value = config.apiKey;
                    const apiUrlEl = document.getElementById('api-url');
                    if (apiUrlEl) apiUrlEl.value = config.apiUrl;
                } else {
                    // 如果没有保存过配置，使用默认配置
                    if (providerConfig) {
                        config.apiUrl = defaultUrl;
                        const apiUrlEl = document.getElementById('api-url');
                        if (apiUrlEl) apiUrlEl.value = defaultUrl;
                    }
                }

                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">已切换到 ' + getProviderName(selectedProvider) + ': ' + currentChatModel + '</span>', true) : showToast(`已切换到 ${getProviderName(selectedProvider)}: ${currentChatModel}`);
            } else {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">已切换到模型: ' + currentChatModel + '</span>', true) : showToast('已切换到模型: ' + currentChatModel);
            }

            // 更新设置中的模型值
            const modelEl = document.getElementById('model');
            if (modelEl) modelEl.value = currentChatModel;
            config.model = currentChatModel;

            // 保存配置
            saveSettings();
        }
    }

    // 思考模式切换处理
    function onThinkingModeChange() {
        const selector = document.getElementById('thinking-mode-select');
        if (!selector) {
            ChatLog.warn('MODEL', '[onThinkingModeChange] thinking-mode-select 元素不存在');
            return;
        }
        const mode = selector.value;
        App.config.thinkingMode = mode;
        localStorage.setItem('thinkingMode', mode);
        ChatLog.info('MODEL', `[思考模式] 已切换为: ${mode}`);
        
        // 合并到状态栏显示，不再用toast
        const modeText = mode === 'AUTO' ? '智能自动选择' : mode;
        if (App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function') {
            App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">思考模式: ' + modeText + '</span>');
        }
    }

    // ==================== 环境检测与适配 ====================

    // 初始化思考模式（从 localStorage加载）
    if (!App.config) App.config = {};
    const savedThinkingMode = localStorage.getItem('thinkingMode');
    if (savedThinkingMode) {
        App.config.thinkingMode = savedThinkingMode;
        const selector = document.getElementById('thinking-mode-select');
        if (selector) {
            selector.value = savedThinkingMode;
        }
        ChatLog.info('MODEL', `[思考模式] 从缓存加载: ${savedThinkingMode}`);
            
        // 在状态栏显示思考模式
        setTimeout(() => {
            const modeText = savedThinkingMode === 'AUTO' ? '智能自动选择' : savedThinkingMode;
            if (App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function') {
                App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">思考模式: ' + modeText + '</span>');
            }
        }, 500);
    } else {
        App.config.thinkingMode = 'AUTO';
    }

    // --- 来自 app.js:L2570-L2578 ---
    const ENV = {
        isWebView2: !!(window.chrome && window.chrome.webview),
        isFileProtocol: window.location.protocol === 'file:',
        isHttpProtocol: window.location.protocol === 'http:' || window.location.protocol === 'https:',
        host: window.location.host,
        href: window.location.href,
        _mockMode: false
    };

    ChatLog.info('MODEL', '🔍 环境检测:', ENV);

    // Mock ExcelBridge（用于外部浏览器调试）
    // --- 来自 app.js:L2581-L2662 ---
    const MOCK_BRIDGE = {
        _mockMode: true,

        async GetSelection() {
            ChatLog.info('MODEL', '[Mock] GetSelection 被调用');
            return JSON.stringify([
                ['姓名', '年龄', '城市'],
                ['张三', '25', '北京'],
                ['李四', '30', '上海']
            ]);
        },

        async GetAllSkills() {
            // 浏览器（非 WebView2）模式下也走本地桥 HTTP 接口，
            // 与 VSTO 端共用仓库 Skills/*.md，保证两端技能清单同步
            ChatLog.info('MODEL', '[Bridge] GetAllSkills 调用 Bridge 服务器');
            try {
                var response = await fetch(App.Endpoints ? App.Endpoints.excelBridge.invoke : 'http://localhost:9527/api/excel/invoke', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ method: 'GetAllSkills', arguments: [] })
                });
                var result = await response.json();
                return JSON.stringify(result.data || { success: false, error: 'Bridge 服务器未响应' });
            } catch (e) {
                ChatLog.warn('MODEL', '[Bridge] GetAllSkills 调用失败，回退到模拟列表:', e.message);
                return JSON.stringify({
                    success: true,
                    skills: [
                        { name: 'MockSkill1', description: '模拟技能1', parameters: [] },
                        { name: 'MockSkill2', description: '模拟技能2', parameters: [] }
                    ]
                });
            }
        },

        async ExecuteSkill(skillName, inputsJson) {
            // 浏览器（非 WebView2）模式下也走本地桥 HTTP 接口真实执行技能，
            // 与 VSTO 端共用仓库 Skills/*.md
            ChatLog.info('MODEL', '[Bridge] ExecuteSkill 调用 Bridge 服务器:', skillName, inputsJson);
            try {
                var response = await fetch(App.Endpoints ? App.Endpoints.excelBridge.invoke : 'http://localhost:9527/api/excel/invoke', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ method: 'ExecuteSkill', arguments: [skillName, inputsJson] })
                });
                var result = await response.json();
                return JSON.stringify(result.data || { success: false, error: 'Bridge 服务器未响应' });
            } catch (e) {
                ChatLog.warn('MODEL', '[Bridge] ExecuteSkill 调用失败:', e.message);
                return JSON.stringify({
                    success: false,
                    mock: true,
                    error: `[Bridge调用失败] 技能 "${skillName}" 未能执行: ${e.message}`
                });
            }
        },

        async ExecuteSkillAsync(skillName, inputsJson) {
            // 与 ExecuteSkill 一致：浏览器模式下经本地桥 HTTP 真实执行
            ChatLog.info('MODEL', '[Bridge] ExecuteSkillAsync 调用 Bridge 服务器:', skillName, inputsJson);
            try {
                var response = await fetch(App.Endpoints ? App.Endpoints.excelBridge.invoke : 'http://localhost:9527/api/excel/invoke', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ method: 'ExecuteSkillAsync', arguments: [skillName, inputsJson] })
                });
                var result = await response.json();
                return JSON.stringify(result.data || { success: false, error: 'Bridge 服务器未响应' });
            } catch (e) {
                ChatLog.warn('MODEL', '[Bridge] ExecuteSkillAsync 调用失败:', e.message);
                return JSON.stringify({
                    success: false,
                    mock: true,
                    error: `[Bridge调用失败] 技能 "${skillName}" 未能执行: ${e.message}`
                });
            }
        },

        async CallLocalTool(toolName, argsJson) {
            // 对齐 VSTO McpManager.CallLocalTool：本地工具统一分发入口
            // （chat-v2 mcp-dispatcher 硬短路路径经此调用 computer_/ppt_/web_/time_ 等工具）
            ChatLog.info('MODEL', '[Bridge] CallLocalTool:', toolName, argsJson);
            try {
                var response = await fetch(App.Endpoints ? App.Endpoints.excelBridge.invoke : 'http://localhost:9527/api/excel/invoke', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ method: 'CallLocalTool', arguments: [toolName, argsJson] })
                });
                var result = await response.json();
                return JSON.stringify(result.data || { success: false, error: 'Bridge 服务器未响应' });
            } catch (e) {
                ChatLog.warn('MODEL', '[Bridge] CallLocalTool 调用失败:', e.message);
                return JSON.stringify({ success: false, error: 'CallLocalTool 调用失败: ' + e.message });
            }
        },

        async CallLocalToolAsync(toolName, argsJson) {
            return this.CallLocalTool(toolName, argsJson);
        },

        async FillDataToExcel(dataJson, typeLabel) {
            ChatLog.info('MODEL', '[Mock] FillDataToExcel:', typeLabel, dataJson.substring(0, 100));
            return JSON.stringify({ success: true, mock: true, message: '模拟填入成功' });
        },

        async InsertDataSmart(payloadJson) {
            ChatLog.info('MODEL', '[Mock] InsertDataSmart:', payloadJson.substring(0, 100));
            return JSON.stringify({ success: true, mock: true, message: '模拟智能插入成功' });
        },

        async SaveAIConfig(configJson) {
            ChatLog.info('MODEL', '[Mock] SaveAIConfig:', configJson.substring(0, 100));
            localStorage.setItem('mock_ai_config', configJson);
            return JSON.stringify({ success: true });
        },

        async GetAIConfig() {
            ChatLog.info('MODEL', '[Mock] GetAIConfig 被调用');
            const cfg = localStorage.getItem('mock_ai_config');
            return cfg || JSON.stringify({ success: true, config: {} });
        },

        async GetExcelStatus() {
            ChatLog.info('MODEL', '[Mock] GetExcelStatus 被调用');
            return JSON.stringify({
                success: true,
                hasWorkbook: true,
                selectionCount: 4,
                activeSheetName: 'MockSheet'
            });
        },

        async GetCardTypesConfig() {
            ChatLog.info('MODEL', '[Mock] GetCardTypesConfig 被调用');
            return JSON.stringify({
                success: true,
                cardTypes: {
                    default: { name: '默认', icon: '💬', color: '#4a90d9' },
                    data: { name: '数据', icon: '📊', color: '#27ae60' },
                    image: { name: '图片', icon: '🖼️', color: '#e67e22' }
                }
            });
        },

        // 通用 fallback：未实现的方法返回模拟成功
        _fallback(methodName, args) {
            ChatLog.info('MODEL', `[Mock] ${methodName} 被调用 (未实现)`, args);
            return JSON.stringify({ success: false, mock: true, method: methodName, error: `${methodName} 未实现` });
        },

        // ========== RAG 知识库（桌面版如实失败，不再 Mock 假成功）==========
        // 桌面版（9527 HTTP 桥）没有 RAG/Qdrant 服务，创建/导入/检索如实报错，
        // 避免"创建成功实际没建"的假成功。列表查询仍返回空列表（成功），保证 tab 可打开。
        async GetKnowledgeBasesAsync() {
            ChatLog.info('MODEL', '[Mock] GetKnowledgeBasesAsync 被调用');
            return JSON.stringify({ success: true, knowledgeBases: [] });
        },

        async CreateKnowledgeBaseAsync(name, description, optionsJson) {
            ChatLog.info('MODEL', '[Mock] CreateKnowledgeBaseAsync（桌面版不支持）:', name);
            return JSON.stringify({
                success: false,
                error: '桌面版暂不支持知识库：需要 RAG 服务（Qdrant + 云端向量化），请在 Excel 插件版中使用。'
            });
        },

        async DeleteKnowledgeBaseAsync(kbId) {
            ChatLog.info('MODEL', '[Mock] DeleteKnowledgeBaseAsync:', kbId);
            return JSON.stringify({ success: true });
        },

        async ImportDocumentAsync(ossUrl, kbId, title, configJson) {
            ChatLog.info('MODEL', '[Mock] ImportDocumentAsync（桌面版不支持）:', title || ossUrl);
            return JSON.stringify({
                success: false,
                error: '桌面版暂不支持文档导入：需要 RAG 服务（Qdrant + 云端向量化），请在 Excel 插件版中使用。'
            });
        },

        async GetDocumentsAsync(kbId) {
            ChatLog.info('MODEL', '[Mock] GetDocumentsAsync:', kbId);
            return JSON.stringify({ success: true, documents: [] });
        },

        async DeleteDocumentAsync(docId) {
            ChatLog.info('MODEL', '[Mock] DeleteDocumentAsync:', docId);
            return JSON.stringify({ success: true });
        },

        async SearchKnowledgeAsync(query, kbId, topK, optionsJson) {
            ChatLog.info('MODEL', '[Mock] SearchKnowledgeAsync（桌面版不支持）:', query);
            return JSON.stringify({
                success: false,
                error: '桌面版暂不支持知识检索：需要 RAG 服务，请在 Excel 插件版中使用。',
                results: []
            });
        },

        async TestRagQueryAsync(question, kbId, topK) {
            ChatLog.info('MODEL', '[Mock] TestRagQueryAsync（桌面版不支持）:', question);
            return JSON.stringify({
                success: false,
                error: '桌面版暂不支持 RAG 问答：需要 RAG 服务，请在 Excel 插件版中使用。',
                shouldFallback: true
            });
        },

        async GetRagConfigAsync() {
            ChatLog.info('MODEL', '[Bridge] GetRagConfigAsync 调用 Bridge 服务器');
            try {
                var response = await fetch(App.Endpoints ? App.Endpoints.excelBridge.invoke : 'http://localhost:9527/api/excel/invoke', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ method: 'GetRagConfigAsync', arguments: [] })
                });
                var result = await response.json();
                return JSON.stringify(result.data || { success: false, error: 'Bridge 服务器未响应' });
            } catch (e) {
                ChatLog.warn('MODEL', '[Bridge] GetRagConfigAsync 调用失败，使用默认配置:', e.message);
                return JSON.stringify({
                    success: true,
                    config: {
                        defaultProvider: 'local_bge',
                        localBgeEndpoint: 'http://localhost:11434',
                        localBgeModel: 'bge-m3',
                        qdrantEndpoint: 'http://localhost:6333',
                        defaultTopK: 5,
                        similarityThreshold: 0.7,
                        enableChatEnhancement: true
                    }
                });
            }
        },
        
        async SaveRagConfigAsync(configJson) {
            ChatLog.info('MODEL', '[Bridge] SaveRagConfigAsync 调用 Bridge 服务器');
            try {
                var response = await fetch(App.Endpoints ? App.Endpoints.excelBridge.invoke : 'http://localhost:9527/api/excel/invoke', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ method: 'SaveRagConfigAsync', arguments: [configJson] })
                });
                var result = await response.json();
                return JSON.stringify(result.data || { success: false, error: 'Bridge 服务器未响应' });
            } catch (e) {
                ChatLog.warn('MODEL', '[Bridge] SaveRagConfigAsync 调用失败:', e.message);
                return JSON.stringify({ success: false, error: e.message });
            }
        },
        
        async TestQdrantConnectionAsync() {
            ChatLog.info('MODEL', '[Bridge] TestQdrantConnectionAsync 调用 Bridge 服务器');
            try {
                var response = await fetch(App.Endpoints ? App.Endpoints.excelBridge.invoke : 'http://localhost:9527/api/excel/invoke', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ method: 'TestQdrantConnectionAsync', arguments: [] })
                });
                var result = await response.json();
                return JSON.stringify(result.data || { success: false, error: 'Bridge 服务器未响应' });
            } catch (e) {
                ChatLog.warn('MODEL', '[Bridge] TestQdrantConnectionAsync 调用失败:', e.message);
                return JSON.stringify({ success: false, error: e.message });
            }
        },
        
        async TestEmbeddingConnectionAsync(provider) {
            ChatLog.info('MODEL', '[Bridge] TestEmbeddingConnectionAsync 调用 Bridge 服务器');
            try {
                var response = await fetch(App.Endpoints ? App.Endpoints.excelBridge.invoke : 'http://localhost:9527/api/excel/invoke', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ method: 'TestEmbeddingConnectionAsync', arguments: [provider || ''] })
                });
                var result = await response.json();
                return JSON.stringify(result.data || { success: false, error: 'Bridge 服务器未响应' });
            } catch (e) {
                ChatLog.warn('MODEL', '[Bridge] TestEmbeddingConnectionAsync 调用失败:', e.message);
                return JSON.stringify({ success: false, error: e.message });
            }
        },
                
        async GetQdrantInstallStatus() {
            ChatLog.info('MODEL', '[Bridge] GetQdrantInstallStatus 调用 Bridge 服务器');
            try {
                var response = await fetch(App.Endpoints ? App.Endpoints.excelBridge.invoke : 'http://localhost:9527/api/excel/invoke', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ method: 'GetQdrantInstallStatus', arguments: [] })
                });
                var result = await response.json();
                return JSON.stringify(result.data || { success: false, error: 'Bridge 服务器未响应' });
            } catch (e) {
                ChatLog.warn('MODEL', '[Bridge] GetQdrantInstallStatus 调用失败:', e.message);
                return JSON.stringify({ success: false, error: e.message });
            }
        },
                
        async InstallQdrantAsync() {
            ChatLog.info('MODEL', '[Bridge] InstallQdrantAsync 调用 Bridge 服务器');
            try {
                var response = await fetch(App.Endpoints ? App.Endpoints.excelBridge.invoke : 'http://localhost:9527/api/excel/invoke', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ method: 'InstallQdrantAsync', arguments: [] })
                });
                var result = await response.json();
                return JSON.stringify(result.data || { success: false, error: 'Bridge 服务器未响应' });
            } catch (e) {
                ChatLog.warn('MODEL', '[Bridge] InstallQdrantAsync 调用失败:', e.message);
                return JSON.stringify({ success: false, error: e.message });
            }
        },
                
        async InstallQdrantFromLocalFileAsync(filePath) {
            ChatLog.info('MODEL', '[Bridge] InstallQdrantFromLocalFileAsync 调用 Bridge 服务器');
            try {
                var response = await fetch(App.Endpoints ? App.Endpoints.excelBridge.invoke : 'http://localhost:9527/api/excel/invoke', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ method: 'InstallQdrantFromLocalFileAsync', arguments: [filePath] })
                });
                var result = await response.json();
                return JSON.stringify(result.data || { success: false, error: 'Bridge 服务器未响应' });
            } catch (e) {
                ChatLog.warn('MODEL', '[Bridge] InstallQdrantFromLocalFileAsync 调用失败:', e.message);
                return JSON.stringify({ success: false, error: e.message });
            }
        },
                
        async StartQdrantAsync() {
            ChatLog.info('MODEL', '[Bridge] StartQdrantAsync 调用 Bridge 服务器');
            try {
                var response = await fetch(App.Endpoints ? App.Endpoints.excelBridge.invoke : 'http://localhost:9527/api/excel/invoke', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ method: 'StartQdrantAsync', arguments: [] })
                });
                var result = await response.json();
                return JSON.stringify(result.data || { success: false, error: 'Bridge 服务器未响应' });
            } catch (e) {
                ChatLog.warn('MODEL', '[Bridge] StartQdrantAsync 调用失败:', e.message);
                return JSON.stringify({ success: false, error: e.message });
            }
        },
                
        async StopQdrantAsync() {
            ChatLog.info('MODEL', '[Bridge] StopQdrantAsync 调用 Bridge 服务器');
            try {
                var response = await fetch(App.Endpoints ? App.Endpoints.excelBridge.invoke : 'http://localhost:9527/api/excel/invoke', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ method: 'StopQdrantAsync', arguments: [] })
                });
                var result = await response.json();
                return JSON.stringify(result.data || { success: false, error: 'Bridge 服务器未响应' });
            } catch (e) {
                ChatLog.warn('MODEL', '[Bridge] StopQdrantAsync 调用失败:', e.message);
                return JSON.stringify({ success: false, error: e.message });
            }
        },
                
        async ShowOpenFileDialog() {
            ChatLog.info('MODEL', '[Bridge] ShowOpenFileDialog 调用 Bridge 服务器');
            try {
                var response = await fetch(App.Endpoints ? App.Endpoints.excelBridge.invoke : 'http://localhost:9527/api/excel/invoke', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ method: 'ShowOpenFileDialog', arguments: [] })
                });
                var result = await response.json();
                return JSON.stringify(result.data || { success: false, message: 'Bridge 服务器未响应' });
            } catch (e) {
                ChatLog.warn('MODEL', '[Bridge] ShowOpenFileDialog 调用失败:', e.message);
                return JSON.stringify({ success: false, message: e.message });
            }
        }
    };

    // --- 来自 app.js:L2665-L2678 ---
    // 为 Mock Bridge 添加未定义方法的代理
    // 【本地桥优先】浏览器网页版（Python 桥 localhost:9527）在线时，白名单方法直接转发到
    // 真实本地桥读写真实 Excel；Mock 仅作为桥离线或方法未实现时的兜底。
    const LOCAL_BRIDGE_URL = (App.Endpoints ? App.Endpoints.excelBridge.baseUrl : (window.ExcelBridgeConfig && window.ExcelBridgeConfig.bridgeUrl)) || 'http://localhost:9527/api/excel';
    let _localBridgeOnline = false;
    (function _watchLocalBridge() {
        // 【守卫修正·B】本地桥(9527)仅网页独立版使用；VSTO(WebView2) 走 COM 桥，
        // 探 9527/api/excel/status 必失败并刷屏。VSTO 环境直接跳过探测。
        if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects) return;
        const probe = function() {
            fetch(LOCAL_BRIDGE_URL + '/status', { method: 'GET', cache: 'no-store' })
                .then(function(r) { return r.json(); })
                .then(function(d) { _localBridgeOnline = !!(d && d.status === 'ok'); })
                .catch(function() { _localBridgeOnline = false; });
        };
        probe();
        setInterval(probe, 3000);
    })();

    // 真实本地桥已实现的方法白名单（其余方法仍走 Mock，不改变现有调试体验）
    // License* 与 AI 配置方法由桌面版桥实现（离线激活/密钥共享），必须放行，
    // 否则网页版被 Mock 拦截后永远显示"未使用激活码"。
    const REAL_BRIDGE_METHODS = {
        GetSelection: 1, GetMultiSelection: 1, StartSelectionMonitor: 1, StopSelectionMonitor: 1,
        GetSelectedRange: 1, GetSelectedRangeAsync: 1, GetSelectedRangeData: 1,
        GetOpenWorkbooks: 1, GetWorkbookInfo: 1, GetActiveSheetName: 1, GetActiveWorkbookName: 1,
        GetSheets: 1, ActivateWorkbook: 1, ActivateSheet: 1, SelectSheet: 1, ActivateSheetAsync: 1,
        GetRangeData: 1, GetAvailableMcpTools: 1,
        WriteToRange: 1, WriteToRangeAsync: 1,
        InsertTableData: 1, InsertDataSmart: 1, InsertDataSmartAsync: 1,
        InsertParsedData: 1, InsertParsedDataAsync: 1, FillDataToExcel: 1, FillDataToExcelAsync: 1,
        ExecuteVBA: 1, ExecuteVBAAsync: 1,
        LicenseActivateKey: 1, LicenseGetLocalActivation: 1, LicenseClearLocalActivation: 1,
        LicenseDeactivateForRefund: 1, LicenseGetEntitlement: 1, LicenseIsSubscriptionActive: 1,
        LicenseGetApiBaseUrl: 1, LicensePingHealth: 1, LicenseRegister: 1, LicenseLogin: 1,
        LicenseLogout: 1, LicenseStartCheckout: 1, LicenseOpenBillingPortal: 1,
        SaveAIConfig: 1, GetAIConfig: 1, SaveLLMConfig: 1, GetLLMConfig: 1,
        GetAllSkills: 1, GetAvailableSkills: 1, GetQuickShortcutSkills: 1,
        ExecuteSkill: 1, ExecuteSkillAsync: 1, CallLocalTool: 1, CallLocalToolAsync: 1,
        SaveBase64Image: 1, StartOcrRecognition: 1, StartOcrRecognitionWithFormat: 1,
        GetOcrResult: 1, GetOcrResultWithFormat: 1, GetOcrStructuredResult: 1,
        GetCurrentOcrEngine: 1, SetOcrEngine: 1, GetOcrEngineStatus: 1,
        GetMcpStatus: 1, StartMcpServer: 1, StopMcpServer: 1,
        ToggleSkillEnabled: 1, SaveSkill: 1, DeleteSkill: 1, GetSkillByName: 1,
        CheckPythonEnvironment: 1, ExecutePython: 1,
        GetCloudOcrConfig: 1, SaveCloudOcrConfig: 1, TestCloudOcrConfig: 1,
        SaveCustomPptTheme: 1, ListCustomPptThemes: 1, DeleteCustomPptTheme: 1
    };

    function _callLocalBridge(methodName, args) {
        return fetch(LOCAL_BRIDGE_URL + '/invoke', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ method: methodName, arguments: args || [] })
        }).then(function(r) { return r.json(); }).then(function(resp) {
            if (resp && resp.success) {
                // 【契约】与 bridge-accessor / VSTO 宿主对象一致：桥方法统一返回 JSON 字符串。
                // 这里若直接回传对象，前端大量 JSON.parse(result) 的调用点会把对象转成
                // "[object Object]" 并抛 "is not valid JSON"（技能开关/保存/删除、OCR 等均受影响）。
                var d = (resp.data !== undefined) ? resp.data : resp;
                return (typeof d === 'string') ? d : JSON.stringify(d);
            }
            throw new Error((resp && resp.error) ? resp.error : ('本地桥调用失败: ' + methodName));
        });
    }

    const MOCK_BRIDGE_PROXY = new Proxy(MOCK_BRIDGE, {
        get(target, prop) {
            if (prop === 'then' || typeof prop !== 'string') return undefined;
            // 非函数属性（如 _mockMode）原样返回，保持原有语义
            if (prop in target && typeof target[prop] !== 'function') {
                return target[prop];
            }
            const mockFn = (prop in target)
                ? function(...args) { return target[prop].apply(target, args); }
                : function(...args) { return Promise.resolve(target._fallback(prop, args)); };
            if (REAL_BRIDGE_METHODS[prop] && _localBridgeOnline) {
                return function(...args) {
                    return _callLocalBridge(prop, args).catch(function(err) {
                        ChatLog.warn('MODEL', '[本地桥] ' + prop + ' 失败，回落 Mock: ' + (err && err.message ? err.message : err));
                        return mockFn.apply(null, args);
                    });
                };
            }
            return mockFn;
        }
    });

    // --- 来自 app.js:L2682-L2704 ---
    // Mock 注入逻辑
    // 如果不在 WebView2 环境中，注入模拟的 chrome.webview 对象
    // 这样现有代码中的 window.chrome.webview.hostObjects.excelBridge 调用无需修改
    if (!ENV.isWebView2) {
        ENV._mockMode = true;
        ChatLog.info('MODEL', '%c[开发模式] 注入 Mock WebView2 环境', 'color: orange; font-weight: bold');

        if (!window.chrome) {
            window.chrome = {};
        }

        window.chrome.webview = {
            _mockMode: true,
            hostObjects: {
                options: {
                    defaultSyncMethod: 'sync'
                },
                excelBridge: MOCK_BRIDGE_PROXY
            },
            postMessage: function(message) {
                ChatLog.info('MODEL', '[Mock] postMessage:', message);
            },
            addEventListener: function(event, handler) {
                ChatLog.info('MODEL', '[Mock] addEventListener:', event);
            }
        };
    }

    // ==================== Bridge 初始化与状态管理 ====================

    // --- 来自 app.js:L2706-L2730 ---
    function initBridge() {
        // @ts-ignore
        if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects) {
            // 配置 COM 桥接选项，使所有方法默认同步执行
            // @ts-ignore
            if (window.chrome.webview.hostObjects.options) {
                // @ts-ignore
                // 合法取值为 'none'（默认·异步）与 'sync'，'async' 非有效值会被忽略。
                // 本项目桥调用均 await，必须用异步模式（'none'），否则 WebView2 会抛 e is not a function。
                window.chrome.webview.hostObjects.options.defaultSyncMethod = 'none';
            }
            // @ts-ignore
            var rawBridge = window.chrome.webview.hostObjects.excelBridge;
            // 必须经 BridgeAccessor 包 Proxy：裸对象上缺方法会异步抛 0x80020006 / e is not a function 并漏成未处理拒绝
            if (window.BridgeAccessor && typeof window.BridgeAccessor.adoptHostBridge === 'function') {
                excelBridge = window.BridgeAccessor.adoptHostBridge(rawBridge);
            } else {
                excelBridge = rawBridge;
                App.excelBridge = excelBridge;
                App.state.excelBridge = excelBridge;
            }
            App.state.isConnected = true;
            updateStatus(true);

            // ===== [诊断] 经 BridgeAccessor 探测（禁止 typeof 裸 host[prop]）=====
            // WebView2 异步代理：访问不存在的成员会返回会 reject 的 Promise；
            // 对 _hb[m] 做 typeof 且不 await → 未处理拒绝（0x80020006 / remoteObjectId）。
            try {
                if (window.BridgeAccessor && typeof window.BridgeAccessor.callAsync === 'function') {
                    ['GetDesktopPath', 'GetMcpStatus', 'GetSelection'].forEach(function (m) {
                        Promise.resolve(window.BridgeAccessor.callAsync(m)).then(function (r) {
                            var preview = (typeof r === 'string') ? r.slice(0, 160) : (r == null ? 'null' : JSON.stringify(r).slice(0, 160));
                            ChatLog.info('DIAG-BRIDGE', '[OK] ' + m + ' preview=' + preview);
                        }).catch(function (e) {
                            ChatLog.warn('DIAG-BRIDGE', '[FAIL] ' + m + ' -> ' + ((e && e.message) || String(e)));
                        });
                    });
                }
            } catch (e) {
                ChatLog.warn('DIAG-BRIDGE', '诊断自身异常: ' + (e && e.message));
            }

            // 初始化 RAG 知识库模块（异步加载配置和列表）
            if (window.App && window.App.rag) {
                window.App.rag.init();
            }

            // 添加测试消息发送
            ChatLog.info('MODEL', 'initBridge: 发送测试消息...');
            window.chrome.webview.postMessage({
                type: 'test',
                message: 'Hello from frontend'
            });
            ChatLog.info('MODEL', 'initBridge: 测试消息已发送');
        } else {
            App.state.excelBridge = null;
            App.state.isConnected = false;
            updateStatus(false);
            ChatLog.info('MODEL', '开发模式：Excel桥接对象未找到');
        }
    }

    // --- 来自 app.js:L2732-L2742 ---
    function updateStatus(connected) {
        const dot = document.getElementById('status-dot');
        const text = document.getElementById('status-text');
        if (!dot || !text) return;  // 【修复】元素不存在时直接返回
        
        if (connected) {
            dot.classList.remove('disconnected');
            text.textContent = '已连接';
        } else {
            dot.classList.add('disconnected');
            text.textContent = '未连接';
        }
        // 同步 tooltip
        const cs = document.getElementById('connection-status');
        if (cs) cs.title = text.textContent;
    }

    // ==================== Excel 上下文感知 ====================

    /**
     * 获取当前Excel表格的表头信息（第一行非空单元格）
     * @returns {Promise<Object>} 表头信息对象
     */
    // --- 来自 app.js:L2750-L2797 ---
    async function getExcelHeaders(sheetName) {
        try {
            // 非 Excel 宿主（Word / PowerPoint）没有工作表概念，直接返回空
            if (window.SmartPilotHost && window.SmartPilotHost.isExcel === false) {
                return null;
            }
            if (!excelBridge) {
                ChatLog.info('MODEL', 'ExcelBridge不可用，无法获取表头');
                return null;
            }

            const selectionJson = await excelBridge.GetSelection();
            const selection = JSON.parse(selectionJson);

            if (!selection || selection.length === 0) {
                ChatLog.info('MODEL', '没有选中区域或选中区域为空');
                return null;
            }

            let headers = [];

            if (Array.isArray(selection)) {
                if (Array.isArray(selection[0])) {
                    headers = selection[0].map(h => h ? h.toString().trim() : '');
                } else {
                    headers = selection.map(h => h ? h.toString().trim() : '');
                }
            } else if (typeof selection === 'object') {
                Object.keys(selection).forEach(key => {
                    headers.push(key);
                });
            }

            headers = headers.filter(h => h && h.length > 0);

            if (headers.length === 0) {
                return null;
            }

            ChatLog.info('MODEL', '📋 检测到Excel表头:', headers);

            return {
                headers: headers,
                headerCount: headers.length,
                timestamp: new Date().toISOString()
            };

        } catch (error) {
            ChatLog.error('MODEL', '获取Excel表头失败:', error);
            return null;
        }
    }

    /**
     * 构建包含Excel上下文的增强消息
     * @param {string} originalMessage - 用户原始消息
     * @returns {Promise<string>} 增强后的消息
     */
    // --- 来自 app.js:L2804-L2824 ---
    async function buildContextAwareMessage(userMessage) {
        const excelInfo = await getExcelHeaders();

        if (!excelInfo || excelInfo.headers.length === 0) {
            return userMessage;
        }

        const contextPrompt = `
【当前Excel表格上下文】
检测到您的Excel表格有以下列（表头）：
${excelInfo.headers.map((h, i) => `${i + 1}. ${h}`).join('\n')}

请根据以上表格结构来理解和处理我的需求。
如果涉及数据填充/录入，请按照这些字段名来组织数据格式。

---
我的原始请求：${userMessage}`;

        ChatLog.info('MODEL', '✅ 已构建带Excel上下文的消息');
        return contextPrompt;
    }

    /**
     * 检测是否为数据填充/录入类请求
     * @param {string} message - 用户消息
     * @returns {boolean} 是否为数据填充请求
     */
    // --- 来自 app.js:L2831-L2833 ---
    function isDataFillRequest(message) {
        return false;
    }

    // ==================== Tab 切换 ====================

    /**
     * 保存当前UI状态到localStorage
     */
    // --- 来自 app.js:L2840-L2851 ---
    function saveUIState() {
        try {
            const state = {
                activeTab: document.querySelector('.tab-btn.active')?.dataset.tab || 'chat',
                chatScrollTop: document.getElementById('chat-messages')?.scrollTop || 0,
                editorCollapsed: document.getElementById('left-panel')?.classList.contains('collapsed') || false
            };
            localStorage.setItem('excelai_ui_state', JSON.stringify(state));
        } catch (e) {
            ChatLog.warn('MODEL', '保存UI状态失败:', e);
        }
    }

    /**
     * 恢复上次UI状态
     */
    // --- 来自 app.js:L2856-L2894 ---
    function restoreUIState() {
        try {
            const savedState = localStorage.getItem('excelai_ui_state');
            if (!savedState) return;

            const state = JSON.parse(savedState);

            // 恢复标签页：默认固定为「对话」视图，不恢复上次标签（避免刷新后进入知识库等其它面板）

            // 恢复滚动位置（延迟执行以确保内容已渲染）
            if (state.chatScrollTop) {
                const chatMessages = document.getElementById('chat-messages');
                if (chatMessages) {
                    chatMessages.scrollTop = state.chatScrollTop;
                }
            }

            // 代码编辑器默认折叠：由 js-app-init 初始化统一强制折叠（刷新后默认显示「对话+欢迎页」），
            // 此处不再恢复旧的编辑器状态，避免单元误用（旧逻辑用不存在的 .collapsed 类判断导致可能误展开）。

            // 刷新后确保对话视图显示像素欢迎大字（如已被移除则重新注入）
            if (typeof window.ensurePixelWelcome === 'function') {
                try { window.ensurePixelWelcome(); } catch (e) { (window.ChatLog || console).warn?.('UI', '[model] 注入像素欢迎大字失败:', e); }
            }

            ChatLog.info('MODEL', 'UI状态已恢复:', state);
        } catch (e) {
            ChatLog.warn('MODEL', '恢复UI状态失败:', e);
        }
    }

    // --- 来自 app.js:L2896-L2922 ---
    function initTabs() {
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', function() {
                document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
                document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));

                this.classList.add('active');
                const tabName = this.getAttribute('data-tab');
                document.getElementById(tabName + '-panel').classList.add('active');

                // 保存当前标签页状态
                saveUIState();

                // 切换到历史、设置、Skill 或知识库标签时，自动收起代码编辑器
                if (tabName === 'history' || tabName === 'settings' || tabName === 'skills' || tabName === 'knowledge') {
                    collapseEditor();
                }

                // 切换到 Skill 管理标签时，加载 Skill 列表和识别类型列表
                if (tabName === 'skills') {
                    loadCustomSkills();
                }

                // 切换到知识库标签时，初始化 RAG 模块
                if (tabName === 'knowledge' && window.App && window.App.rag) {
                    window.App.rag.init();
                }

                // 输入框和模型选择栏只在对话面板显示
                var inputArea = document.getElementById('input-area');
                if (inputArea) {
                    inputArea.style.display = (tabName === 'chat') ? '' : 'none';
                }
                var modelSelector = document.getElementById('model-selector-area');
                if (modelSelector) {
                    modelSelector.style.display = (tabName === 'chat') ? '' : 'none';
                }
                // 快捷技能栏只在对话面板显示
                var quickShortcutBar = document.getElementById('quick-shortcut-bar');
                if (quickShortcutBar) {
                    quickShortcutBar.style.display = (tabName === 'chat') ? '' : 'none';
                }
            });
        });

        // 扩展面板已合并为单一 Skill 管理界面
    }

    // ==================== 扩展面板子标签页切换（已废弃，保留兼容） ====================
    // --- 来自 app.js:L2925-L2928 ---
    function initSkillsSubTabs() {
        // 识别类型已合并到 Skill 管理，此函数保留用于兼容
        ChatLog.info('MODEL', '[Deprecated] initSkillsSubTabs 已废弃，识别类型已合并到 Skill 管理');
    }

    // 收起代码编辑器
    // --- 来自 app.js:L2931-L2936 ---
    function collapseEditor() {
        const leftPanel = document.getElementById('left-panel');
        if (leftPanel && leftPanel.classList.contains('expanded')) {
            // 直接收起左侧面板，不调用 toggleEditor()，避免 activateChatView() 切回对话标签
            // 确保切换历史/扩展/知识库等标签时，历史列表等正确显示而非被切回对话
            var tpPanel = document.getElementById('task-planner-panel');
            if (tpPanel && tpPanel.classList.contains('active')) {
                tpPanel.classList.remove('active');
                tpPanel.style.display = 'none';
            }
            var editorPanel = document.getElementById('editor-panel');
            if (editorPanel) editorPanel.style.display = '';

            leftPanel.classList.remove('expanded');

            var codeContainer = document.getElementById('code-with-lines');
            var actions = document.getElementById('editor-actions');
            var title = document.getElementById('editor-title');
            var langSelect = document.getElementById('editor-language');

            if (codeContainer) {
                codeContainer.classList.add('is-hidden');
                codeContainer.style.display = 'none';
            }
            if (actions) {
                actions.classList.add('is-hidden');
                actions.style.display = 'none';
            }
            if (title) {
                title.classList.add('is-hidden');
                title.style.display = 'none';
            }
            if (langSelect) {
                langSelect.classList.add('is-hidden');
                langSelect.style.display = 'none';
            }
        }
    }

    // ==================== 提供商切换 ====================
    // --- 来自 app.js:L2939-L3006 ---
    function onProviderChange(provider) {
        const providerEl = document.getElementById('api-type');
        // 设置面板元素可能不存在（例如在 index.html 中调用时）
        if (!providerEl && !provider) {
            ChatLog.debug('MODEL', '[onProviderChange] api-type 元素不存在且无 provider 参数，跳过');
            return;
        }
        const providerValue = provider || (providerEl ? providerEl.value : '');
        const providerConfig = PROVIDER_CONFIGS[providerValue];

        // 更新 config.apiType
        config.apiType = providerValue;
        ChatLog.info('MODEL', 'onProviderChange: 设置 config.apiType =', providerValue);

        // 加载该提供商已保存的配置（如果有）
        const savedProviderConfig = config.providers?.[providerValue];
        if (savedProviderConfig) {
            // 使用已保存的配置
            config.apiKey = savedProviderConfig.apiKey || '';
            config.apiUrl = savedProviderConfig.apiUrl || providerConfig?.apiUrl || '';
            config.model = savedProviderConfig.model || providerConfig?.model || '';

            document.getElementById('api-key').value = config.apiKey;
            document.getElementById('model').value = config.model;

            ChatLog.info('MODEL', 'onProviderChange: 加载已保存的配置', savedProviderConfig);
        } else {
            // 使用默认配置
            config.apiKey = '';
            config.apiUrl = providerConfig?.apiUrl || '';
            config.model = providerConfig?.model || '';

            document.getElementById('api-key').value = '';
            document.getElementById('model').value = config.model;

            ChatLog.info('MODEL', 'onProviderChange: 使用默认配置');
        }

        // 更新 API URL 输入框
        if (config.apiUrl) {
            document.getElementById('api-url').value = config.apiUrl;
            document.getElementById('api-url').placeholder = '';
        } else if (providerConfig?.placeholder) {
            // 如果没有默认 URL 但有占位符（如 Azure）
            document.getElementById('api-url').value = '';
            document.getElementById('api-url').placeholder = providerConfig.placeholder;
            ChatLog.info('MODEL', 'onProviderChange: 设置 API URL 占位符为', providerConfig.placeholder);
        }

        // 显示提供商描述（如果有）
        if (providerConfig?.description) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">' + providerConfig.description + '</span>', true) : showToast(providerConfig.description);
        }

        // 显示/隐藏 API 密钥输入
        const apiKeyGroup = document.getElementById('api-key-group');
        apiKeyGroup.style.display = providerConfig.needsKey ? 'block' : 'none';

        // 显示/隐藏检测按钮（在 flex 容器中，使用空字符串恢复默认 display）
        const detectBtn = document.getElementById('detect-btn');
        detectBtn.style.display = providerConfig.isLocal ? '' : 'none';

        // 本地模型安装引导卡片：云端提供商直接隐藏；本地提供商只切换对应引导块，
        // 卡片整体显隐由 updateLocalGuideStatus 按检测结果决定（检测到模型 → 隐藏）
        const guideCard = document.getElementById('local-model-guide');
        if (guideCard && !providerConfig.isLocal) {
            guideCard.style.display = 'none';
        }
        if (providerConfig.isLocal) {
            const ollamaGuide = document.getElementById('local-guide-ollama');
            const llamacppGuide = document.getElementById('local-guide-llamacpp');
            if (ollamaGuide) ollamaGuide.style.display = providerValue === 'llamacpp' ? 'none' : '';
            if (llamacppGuide) llamacppGuide.style.display = providerValue === 'llamacpp' ? '' : 'none';
        }

        // 渲染当前平台的模型列表（本地 → 本地检测模型，云端 → 该平台预设模型）
        const modelListEl = document.getElementById('model-list');
        if (modelListEl) {
            renderAllModelList(modelListEl, detectedModels);
        }

        // 更新对话框模型选择器
        updateChatModelSelector();

        // 重置当前对话模型（仅当 chat-model-select 存在时）
        currentChatModel = '';
        const chatModelSelect = document.getElementById('chat-model-select');
        if (chatModelSelect) {
            chatModelSelect.value = '';
        }
    }

    // ==================== 本地模型检测 ====================
    // --- 来自 app.js:L3009-L3013 ---
    // 在途锁：同一时间只允许一趟本地模型检测。
    // 独立于死循环之外的冗余：init 阶段有多个入口（js-app-init 后台任务、本地模型自动检测、
    // refreshModelsList 等）都会并发调用 detectLocalModels。没有新操作时这些调用会被重复执行，
    // 重复走 Bridge 检测并反复刷新下拉框、刷屏打印诊断日志。此锁把并发调用合并为一次。
    let _detectLocalModelsInFlight = false;

    async function detectLocalModels() {
        // 始终检测 Ollama，不依赖于当前选择的提供商
        // 这样即使用户当前选择了 Moonshot，也能检测到本地 Ollama 模型
        if (_detectLocalModelsInFlight) {
            ChatLog.debug('MODEL', '本地模型检测已在执行中，合并本次触发');
            return;
        }
        _detectLocalModelsInFlight = true;
        try {
            await detectOllamaModels();
        } finally {
            _detectLocalModelsInFlight = false;
        }
    }

    /**
     * 本地模型的展示名：llama.cpp 返回的 id 常带路径（如 models\xxx.gguf），
     * 剥离路径前缀和 .gguf 后缀，保持与 Ollama 简洁命名的视觉一致性。
     * 模型标识（config.model / 请求 model 字段）仍用原始 name，不受影响。
     */
    function getLocalModelDisplayName(model) {
        const raw = (model && model.name) || '';
        if ((model && model.provider) !== 'llamacpp') return raw;
        let name = raw.split(/[\\/]/).pop() || raw;
        return name.replace(/\.gguf$/i, '');
    }

    /**
     * 渲染设置面板的模型列表（跟随当前选中的平台）：
     *  - 本地平台（Ollama / llama.cpp）：仅显示本地检测模型
     *  - 云端平台：仅显示该平台的预设模型
     * 点击云端模型自动切换到对应平台；点击本地模型仅切换模型。
     * 模型名通过 data-* 属性传递，避免内联 onclick 单引号转义破坏含反斜杠的 GGUF 路径。
     */
    function getShortProviderName(providerKey) {
        const short = {
            'openai': 'OpenAI', 'azure': 'Azure', 'wenxin': '文心', 'qianwen': '通义',
            'zhipu': '智谱', 'xunfei': '星火', 'moonshot': 'Kimi', 'deepseek': 'DeepSeek', 'volcengine': '火山', 'custom': '自定义',
            'ollama': 'Ollama', 'llamacpp': 'llama.cpp'
        };
        return short[providerKey] || providerKey;
    }

    function renderModelItemHtml(name, displayName, provider, isSelected) {
        return `<div class="model-item ${isSelected ? 'selected' : ''}" data-model="${name}" data-provider="${provider}" onclick="selectModel(this.dataset.model, event, this.dataset.provider)">
            <span class="model-name" title="${name}">${displayName}</span>
            <span class="model-tag ${provider === 'ollama' ? 'local' : provider === 'llamacpp' ? 'llamacpp' : 'cloud'}">${getShortProviderName(provider)}</span>
        </div>`;
    }

    /**
     * 刷新「本地模型安装引导」卡片：已检测到本地模型 → 整卡隐藏（装好不用再看引导）；
     * 未检测到（含检测失败）→ 显示引导并在状态行提示按步骤安装后再检测。
     * 卡片不存在（非设置面板场景）时静默跳过。
     */
    function updateLocalGuideStatus(count) {
        const card = document.getElementById('local-model-guide');
        if (card) card.style.display = count > 0 ? 'none' : '';
        const el = document.getElementById('local-guide-status');
        if (!el) return;
        let text = '';
        try {
            if (window.App && App.i18n && typeof App.i18n.t === 'function') {
                text = count > 0
                    ? App.i18n.t('settings.localGuide.statusOk', { count: count })
                    : App.i18n.t('settings.localGuide.statusNone');
            }
        } catch (_) { /* ignore */ }
        // 词条缺失时 t() 返回 key 本身，回退到内置中文
        if (!text || text.indexOf('settings.localGuide.') === 0) {
            text = count > 0
                ? `已检测到 ${count} 个本地模型，可在上方模型列表选择。`
                : '尚未检测到本地模型。安装完成后点「检测本地模型」。';
        }
        el.textContent = text;
    }

    function renderAllModelList(modelList, models) {
        if (!modelList) return;
        const current = config.apiType;
        const currentCfg = PROVIDER_CONFIGS[current];
        let html = '';

        if (!currentCfg) {
            // 平台配置未加载时给出醒目诊断，便于定位问题
            html += '<div style="padding:8px 14px;font-size:12px;color:#d32f2f;background:#fff3cd;border:1px solid #f0c36d;border-radius:4px;margin:8px 10px;">⚠️ 平台模型配置未加载（config/models-config.js 未生效）</div>';
            modelList.innerHTML = html;
            return;
        }

        if (!currentCfg.isLocal) {
            // 云端平台：仅显示当前平台的预设模型
            const presetModels = currentCfg.models || [];
            if (presetModels.length > 0) {
                html += `<div style="padding:6px 14px;font-size:12px;font-weight:600;color:#555;background:#f2f4f7;border-bottom:1px solid #e5e7eb;border-left:3px solid #4a90d9;">${getProviderName(current)} · 云端</div>`;
                presetModels.forEach(m => {
                    const mName = typeof m === 'object' ? m.name : m;
                    html += renderModelItemHtml(mName, mName, current, config.model === mName);
                });
            } else {
                html += '<div style="text-align:center;padding:10px;color:#999;">该平台暂无预设模型</div>';
            }
        } else {
            // 本地平台：仅显示当前本地服务检测到的模型
            // （Ollama 平台只显示 Ollama 模型，llama.cpp 平台只显示 llama.cpp 模型）
            const targetProvider = current; // 'ollama' 或 'llamacpp'
            const localModels = (models || []).filter(m => (m.provider || 'ollama') === targetProvider);
            if (localModels.length > 0) {
                updateLocalGuideStatus(localModels.length);
                html += `<div style="padding:6px 14px;font-size:12px;font-weight:600;color:#555;background:#f2f4f7;border-bottom:1px solid #e5e7eb;border-left:3px solid #6c757d;">本地模型</div>`;
                localModels.forEach(m => {
                    const disp = getLocalModelDisplayName(m);
                    html += renderModelItemHtml(m.name, disp, targetProvider, config.model === m.name);
                });
            } else {
                updateLocalGuideStatus(0);
                html += targetProvider === 'llamacpp'
                    ? '<div style="text-align:center;padding:10px;color:#999;">未检测到 llama.cpp 模型<br>请确认 llama.cpp 服务已启动（默认 http://localhost:8080）<br>然后点击"检测本地模型"</div>'
                    : '<div style="text-align:center;padding:10px;color:#999;">未找到 Ollama 模型<br>请先运行: ollama pull llama2<br>然后点击"检测本地模型"</div>';
            }
        }

        modelList.innerHTML = html;
    }

    // --- 来自 app.js:L3015-L3174 ---
    async function detectOllamaModels() {
        const modelList = document.getElementById('model-list');
        const isLocalProvider = PROVIDER_CONFIGS[config.apiType]?.isLocal;
        if (!modelList) {
            // 设置面板未渲染，跳过 UI 渲染，后台仍然尝试检测
            ChatLog.info('MODEL', 'model-list 元素不存在，跳过 UI 渲染，继续后台检测');
        } else if (isLocalProvider) {
            // 仅本地平台显示"检测中"，云端平台直接渲染预设模型，避免闪烁
            modelList.innerHTML = '<div style="text-align:center;padding:10px;">检测中</div>';
        }

        try {
            // ============================================================
            // 【方案一】优先通过后端 Bridge 获取 Ollama 模型列表（推荐）
            // 后端 .NET HttpClient 访问 localhost 无 CORS 限制
            // ============================================================
            try {
                if (window.chrome?.webview?.hostObjects?.excelBridge) {
                    ChatLog.info('MODEL', '尝试通过后端 Bridge 获取 Ollama 模型列表...');
                    // 加 5 秒超时保护：Bridge 方法不可用或挂起时降级到直连，避免一直卡在"检测中"
                    const bridgeResult = await Promise.race([
                        window.chrome.webview.hostObjects.excelBridge.GetOllamaModelsAsync(),
                        new Promise((_, reject) => setTimeout(() => reject(new Error('Bridge调用超时')), 5000))
                    ]);
                    const bridgeData = JSON.parse(bridgeResult);

                    if (bridgeData.success && bridgeData.models && bridgeData.models.length > 0) {
                        ChatLog.info('MODEL', `后端Bridge返回 ${bridgeData.count} 个本地模型:`, bridgeData.models);

                        // 兼容两种返回格式：旧版字符串数组（全部视为 Ollama），
                        // 新版 {name, source} 对象数组（source 区分 ollama / llamacpp）
                        detectedModels = bridgeData.models.map(m => {
                            if (typeof m === 'string') {
                                return { name: m, provider: 'ollama' };
                            }
                            return {
                                name: m.name || m.model || '',
                                provider: (m.source === 'llamacpp') ? 'llamacpp' : 'ollama'
                            };
                        }).filter(m => m.name);

                        // 渲染模型列表
                        renderAllModelList(modelList, detectedModels);

                        ChatLog.info('MODEL', `✅ 通过后端 Bridge 检测到 ${detectedModels.length} 个 Ollama 本地模型`);
                        updateChatModelSelector();
                        return detectedModels;
                    }
                    ChatLog.info('MODEL', '后端Bridge未返回模型，将尝试直连...');
                }
            } catch (bridgeError) {
                ChatLog.info('MODEL', '后端Bridge获取模型失败，降级到直连模式:', bridgeError.message);
            }

            // ============================================================
            // 【方案二】直接 fetch 访问 Ollama（可能受 CORS 限制）
            // ============================================================
            // 检查 fetch 是否可用
            if (typeof fetch === 'undefined') {
                throw new Error('浏览器不支持 fetch API');
            }

            let responseText;

            // 按当前本地平台选择检测端点：
            //  - Ollama：http://localhost:11434/api/tags（返回 { models: [{ name }] }）
            //  - llama.cpp：http://localhost:8080/v1/models（OpenAI 兼容，返回 { data: [{ id }] }）
            const isLlamacpp = config.apiType === 'llamacpp';
            const localEndpoint = isLlamacpp ? 'http://localhost:8080/v1/models' : 'http://localhost:11434/api/tags';
            const localLabel = isLlamacpp ? 'llama.cpp' : 'Ollama';
            ChatLog.info('MODEL', `使用 fetch 访问 ${localLabel} (${localEndpoint})`);

            // 首先尝试正常模式
            let response;
            try {
                response = await fetch(localEndpoint, {
                    method: 'GET',
                    headers: { 'Content-Type': 'application/json' }
                });
            } catch (corsError) {
                // 如果失败，可能是 CORS 问题，尝试 no-cors 模式
                ChatLog.info('MODEL', '正常模式失败，尝试 no-cors 模式');
                try {
                    response = await fetch(localEndpoint, {
                        method: 'GET',
                        mode: 'no-cors',
                        headers: { 'Content-Type': 'application/json' }
                    });
                    // no-cors 模式下无法读取响应内容
                    if (!response || response.type === 'opaque') {
                        throw new Error('CORS 限制：无法读取响应。请使用方案三设置 OLLAMA_ORIGINS');
                    }
                } catch (noCorsError) {
                    ChatLog.error('MODEL', 'no-cors模式也失败了:', noCorsError);
                    // 两种方式都失败，静默降级：仍渲染当前平台的模型分组，避免列表卡在"检测中"
                    if (modelList) renderAllModelList(modelList, []);
                    return [];
                }
            }

            // 防御性检查：确保 response 存在
            if (!response) {
                ChatLog.warn('MODEL', `${localLabel}响应为空，返回空模型列表`);
                // 静默降级：仍渲染当前平台的模型分组，避免列表卡在"检测中"
                if (modelList) renderAllModelList(modelList, []);
                return [];
            }

            if (!response.ok) {
                throw new Error(`无法连接到 ${localLabel} (HTTP ${response.status})`);
            }
            responseText = await response.text();

            // 调试：显示原始响应的前 200 个字符
            ChatLog.info('MODEL', `${localLabel} 原始响应:`, responseText.substring(0, 200));

            // 检查返回的是否是 JSON
            const trimmedResponse = responseText.trim();
            if (!trimmedResponse.startsWith('{') && !trimmedResponse.startsWith('[')) {
                // 可能是 HTML 错误页面或其他错误
                throw new Error(`${localLabel} 返回了非 JSON 响应: ` + trimmedResponse.substring(0, 100));
            }

            const result = JSON.parse(responseText);
            if (result.error) {
                throw new Error(result.error);
            }
            // Ollama 用 models[].name；llama.cpp（OpenAI 兼容）用 data[].id
            const models = isLlamacpp ? (result.data || []) : (result.models || []);
            ChatLog.info('MODEL', `${localLabel} 返回模型数量:`, models.length);

            // WebView2 调试：在界面上显示日志
            const debugInfo = document.getElementById('debug-info');
            if (debugInfo) {
                debugInfo.innerHTML += '<div>返回模型数量: ' + models.length + '</div>';
            }

            detectedModels = models.map(m => ({
                name: isLlamacpp ? (m.id || m.name || '') : (m.model || m.name || ''),
                provider: isLlamacpp ? 'llamacpp' : 'ollama',
                size: m.size,
                modified_at: m.modified_at
            }));
            ChatLog.info('MODEL', '转换后的 detectedModels:', detectedModels);
            updateChatModelSelector();

            // 渲染模型列表（空列表时显示空态提示）
            renderAllModelList(modelList, detectedModels);

            // 进度条已显示,无需 toast
        } catch (error) {
            // 使用统一错误处理（仅本地平台弹窗，云端平台静默降级为渲染该平台预设模型）
            if (PROVIDER_CONFIGS[config.apiType]?.isLocal) {
                window.ErrorHandler.showError(error, '模型检测');
            }
            ChatLog.error('MODEL', '检测失败:', error);
            let errorDetail = error.message;

            // 针对常见错误提供更详细的帮助
            if (error.message.includes('Failed to fetch') || error.message.includes('CORS')) {
                const psCommand = '$env:OLLAMA_ORIGINS="http://localhost,https://localhost,http://127.0.0.1,https://127.0.0.1,http://*,https://*,*"; ollama serve';
                const cmdCommand = 'set OLLAMA_ORIGINS=http://localhost,https://localhost,http://127.0.0.1,https://127.0.0.1,http://*,https://*,* && ollama serve';

                errorDetail += '<br><br><b style="color:#ff9800;">方案三：完整 OLLAMA_ORIGINS 配置</b><br><br>';
                errorDetail += '<b>步骤 1：关闭所有 Ollama 进程</b><br>';
                errorDetail += '在 PowerShell 中运行：<br>';
                errorDetail += '<div style="display:flex;align-items:center;gap:8px;margin:5px 0;">';
                errorDetail += '<code style="background:#f5f5f5;padding:4px 8px;border-radius:3px;flex:1;">taskkill /F /IM ollama.exe</code>';
                errorDetail += '<button onclick="copyToClipboard(\'taskkill /F /IM ollama.exe\')" style="padding:4px 8px;background:var(--primary-color);color:white;border:none;border-radius:3px;cursor:pointer;font-size:11px;">复制</button>';
                errorDetail += '</div><br>';

                errorDetail += '<b>步骤 2：设置环境变量并启动 Ollama</b><br>';
                errorDetail += '在 PowerShell 中运行（推荐）：<br>';
                errorDetail += '<div style="display:flex;align-items:center;gap:8px;margin:5px 0;">';
                errorDetail += '<code style="background:#f5f5f5;padding:4px 8px;border-radius:3px;flex:1;word-break:break-all;font-size:11px;">' + psCommand + '</code>';
                errorDetail += '<button onclick="copyToClipboard(\'' + psCommand.replace(/"/g, '&quot;') + '\')" style="padding:4px 8px;background:var(--primary-color);color:white;border:none;border-radius:3px;cursor:pointer;font-size:11px;">复制</button>';
                errorDetail += '</div><br>';

                errorDetail += '或在 CMD 中运行：<br>';
                errorDetail += '<div style="display:flex;align-items:center;gap:8px;margin:5px 0;">';
                errorDetail += '<code style="background:#f5f5f5;padding:4px 8px;border-radius:3px;flex:1;word-break:break-all;font-size:11px;">' + cmdCommand + '</code>';
                errorDetail += '<button onclick="copyToClipboard(\'' + cmdCommand.replace(/"/g, '&quot;') + '\')" style="padding:4px 8px;background:var(--primary-color);color:white;border:none;border-radius:3px;cursor:pointer;font-size:11px;">复制</button>';
                errorDetail += '</div><br><br>';
                errorDetail += '<b style="color:#ff9800;">注意：</b>如果上述命令报错 "ollama 不是内部或外部命令"，<br>';
                errorDetail += '请先安装 Ollama：<a href="https://ollama.ai" target="_blank" style="color:var(--primary-color);">https://ollama.ai</a>';
            }

            // 检测失败时仍渲染当前平台的模型列表（本地 → 空态提示，云端 → 该平台预设模型）；
            // 失败横幅仅本地平台显示（云端平台渲染的是预设模型，与本地检测无关）
            if (modelList) {
                renderAllModelList(modelList, []);
                if (PROVIDER_CONFIGS[config.apiType]?.isLocal) {
                    const banner = document.createElement('div');
                    banner.style.cssText = 'padding:8px 14px;font-size:12px;color:#d32f2f;background:#fff3cd;border-bottom:1px solid #f0f0f0;';
                    banner.textContent = '⚠️ 本地模型检测失败：' + (error.message || '未知错误');
                    modelList.prepend(banner);
                }
            }
        }

        // 显示检测结果摘要
        if (detectedModels.length > 0) {
            ChatLog.info('MODEL', `✅ 检测到 ${detectedModels.length} 个 Ollama 本地模型:`);
            detectedModels.forEach(m => ChatLog.info('MODEL', `  - ${m.name}`));
        } else {
            ChatLog.info('MODEL', '⚠️ 未检测到 Ollama 本地模型');
        }

        // 检测完成后刷新对话模型下拉框，避免首次加载时本地模型不显示
        updateChatModelSelector();
    }

    // --- 来自 app.js:L3176-L3187 ---
    function selectModel(modelName, event, provider) {
        if (!modelName) return;
        // 云端模型：点击时切换到对应平台（onProviderChange 会加载该平台已保存的 apiKey/apiUrl/model）
        if (provider && provider !== 'ollama' && provider !== 'llamacpp' && config.apiType !== provider) {
            onProviderChange(provider);
        }
        const modelEl = document.getElementById('model');
        if (modelEl) modelEl.value = modelName;
        config.model = modelName;

        // 更新选中状态
        document.querySelectorAll('.model-item').forEach(item => {
            item.classList.remove('selected');
        });
        if (event && event.currentTarget) {
            event.currentTarget.classList.add('selected');
        }
    }

    // ==================== 配置管理 ====================
    // --- 来自 app.js:L3190-L3285 ---
    async function loadConfig() {
        const saved = localStorage.getItem(CONFIG_KEY);
        ChatLog.info('MODEL', 'loadConfig - localStorage 保存的配置:', saved);
        if (saved) {
            const savedConfig = JSON.parse(saved);
            ChatLog.info('MODEL', 'loadConfig - 解析后的配置:', savedConfig);
            
            // 【修复】直接修改 App.config 而不是创建新对象，避免 config 与 App.config 分裂
            Object.assign(App.config, savedConfig);
            config = App.config; // 同步全局 config 引用

            if (config.apiUrl) config.apiUrl = config.apiUrl.trim();
            if (config.apiKey) config.apiKey = config.apiKey.trim();

            // 如果保存了 providers 对象，恢复当前提供商的配置
            if (config.providers && config.apiType && config.providers[config.apiType]) {
                const providerSettings = config.providers[config.apiType];
                config.apiKey = (providerSettings.apiKey || '').trim();
                const defaultUrl = PROVIDER_CONFIGS[config.apiType]?.apiUrl || '';
                config.apiUrl = (providerSettings.apiUrl || defaultUrl).trim();
                config.model = providerSettings.model || '';
            }
        } else {
            // 【修复】localStorage 为空时（WebView2 实例间不共享），从后端 Bridge 读取配置
            try {
                if (window.chrome?.webview?.hostObjects?.excelBridge) {
                    ChatLog.info('MODEL', 'loadConfig - localStorage 为空，尝试从后端 Bridge 读取配置...');
                    const bridgeResult = await Promise.race([
                        window.chrome.webview.hostObjects.excelBridge.GetAIConfig(),
                        new Promise((_, reject) => setTimeout(() => reject(new Error('Bridge调用超时')), 5000))
                    ]);
                    ChatLog.info('MODEL', 'loadConfig - Bridge 返回原始数据:', bridgeResult);
                    const bridgeConfig = JSON.parse(bridgeResult);
                    ChatLog.info('MODEL', 'loadConfig - Bridge 解析后配置:', bridgeConfig);
                    if (!bridgeConfig.error && (bridgeConfig.apiKey || bridgeConfig.provider === 'ollama' || bridgeConfig.apiType === 'ollama')) {
                        ChatLog.info('MODEL', 'loadConfig - 从后端 Bridge 读取到配置:', bridgeConfig.provider || bridgeConfig.apiType);
                        config.apiType = bridgeConfig.provider || bridgeConfig.apiType || 'ollama';
                        config.model = bridgeConfig.model || '';
                        config.apiKey = (bridgeConfig.apiKey || '').trim();
                        config.apiUrl = (bridgeConfig.apiUrl || '').trim();
                        // 构建 providers 结构
                        if (!config.providers) config.providers = {};
                        config.providers[config.apiType] = {
                            apiKey: config.apiKey,
                            apiUrl: config.apiUrl,
                            model: config.model
                        };
                    } else {
                        ChatLog.warn('MODEL', 'loadConfig - Bridge 返回数据无效:', bridgeConfig.error || 'apiKey 为空');
                        
                        // Ollama 模式下允许空 apiKey
                        if (!bridgeConfig.error && (!bridgeConfig.apiKey || bridgeConfig.apiKey === '')) {
                            const apiType = bridgeConfig.provider || bridgeConfig.apiType || 'ollama';
                            if (PROVIDER_CONFIGS[apiType] && !PROVIDER_CONFIGS[apiType].needsKey) {
                                ChatLog.info('MODEL', 'loadConfig - Ollama 模式允许空 apiKey，使用默认配置');
                                config.apiType = apiType;
                                config.model = bridgeConfig.model || PROVIDER_CONFIGS[apiType].model || '';
                                config.apiKey = 'local';
                                config.apiUrl = bridgeConfig.apiUrl || PROVIDER_CONFIGS[apiType].apiUrl || '';
                                if (!config.providers) config.providers = {};
                                config.providers[config.apiType] = {
                                    apiKey: config.apiKey,
                                    apiUrl: config.apiUrl,
                                    model: config.model
                                };
                            }
                        }
                    }
                } else {
                    ChatLog.warn('MODEL', 'loadConfig - excelBridge 对象不存在');
                }
            } catch (bridgeErr) {
                ChatLog.warn('MODEL', 'loadConfig - 从后端 Bridge 读取配置失败:', bridgeErr.message);
            }
        }

        // 验证和清理配置 - 确保 apiUrl 与当前提供商匹配
        if (config.apiType && PROVIDER_CONFIGS[config.apiType]) {
            const providerConfig = PROVIDER_CONFIGS[config.apiType];
            const defaultUrl = providerConfig.apiUrl;

            // 如果当前 apiUrl 与默认 URL 不匹配（可能是之前切换提供商留下的）
            // 且不是用户自定义的 URL，则重置为默认 URL
            if (config.apiUrl && config.apiUrl !== defaultUrl) {
                // 检查是否可能是其他提供商的 URL
                const isOtherProviderUrl = Object.keys(PROVIDER_CONFIGS).some(key => {
                    if (key === config.apiType) return false;
                    return config.apiUrl === PROVIDER_CONFIGS[key].apiUrl;
                });

                if (isOtherProviderUrl) {
                    ChatLog.info('MODEL', `loadConfig - 检测到 URL 不匹配，重置为 ${config.apiType} 的默认 URL`);
                    ChatLog.info('MODEL', `  当前 URL: ${config.apiUrl}`);
                    ChatLog.info('MODEL', `  默认 URL: ${defaultUrl}`);
                    config.apiUrl = defaultUrl;
                    // 【修复】同步更新 providers 中的 URL 和 Key，避免 selectModel 使用错误的配置
                    if (config.providers && config.providers[config.apiType]) {
                        config.providers[config.apiType].apiUrl = defaultUrl;
                        config.providers[config.apiType].apiKey = '';
                    }
                    config.apiKey = '';
                }
            }

            // 确保 apiUrl 有值
            if (!config.apiUrl) {
                config.apiUrl = defaultUrl;
                ChatLog.info('MODEL', 'loadConfig - 使用默认 API URL:', config.apiUrl);
            }

            if (config.model && typeof config.model === 'string' && providerConfig.models) {
                const modelName = config.model;
                const isValidModel = providerConfig.models.some(m => {
                    const mName = typeof m === 'string' ? m : (m && m.name ? m.name : '');
                    if (!mName) return false;
                    if (modelName === mName) return true;
                    if (providerConfig.isLocal) {
                        return modelName.startsWith(mName + ':') || mName.startsWith(modelName + ':');
                    }
                    return false;
                }) || hasDetectedLocalModels(config.apiType) && detectedModels
                        .some(m => m.provider === config.apiType && m.name === modelName);
                if (!isValidModel) {
                    ChatLog.info('MODEL', `loadConfig - 模型 ${config.model} 不属于 ${config.apiType}，重置为默认模型`);
                    config.model = providerConfig.model;
                    // 【修复】同步更新已保存的提供商配置，否则下方 savedProviderConfig 覆盖会把旧模型名盖回来
                    // （例如 moonshot-v1-32k 已下线，必须被重置为 kimi-k3）
                    if (config.providers && config.providers[config.apiType]) {
                        config.providers[config.apiType].model = providerConfig.model;
                    }
                }
            }
        }

        ChatLog.info('MODEL', 'loadConfig - 合并后 config.apiType:', config.apiType);
        ChatLog.info('MODEL', 'loadConfig - 合并后 config.apiUrl:', config.apiUrl);
        ChatLog.info('MODEL', 'loadConfig - 合并后 config.model:', config.model);

        // 加载当前选中提供商的完整配置
        const providerConfig = PROVIDER_CONFIGS[config.apiType];
        const savedProviderConfig = config.providers?.[config.apiType];

        // 优先使用已保存的提供商配置
        if (savedProviderConfig) {
            config.apiKey = savedProviderConfig.apiKey || '';
            config.apiUrl = savedProviderConfig.apiUrl || providerConfig?.apiUrl || '';
            config.model = savedProviderConfig.model || providerConfig?.model || '';
        }

        // 仅当设置面板 DOM 元素存在时才更新 UI（index.html 中没有这些元素）
        var apiTypeEl = document.getElementById('api-type');
        if (apiTypeEl) apiTypeEl.value = config.apiType;

        var apiUrlEl = document.getElementById('api-url');
        if (apiUrlEl) {
            if (!config.apiUrl && providerConfig && providerConfig.placeholder) {
                apiUrlEl.value = '';
                apiUrlEl.placeholder = providerConfig.placeholder;
            } else {
                apiUrlEl.value = config.apiUrl || '';
                apiUrlEl.placeholder = '';
            }
        }

        var apiKeyEl = document.getElementById('api-key');
        if (apiKeyEl) apiKeyEl.value = config.apiKey;
        var modelEl = document.getElementById('model');
        if (modelEl) modelEl.value = config.model;
        var tempEl = document.getElementById('temperature');
        if (tempEl) tempEl.value = config.temperature;
        var tempRange = document.getElementById('temperature-range');
        if (tempRange) tempRange.value = config.temperature;
        var ctxModeEl = document.getElementById('context-mode');
        if (ctxModeEl) ctxModeEl.value = config.contextMode || 'auto';
        var dangerAskEl = document.getElementById('danger-always-ask');
        if (dangerAskEl) dangerAskEl.checked = !!config.dangerAlwaysAsk;
        var localeEl = document.getElementById('ui-locale');
        if (localeEl) {
            var loc = config.uiLocale || (window.App && App.i18n && App.i18n.getLocale && App.i18n.getLocale()) || 'zh';
            localeEl.value = normalizeUiLocale(loc);
        }

        updateContextStats();
    }

    /**
     * 测试 AI 配置是否可用
     */
    // --- 来自 app.js:L3289-L3414 ---
    async function testAIConfig() {
        const resultDiv = document.getElementById('test-result');
        const emptyMsg = document.getElementById('test-result-empty');
        const provider = document.getElementById('api-type').value;
        const apiUrl = document.getElementById('api-url').value;
        const apiKey = document.getElementById('api-key').value;
        const model = document.getElementById('model').value;

        if (emptyMsg) emptyMsg.style.display = 'none';
        resultDiv.style.display = 'block';
        resultDiv.className = '';
        resultDiv.innerHTML = '🔄 正在测试连接...';
        resultDiv.style.padding = '10px 12px';
        resultDiv.style.borderRadius = '6px';
        resultDiv.style.background = '#fff3cd';
        resultDiv.style.color = '#856404';
        resultDiv.style.border = '1px solid #ffeaa7';

        ChatLog.info('MODEL', '测试配置:', { provider, apiUrl, model, hasKey: !!apiKey });

        // 基本验证
        if (!apiUrl) {
            resultDiv.style.background = '#f8d7da';
            resultDiv.style.color = '#721c24';
            resultDiv.style.border = '1px solid #f5c6cb';
            resultDiv.innerHTML = '❌ 错误：API 地址不能为空';
            return;
        }

        const providerConfig = PROVIDER_CONFIGS[provider];
        if (providerConfig.needsKey && !apiKey) {
            resultDiv.style.background = '#f8d7da';
            resultDiv.style.color = '#721c24';
            resultDiv.style.border = '1px solid #f5c6cb';
            resultDiv.innerHTML = '❌ 错误：该提供商需要 API Key';
            return;
        }

        // 保存原始配置（在 try 外部定义，确保 catch 中也能访问）
        const originalConfig = { ...config };

        // 创建 AbortController 用于超时控制
        // Ollama 升级后首次加载模型可能需要更长时间，增加超时时间
        const controller = new AbortController();
        const timeoutId = setTimeout(() => {
            controller.abort();
        }, 30000); // 30秒超时

        try {
            // 构建测试请求
            const testMessages = [
                { role: 'system', content: '你是一个测试助手，请回复"连接成功"即可。' },
                { role: 'user', content: '测试连接' }
            ];

            // 临时使用当前表单中的配置进行测试
            const tempConfig = {
                apiType: provider,
                apiUrl: apiUrl,
                apiKey: apiKey,
                model: model,
                temperature: 0.1
            };

            // 使用临时配置
            Object.assign(config, tempConfig);

            // ===== 直接使用 fetch 调用 API（不依赖 ChatV2/llmClient，确保设置面板独立可用）=====
            const pc = PROVIDER_CONFIGS[provider] || {};
            let resolvedUrl = (apiUrl || '').trim().replace(/\/$/, '');
            if (!resolvedUrl && pc.apiUrl) resolvedUrl = pc.apiUrl.replace(/\/$/, '');

            // 1. 解析最终 URL（处理用户可能填了完整路径或仅域名）
            let finalUrl;
            if (provider === 'ollama') {
                if (/\/api\/chat$/i.test(resolvedUrl)) finalUrl = resolvedUrl;
                else finalUrl = (resolvedUrl || 'http://localhost:11434') + '/api/chat';
            } else {
                // Azure 通常已经包含完整部署路径，不强制补后缀
                if (provider === 'azure' && resolvedUrl) {
                    finalUrl = resolvedUrl;
                } else if (/\/chat\/completions$/i.test(resolvedUrl)) {
                    finalUrl = resolvedUrl;
                } else {
                    finalUrl = (resolvedUrl || 'https://api.openai.com/v1') + '/chat/completions';
                }
            }

            // 2. 构造 headers（按各提供商鉴权方式）
            const headers = { 'Content-Type': 'application/json' };
            const headerName = pc.headerName || 'Authorization';
            const headerPrefix = pc.headerPrefix != null ? pc.headerPrefix : 'Bearer ';
            if (apiKey) {
                headers[headerName] = headerPrefix + apiKey;
            }

            // 3. 构造 body（区分 Ollama 与 OpenAI 兼容格式）
            // kimi-k3 / deepseek-v4 等模型只允许 temperature=1（与 llm-client.js buildBody 决策一致），
            // 否则 API 返回 400 invalid temperature
            const testModel = model || tempConfig.model;
            const testTemperature = /^(kimi-k[2-9]|deepseek-v4)\b/i.test(testModel) ? 1 : 0.1;

            let body;
            if (provider === 'ollama') {
                body = JSON.stringify({
                    model: testModel,
                    stream: false,
                    messages: testMessages,
                    options: { temperature: testTemperature, num_predict: 50 }
                });
            } else {
                body = JSON.stringify({
                    model: testModel,
                    stream: false,
                    messages: testMessages,
                    temperature: testTemperature,
                    max_tokens: 50
                });
            }

            // 4. 发起请求（带超时）
            const response = await fetch(finalUrl, {
                method: 'POST',
                headers,
                body,
                signal: controller.signal
            });

            // 清除超时定时器
            clearTimeout(timeoutId);

            // 恢复原始配置
            Object.assign(config, originalConfig);

            if (response.ok) {
                const data = await response.json();
                let content = '';

                // 解析响应
                if (provider === 'ollama') {
                    content = data.message?.content || data.response || '';
                } else if (provider === 'qianwen') {
                    content = data.output?.text || data.output?.choices?.[0]?.message?.content
                        || data.choices?.[0]?.message?.content || '';
                } else if (provider === 'wenxin') {
                    content = data.result || data.choices?.[0]?.message?.content || '';
                } else {
                    content = data.choices?.[0]?.message?.content || '';
                }

                resultDiv.style.background = '#d4edda';
                resultDiv.style.color = '#155724';
                resultDiv.style.border = '1px solid #c3e6cb';
                resultDiv.innerHTML = `✅ 连接成功！<br>模型响应: ${content.substring(0, 100)}${content.length > 100 ? '...' : ''}`;
            } else {
                const errorText = await response.text();
                resultDiv.style.background = '#f8d7da';
                resultDiv.style.color = '#721c24';
                resultDiv.style.border = '1px solid #f5c6cb';
                resultDiv.innerHTML = `❌ 连接失败<br>状态码: ${response.status}<br>错误: ${errorText.substring(0, 200)}`;
            }
        } catch (error) {
            // 清除超时定时器（如果还没清除）
            clearTimeout(timeoutId);

            // 恢复原始配置
            Object.assign(config, originalConfig);

            resultDiv.style.background = '#f8d7da';
            resultDiv.style.color = '#721c24';
            resultDiv.style.border = '1px solid #f5c6cb';

            // 特殊处理超时错误
            if (error.name === 'AbortError') {
                resultDiv.innerHTML = `⏱️ 连接超时<br>无法在 30 秒内连接到 API 服务器<br><br>可能原因：<br>• 网络连接问题<br>• API 地址无法访问（国内可能需要代理）<br>• 服务器响应缓慢<br>• Ollama 升级后首次加载模型需要较长时间`;
            } else {
                resultDiv.innerHTML = `❌ 测试失败<br>错误: ${error.message}`;
            }
        }
    }

    /**
     * 切换设置标签页
     * @param {string} tab - 标签页名称 ('ai' 或 'ocr')
     */
    // --- 来自 app.js:L3420-L3449 ---
    function switchSettingsTab(tabId) {
        const tabAi = document.getElementById('tab-ai');
        const tabOcr = document.getElementById('tab-ocr');
        const tabMobile = document.getElementById('tab-mobile');

        [tabAi, tabOcr, tabMobile].forEach(t => {
            if (t) {
                t.classList.remove('active');
                t.style.borderBottomColor = 'transparent';
                t.style.color = '#888';
                t.style.fontWeight = '400';
            }
        });

        document.getElementById('settings-tab-ai').style.display = 'none';
        document.getElementById('settings-tab-ocr').style.display = 'none';
        const tabMobilePanel = document.getElementById('settings-tab-mobile');
        if (tabMobilePanel) tabMobilePanel.style.display = 'none';

        if (tabId === 'ai') {
            tabAi.classList.add('active');
            tabAi.style.borderBottomColor = '#1976d2';
            tabAi.style.color = '#1976d2';
            tabAi.style.fontWeight = '600';
            document.getElementById('settings-tab-ai').style.display = 'block';
        } else if (tabId === 'ocr') {
            tabOcr.classList.add('active');
            tabOcr.style.borderBottomColor = '#9c27b0';
            tabOcr.style.color = '#9c27b0';
            tabOcr.style.fontWeight = '600';
            document.getElementById('settings-tab-ocr').style.display = 'block';
        } else if (tabId === 'mobile') {
            if (tabMobile) {
                tabMobile.classList.add('active');
                tabMobile.style.borderBottomColor = '#4caf50';
                tabMobile.style.color = '#4caf50';
                tabMobile.style.fontWeight = '600';
            }
            if (tabMobilePanel) tabMobilePanel.style.display = 'block';
            // 初始化手机遥控模块
            if (typeof window.initMobileControl === 'function') {
                setTimeout(function() { window.initMobileControl(); }, 100);
            }
        }
    }

    // 扩展页签切换（Skills / MCP）
    // --- 来自 app.js:L3452-L3473 ---
    function switchExtensionTab(tabId) {
        const tabSkills = document.getElementById('ext-tab-skills');
        const tabMcp = document.getElementById('ext-tab-mcp');
        const panelSkills = document.getElementById('ext-panel-skills');
        const panelMcp = document.getElementById('ext-panel-mcp');

        [tabSkills, tabMcp].forEach(t => {
            if (t) { t.style.borderBottomColor = 'transparent'; t.style.color = '#666'; }
        });

        // 【修复】is-hidden 类带 !important，仅设置 style.display 无法显示面板
        // 必须同时通过 classList 切换 is-hidden 类才能正确控制显隐
        if (panelSkills) { panelSkills.style.display = 'none'; panelSkills.classList.add('is-hidden'); }
        if (panelMcp) { panelMcp.style.display = 'none'; panelMcp.classList.add('is-hidden'); }

        if (tabId === 'skills') {
            if (tabSkills) { tabSkills.style.borderBottomColor = '#ff9800'; tabSkills.style.color = '#ff9800'; }
            if (panelSkills) { panelSkills.classList.remove('is-hidden'); panelSkills.style.display = 'block'; }
            
            // 【修复】每次切换到Skills标签都重新渲染（带重试机制）
            // 数据源：管理面板必须用完整列表（GetAllSkills，含停用项 + isEnabled），
            // 否则停用的 Skill 会被 GetAvailableSkills 过滤，导致开关状态回弹/消失。
            function tryRenderSkills(maxRetries = 10) {
                const container = document.getElementById('skill-list-container');
                if (!container) return;

                // 优先走管理面板专用加载器（内部已含 GetAllSkills → cachedSkills → __sharedSkillsData 回退链）
                if (typeof App.loadSkillsForManagement === 'function') {
                    ChatLog.info('MODEL', '[switchExtensionTab] 加载 Skill 管理完整列表...');
                    App.loadSkillsForManagement();
                    return;
                }

                const skills = (App.cachedSkills && App.cachedSkills.length > 0)
                    ? App.cachedSkills
                    : (window.__sharedSkillsData && window.__sharedSkillsData.length > 0)
                        ? window.__sharedSkillsData
                        : null;

                if (skills && typeof renderSkillListDirect === 'function') {
                    ChatLog.info('MODEL', '[switchExtensionTab] 重新渲染Skill列表...', skills.length, '个');
                    renderSkillListDirect(container, skills);
                    return;
                }

                if (maxRetries > 0) {
                    setTimeout(() => tryRenderSkills(maxRetries - 1), 300);
                }
            }
            tryRenderSkills();
        } else if (tabId === 'mcp') {
            if (tabMcp) { tabMcp.style.borderBottomColor = '#7c4dff'; tabMcp.style.color = '#7c4dff'; }
            if (panelMcp) { panelMcp.classList.remove('is-hidden'); panelMcp.style.display = 'block'; }

            // 【修复】app-mcp.js 是 defer 加载，且 initBridge 可能在后台并行执行。
            // 切换 MCP 标签时，先等 bridge 和 MCP 函数都就绪，再顺序刷新状态/工具。
            async function waitForBridgeAndMcp(maxRetries) {
                while (maxRetries > 0) {
                    const hasBridge = !!(App.excelBridge || (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge));
                    const hasRefresh = typeof window.refreshMcpStatus === 'function' && String(window.refreshMcpStatus).indexOf('模块加载中') === -1;
                    const hasLoadTools = typeof window.loadMcpToolsForAI === 'function' && String(window.loadMcpToolsForAI).indexOf('模块加载中') === -1;
                    if (hasBridge && hasRefresh) {
                        try {
                            await window.refreshMcpStatus();
                            if (hasLoadTools && (!window._mcpTools || window._mcpTools.length === 0)) {
                                await window.loadMcpToolsForAI();
                            }
                        } catch (e) {
                            ChatLog.warn('MODEL', '[switchExtensionTab] MCP 刷新失败:', e);
                        }
                        return;
                    }
                    maxRetries--;
                    if (maxRetries > 0) await new Promise(r => setTimeout(r, 150));
                }
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function'
                    ? App.fixedStatusBar.setExtraStatus(' <span style="color:#f44336;font-size:11px;margin-left:8px;">MCP 模块加载失败，请刷新页面重试</span>', true)
                    : (App.showToast && App.showToast('MCP 模块加载失败，请刷新页面重试', 'error'));
            }
            waitForBridgeAndMcp(40); // 最多等待约 6 秒
        }
    }

    // ============================================
    // 识别类型管理功能
    // ============================================

    /**
     * 从后端重新加载识别类型配置
     */
    // --- 来自 app.js:L3482-L3505 ---
    async function reloadCardTypesConfigFromBackend() {
        const btn = event.target;
        const originalText = btn.innerHTML;
        btn.innerHTML = '🔄 加载中';
        btn.disabled = true;

        try {
            const success = await reloadCardTypesConfig();

            if (success) {
                // 重新渲染列表
                renderRecognitionTypesList();
                // 进度条已显示,无需 toast
            } else {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">️ 后端加载失败，使用默认配置</span>', true) : showToast('️ 后端加载失败，使用默认配置', 'warning');
            }
        } catch (error) {
            // 使用统一错误处理
            window.ErrorHandler.showError(error, '配置加载');
        } finally {
            btn.innerHTML = originalText;
            btn.disabled = false;
        }
    }

    // ==================== 识别类型管理（已废弃，功能合并到 Skill）====================
    // 以下函数保留用于向后兼容，但不再在 UI 中显示
    // 识别类型配置现在通过 DefaultSkills.json 中的 Skill 定义管理

    // --- 来自 app.js:L3511-L3513 ---
    function renderRecognitionTypesList() {
        ChatLog.info('MODEL', '[Deprecated] renderRecognitionTypesList 已废弃，识别类型已合并到 Skill 管理');
    }

    // --- 来自 app.js:L3515-L3518 ---
    function showAddRecognitionTypeDialog() {
        ChatLog.info('MODEL', '[Deprecated] 识别类型已合并到 Skill，请使用 Skill 管理');
        App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">识别类型已合并到 Skill 管理，请直接添加 Skill</span>', true) : showToast('识别类型已合并到 Skill 管理，请直接添加 Skill', 'info');
    }

    // --- 来自 app.js:L3520-L3522 ---
    function editRecognitionType(typeName) {
        ChatLog.info('MODEL', '[Deprecated] 识别类型已合并到 Skill');
    }

    // --- 来自 app.js:L3524-L3526 ---
    function saveRecognitionType(originalTypeName) {
        ChatLog.info('MODEL', '[Deprecated] 识别类型已合并到 Skill');
    }

    // --- 来自 app.js:L3528-L3530 ---
    function deleteRecognitionType(typeName) {
        ChatLog.info('MODEL', '[Deprecated] 识别类型已合并到 Skill');
    }

    // --- 来自 app.js:L3532-L3534 ---
    function saveCustomTypesToStorage(types) {
        // 不再保存识别类型到 localStorage，统一使用 Skill 管理
    }

    // --- 来自 app.js:L3536-L3538 ---
    function getDefaultBuiltInConfig(typeName) {
        return {};
    }

    // --- 来自 app.js:L3540-L3542 ---
    function getModifiedFields(original, current) {
        return {};
    }

    // --- 来自 app.js:L3544-L3547 ---
    function loadCustomTypesFromStorage() {
        // 识别类型已合并到 Skill，不再从 localStorage 加载
        ChatLog.info('MODEL', '[Deprecated] 识别类型已合并到 Skill 管理');
    }

    // --- 来自 app.js:L3549-L3579 ---
    function saveSettings() {
        // 【修复】所有getElementById可能返回null
        const apiTypeEl = document.getElementById('api-type');
        const apiUrlEl = document.getElementById('api-url');
        const apiKeyEl = document.getElementById('api-key');
        const modelEl = document.getElementById('model');
        const tempEl = document.getElementById('temperature');
        const contextEl = document.getElementById('context-mode');
        
        if (!apiTypeEl || !apiUrlEl || !apiKeyEl || !modelEl) {
            ChatLog.warn('MODEL', '[saveSettings] 设置面板元素不存在');
            return;
        }
        
        const apiType = apiTypeEl.value;
        const apiUrl = apiUrlEl.value.trim();
        const apiKey = apiKeyEl.value.trim();
        const model = modelEl.value.trim();

        if (!config.providers) {
            config.providers = {};
        }

        config.providers[apiType] = {
            apiKey: apiKey,
            apiUrl: apiUrl.trim(),
            model: model
        };

        config.apiType = apiType;
        config.apiUrl = apiUrl.trim();
        config.apiKey = apiKey;
        config.model = model;
        config.temperature = tempEl ? parseFloat(tempEl.value) : 0.7;
        config.contextMode = contextEl ? contextEl.value : 'auto';
        var dangerAskEl = document.getElementById('danger-always-ask');
        if (dangerAskEl) config.dangerAlwaysAsk = !!dangerAskEl.checked;
        var localeEl = document.getElementById('ui-locale');
        if (localeEl) {
            config.uiLocale = normalizeUiLocale(localeEl.value);
            try {
                if (window.App && App.i18n && typeof App.i18n.setLocale === 'function') {
                    App.i18n.setLocale(config.uiLocale, { apply: true, reloadWelcome: true });
                }
                if (window.parent && window.parent !== window && window.parent.App && window.parent.App.i18n) {
                    window.parent.App.i18n.setLocale(config.uiLocale, { apply: true, reloadWelcome: true });
                }
            } catch (_) { /* ignore */ }
        }

        localStorage.setItem(CONFIG_KEY, JSON.stringify(config));
        var savedMsg = '设置已保存';
        try {
            if (window.App && App.i18n && typeof App.i18n.t === 'function') savedMsg = App.i18n.t('settings.saved');
        } catch (_) { /* ignore */ }
        App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">' + savedMsg + '</span>', true) : showToast(savedMsg);

        // 同步配置到后端
        syncConfigToBackend();

        updateContextStats();
        try {
            if (typeof window.refreshWelcomeSetupCta === 'function') window.refreshWelcomeSetupCta();
            if (window.opener && typeof window.opener.refreshWelcomeSetupCta === 'function') {
                window.opener.refreshWelcomeSetupCta();
            }
            if (typeof window.notifyParentSettingsSaved === 'function') {
                window.notifyParentSettingsSaved();
            } else if (window.parent && window.parent !== window) {
                window.parent.postMessage({ type: 'smartpilot:ai-settings-saved' }, '*');
            }
        } catch (_) { /* ignore */ }
    }

    // --- 来自 app.js:L4299-L4315 ---
    // 更新上下文统计信息显示（在设置面板中展示当前对话的 token 使用情况）
    function updateContextStats() {
        if (typeof App.ContextManager === 'undefined' || !App.ContextManager) {
            ChatLog.info('MODEL', '[Context] ContextManager 尚未初始化，跳过统计更新');
            return;
        }
        var stats = App.ContextManager.getContextStats();
        var statsDiv = document.getElementById('context-stats');
        if (statsDiv) {
            statsDiv.innerHTML =
                '当前对话: ' + stats.messageCount + ' 条消息 | ' +
                '预估 Token: ' + stats.totalTokens.toLocaleString() + ' / ' + stats.limit.toLocaleString() + ' | ' +
                '使用率: ' + stats.usage + '%';
        }
    }

    /**
     * 同步前端配置到后端
     */
    // --- 来自 app.js:L3584-L3607 ---
    async function syncConfigToBackend() {
        try {
            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                const configToSync = {
                    apiType: config.apiType,
                    apiUrl: config.apiUrl,
                    model: config.model,
                    apiKey: config.apiKey
                };

                ChatLog.info('MODEL', '正在同步配置到后端...');
                // 与 loadConfig 相同：走新鲜 hostObjects 调用（BridgeAccessor.callAsync 亦然）
                let result;
                if (window.BridgeAccessor && typeof BridgeAccessor.callAsync === 'function') {
                    result = await BridgeAccessor.callAsync('SaveAIConfig', JSON.stringify(configToSync));
                } else {
                    result = await window.chrome.webview.hostObjects.excelBridge.SaveAIConfig(JSON.stringify(configToSync));
                }
                const resultObj = typeof result === 'string' ? JSON.parse(result) : (result || {});

                if (resultObj.success) {
                    ChatLog.info('MODEL', '配置已同步到后端');
                } else if (resultObj.error === 'bridge-call-failed') {
                    ChatLog.warn('MODEL', '配置同步跳过（桥方法暂不可用）:', resultObj.detail || resultObj.method);
                } else {
                    ChatLog.warn('MODEL', '配置同步失败:', resultObj.error);
                }
            }
        } catch (error) {
            const msg = (error && error.message) ? error.message : String(error);
            if (!/e is not a function|0x80020006|未知名称/i.test(msg)) {
                ChatLog.warn('MODEL', '同步配置到后端失败:', msg);
            }
        }
    }

    // 导出模块
    window.App = window.App || {};

    // Provider/模型管理相关
    App.isProviderAvailable = isProviderAvailable;
    App.getControlFlowHint = getControlFlowHint;
    App.getAvailableProviders = getAvailableProviders;
    App.updateChatModelSelector = updateChatModelSelector;
    App.autoSelectModel = autoSelectModel;
    App.isAutoMode = isAutoMode;
    App.getEffectiveModel = getEffectiveModel;
    App.getProviderName = getProviderName;
    App.refreshModelsList = refreshModelsList;
    App.onChatModelChange = onChatModelChange;
    App.onThinkingModeChange = onThinkingModeChange;
    // 【暴露到全局】使 HTML 的 onchange 可以直接调用
    window.onThinkingModeChange = onThinkingModeChange;

    // 常量
    App.AUTO_MODEL_CONFIG = AUTO_MODEL_CONFIG;
    App.ENV = ENV;
    App.MOCK_BRIDGE = MOCK_BRIDGE;
    App.MOCK_BRIDGE_PROXY = MOCK_BRIDGE_PROXY;

    // Bridge通信相关
    App.initBridge = initBridge;
    App.updateStatus = updateStatus;
    App.getExcelHeaders = getExcelHeaders;
    App.buildContextAwareMessage = buildContextAwareMessage;
    App.isDataFillRequest = isDataFillRequest;

    // UI状态管理
    App.saveUIState = saveUIState;
    App.restoreUIState = restoreUIState;
    App.initTabs = initTabs;
    App.initSkillsSubTabs = initSkillsSubTabs;
    App.collapseEditor = collapseEditor;

    // 提供商切换
    App.onProviderChange = onProviderChange;
    App.selectModel = selectModel;

    // 本地模型检测
    App.detectLocalModels = detectLocalModels;
    App.detectOllamaModels = detectOllamaModels;

    // 配置管理
    App.loadConfig = loadConfig;
    App.testAIConfig = testAIConfig;
    App.saveSettings = saveSettings;
    App.syncConfigToBackend = syncConfigToBackend;
    App.updateContextStats = updateContextStats;

    // 设置面板标签切换
    App.switchSettingsTab = switchSettingsTab;
    App.switchExtensionTab = switchExtensionTab;
    // 【修复】index.html 中 onclick="switchExtensionTab(...)" 直接调用全局函数
    window.switchExtensionTab = switchExtensionTab;

    // 识别类型管理（已废弃）
    App.reloadCardTypesConfigFromBackend = reloadCardTypesConfigFromBackend;
    App.renderRecognitionTypesList = renderRecognitionTypesList;
    App.showAddRecognitionTypeDialog = showAddRecognitionTypeDialog;
    App.editRecognitionType = editRecognitionType;
    App.saveRecognitionType = saveRecognitionType;
    App.deleteRecognitionType = deleteRecognitionType;
    App.saveCustomTypesToStorage = saveCustomTypesToStorage;
    App.getDefaultBuiltInConfig = getDefaultBuiltInConfig;
    App.getModifiedFields = getModifiedFields;
    App.loadCustomTypesFromStorage = loadCustomTypesFromStorage;

    // 模型能力检测工具函数
    App.isVisionModel = isVisionModel;
    App.isReasoningModel = isReasoningModel;

    /**
     * 是否已具备可发起对话的 LLM 配置（云端有 Key，或本地已选模型）
     * 用于发送前软拦截与欢迎页引导条。
     */
    function isLlmConfigured() {
        try {
            const cfg = (typeof App !== 'undefined' && App.config) || (typeof config !== 'undefined' && config) || {};
            const apiType = String(cfg.apiType || '').trim();
            const model = String(cfg.model || '').trim();
            const providersMap = (typeof PROVIDER_CONFIGS !== 'undefined' && PROVIDER_CONFIGS)
                || (App && App.PROVIDER_CONFIGS) || {};
            const pc = providersMap[apiType] || {};
            const isLocal = apiType === 'ollama' || apiType === 'llamacpp' || !!pc.isLocal;

            if (isLocal) {
                // 本地：必须已选具体模型，避免空 model 打出去只换来连接错误
                return !!model;
            }

            let key = String(cfg.apiKey || '').trim();
            if ((!key || key === 'local') && cfg.providers && apiType && cfg.providers[apiType]) {
                key = String(cfg.providers[apiType].apiKey || '').trim();
            }
            if (!key || key === 'local') return false;
            return true;
        } catch (e) {
            ChatLog.warn?.('CONFIG', 'isLlmConfigured 检查失败:', e && e.message);
            return false; // 检查异常时 fail-closed，引导用户打开设置
        }
    }

    /**
     * 发送前检查：未配置则 toast + 刷新欢迎引导，并尽量打开设置。
     * @returns {boolean} true=可以发送
     */
    function ensureLlmConfiguredForSend() {
        if (isLlmConfigured()) return true;
        let msg = '请先配置 AI 模型：打开设置填写 API 密钥，或启动 Ollama 并选择模型。';
        try {
            if (App && App.i18n && typeof App.i18n.t === 'function') msg = App.i18n.t('gate.needLlm');
        } catch (_) { /* ignore */ }
        try {
            if (typeof App.showToast === 'function') App.showToast(msg, 'warning');
            else if (typeof window.showToast === 'function') window.showToast(msg, 'warning');
        } catch (_) { /* ignore */ }
        try {
            if (typeof window.refreshWelcomeSetupCta === 'function') window.refreshWelcomeSetupCta();
            else if (typeof window.showWelcomePanel === 'function') window.showWelcomePanel();
        } catch (_) { /* ignore */ }
        try {
            // 欢迎页已有「打开 AI 设置」按钮时不强制弹窗；仅在无欢迎页时打开设置
            const hasWelcome = !!document.querySelector('.welcome-panel');
            if (!hasWelcome) {
                setTimeout(function() {
                    if (typeof window.openAiSettings === 'function') window.openAiSettings();
                }, 80);
            }
        } catch (_) { /* ignore */ }
        return false;
    }

    App.isLlmConfigured = isLlmConfigured;
    App.ensureLlmConfiguredForSend = ensureLlmConfiguredForSend;

})();