/**
 * ============================================================
 * skill-registry.js — Declarative Skill Registry
 * ============================================================
 * 声明式 Skill 注册中心。Skill 通过 manifest 描述所需上下文、输出类型、
 * 系统提示词与关键词，注册中心负责校验、查询与意图匹配。
 *
 * 核心类：
 *   - SkillRegistry
 *     - register(skillManifest)   注册并校验 manifest
 *     - get(skillName)            按名称获取 skill
 *     - match(userMessage, ctx)   根据用户消息与上下文返回候选 skill 列表
 *     - loadManifests()           动态加载本地 manifest（含降级兜底）
 *     - getAll()                  获取所有已注册 skill
 *
 * 依赖：
 *   - App 命名空间（由 app-bootstrap.js 提供）
 * ============================================================
 */

(function() {
    'use strict';

    const ChatLog = (typeof window !== 'undefined' && window.ChatLog) ? window.ChatLog : console;

    /**
     * Skill manifest 必填字段
     */
    const REQUIRED_FIELDS = ['name', 'description', 'inputContext', 'outputType', 'systemPrompt'];

    /**
     * 本地 manifest 文件名列表
     */
    const MANIFEST_NAMES = [
        'excel-compare-regions',
        'excel-split-sheets',
        'general-vba',
        'general-python'
    ];

    /**
     * 内置示例 manifest（当 fetch 本地 JSON 失败时的兜底）
     */
    const BUILTIN_MANIFESTS = [
        {
            "name": "excel-compare-regions",
            "displayName": "比较多个引用区域",
            "description": "比较用户添加的多个 Excel 引用区域，生成差异分析或合并报告。",
            "version": 1,
            "inputContext": ["referenceRegions"],
            "outputType": "python_script",
            "pipeline": "python",
            "keywords": ["比较", "对比", "差异", "多个区域", "引用区域"],
            "examples": ["比较这两个区域的数据差异", "对比 A1:D10 和 F1:I10"],
            "systemPrompt": "你是 SmartPilot 的 Python 数据分析专家。用户要求比较多个 Excel 引用区域。\n\n【强制规则】\n1. 必须使用名为 `reference_data` 的变量访问引用区域数据。`reference_data` 是一个列表，每个元素是一个字典，包含 `headers` 和 `rows`。\n2. 不要尝试通过文件路径读取 Excel 文件；所有数据都已通过 `reference_data` 注入。\n3. 分析区域之间的差异、重复项、唯一项或汇总指标，并输出清晰的结论。\n4. 输出应为可直接执行的 Python 脚本。",
            "tags": ["excel", "compare", "python"],
            "ui": { "showCopy": true, "showRegenerate": true }
        },
        {
            "name": "excel-split-sheets",
            "displayName": "拆分工作表",
            "description": "根据当前选区或工作表内容，将 Excel 工作簿拆分为多个独立文件。",
            "version": 1,
            "inputContext": ["selection"],
            "outputType": "vba_code",
            "keywords": ["拆分", "分割", "工作表", "拆成多个文件", "按表拆分"],
            "examples": ["把工作表按月份拆分成多个文件", "拆分当前工作簿"],
            "systemPrompt": "你是 SmartPilot 的 VBA 专家。用户要求拆分 Excel 工作表。\n\n【强制规则】\n1. 生成完整、可直接运行的 VBA 宏（Sub/End Sub）。\n2. 基于用户提供的当前选区或上下文信息推断拆分规则。\n3. 代码需包含错误处理和进度提示。\n4. 不要输出 Markdown 代码块标记，只输出纯 VBA 代码。",
            "tags": ["excel", "split", "vba"],
            "ui": { "showCopy": true, "showRegenerate": true }
        },
        {
            "name": "general-vba",
            "displayName": "通用 VBA 代码生成",
            "description": "根据用户意图和当前上下文生成通用 VBA 宏代码。",
            "version": 1,
            "inputContext": [],
            "outputType": "vba_code",
            "pipeline": "code",
            "keywords": ["VBA", "宏", "代码", "Excel 宏", "写个宏"],
            "examples": ["帮我写一个批量替换的宏", "生成一个格式化选中区域的 VBA"],
            "systemPrompt": "你是 SmartPilot 的通用 VBA 代码生成助手。\n\n【强制规则】\n1. 根据用户意图和提供的上下文（当前选区、引用区域）生成完整 VBA 宏。\n2. 代码必须是 Sub/End Sub 结构，可直接运行。\n3. 若上下文为空，应基于用户描述生成通用代码。\n4. 引用工作表时必须使用名称（如 Sheets(\"Sheet1\")），禁止使用索引（如 Sheets(1)）。\n5. 不要输出 Markdown 代码块标记，只输出纯 VBA 代码。",
            "tags": ["excel", "vba", "general"],
            "ui": { "showCopy": true, "showRegenerate": true }
        },
        {
            "name": "general-python",
            "displayName": "通用 Python 代码生成",
            "description": "根据用户意图和当前上下文生成通用 Python 分析脚本。",
            "version": 1,
            "inputContext": [],
            "outputType": "python_script",
            "pipeline": "python",
            "keywords": ["Python", "脚本", "分析", "pandas", "写个脚本"],
            "examples": ["用 Python 分析选中数据", "生成一个统计脚本"],
            "systemPrompt": "你是 SmartPilot 的通用 Python 代码生成助手。\n\n【强制规则】\n1. 根据用户意图和提供的上下文（当前选区、引用区域）生成可直接执行的 Python 脚本。\n2. 当存在引用区域时，优先使用名为 `reference_data` 的变量访问数据；当只有当前选区时，使用 `selection_data`。\n3. 不要尝试通过相对路径读取 Excel 文件。\n4. 输出应为可直接执行的 Python 代码。",
            "tags": ["excel", "python", "general"],
            "ui": { "showCopy": true, "showRegenerate": true }
        }
    ];

    /**
     * 声明式 Skill 注册中心
     */
    class SkillRegistry {
        constructor() {
            this.skills = new Map();
        }

        /**
         * 注册一个 skill manifest
         * @param {Object} manifest
         * @returns {Object} 注册后的 manifest
         */
        register(manifest) {
            if (!manifest || typeof manifest !== 'object') {
                throw new Error('Skill manifest 必须是一个对象');
            }

            for (let i = 0; i < REQUIRED_FIELDS.length; i++) {
                const field = REQUIRED_FIELDS[i];
                if (manifest[field] === undefined || manifest[field] === null || manifest[field] === '') {
                    throw new Error('Skill manifest 缺少必填字段: ' + field);
                }
            }

            if (!Array.isArray(manifest.inputContext)) {
                throw new Error('Skill manifest 的 inputContext 必须是数组');
            }

            if (typeof manifest.name !== 'string' || manifest.name.trim().length === 0) {
                throw new Error('Skill manifest 的 name 必须是非空字符串');
            }

            const normalized = this._normalize(manifest);
            this.skills.set(normalized.name, normalized);

            if (ChatLog && typeof ChatLog.info === 'function') {
                ChatLog.info('SkillRegistry', '已注册 skill:', normalized.name);
            }

            return normalized;
        }

        /**
         * 标准化 manifest
         * @param {Object} manifest
         * @returns {Object}
         */
        _normalize(manifest) {
            return {
                name: String(manifest.name).trim(),
                displayName: (manifest.displayName || manifest.name),
                description: String(manifest.description || ''),
                version: typeof manifest.version === 'number' ? manifest.version : 1,
                inputContext: Array.isArray(manifest.inputContext) ? manifest.inputContext.slice() : [],
                outputType: String(manifest.outputType || ''),
                systemPrompt: String(manifest.systemPrompt || ''),
                keywords: Array.isArray(manifest.keywords) ? manifest.keywords.slice() : [],
                examples: Array.isArray(manifest.examples) ? manifest.examples.slice() : [],
                tags: Array.isArray(manifest.tags) ? manifest.tags.slice() : [],
                ui: manifest.ui && typeof manifest.ui === 'object' ? manifest.ui : {}
            };
        }

        /**
         * 按名称获取 skill
         * @param {string} skillName
         * @returns {Object|null}
         */
        get(skillName) {
            if (typeof skillName !== 'string') return null;
            return this.skills.get(skillName.trim()) || null;
        }

        /**
         * 获取所有已注册 skill
         * @returns {Object[]}
         */
        getAll() {
            const result = [];
            this.skills.forEach(function(skill) {
                result.push(skill);
            });
            return result;
        }

        /**
         * 根据用户消息与上下文匹配候选 skill
         * @param {string} userMessage
         * @param {Object} context
         * @returns {Object[]} 按相关度排序的候选列表（包含 skill 与 score）
         */
        match(userMessage, context) {
            const message = (userMessage || '').toLowerCase();
            context = context || {};

            const candidates = [];
            const allSkills = this.getAll();

            for (let i = 0; i < allSkills.length; i++) {
                const skill = allSkills[i];
                const keywordScore = this._calcKeywordScore(message, skill.keywords);
                const contextScore = this._calcContextScore(context, skill.inputContext);
                const outputScore = this._calcOutputScore(context, skill.outputType);

                // 综合得分：关键词命中权重最高，上下文可用性次之，输出类型偏好再次
                const score = keywordScore * 3 + contextScore * 1.5 + outputScore * 0.5;

                // 只要关键词命中或上下文部分可用即纳入候选
                if (keywordScore > 0 || contextScore > 0) {
                    candidates.push({
                        skill: skill,
                        score: score,
                        keywordScore: keywordScore,
                        contextScore: contextScore,
                        outputScore: outputScore
                    });
                }
            }

            // 按综合得分降序排序
            candidates.sort(function(a, b) {
                return b.score - a.score;
            });

            return candidates;
        }

        /**
         * 计算关键词得分
         * @param {string} message
         * @param {string[]} keywords
         * @returns {number}
         */
        _calcKeywordScore(message, keywords) {
            if (!message || !Array.isArray(keywords) || keywords.length === 0) return 0;

            let score = 0;
            for (let i = 0; i < keywords.length; i++) {
                const kw = String(keywords[i] || '').toLowerCase().trim();
                if (!kw) continue;

                let pos = message.indexOf(kw);
                while (pos !== -1) {
                    score += 1;
                    pos = message.indexOf(kw, pos + kw.length);
                }
            }

            return score;
        }

        /**
         * 计算上下文可用性得分
         * @param {Object} context
         * @param {string[]} requiredInputs
         * @returns {number}
         */
        _calcContextScore(context, requiredInputs) {
            if (!Array.isArray(requiredInputs) || requiredInputs.length === 0) return 0;

            let available = 0;
            for (let i = 0; i < requiredInputs.length; i++) {
                const name = requiredInputs[i];
                const value = context[name];

                if (name === 'referenceRegions') {
                    if (Array.isArray(value) && value.length > 0) available += 1;
                } else if (name === 'selection') {
                    if (value && typeof value === 'object' && !value.error) available += 1;
                } else if (name === 'attachments') {
                    if (Array.isArray(value) && value.length > 0) available += 1;
                } else if (name === 'history') {
                    if (Array.isArray(value) && value.length > 0) available += 1;
                } else if (value !== undefined && value !== null) {
                    available += 1;
                }
            }

            return available / requiredInputs.length;
        }

        /**
         * 计算输出类型偏好得分
         * @param {Object} context
         * @param {string} outputType
         * @returns {number}
         */
        _calcOutputScore(context, outputType) {
            if (!context.outputType) return 0;
            return String(context.outputType).toLowerCase() === String(outputType).toLowerCase() ? 1 : 0;
        }

        /**
         * 动态加载本地 manifest 并注册。
         * 若浏览器 fetch 本地 JSON 失败，则回退到内置 manifest。
         * @returns {Promise<number>} 成功注册的 manifest 数量
         */
        async loadManifests() {
            const manifests = await this._fetchManifests();
            let registeredCount = 0;

            for (let i = 0; i < manifests.length; i++) {
                try {
                    this.register(manifests[i]);
                    registeredCount += 1;
                } catch (err) {
                    if (ChatLog && typeof ChatLog.warn === 'function') {
                        ChatLog.warn('SkillRegistry', '注册 manifest 失败:', err && err.message ? err.message : err);
                    }
                }
            }

            if (ChatLog && typeof ChatLog.info === 'function') {
                ChatLog.info('SkillRegistry', 'loadManifests 完成，共注册 ' + registeredCount + ' 个 skill');
            }

            return registeredCount;
        }

        /**
         * 尝试 fetch 本地 manifest；失败时返回内置清单
         * @returns {Promise<Object[]>}
         */
        async _fetchManifests() {
            if (typeof fetch !== 'function') {
                if (ChatLog && typeof ChatLog.info === 'function') {
                    ChatLog.info('SkillRegistry', '当前环境不支持 fetch，使用内置 manifest');
                }
                return BUILTIN_MANIFESTS.slice();
            }

            const builtinByName = {};
            for (let i = 0; i < BUILTIN_MANIFESTS.length; i++) {
                const manifest = BUILTIN_MANIFESTS[i];
                builtinByName[manifest.name] = manifest;
            }

            try {
                const results = [];
                for (let i = 0; i < MANIFEST_NAMES.length; i++) {
                    const name = MANIFEST_NAMES[i];
                    const fallback = builtinByName[name];
                    try {
                        const response = await fetch('config/skills/' + name + '.json');
                        if (response && response.ok) {
                            const data = await response.json();
                            results.push(data);
                        } else if (fallback) {
                            results.push(fallback);
                        }
                    } catch (e) {
                        if (fallback) {
                            results.push(fallback);
                        }
                    }
                }
                return results;
            } catch (e) {
                if (ChatLog && typeof ChatLog.warn === 'function') {
                    ChatLog.warn('SkillRegistry', 'fetch manifest 失败，使用内置清单:', e && e.message ? e.message : e);
                }
                return BUILTIN_MANIFESTS.slice();
            }
        }
    }

    // ==================== 挂载 ====================
    const root = (typeof globalThis !== 'undefined') ? globalThis :
                 (typeof window !== 'undefined') ? window :
                 (typeof global !== 'undefined') ? global : {};

    if (typeof App !== 'undefined') {
        App.SkillRegistry = new SkillRegistry();
        App.SkillRegistry.loadManifests().catch(function(err) {
            if (ChatLog && typeof ChatLog.warn === 'function') {
                ChatLog.warn('SkillRegistry', '初始加载 manifest 失败:', err && err.message ? err.message : err);
            }
        });
        if (ChatLog && typeof ChatLog.info === 'function') {
            ChatLog.info('SkillRegistry', 'Skill 注册中心已挂载到 App.SkillRegistry');
        }
    } else if (typeof module !== 'undefined' && module.exports) {
        module.exports = SkillRegistry;
    } else {
        root.SkillRegistry = SkillRegistry;
    }
})();
