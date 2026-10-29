/**
 * ============================================================
 * capability-profile.js — Model Capability Profile
 * ============================================================
 * 功能：根据当前模型能力自动选择执行策略，避免在调用点硬编码模型名称。
 *
 * 核心内容：
 *   - CAPABILITY_PROFILES — 预定义能力档案常量
 *   - getProfile(modelId) — 根据 modelId 返回对应 profile
 *   - CapabilityProfile — 提供 getStrategy / supportsToolUse / get 方法
 *
 * 依赖：
 *   - App 全局命名空间（由 app-bootstrap.js 提供）
 * ============================================================
 */

(function() {
    'use strict';

    const ChatLog = window.ChatLog || console;

    /**
     * 预定义能力档案
     */
    const CAPABILITY_PROFILES = Object.freeze({
        'function-calling': Object.freeze({
            name: 'function-calling',
            supportsToolUse: true,
            strategy: 'LLM_FUNCTION_CALLING',
            maxToolRounds: 5
        }),
        'chat-only': Object.freeze({
            name: 'chat-only',
            supportsToolUse: false,
            strategy: 'MCP_TOOL_CHAIN'
        }),
        'local-weak': Object.freeze({
            name: 'local-weak',
            supportsToolUse: false,
            strategy: 'DIRECT_AI_ANALYSIS',
            maxContextTokens: 4000
        })
    });

    /**
     * 返回 modelId 对应的能力档案
     * 默认档案为 function-calling：与"云端 OpenAI 兼容 API 默认下发原生 tools"的既有行为一致，
     * 避免未知模型因档案误判而丢失原生工具调用能力。
     * @param {string} modelId - 模型标识
     * @returns {Object} 能力档案副本
     */
    function getProfile(modelId) {
        const id = (modelId || '').toString().toLowerCase().trim();

        // Function-calling 模型：支持工具使用
        if (id.includes('function') ||
            id.includes('tool') ||
            id.includes('claude') ||
            id.includes('gpt-4')) {
            return _cloneProfile('function-calling');
        }

        // Qwen2.5 系列：原生 function calling 不可靠（尤其本地 7B/8B），
        // 统一按 chat-only 处理，走文本标记 [tool_name]{args} 模式；
        // 云端如需原生工具可通过 config.nativeTools = true 显式开启
        if (id.includes('qwen')) {
            return _cloneProfile('chat-only');
        }

        // 本地/轻量模型：能力有限，直接 AI 分析并限制上下文
        if (id.includes('local') || id.includes('ollama')) {
            return _cloneProfile('local-weak');
        }

        // 兜底：无法识别时按 function-calling 处理（保持云端行为不变）
        return _cloneProfile('function-calling');
    }

    /**
     * 解析当前调用是否应下发原生 tools（function calling 的单一决策点）
     * 决策优先级：
     *   1. Ollama 本地通道 → 一律 false（与 buildBody 既有决策一致：小模型原生 tool calling 不可靠，走文本标记）
     *   2. config.nativeTools 三态显式覆盖：true 强制开 / false 强制关 / undefined 自动
     *   3. config.disableNativeTools === true → false（兼容旧配置字段）
     *   4. 模型能力档案 supportsToolUse
     * @param {string} modelId - 模型标识（可空，取 config.model）
     * @param {Object} config - App.config
     * @returns {{nativeTools:boolean, profileName:string, maxToolRounds:number}}
     */
    function resolveToolMode(modelId, config) {
        const cfg = config || (typeof App !== 'undefined' ? App.config : null) || {};
        const apiType = cfg.apiType || 'ollama';
        const profile = getProfile(modelId || cfg.model);
        const maxToolRounds = profile.maxToolRounds || 5;

        // Ollama 本地通道：Ollama /api/chat 原生支持 tools（qwen2.5 等模型原生 function calling）。
        // 工具已按意图瘦身（每轮仅注入 12~30 个相关工具），原生 tool calling 可控可靠；
        // 仍保留 config.nativeTools=false / disableNativeTools=true 显式关闭通道。
        if (apiType === 'ollama') {
            if (cfg.nativeTools === false || cfg.disableNativeTools === true) {
                return { nativeTools: false, profileName: profile.name, maxToolRounds };
            }
            return { nativeTools: true, profileName: profile.name, maxToolRounds };
        }

        // 配置显式覆盖（三态）
        if (cfg.nativeTools === true) {
            return { nativeTools: true, profileName: profile.name, maxToolRounds };
        }
        if (cfg.nativeTools === false || cfg.disableNativeTools === true) {
            return { nativeTools: false, profileName: profile.name, maxToolRounds };
        }

        return { nativeTools: !!profile.supportsToolUse, profileName: profile.name, maxToolRounds };
    }

    /**
     * 复制一份档案对象，防止外部修改常量
     * @param {string} profileName - 档案名称
     * @returns {Object} 档案副本
     */
    function _cloneProfile(profileName) {
        const profile = CAPABILITY_PROFILES[profileName];
        if (!profile) {
            // 极端兜底，理论上不会发生
            return Object.assign({}, CAPABILITY_PROFILES['chat-only']);
        }
        return Object.assign({}, profile);
    }

    /**
     * 模型能力档案类
     */
    class CapabilityProfile {
        /**
         * 获取完整档案
         * @param {string} modelId
         * @returns {Object}
         */
        get(modelId) {
            return getProfile(modelId);
        }

        /**
         * 获取执行策略
         * @param {string} modelId
         * @returns {string}
         */
        getStrategy(modelId) {
            return getProfile(modelId).strategy;
        }

        /**
         * 是否支持工具调用
         * @param {string} modelId
         * @returns {boolean}
         */
        supportsToolUse(modelId) {
            return !!getProfile(modelId).supportsToolUse;
        }

        /**
         * 解析本次调用是否下发原生 tools（含通道/配置/档案三层决策）
         * @param {string} modelId
         * @param {Object} config
         * @returns {{nativeTools:boolean, profileName:string, maxToolRounds:number}}
         */
        resolveToolMode(modelId, config) {
            return resolveToolMode(modelId, config);
        }
    }

    // 注册到 App 命名空间
    if (typeof App !== 'undefined') {
        App.CAPABILITY_PROFILES = CAPABILITY_PROFILES;
        App.getProfile = getProfile;
        App.resolveToolMode = resolveToolMode;
        App.CapabilityProfile = new CapabilityProfile();
        ChatLog.info('CapabilityProfile', '能力档案模块已挂载到 App.CapabilityProfile（含 resolveToolMode 工具模式决策）');
    } else {
        ChatLog.warn('CapabilityProfile', 'App 命名空间未找到，CapabilityProfile 未挂载');
    }
})();
