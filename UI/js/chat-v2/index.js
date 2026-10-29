/**
 * ============================================================
 * chat-v2/index.js — 聚合入口 & 模块加载校验
 * ============================================================
 * 职责：
 *   - 校验所有 chat-v2 子模块已加载
 *   - 提供版本号与启动诊断
 *   - 显式触发 renderer / chatEntry 初始化（确保依赖就绪）
 *
 * 加载顺序（由 index.html 中的 <script> 顺序保证）：
 *   1. services/session-store.js
 *   2. services/stream-processor.js
 *   3. services/content-extractor.js
 *   4. services/llm-client.js         （依赖 StreamProcessor）
 *   5. services/tool-registry.js
 *   6. services/prompt-builder.js
 *   7. services/intent-detector.js    ← 新增：意图识别正则 + detect*
 *   8. services/task-planner.js       ← 新增：任务拆分（依赖 intentDetector）
 *   9. services/render-html.js        ← 新增：聚合数据 → HTML 渲染
 *  10. services/task-execution.js     ← 新增：计数器/summary/徽章/直接调用
 *  11. services/script-generator.js   ← 新增：Python 脚本（JSON 参数结构化注入）
 *  12. services/json-utils.js         ← 新增：统一 JSON 解析（safeParse/extractJson*）
 *  13. services/llm-io.js             ← 新增：LLM 流式消费封装（streamText/askJson，依赖 json-utils）
 *  12. agent-loop.js                  （依赖 llmClient/toolRegistry/sessionStore）
 *  13. pipelines/text-pipeline.js    （消费上述 5 个 service）
 *  14. pipelines/image-pipeline.js
 *  15. pipelines/doc-pipeline.js
 *  16. chat-orchestrator.js           （依赖 sessionStore/pipelines/routingConstants）
 *  17. ui/chat-renderer.js            （依赖 sessionStore/chatEntry）
 *  18. chat-entry.js                  （依赖 orchestrator/contextEngine/attachmentManager）
 *  19. index.js（本文件）
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    App.ChatV2.VERSION = 'v202609090050';
    App.ChatV2.LOADED_AT = new Date().toISOString();

    /**
     * 校验所有模块已加载
     * @returns {{ok: boolean, missing: string[], loaded: string[]}}
     */    function verifyModules() {
        const required = [
            'sessionStore',  'compaction',      'trajectoryReplay', 'StreamProcessor', 'ContentExtractor',
            'llmClient',     'agentEvents',      'toolRegistry',     'promptBuilder',
            'pipelineRegistry',                 'services.skillSpecEngine',
            'intentDetector', 'taskPlanner',    'renderHtml',        'taskExecution',  'scriptGenerator',
            'jsonUtils', 'llmIo', 'intentAnalyzer', 'routingConstants',
            'agentLoop',
            'pipelines.text', 'pipelines.image', 'pipelines.doc',
            'orchestrator',  'renderer',         'chatEntry'
        ];
        const loaded = [];
        const missing = [];
        for (const path of required) {
            const ok = path.split('.').reduce((obj, key) => obj && obj[key], App.ChatV2);
            if (ok) {
                loaded.push(path);
            } else {
                missing.push(path);
            }
        }
        return { ok: missing.length === 0, missing, loaded };
    }

    /**
     * 启动：校验 + 初始化 renderer + chatEntry
     * chatEntry 和 renderer 已有 DOMContentLoaded 自动初始化，本函数确保依赖就绪后再次触发
     */
    function start() {
        const status = verifyModules();
        if (status.ok) {
            ChatLog.info?.('CHAT-V2', `✅ chat-v2 ${App.ChatV2.VERSION} 全部 ${status.loaded.length} 个模块就绪`);
        } else {
            ChatLog.error?.('CHAT-V2', `❌ chat-v2 模块缺失: ${status.missing.join(', ')}`);
            ChatLog.info?.('CHAT-V2', `已加载: ${status.loaded.join(', ')}`);
        }

        // 显式触发 renderer 初始化（确保 sessionStore 已就绪时才订阅）
        try {
            App.ChatV2.renderer?.init?.();
        } catch (e) {
            ChatLog.error?.('CHAT-V2', '[index] renderer init failed', e);
        }

        // 显式触发 chatEntry 初始化（重新绑定 UI 以防 DOMContentLoaded 已错过）
        try {
            App.ChatV2.chatEntry?.init?.();
        } catch (e) {
            ChatLog.error?.('CHAT-V2', '[index] chatEntry init failed', e);
        }

        // 简短诊断
        ChatLog.info?.('CHAT-V2', '[index] tools=' + (App.ChatV2.toolRegistry?.names?.()?.length || 0)
            + ' history=' + (App.ChatV2.sessionStore?.getHistory?.()?.length || 0));

        return status;
    }

    // 在所有 defer 脚本加载完成后启动
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', start);
    } else {
        // 'interactive' 或 'complete' 时立即启动
        start();
    }

    App.ChatV2.verifyModules = verifyModules;
    App.ChatV2.start = start;

    ChatLog.info?.('CHAT-V2', `✅ chat-v2 index loaded (version ${App.ChatV2.VERSION})`);
})();
