/*
 * chat-v2/ui/renderers/suggestions.js - Follow-up suggestion pure-data submodule
 * Extracted from chat-renderer.js (getTurnPipeline + getFollowUpSuggestions).
 * Pure data functions, unit-testable with node.
 */

(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.renderers) App.ChatV2.renderers = {};

    const ChatLog = window.ChatLog || console;
    /** 推断当前 turn 的管道（优先 _execSpec.pipeline，其次 _matchedSkill） */
    function getTurnPipeline(turn) {
        if (turn && turn._execSpec && turn._execSpec.pipeline) return turn._execSpec.pipeline;
        if (turn && turn.context && turn.context._matchedSkill) return 'skill';
        return '';
    }

    /** 管道类型兜底建议（仅当 turn 无 _produce 实际产出信息时使用） */
    function buildByPipeline(pipeline) {
        const map = {
            python: [
                { label: '再对比一下', prompt: '再对比一下刚才的分析结果' },
                { label: '把趋势细化', prompt: '把刚才的趋势分析进一步细化' },
                { label: '生成图表', prompt: '把刚才的分析结果生成图表' },
                { label: '做成PPT', prompt: '把刚才的分析结果做成PPT' }
            ],
            text: [
                { label: '调整图表', prompt: '调整一下刚才的图表' },
                { label: '换个图表类型', prompt: '把刚才的图表换成其他类型' },
                { label: '做成PPT', prompt: '把刚才的图表结果做成PPT' }
            ],
            // skill 管道不再兜底写死"调整图表/做成PPT"等无关建议：
            // 有 uiActions 的技能由 chat-renderer 在底部渲染专门的脚本盒子 + 操作按钮；
            // 其余技能依赖数据驱动的 _produce（hasChart/rowCount）或技能自身 quickPrompts。
            code: [
                { label: '调整配色', prompt: '调整一下刚才PPT的配色' },
                { label: '增加章节', prompt: '在刚才的PPT里增加一个章节' },
                { label: '生成图表', prompt: '为刚才的PPT生成图表' }
            ]
        };
        return map[pipeline] || [];
    }

    /** 数据驱动：根据 turn._produce 实际产出信息动态生成后续建议 */
    function buildFromTurnProduce(turn) {
        const p = turn && turn._produce;
        if (!p || typeof p !== 'object') return null;
        const out = [];
        const seen = new Set();
        const add = (label, prompt) => {
            const key = `${label}|${prompt}`;
            if (!seen.has(key)) { seen.add(key); out.push({ label, prompt }); }
        };
        // 结果本身是本地文件 → 提供文件类操作
        if (p.filePath) add('重新打开报表', '重新打开刚才生成的网页报表');
        // 有图 → 图表类建议（贴合报表场景的措辞）；无图但有多行数据 → 引导生成图表
        if (p.hasChart) {
            add('调整图表', '调整刚才报表里的图表（例如换一个分类口径）');
            add('换个图表类型', '把刚才报表的图表换成其他类型');
        } else if (p.rowCount) {
            add('生成图表', '把刚才读取的数据生成图表');
        }
        // 有数据 → 沉稳追加一个汇报出口
        if (p.rowCount) add('做成PPT', '把刚才的报表和图表整理成PPT汇报');
        return out.slice(0, 4);
    }

    /** 按管道/产物返回后续建议（数据驱动优先，管道分桶兜底） */
    function getFollowUpSuggestions(input) {
        // 传入 turn 对象（含 id）：优先根据 _produce 实际产出生成；无产物则退回管道分桶
        if (input && typeof input === 'object' && input.id) {
            const fromProduce = buildFromTurnProduce(input);
            if (fromProduce && fromProduce.length) return fromProduce;
            return buildByPipeline(getTurnPipeline(input));
        }
        // 兼容旧调用：直接传入 pipeline 字符串
        return buildByPipeline(input);
    }

    App.ChatV2.renderers.suggestions = {
        getTurnPipeline,
        getFollowUpSuggestions
    };
    ChatLog.info?.('CHAT-V2', 'renderers/suggestions loaded');
})();