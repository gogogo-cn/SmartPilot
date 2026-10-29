/**
 * ============================================================
 * chat-v2/services/intent-detector.js — 意图识别引擎
 * ============================================================
 * 职责：从 text-pipeline 抽离的全部意图识别正则与 detect* 函数。
 *   - 图表类型识别（chartTypeDefs / chartWordRe）
 *   - 浏览器展示意图 / 展示模式
 *   - 写入工作表意图
 *   - 聚合意图 / 聚合方式 / 维度列 / 数值列
 *   - 高级场景 / 筛选条件
 * 纯函数，无副作用，可独立测试。
 *
 * 关键词唯一事实源：App.ChatV2.intentConfig（Config/intent_config.json）。
 * 本文件不再内嵌任何关键词字面量；每次调用经 defs() 取最新配置。
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    // 配置访问入口（加载器在 intent-config.js，须先于本文件加载）
    function defs() {
        const ic = App.ChatV2.intentConfig;
        return (ic && typeof ic.defs === 'function') ? ic.defs() : {};
    }

    // 兼容旧导出：常量改为 getter，实时反映当前配置（task-planner.js 等消费方无需改动）
    // 旧 UPPER_SNAKE 键 → defs() 驼峰键映射（配置化后 defs() 只暴露驼峰键）
    const LEGACY_KEY_MAP = {
        CHART_TYPE_DEFS: 'chartTypeDefs',
        CHART_WORD_RE: 'chartWordRe',
        CHART_WORD_COUNT_RE: 'chartWordCountRe',
        CHART_SPLIT_WORDS: 'chartSplitWords',
        BROWSER_DISPLAY_RE: 'browserDisplayRe',
        AGGREGATE_INTENT_RE: 'aggregateIntentRe',
        ADVANCED_PATTERNS: 'advancedPatterns',
        AGG_FUNC_DEFS: 'aggFuncDefs',
        DIM_COL_PATTERNS: 'dimColPatterns',
        VALUE_COL_PATTERNS: 'valueColPatterns'
    };
    function getter(key) {
        return {
            get() {
                const d = defs();
                return d[LEGACY_KEY_MAP[key] || key];
            }
        };
    }

    function countChartWords(text) {
        if (!text) return 0;
        const re = defs().chartWordCountRe;
        if (!re) return 0;
        const m = String(text).match(re);
        return m ? m.length : 0;
    }

    function detectChartType(text) {
        if (!text) return null;
        for (const def of defs().chartTypeDefs || []) {
            if (def.re && def.re.test(text)) return def;
        }
        return null;
    }

    function isBrowserDisplayIntent(text) {
        if (!text) return false;
        const re = defs().browserDisplayRe;
        return !!(re && re.test(text));
    }

    function detectDisplayMode(text) {
        if (!text) return null;
        const dm = defs().displayMode || {};
        if (dm.bothRe && dm.bothRe.test(text)) return 'both';
        if (dm.chartRe && dm.chartRe.test(text)) return 'chart';
        if (dm.tableRe && dm.tableRe.test(text)) return 'table';
        return null;
    }

    function extractTargetSheet(text) {
        if (!text) return null;
        const ts = defs().targetSheet || {};
        const verbs = (ts.verbs || []).join('|');
        const nouns = (ts.nouns || []).join('|');
        const exclude = (ts.exclude || []).join('|');
        if (!verbs) return null;
        const m = text.match(new RegExp(
            '(?:' + verbs + ')(?:一个|新的)?(?:' + nouns + ')?[「"\']?(?!' + exclude + '|[=()｛｝【】\\s])' +
            '([^」"\'"，,。；;\\s=()｛｝【】]{1,20})[」"\']?'
        ));
        if (m && m[1]) return m[1].trim();
        const m2 = text.match(new RegExp('(?:' + verbs + ')["\'「」]([^"\'「」]{1,20})["\'「」]'));
        if (m2 && m2[1]) return m2[1].trim();
        return null;
    }

    // ================================================================
    // 聚合意图识别
    // ================================================================

    function detectAggregateIntent(text) {
        if (!text) return false;
        const re = defs().aggregateIntentRe;
        return !!(re && re.test(text));
    }

    function detectAggFunc(text) {
        if (!text) return null;
        for (const def of defs().aggFuncDefs || []) {
            if (def.re && def.re.test(text)) return def;
        }
        return { func: 'sum', label: '求和' };
    }

    // 聚合语法兜底正则：由配置 aggregateIntent 段拼装（groupPrefixes + 动作词），不内嵌关键词
    function aggGrammarRe() {
        const d = defs();
        const ai = d._aggregateIntent || {};
        const join = arr => (arr || []).join('|');
        const prefixes = join(ai.groupPrefixes);
        const actions = join([].concat(
            ai.aggWords || [], ai.avgWords || [], ai.minMaxWords || [],
            ai.countWords || [], ai.groupWords || [], ai.displayWords || []
        ));
        if (!prefixes || !actions) return null;
        return new RegExp('(?:' + prefixes + ')([^，,。\\s]+?)(?:' + actions + '|的)');
    }

    function detectDimensionColumns(text) {
        if (!text) return [];
        const dims = [];
        for (const p of defs().dimColPatterns || []) {
            if (p.re && p.re.test(text)) {
                // 返回用户实际输入的关键词（如「区域」）而非规范化标签（如「地区」），
                // 这样 dimColumn 能精确命中表头真实列名；后端另有语义同义词兜底。
                const m = text.match(p.re);
                dims.push((m && m[0]) || p.label);
            }
        }
        if (dims.length === 0) {
            const byMatch = text.match(aggGrammarRe());
            if (byMatch && byMatch[1]) {
                dims.push(byMatch[1]);
            }
        }
        return dims;
    }

    function detectValueColumns(text) {
        if (!text) return [];
        const vals = [];
        for (const p of defs().valueColPatterns || []) {
            if (p.re && p.re.test(text)) vals.push(p.label);
        }
        if (vals.length === 0) {
            const d = defs();
            const ai = d._aggregateIntent || {};
            const valueWords = (ai.valueWords || []).join('|');
            if (valueWords) {
                const ofMatch = text.match(new RegExp('([^，,。\\s]+?)的(?:' + valueWords + ')'));
                if (ofMatch) {
                    const colMatch = text.match(new RegExp('(?:' + valueWords + ')'));
                    if (colMatch) vals.push(colMatch[1]);
                }
            }
        }
        return vals;
    }

    function detectAdvancedScenario(text) {
        if (!text) return null;
        for (const p of defs().advancedPatterns || []) {
            if (p.re && p.re.test(text)) return p;
        }
        return null;
    }

    function detectFilterConditions(text) {
        if (!text) return null;
        // 筛选词来自配置 advancedPatterns 中 type=filtered 的正则源串（不内嵌关键词）
        const filtered = (defs().advancedPatterns || []).find(p => p.type === 'filtered');
        const filterRe = filtered && filtered.re;
        const filterMatch = filterRe && text.match(new RegExp(filterRe.source + '\\s*(.+?)(?:的|数据|记录|$)'));
        if (filterMatch && filterMatch[1]) {
            const filterText = filterMatch[1].trim();
            let filterColumn = null;
            let filterValue = filterText;
            const kvMatch = filterText.match(/^(.+?)[=:：]\s*(.+)$/);
            if (kvMatch) {
                filterColumn = kvMatch[1].trim();
                filterValue = kvMatch[2].trim();
            } else {
                for (const p of defs().dimColPatterns || []) {
                    if (p.re && (p.re.test(filterText) || filterText.includes(p.label))) {
                        filterColumn = p.label;
                        filterValue = filterText.replace(new RegExp(p.label, 'i'), '').trim() || filterText;
                        break;
                    }
                }
                if (!filterColumn) {
                    const valueMatch = filterText.match(/[\u4e00-\u9fa5]+/g);
                    if (valueMatch && valueMatch.length > 0) {
                        filterValue = filterText;
                    }
                }
            }
            return { filterText, filterColumn, filterValue };
        }
        return null;
    }

    // 导出：函数 + 常量 getter（常量实时反映配置，供 task-planner.js 等消费方使用）
    const exported = {
        countChartWords,
        detectChartType,
        isBrowserDisplayIntent,
        detectDisplayMode,
        extractTargetSheet,
        detectAggregateIntent,
        detectAggFunc,
        detectDimensionColumns,
        detectValueColumns,
        detectAdvancedScenario,
        detectFilterConditions,
        defs
    };
    for (const key of ['CHART_TYPE_DEFS', 'CHART_WORD_RE', 'CHART_WORD_COUNT_RE', 'CHART_SPLIT_WORDS',
                       'BROWSER_DISPLAY_RE', 'AGGREGATE_INTENT_RE', 'ADVANCED_PATTERNS',
                       'AGG_FUNC_DEFS', 'DIM_COL_PATTERNS', 'VALUE_COL_PATTERNS']) {
        Object.defineProperty(exported, key, getter(key));
    }

    App.ChatV2.intentDetector = exported;

    (window.ChatLog || console).info?.('CHAT-V2', '✅ intentDetector loaded');
})();
