/**
 * ============================================================
 * shared-utils.js — Shared Utilities
 * ============================================================
 * 功能：多个路由模块共享的意图检测、图表工具和错误处理函数。
 *
 * 核心函数：
 *   - detectChartTypeFromMessage() — 从消息检测图表类型
 *   - getChartTypeDisplayName() — 图表类型常量转中文名
 *   - getAggFuncDisplayName() — 聚合函数常量转中文名
 *   - hasAggregationIntent() — 检测聚合/分组统计语义
 *   - isReadDataIntent() — 检测读取/展示数据意图
 *   - hasChartIntent() — 检测图表意图
 *   - detectDualAxisIntent() — 检测双轴/组合图意图
 *   - extractComboConfig() — 解析主次轴分配
 *   - safeErrorText() — 通用错误值转可读文本（消除各模块重复实现）
 *   - safeDisplayText() — 通用展示文本化（工具结果/管道输出）
 *
 * 依赖：无
 * ============================================================
 */

(function() {
    'use strict';

    const ChatLog = window.ChatLog || console;

    /**
     * 从用户消息中提取图表类型（Excel XlChartType 常量）
     * 优先匹配更具体的类型，避免误匹配
     */
    function detectChartTypeFromMessage(message) {
        const lower = (message || '').toLowerCase();
        // 饼图/环形图（必须在柱形图之前，因为"饼"不含"柱"关键词）
        if (/饼图|扇形图|占比图|pie chart|\bpie\b/.test(lower)) return 'xlPie';
        if (/环形图|圆环图|doughnut|donut/.test(lower)) return 'xlDoughnut';
        // 条形图（横向柱）— 必须在 column 之前检查
        if (/条形图|横向柱|bar chart|\bbar\b/.test(lower)) return 'xlBarClustered';
        // 堆叠/百分比柱
        if (/堆叠柱|累积柱/.test(lower)) return 'xlColumnStacked';
        if (/百分比柱|100.?堆叠/.test(lower)) return 'xlColumnStacked100';
        // 柱形图（最常见）
        if (/柱状图|柱形图|分组柱|对比图|column chart|\bcolumn\b/.test(lower)) return 'xlColumnClustered';
        // 折线图
        if (/折线图|折线带点|趋势图|line chart|\bline\b/.test(lower)) return 'xlLineMarkers';
        // 散点图
        if (/散点图|scatter plot|\bscatter\b|xy图/.test(lower)) return 'xlXYScatter';
        // 面积图
        if (/面积图|area chart|\barea\b/.test(lower)) return 'xlArea';
        // 其他高级类型
        if (/瀑布图|waterfall/.test(lower)) return 'xlWaterfall';
        if (/漏斗图|funnel/.test(lower)) return 'xlFunnel';
        if (/雷达图|radar/.test(lower)) return 'xlRadar';
        if (/箱线图|boxwhisker|box plot/.test(lower)) return 'xlBoxwhisker';
        if (/树状图|treemap/.test(lower)) return 'xlTreemap';
        if (/旭日图|sunburst/.test(lower)) return 'xlSunburst';
        if (/直方图|histogram/.test(lower)) return 'xlHistogram';
        if (/气泡图|bubble/.test(lower)) return 'xlBubble';
        // 默认：柱形图
        return 'xlColumnClustered';
    }

    function getChartTypeDisplayName(chartType) {
        const map = {
            'xlColumnClustered': '柱状图',
            'xlColumnStacked': '堆叠柱',
            'xlColumnStacked100': '100%堆叠柱',
            'xlBarClustered': '条形图',
            'xlBarStacked': '堆叠条形图',
            'xlBarStacked100': '100%堆叠条形图',
            'xlLine': '折线图',
            'xlLineMarkers': '折线带点',
            'xlArea': '面积图',
            'xlPie': '饼图',
            'xlDoughnut': '圆环图',
            'xlXYScatter': '散点图',
            'xlRadar': '雷达图',
            'xlRadarFilled': '填充雷达图',
            'xlBubble': '气泡图',
            'xlSunburst': '旭日图',
            'xlTreemap': '树状图',
            'xlHistogram': '直方图',
            'xlBoxwhisker': '箱线图',
            'xlFunnel': '漏斗图',
            'xlWaterfall': '瀑布图'
        };
        return map[chartType] || '柱状图';
    }

    function getAggFuncDisplayName(aggFunc) {
        const map = { 'sum': '求和', 'avg': '平均值', 'count': '计数', 'max': '最大值', 'min': '最小值' };
        return map[aggFunc] || '求和';
    }

    /**
     * 检测消息是否包含聚合/分组统计语义
     * 支持：按部门统计、每个项目单量、各项目数量、统计每个项目等
     */
    function hasAggregationIntent(message) {
        const msg = message || '';
        return /按\s*[^\s，,、按]+?(?:\s*分组|\s*统计|\s*求和|\s*汇总|\s*计算|\s*聚合|\s*分析|\s*展示|\s*显示)/.test(msg) ||
               /每个\s*[^\s，,、]+?(?:\s*统计|\s*求和|\s*汇总|\s*计算|\s*聚合|\s*数量|\s*单量|\s*个数|\s*笔数)/.test(msg) ||
               /各\s*[^\s，,、]+?(?:\s*统计|\s*求和|\s*汇总|\s*计算|\s*聚合|\s*数量|\s*单量|\s*个数|\s*笔数)/.test(msg) ||
               /统计\s*每个\s*[^\s，,、]+/.test(msg) ||
               /统计\s*各\s*[^\s，,、]+/.test(msg);
    }

    /**
     * 检测消息是否仅要求读取/展示选中数据
     */
    function isReadDataIntent(message) {
        return /(?:列出|读取|查看|展示|显示|获取|提取)\s*(?:选中|数据|表格|区域|范围|单元格|内容)/.test(message || '');
    }

    /**
     * 检测消息是否包含图表意图
     */
    function hasChartIntent(message) {
        return /(?:画|绘制|生成|做|创建|画一个|画个)?\s*(?:图表|柱状图|条形图|折线图|趋势图|饼图|散点图|面积图|可视化|chart|graph|plot)/.test(message || '');
    }

    /**
     * 检测是否要求双轴/组合图
     */
    function detectDualAxisIntent(message) {
        return /(主轴|次轴|主次轴|双轴|组合图|左右轴|secondary axis|combo chart|次坐标轴|第二轴)/i.test(message || '');
    }

    /**
     * 通用错误值转可读文本 —— 消除各模块重复实现
     * 适用场景：把任意错误值（函数/Promise/对象/null）转成用户可读的中文错误描述
     * 避免 String() 成 "[object Object]" 或 "function () { [native code] }"
     * @param {*} v - 任意错误值
     * @param {string} [defaultMsg='操作失败'] - 默认错误消息
     * @param {number} [maxLen=500] - 最大长度
     * @returns {string} 可读错误文本
     */
    function safeErrorText(v, defaultMsg, maxLen) {
        const def = defaultMsg || '操作失败';
        const limit = typeof maxLen === 'number' && maxLen > 0 ? maxLen : 500;
        if (v == null) return def;
        if (typeof v === 'string') return v || def;
        if (typeof v === 'function') return '后端调用返回了函数对象（可能未正确 await 或桥接异常）';
        if (v instanceof Error) return v.message || String(v);
        if (typeof v === 'object') {
            if (typeof v.message === 'string') return v.message;
            try {
                const s = JSON.stringify(v);
                return s && s.length <= limit ? s : (s ? s.slice(0, limit) + '…' : def);
            } catch (e) { return String(v); }
        }
        return String(v);
    }

    /**
     * 通用展示文本化 —— 把任意结果值转成可靠的可读文本
     * 适用场景：工具执行结果/管道输出的前端展示
     * 绝不 String() 成 "[object Object]" 或 "function () { [native code] }"
     * @param {*} value - 任意结果值
     * @param {number} [maxLen=500] - 最大长度
     * @returns {string} 可读展示文本
     */
    function safeDisplayText(value, maxLen) {
        const limit = typeof maxLen === 'number' && maxLen > 0 ? maxLen : 500;
        if (value == null) return '';
        const ty = typeof value;
        if (ty === 'string') return value.length > limit ? value.slice(0, limit) + '…' : value;
        if (ty === 'number' || ty === 'boolean' || ty === 'bigint') return String(value);
        if (ty === 'function') return '后端返回了函数对象（可能未正确 await 或桥接异常）';
        if (value instanceof Error) return value.message || String(value);
        if (ty === 'object') {
            // 顶层常见可读字段优先
            for (const k of ['resultText', 'result', 'message', 'msg', 'text', 'summary', 'content']) {
                const v = value[k];
                if (typeof v === 'string' && v.trim()) {
                    return v.length > limit ? v.slice(0, limit) + '…' : v;
                }
            }
            try {
                const s = JSON.stringify(value);
                if (s && s.length > limit) return s.slice(0, limit) + '…';
                return s || String(value);
            } catch (e) {
                return '[无法序列化的结果]';
            }
        }
        return String(value);
    }

    /**
     * 从已提取的指标中解析主次轴分配
     */
    function extractComboConfig(message, metrics) {
        const hasDualAxis = detectDualAxisIntent(message);
        ChatLog.info('SharedUtils', '[组合图检测] hasDualAxis:', hasDualAxis, 'metrics.length:', metrics ? metrics.length : 0);
        if (!hasDualAxis || !metrics || metrics.length < 2) return null;

        let primaryIndex = 0;
        let secondaryIndex = 1;

        // 尝试从"主轴...显示/表示..."中提取指标名
        const primaryMatch = message.match(/主轴[^，,、]*?(?:显示|表示|画|用|为)?\s*([^\s，,、]+)/i);
        const secondaryMatch = message.match(/次轴[^，,、]*?(?:显示|表示|画|用|为)?\s*([^\s，,、]+)/i);

        if (primaryMatch) {
            const name = primaryMatch[1].trim();
            const idx = metrics.findIndex(function(m) {
                return m.name.indexOf(name) >= 0 || name.indexOf(m.name) >= 0;
            });
            if (idx >= 0) primaryIndex = idx;
        }
        if (secondaryMatch) {
            const name = secondaryMatch[1].trim();
            const idx = metrics.findIndex(function(m) {
                return m.name.indexOf(name) >= 0 || name.indexOf(m.name) >= 0;
            });
            if (idx >= 0) secondaryIndex = idx;
        }

        // 确保主次轴不是同一个指标
        if (secondaryIndex === primaryIndex) {
            secondaryIndex = metrics.findIndex(function(m, i) { return i !== primaryIndex; });
        }

        const primary = metrics[primaryIndex];
        const secondary = metrics[secondaryIndex];
        if (!primary || !secondary) return null;

        // 默认主轴柱状图、次轴线形图；可按用户指定调整
        const primaryChartType = /主轴[^，,、]*?(?:折线|线形|line)/i.test(message) ? 'xlLine' :
                                 /主轴[^，,、]*?(?:条形|bar)/i.test(message) ? 'xlBarClustered' : 'xlColumnClustered';
        const secondaryChartType = /次轴[^，,、]*?(?:柱形|柱状|column)/i.test(message) ? 'xlColumnClustered' :
                                   /次轴[^，,、]*?(?:条形|bar)/i.test(message) ? 'xlBarClustered' : 'xlLine';

        return {
            primary: Object.assign({}, primary, { chartType: primaryChartType, axisGroup: 'primary' }),
            secondary: Object.assign({}, secondary, { chartType: secondaryChartType, axisGroup: 'secondary' })
        };
    }

    // ==================== 初始化和注册 ====================

    if (typeof App !== 'undefined') {
        App.SharedUtils = {
            detectChartTypeFromMessage: detectChartTypeFromMessage,
            getChartTypeDisplayName: getChartTypeDisplayName,
            getAggFuncDisplayName: getAggFuncDisplayName,
            hasAggregationIntent: hasAggregationIntent,
            isReadDataIntent: isReadDataIntent,
            hasChartIntent: hasChartIntent,
            detectDualAxisIntent: detectDualAxisIntent,
            extractComboConfig: extractComboConfig,
            // 通用错误/展示文本工具（P0-3 新增）
            safeErrorText: safeErrorText,
            safeDisplayText: safeDisplayText
        };
        ChatLog.info('SharedUtils', '✅ 共享工具模块已加载');
    }

})();