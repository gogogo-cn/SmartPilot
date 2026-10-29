/*
 * chat-v2/ui/renderers/inline-chart.js - Inline ECharts renderer submodule
 * Extracted from chat-renderer.js (initInlineCharts + adaptCategoryLabels).
 * Self-contained ECharts initializer over [data-echarts] nodes; DOM-bound.
 */

(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.renderers) App.ChatV2.renderers = {};

    const ChatLog = window.ChatLog || console;
    /**
     * 扫描容器内所有 data-echarts 节点，初始化为 ECharts。
     * 每个节点：<div data-echarts='{...json配置}'></div>
     * 配置包含: { title, type(pie/bar/line), labels, values, isHorizontal }
     * 由前端(如 text-pipeline 浏览器展示子任务)产出，用于在面板内嵌渲染，不弹系统浏览器/不写临时文件。
     */
    /**
     * 分类轴标签自适应（纯函数）：
     *  - 标签过长 → 旋转分类轴标签 + 自动间隔/隐藏重叠，避免"标签太长"排不下
     *  - 柱状图 → 柱宽封顶，避免分类少时"柱子太大"
     * 仅对"标签在 X 轴"的柱/折线图生效（饼/热力/水平柱状天然无需旋转）。
     */
    function adaptCategoryLabels(type, isHorizontal, labels) {
        const out = { axisLabel: null, rotateBottom: 0, barMaxWidth: null };
        if (type === 'pie' || type === 'heatmap' || isHorizontal) return out;
        if (!Array.isArray(labels) || labels.length === 0) return out;
        let maxLen = 0;
        for (let k = 0; k < labels.length; k++) { const s = String(labels[k]); if (s.length > maxLen) maxLen = s.length; }
        const n = labels.length;
        out.barMaxWidth = Math.max(24, Math.min(48, Math.floor(600 / Math.max(n, 1))));
        if (maxLen >= 10) {
            out.axisLabel = { rotate: 40, interval: 'auto', hideOverlap: true };
            out.rotateBottom = 72;
        } else if (maxLen >= 6) {
            out.axisLabel = { rotate: 30, interval: 'auto', hideOverlap: true };
            out.rotateBottom = 60;
        } else if (n >= 10) {
            out.axisLabel = { interval: 'auto', hideOverlap: true };
        }
        return out;
    }

    function initInlineCharts(container, retryCount) {
        retryCount = retryCount || 0;
        if (!container) { console.warn('[ECharts] initInlineCharts: container为空'); return; }
        if (typeof echarts === 'undefined') {
            console.warn('[ECharts] initInlineCharts: echarts未加载 (typeof echarts=undefined), retry=' + retryCount);
            if (retryCount < 8) {
                setTimeout(function() { initInlineCharts(container, retryCount + 1); }, 200 * (retryCount + 1));
            }
            return;
        }
        const nodes = container.querySelectorAll('[data-echarts]');
        console.log('[ECharts] initInlineCharts: 找到 ' + nodes.length + ' 个data-echarts节点, retry=' + retryCount);
        if (nodes.length === 0) return;
        const MAX_RETRY = 5;
        let needsRetry = false;
        nodes.forEach(function(node, idx) {
            // 跳过已初始化的节点
            if (node._echartsInstance) { console.log('[ECharts] 节点' + idx + ' 已初始化，跳过'); return; }
            let cfg;
            try {
                const raw = node.getAttribute('data-echarts') || '{}';
                cfg = JSON.parse(raw);
                console.log('[ECharts] 节点' + idx + ' 配置解析成功: type=' + cfg.type + ', labels=' + (cfg.labels ? cfg.labels.length : 0));
            } catch (e) {
                console.error('[ECharts] 节点' + idx + ' 配置解析失败:', e.message, node.getAttribute('data-echarts'));
                return;
            }
            if (!Array.isArray(cfg.labels) || cfg.labels.length === 0) {
                console.warn('[ECharts] 节点' + idx + ' labels为空，跳过');
                return;
            }
            // 宽度为 0 时延迟重试
            const cw = node.clientWidth, ch = node.clientHeight;
            if (!cw || cw < 10) {
                console.warn('[ECharts] 节点' + idx + ' 宽度不足: clientWidth=' + cw + ', clientHeight=' + ch + ', retry=' + retryCount);
                if (retryCount < MAX_RETRY) needsRetry = true;
                return;
            }
            const type = cfg.type || 'bar';
            const isHorizontal = !!cfg.isHorizontal;
            // 注意：热力图分支会对 option 整体重赋值，故声明为 let 而非 const，否则抛 "Assignment to constant variable"
            let option = {
                tooltip: { trigger: type === 'pie' ? 'item' : 'axis' },
                grid: type === 'pie' ? {} : { top: 40, left: 8, right: 16, bottom: 32, containLabel: true },
                series: [{
                    type: type,
                    data: type === 'pie'
                        ? cfg.labels.map(function(l, i) { return { name: String(l), value: cfg.values[i] }; })
                        : cfg.values,
                    label: type === 'pie' ? { show: true, formatter: '{b}: {d}%' } : { show: false },
                    itemStyle: type === 'pie' ? { borderRadius: 4 } : {}
                }]
            };
            // 饼图：title 在顶部、legend 在底部，避免堆叠
            if (type === 'pie') {
                option.title = cfg.title ? {
                    text: cfg.title,
                    left: 'center',
                    top: 0,
                    textStyle: { fontSize: 14 }
                } : { show: false };
                option.legend = {
                    bottom: 0,
                    left: 'center',
                    itemGap: 12,
                    itemWidth: 12,
                    itemHeight: 8,
                    textStyle: { fontSize: 12 }
                };
                // 饼图半径缩小让出空间给 title 和 legend
                option.series[0].radius = '55%';
                option.series[0].center = ['50%', '50%'];
            } else {
                // 柱状图等：title 顶部、legend 底部
                option.title = cfg.title ? { text: cfg.title, left: 'center', textStyle: { fontSize: 14 } } : { show: false };
                option.legend = { bottom: 0 };
            }
            if (type === 'bar' || type === 'line') {
                // 柱状图/折线图都是笛卡尔坐标，必须显式提供 xAxis/yAxis；
                // 否则 ECharts 构建坐标系时访问空 axis 会抛 "Cannot read ... 'get'"。
                const adapt = adaptCategoryLabels(type, isHorizontal, cfg.labels);
                option.xAxis = isHorizontal
                    ? { type: 'value' }
                    : Object.assign({ type: 'category', data: cfg.labels }, adapt.axisLabel ? { axisLabel: adapt.axisLabel } : {});
                option.yAxis = isHorizontal ? { type: 'category', data: cfg.labels } : { type: 'value' };
                // 标签旋转时给底部让出空间，避免文字被裁切
                if (adapt.rotateBottom) option.grid.bottom = adapt.rotateBottom;
                // 柱状图柱宽封顶，避免分类少时柱子过壮
                if (type === 'bar' && adapt.barMaxWidth) {
                    (Array.isArray(option.series) ? option.series : [option.series])
                        .forEach(function(s) { if (s && s.barMaxWidth == null) s.barMaxWidth = adapt.barMaxWidth; });
                }
                // 多系列支持：cfg.series 为数组时，用多系列替换默认单系列（如趋势图：历史+预测）
                if (Array.isArray(cfg.series) && cfg.series.length > 1) {
                    option.series = cfg.series.map(function (s) {
                        return {
                            type: type,
                            name: s.name || '',
                            data: s.data,
                            smooth: s.smooth === true,
                            lineStyle: s.lineStyle || {},
                            itemStyle: s.itemStyle || {}
                        };
                    });
                    option.legend = {
                        bottom: 0,
                        left: 'center',
                        itemGap: 16,
                        itemWidth: 18,
                        itemHeight: 10,
                        textStyle: { fontSize: 12 }
                    };
                    // 多系列时给图例让出底部空间
                    option.grid.bottom = 44;
                }
            } else if (type === 'heatmap' && Array.isArray(cfg.data) && Array.isArray(cfg.xLabels)) {
                // 相关性矩阵热力图：data=[[row,col,value],...]，xLabels/yLabels 为列名
                option = {
                    tooltip: { position: 'top' },
                    grid: { top: 40, left: 8, right: 16, bottom: 64, containLabel: true },
                    xAxis: { type: 'category', data: cfg.xLabels, splitArea: { show: true } },
                    yAxis: { type: 'category', data: cfg.yLabels || cfg.xLabels, splitArea: { show: true } },
                    visualMap: {
                        min: -1, max: 1, calculable: true, orient: 'horizontal',
                        left: 'center', bottom: 0,
                        inRange: { color: ['#1a9850', '#d9ef8b', '#ffffbf', '#fee08b', '#d73027'] }
                    },
                    series: [{
                        type: 'heatmap',
                        data: cfg.data.map(function (p) { return [p[0], p[1], p[2]]; }),
                        label: { show: false },
                        emphasis: { itemStyle: { shadowBlur: 10, shadowColor: 'rgba(0,0,0,0.5)' } }
                    }],
                    title: cfg.title ? { text: cfg.title, left: 'center', textStyle: { fontSize: 14 } } : { show: false }
                };
            }
            try {
                const chart = echarts.init(node);
                chart.setOption(option);
                node._echartsInstance = chart;
                console.log('[ECharts] 节点' + idx + ' 渲染成功: type=' + type + ', data points=' + cfg.labels.length);
                if (window.ResizeObserver) {
                    new ResizeObserver(function() { chart.resize(); }).observe(node);
                } else {
                    window.addEventListener('resize', function() { chart.resize(); });
                }
            } catch (e) {
                console.error('[ECharts] 节点' + idx + ' 渲染异常:', e.message);
                node.innerHTML = '图表渲染失败: ' + e.message;
            }
        });
        // 如果有节点需要重试
        if (needsRetry) {
            console.log('[ECharts] 有节点需要重试, 延迟后重试 retry=' + (retryCount + 1));
            setTimeout(function() {
                initInlineCharts(container, retryCount + 1);
            }, 100 * (retryCount + 1));
        }
    }


    App.ChatV2.renderers.inlineChart = {
        initInlineCharts,
        adaptCategoryLabels
    };
    ChatLog.info?.('CHAT-V2', 'renderers/inline-chart loaded');
})();