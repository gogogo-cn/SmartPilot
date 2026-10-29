/**
 * ============================================================
 * chat-v2/services/render-html.js — 聚合数据 → HTML 渲染
 * ============================================================
 * 职责：从 text-pipeline 抽离的全部 HTML 生成函数。
 *   - enumToEchartsType：Excel 图表枚举 → ECharts 类型映射
 *   - buildAggChartHtml：聚合数据 → 完整 ECharts 页面 HTML
 *   - buildInlineChartHtml：聚合数据 → 面板内嵌 data-echarts div
 *   - buildAggTableHtml：聚合数据 → 表格 HTML 片段
 *   - buildCombinedHtml：聚合数据 → 表格+图表组合 HTML
 *   - esc：统一的 HTML 转义函数（消除重复定义）
 * 纯函数，无副作用。
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    function esc(s) {
        return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;')
            .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    /**
     * 统一 HTML 转义（含单引号），作为各 pipeline / UI 中 escapeHtml 的单一事实源
     */
    function escapeHtml(str) {
        if (str == null) return '';
        return String(str).replace(/[&<>"']/g, c => ({
            '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
        }[c]));
    }

    /**
     * 统一字节大小格式化，收敛 chat-renderer / chat-entry 中的重复实现
     */
    function formatSize(bytes) {
        if (!bytes) return '0 B';
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
        return (bytes / 1024 / 1024).toFixed(1) + ' MB';
    }

    function enumToEchartsType(enumVal) {
        if (!enumVal) return 'bar';
        const str = String(enumVal).trim().toLowerCase();
        // VB 枚举 → ECharts 类型
        const xlMap = {
            'xlpie': 'pie', 'xldoughnut': 'pie',
            'xlbarclustered': 'bar', 'xlbarstacked': 'bar', 'xlbarstacked100': 'bar',
            'xlcolumnclustered': 'bar', 'xlcolumnstacked': 'bar',
            'xlline': 'line', 'xllinemarkers': 'line', 'xlarea': 'line',
            'xlradar': 'radar', 'xlfunnel': 'funnel', 'xltreemap': 'treemap', 'xlsunburst': 'treemap'
        };
        if (xlMap[str]) return xlMap[str];
        // 中文 → ECharts 类型
        if (/饼|环形|pie/.test(str)) return 'pie';
        if (/柱|条形|bar|column/.test(str)) return 'bar';
        if (/折线|线|line/.test(str)) return 'line';
        if (/雷达|radar/.test(str)) return 'radar';
        if (/漏斗|funnel/.test(str)) return 'funnel';
        // 已小写的 echarts 类型直接返回
        const knownTypes = ['pie', 'bar', 'line', 'radar', 'funnel', 'treemap', 'scatter', 'heatmap'];
        if (knownTypes.includes(str)) return str;
        // 兜底
        return 'bar';
    }

    /**
     * 分类轴标签自适应（纯函数）：标签过长旋转 + 自动间隔，柱状图柱宽封顶。
     * 与 chat-renderer.initInlineCharts 的 adaptCategoryLabels 保持一致。
     */
    function categoryAxisAdapt(isHorizontal, labels) {
        const out = { axisLabel: null, barMaxWidth: null };
        if (isHorizontal || !Array.isArray(labels) || labels.length === 0) return out;
        let maxLen = 0;
        for (let i = 0; i < labels.length; i++) { const s = String(labels[i]); if (s.length > maxLen) maxLen = s.length; }
        const n = labels.length;
        out.barMaxWidth = Math.max(24, Math.min(48, Math.floor(600 / Math.max(n, 1))));
        if (maxLen >= 10) {
            out.axisLabel = { rotate: 40, interval: 'auto', hideOverlap: true };
        } else if (maxLen >= 6) {
            out.axisLabel = { rotate: 30, interval: 'auto', hideOverlap: true };
        } else if (n >= 10) {
            out.axisLabel = { interval: 'auto', hideOverlap: true };
        }
        return out;
    }

    function buildAggChartHtml(aggData) {
        if (!aggData || !Array.isArray(aggData.labels)) return '';
        const isHorizontal = /Bar/.test(aggData.chartType || '');
        const type = enumToEchartsType(aggData.chartType);
        const rawTitle = String(aggData.title || '数据报表');
        const title = esc(rawTitle);

        const option = {
            title: { text: rawTitle, left: 'center', textStyle: { fontSize: 16 } },
            tooltip: { trigger: type === 'pie' ? 'item' : 'axis' },
            legend: type === 'pie' ? {} : { bottom: 0 },
            series: [{ type, data: type === 'pie'
                ? aggData.labels.map((l, i) => ({ name: String(l), value: aggData.values[i] }))
                : aggData.values }]
        };
        if (type === 'bar' || type === 'line') {
            const adapt = categoryAxisAdapt(isHorizontal, aggData.labels);
            option.xAxis = isHorizontal
                ? { type: 'value' }
                : Object.assign({ type: 'category', data: aggData.labels }, adapt.axisLabel ? { axisLabel: adapt.axisLabel } : {});
            option.yAxis = isHorizontal ? { type: 'category', data: aggData.labels } : { type: 'value' };
            if (type === 'bar' && option.series && option.series[0] && adapt.barMaxWidth) {
                option.series[0].barMaxWidth = adapt.barMaxWidth;
            }
        }
        const optionJson = JSON.stringify(option);

        return `<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="UTF-8"><title>${title}</title>
<script src="https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js"></script>
<style>body{margin:0;padding:24px;font-family:-apple-system,'Segoe UI',Roboto,Arial,sans-serif;background:#f8fafc;color:#1e293b;}
.container{max-width:1000px;margin:0 auto;background:#fff;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,.1);padding:24px;}
#chart{width:100%;height:480px;}.footer{margin-top:12px;color:#94a3b8;font-size:12px;text-align:right;}</style>
</head><body><div class="container"><div id="chart"></div>
<div class="footer">Generated by Excel AI · 共 ${aggData.labels.length} 组</div></div>
<script>var chart=echarts.init(document.getElementById('chart'));
chart.setOption(${optionJson});
window.addEventListener('resize',function(){chart.resize();});</script></body></html>`;
    }

    function buildInlineChartHtml(aggData) {
        if (!aggData || !Array.isArray(aggData.labels)) return '';
        const type = enumToEchartsType(aggData.chartType);
        const isHorizontal = /Bar/.test(aggData.chartType || '');
        const cfg = {
            title: String(aggData.title || ''),
            type,
            isHorizontal: isHorizontal && type === 'bar',
            labels: aggData.labels.map(l => String(l)),
            values: aggData.values.map(v => Number(v) || 0)
        };
        const cfgJson = JSON.stringify(cfg)
            .replace(/</g, '\\u003c').replace(/>/g, '\\u003e').replace(/&/g, '\\u0026');
        return `<div data-echarts='${cfgJson}' style="width:100%;height:320px;margin:8px 0;"></div>`;
    }

    function buildAggTableHtml(aggData, title) {
        if (!aggData || !Array.isArray(aggData.labels)) return '';
        const dim = aggData.dimColumn || '维度';
        const val = aggData.valueColumn || '数值';
        const n = aggData.labels.length;
        let rowsHtml = '';
        for (let i = 0; i < n; i++) {
            rowsHtml += `<tr><td>${esc(aggData.labels[i])}</td><td>${esc(aggData.values[i])}</td></tr>`;
        }
        return `<div style="margin:8px 0;overflow-x:auto;">
${title ? `<h4 style="margin:0 0 8px;font-size:15px;">${esc(title)}</h4>` : ''}
<table style="width:100%;border-collapse:collapse;font-size:13px;">
<thead><tr><th style="border:1px solid #e2e8f0;padding:6px 10px;background:#f1f5f9;text-align:left;">${esc(dim)}</th>
<th style="border:1px solid #e2e8f0;padding:6px 10px;background:#f1f5f9;text-align:left;">${esc(val)}</th></tr></thead>
<tbody>${rowsHtml}</tbody></table>
<div style="margin-top:6px;color:#94a3b8;font-size:12px;">Generated by Excel AI · 共 ${n} 组</div></div>`;
    }

    function buildCombinedHtml(aggData, title) {
        if (!aggData || !Array.isArray(aggData.labels)) return '';
        const dim = aggData.dimColumn || '维度';
        const val = aggData.valueColumn || '数值';
        const n = aggData.labels.length;
        const chartType = enumToEchartsType(aggData.chartType);
        const option = {
            title: { text: String(aggData.title || title || ''), left: 'center', textStyle: { fontSize: 16 } },
            tooltip: { trigger: chartType === 'pie' ? 'item' : 'axis' },
            legend: chartType === 'pie' ? {} : { bottom: 0 },
            series: [{ type: chartType, data: chartType === 'pie'
                ? aggData.labels.map((l, i) => ({ name: String(l), value: aggData.values[i] }))
                : aggData.values }]
        };
        const isHorizontal = /Bar/.test(aggData.chartType || '');
        if (type === 'bar' || type === 'line') {
            const adapt = categoryAxisAdapt(isHorizontal, aggData.labels);
            option.xAxis = isHorizontal
                ? { type: 'value' }
                : Object.assign({ type: 'category', data: aggData.labels }, adapt.axisLabel ? { axisLabel: adapt.axisLabel } : {});
            option.yAxis = isHorizontal ? { type: 'category', data: aggData.labels } : { type: 'value' };
            if (type === 'bar' && option.series && option.series[0] && adapt.barMaxWidth) {
                option.series[0].barMaxWidth = adapt.barMaxWidth;
            }
        }
        const optionJson = JSON.stringify(option);
        let rowsHtml = '';
        for (let i = 0; i < n; i++) {
            rowsHtml += `<tr><td>${esc(aggData.labels[i])}</td><td>${esc(aggData.values[i])}</td></tr>`;
        }
        return `<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="UTF-8"><title>${esc(title || '数据报表')}</title>
<script src="https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js"></script>
<style>body{margin:0;padding:24px;font-family:-apple-system,'Segoe UI',Roboto,Arial,sans-serif;background:#f8fafc;color:#1e293b;}
.container{max-width:1000px;margin:0 auto;background:#fff;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,.1);padding:24px;}
h1{font-size:18px;margin:0 0 16px;}#chart{width:100%;height:360px;margin-bottom:20px;}
table{width:100%;border-collapse:collapse;font-size:14px;}th,td{border:1px solid #e2e8f0;padding:8px 12px;text-align:left;}
th{background:#f1f5f9;font-weight:600;}tr:nth-child(even){background:#f8fafc;}
.footer{margin-top:12px;color:#94a3b8;font-size:12px;text-align:right;}</style>
</head><body><div class="container"><h1>${esc(title || '聚合结果')}</h1>
<div id="chart"></div>
<table><thead><tr><th>${esc(dim)}</th><th>${esc(val)}</th></tr></thead><tbody>${rowsHtml}</tbody></table>
<div class="footer">Generated by Excel AI · 共 ${n} 组</div></div>
<script>var chart=echarts.init(document.getElementById('chart'));
chart.setOption(${optionJson});
window.addEventListener('resize',function(){chart.resize();});</script></body></html>`;
    }

    App.ChatV2.renderHtml = {
        esc,
        escapeHtml,
        formatSize,
        enumToEchartsType,
        buildAggChartHtml,
        buildInlineChartHtml,
        buildAggTableHtml,
        buildCombinedHtml
    };

    ChatLog.info?.('CHAT-V2', '✅ renderHtml loaded');
})();
