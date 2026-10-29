/**
 * ============================================================
 * chat-v2/ui/renderers/json.js — JSON 渲染/转表（纯函数子模块）
 * ============================================================
 * 从 chat-renderer.js 拆出的独立子模块，全部 DOM-free、字符串入/字符串出：
 *   - renderJsonHtml(jsonText)              ：JSON 语法高亮
 *   - renderCollapsibleJson(jsonText)       ：可折叠 JSON 面板（<details>）
 *   - tryConvertToTable(data) + render*Table：把结构化数据转成表格 HTML
 *
 * 自包含 escapeHtml（委托到 App.ChatV2.renderHtml.escapeHtml），避免跨模块依赖。
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.renderers) App.ChatV2.renderers = {};

    const ChatLog = window.ChatLog || console;

    function escapeHtml(str) {
        if (str == null) return '';
        return App.ChatV2.renderHtml ? App.ChatV2.renderHtml.escapeHtml(str) : '';
    }

    /**
     * JSON 语法高亮：键、字符串、数字、布尔/null
     * 先整体转义（保护 <>&'），再还原 &quot; → "，让按引号结构的高亮正则真正生效
     * （引号在 <pre> 文本上下文安全，无需转义）。
     */
    function renderJsonHtml(jsonText) {
        const safe = escapeHtml(jsonText).replace(/&quot;/g, '"');
        const html = safe
            .replace(/"([^"\\]*(?:\\.[^"\\]*)*)":/g, '<span class="json-key">"$1"</span>:')
            .replace(/: ("(?:[^"\\]|\\.)*")/g, ': <span class="json-string">$1</span>')
            .replace(/: (-?\d+(?:\.\d+)?)/g, ': <span class="json-number">$1</span>')
            .replace(/: (true|false|null)/g, ': <span class="json-bool">$1</span>');
        return `<pre class="json-block">${html}</pre>`;
    }

    /**
     * 把 JSON 文本渲染为可折叠面板（默认收起），
     * 使用原生 <details>/<summary>，无需额外 JS 事件绑定。
     */
    function renderCollapsibleJson(jsonText, open) {
        return '<details class="json-collapsible"' + (open ? ' open' : '') + ' style="margin-top:6px;">' +
                '<summary style="cursor:pointer;font-size:12px;color:#1a73e8;user-select:none;padding:2px 0;">' +
                    '查看原始结果（JSON）' +
                '</summary>' +
                renderJsonHtml(jsonText) +
            '</details>';
    }

    /**
     * 尝试将 JSON 数据转换为表格 HTML。
     * 返回 null 表示无法识别的格式，保持原样。
     */
    function tryConvertToTable(data, rawText) {
        if (!data || typeof data !== 'object') return null;

        // 1. 描述统计格式: {"列名": {count, mean, median, stdev, min, max, q1, q3}}
        const keys = Object.keys(data);
        if (keys.length === 0) return null;

        const firstVal = data[keys[0]];
        if (firstVal && typeof firstVal === 'object' && !Array.isArray(firstVal)) {
            const subKeys = Object.keys(firstVal);
            // 描述统计特征: count + mean + median
            if (subKeys.indexOf('count') >= 0 && subKeys.indexOf('mean') >= 0 && subKeys.indexOf('median') >= 0) {
                return renderDescriptiveStatsTable(data);
            }
            // 分组聚合特征: value + count + agg
            if (subKeys.indexOf('value') >= 0 && subKeys.indexOf('agg') >= 0) {
                return renderAggregationTable(data);
            }
            // 相关性矩阵特征: 所有 value 都是数字
            if (subKeys.length >= 2 && subKeys.every(function(k) { return typeof firstVal[k] === 'number'; })) {
                // 检查是否是对称矩阵（key 本身也是子 key）
                const isMatrix = keys.every(function(k) {
                    const v = data[k];
                    return v && typeof v === 'object' && !Array.isArray(v) &&
                        Object.keys(v).every(function(sk) { return typeof v[sk] === 'number'; });
                });
                if (isMatrix) {
                    return renderCorrelationMatrix(data);
                }
            }
            // 异常值检测特征: anomaly_count 或 anomalies
            if (subKeys.indexOf('anomaly_count') >= 0 || subKeys.indexOf('anomalies') >= 0) {
                return renderAnomalyTable(data);
            }
            // 数据分布特征: distribution 数组
            if (subKeys.indexOf('distribution') >= 0 && Array.isArray(firstVal.distribution)) {
                return renderDistributionTable(data);
            }
            // 趋势分析特征: direction, forecast
            if (subKeys.indexOf('direction') >= 0 || subKeys.indexOf('forecast') >= 0) {
                return renderTrendTable(data);
            }
        }

        // 2. 简单数组对象 → 标准表格
        if (Array.isArray(data) && data.length > 0 && typeof data[0] === 'object') {
            return renderArrayTable(data);
        }

        return null;
    }

    function renderDescriptiveStatsTable(data, orientation) {
        const statLabels = {
            count: '样本数', mean: '均值', median: '中位数', stdev: '标准差',
            min: '最小值', max: '最大值', q1: 'Q1(25%)', q3: 'Q3(75%)'
        };
        const allStats = ['count', 'mean', 'median', 'stdev', 'min', 'max', 'q1', 'q3'];
        const cols = Object.keys(data);

        // AI 智能选择默认方向：列数少(≤3)→统计量为列(宽表)，列数多(>3)→统计量为行(长表)
        // orientation: 'row' = 统计量为行 / 'col' = 统计量为列
        const orient = orientation || (cols.length <= 3 ? 'col' : 'row');
        const orientAttr = 'data-stats-orient="' + orient + '"';

        if (orient === 'col') {
            // 统计量为列（横向）：每行一个数据列，统计量展开为列
            const activeStats = allStats.filter(function(stat) {
                return cols.some(function(c) { return data[c] && data[c][stat] !== undefined && data[c][stat] !== null; });
            });
            let html = '<table class="py-data-table py-stats-table" ' + orientAttr + '><thead><tr>';
            activeStats.forEach(function(s) { html += '<th>' + (statLabels[s] || s) + '</th>'; });
            html += '</tr></thead><tbody>';
            cols.forEach(function(c) {
                html += '<tr><td><strong>' + escapeHtml(c) + '</strong></td>';
                activeStats.forEach(function(s) {
                    const v = data[c] && data[c][s];
                    html += '<td>' + (v !== undefined && v !== null ? escapeHtml(String(v)) : '-') + '</td>';
                });
                html += '</tr>';
            });
            html += '</tbody></table>';
            return html;
        }

        // 统计量为行（纵向，默认长表）：行为统计量，列为列名
        let html = '<table class="py-data-table py-stats-table" ' + orientAttr + '><thead><tr><th>统计量</th>';
        cols.forEach(function(c) { html += '<th>' + escapeHtml(c) + '</th>'; });
        html += '</tr></thead><tbody>';

        allStats.forEach(function(stat) {
            const rowVals = cols.map(function(c) { return data[c] && data[c][stat]; });
            if (rowVals.some(function(v) { return v !== undefined && v !== null; })) {
                html += '<tr><td><strong>' + (statLabels[stat] || stat) + '</strong></td>';
                rowVals.forEach(function(v) {
                    html += '<td>' + (v !== undefined && v !== null ? escapeHtml(String(v)) : '-') + '</td>';
                });
                html += '</tr>';
            }
        });
        html += '</tbody></table>';
        return html;
    }

    function renderAggregationTable(data) {
        const aggLabels = { avg: '平均', sum: '求和', count: '计数', min: '最小', max: '最大' };
        const aggHead = { avg: '平均值', sum: '求和', count: '计数', min: '最小值', max: '最大值' };
        // 指标列名用真实聚合方式（avg→平均值…），避免笼统叫 "值"，也避免裸英文 value
        const firstVal = data[Object.keys(data)[0]];
        const metricHead = (firstVal && (aggHead[firstVal.agg] || aggLabels[firstVal.agg])) || firstVal && firstVal.agg || '值';
        let html = '<table class="py-data-table py-agg-table"><thead><tr><th>分组</th><th>' + escapeHtml(String(metricHead)) + '</th><th>计数</th><th>聚合方式</th></tr></thead><tbody>';
        const keys = Object.keys(data);
        keys.forEach(function(k) {
            const v = data[k];
            if (v && typeof v === 'object') {
                html += '<tr>';
                html += '<td>' + escapeHtml(k) + '</td>';
                html += '<td>' + escapeHtml(String(v.value ?? '-')) + '</td>';
                html += '<td>' + escapeHtml(String(v.count ?? '-')) + '</td>';
                html += '<td>' + escapeHtml(aggLabels[v.agg] || v.agg || '-') + '</td>';
                html += '</tr>';
            }
        });
        html += '</tbody></table>';
        return html;
    }

    function renderCorrelationMatrix(data) {
        const keys = Object.keys(data);
        let html = '<table class="py-data-table py-corr-table"><thead><tr><th></th>';
        keys.forEach(function(k) { html += '<th>' + escapeHtml(k) + '</th>'; });
        html += '</tr></thead><tbody>';
        keys.forEach(function(rowKey) {
            html += '<tr><td><strong>' + escapeHtml(rowKey) + '</strong></td>';
            const row = data[rowKey] || {};
            keys.forEach(function(colKey) {
                const val = row[colKey];
                let cls = 'py-corr-cell';
                if (typeof val === 'number') {
                    if (Math.abs(val) >= 0.7) cls += ' py-corr-strong';
                    else if (Math.abs(val) >= 0.4) cls += ' py-corr-medium';
                    else cls += ' py-corr-weak';
                }
                html += '<td class="' + cls + '">' + (typeof val === 'number' ? val.toFixed(3) : escapeHtml(String(val ?? '-'))) + '</td>';
            });
            html += '</tr>';
        });
        html += '</tbody></table>';
        return html;
    }

    function renderAnomalyTable(data) {
        const cols = Object.keys(data);
        let html = '<table class="py-data-table py-anomaly-table"><thead><tr><th>列名</th><th>样本数</th><th>异常数</th><th>异常率</th><th>异常详情</th></tr></thead><tbody>';
        cols.forEach(function(c) {
            const v = data[c];
            if (v && typeof v === 'object') {
                html += '<tr>';
                html += '<td><strong>' + escapeHtml(c) + '</strong></td>';
                html += '<td>' + escapeHtml(String(v.total ?? '-')) + '</td>';
                html += '<td class="py-anomaly-count">' + escapeHtml(String(v.anomaly_count ?? '-')) + '</td>';
                html += '<td>' + escapeHtml(String(v.anomaly_rate != null ? v.anomaly_rate + '%' : '-')) + '</td>';
                html += '<td>';
                if (Array.isArray(v.anomalies) && v.anomalies.length > 0) {
                    v.anomalies.slice(0, 5).forEach(function(a) {
                        html += '<span class="py-anomaly-item">值:' + escapeHtml(String(a.value ?? '')) +
                            ' (Z=' + escapeHtml(String(a.z_score ?? '')) + ')</span>';
                    });
                    if (v.anomalies.length > 5) html += '<span class="py-anomaly-more">...等' + v.anomalies.length + '个</span>';
                }
                html += '</td></tr>';
            }
        });
        html += '</tbody></table>';
        return html;
    }

    function renderDistributionTable(data) {
        const cols = Object.keys(data);
        let html = '<table class="py-data-table py-dist-table"><thead><tr><th>列名</th><th>最小</th><th>最大</th><th>分箱数</th><th>分布</th></tr></thead><tbody>';
        cols.forEach(function(c) {
            const v = data[c];
            if (v && typeof v === 'object') {
                html += '<tr>';
                html += '<td><strong>' + escapeHtml(c) + '</strong></td>';
                html += '<td>' + escapeHtml(String(v.min ?? '-')) + '</td>';
                html += '<td>' + escapeHtml(String(v.max ?? '-')) + '</td>';
                html += '<td>' + escapeHtml(String(v.bins ?? '-')) + '</td>';
                html += '<td>';
                if (Array.isArray(v.distribution)) {
                    v.distribution.forEach(function(d) {
                        if (d && typeof d === 'object') {
                            html += '<div class="py-dist-bar">';
                            html += '<span class="py-dist-range">' + escapeHtml(String(d.range || '')) + '</span>';
                            html += '<span class="py-dist-bar-track"><span class="py-dist-bar-fill" style="width:' +
                                Math.min(100, parseFloat(d.percentage) || 0) + '%"></span></span>';
                            html += '<span class="py-dist-pct">' + escapeHtml(String(d.percentage || 0)) + '%</span>';
                            html += '</div>';
                        }
                    });
                }
                html += '</td></tr>';
            }
        });
        html += '</tbody></table>';
        return html;
    }

    function renderTrendTable(data) {
        if (!data || typeof data !== 'object') return null;
        const keys = Object.keys(data);
        let html = '<table class="py-data-table py-trend-table"><tbody>';
        keys.forEach(function(k) {
            if (k === 'forecast' && Array.isArray(data[k])) {
                html += '<tr><td><strong>预测值</strong></td><td>' + escapeHtml(data[k].join(', ')) + '</td></tr>';
            } else if (typeof data[k] !== 'object' && !Array.isArray(data[k])) {
                html += '<tr><td><strong>' + escapeHtml(k) + '</strong></td><td>' + escapeHtml(String(data[k])) + '</td></tr>';
            }
        });
        html += '</tbody></table>';
        return html;
    }

    /** 全量数据 JSON 岛：<script type="application/json"> 不执行、不显示，
     *  下载/复制/插入Excel 的处理器优先读它，避免"DOM 只渲染了前 N 条 → 导出也只剩 N 条"。
     *  </script 序列需转义，防止数据内容提前闭合标签。 */
    function buildFullDataIsland(headers, rows) {
        try {
            const payload = JSON.stringify({ headers: headers, rows: rows })
                .replace(/<\//g, '<\\/');
            return '<script type="application/json" class="table-full-data">' + payload + '</script>';
        } catch (e) { return ''; }
    }

    function renderArrayTable(data) {
        if (!Array.isArray(data) || data.length === 0) return null;
        const headers = Object.keys(data[0]);
        let html = '<table class="py-data-table"><thead><tr>';
        headers.forEach(function(h) { html += '<th>' + escapeHtml(h) + '</th>'; });
        html += '</tr></thead><tbody>';
        // 展示上限 200：截断只应影响"显示多少"，绝不应影响"导出多少"（导出走下方全量数据岛）
        const DISPLAY_MAX = 200;
        data.slice(0, DISPLAY_MAX).forEach(function(row) {
            html += '<tr>';
            headers.forEach(function(h) {
                html += '<td>' + escapeHtml(String(row[h] ?? '')) + '</td>';
            });
            html += '</tr>';
        });
        if (data.length > DISPLAY_MAX) {
            html += '<tr><td colspan="' + headers.length + '" style="text-align:center;color:#999;">...共 ' + data.length + ' 条，仅显示前 ' + DISPLAY_MAX + ' 条（下载/插入Excel 为全量）</td></tr>';
        }
        html += '</tbody></table>';
        // 全量数据岛（含被截断的行）：下载/复制/插入Excel 用它拿完整数据
        const fullRows = data.map(function(row) {
            return headers.map(function(h) { return row[h] == null ? '' : row[h]; });
        });
        html += buildFullDataIsland(headers, fullRows);
        return html;
    }

    /**
     * 尝试把"行矩阵/读取结果"渲染为表格。覆盖 excel_read_range 等读取类工具的返回形态：
     *   { headers:[...], result:[[...],[...]] }  /  { headers, rows:[[...]] }  /  纯二维数组 [[...]]
     *   { success:true, data:[[...],[...]] }（excel_merge_sheets 等操作类结果，首行即表头）
     * 返回 null 表示不是行矩阵（调用方据此回落）。
     * 与 tryConvertToTable（统计/聚合/对象数组专用）互补，不改动其既有行为。
     */
    function tryRowsMatrixTable(data) {
        if (!data) return null;
        let headers = null;
        let rows = null;
        if (Array.isArray(data)) {
            if (data.length && Array.isArray(data[0])) {
                rows = data;
            } else if (data.length === 1 && data[0] && typeof data[0] === 'object') {
                // MCP 单元素包裹：[{ headers, result }]
                const o = data[0];
                if (Array.isArray(o.headers)) headers = o.headers.map(function (h) { return String(h); });
                if (Array.isArray(o.result)) rows = o.result;
                else if (Array.isArray(o.rows)) rows = o.rows;
            }
        } else if (typeof data === 'object') {
            // 兼容 {headers,result} / {success,data:{headers,result}} / {success,data:[[...]]} 三种包裹
            const cands = [data];
            if (data.data && typeof data.data === 'object' && !Array.isArray(data.data)) cands.push(data.data);
            for (let i = 0; i < cands.length && !rows; i++) {
                const o = cands[i];
                if (Array.isArray(o.headers)) headers = o.headers.map(function (h) { return String(h); });
                if (Array.isArray(o.result)) rows = o.result;
                else if (Array.isArray(o.rows)) rows = o.rows;
                // 操作类结果把数据矩阵放在 data/values 里（首行即表头，由下方统一拆分）；
                // previewRows：合并结果瘦身后（SplitMerge.vb 2026-10-09）只回 10 行预览
                else if (Array.isArray(o.data)) rows = o.data;
                else if (Array.isArray(o.previewRows)) rows = o.previewRows;
                else if (Array.isArray(o.values)) rows = o.values;
            }
        }
        if (!Array.isArray(rows) || rows.length === 0) return null;
        // 必须是二维（每行都是数组），否则不是行矩阵
        if (!rows.every(function (r) { return Array.isArray(r); })) return null;

        let body = rows;
        if (!headers) {
            headers = (rows[0] || []).map(function (v) { return String(v == null ? '' : v); });
            body = rows.slice(1);
        }
        // 列数：不用 Math.max(...spread)，避免大数组爆栈
        let colCount = headers.length;
        for (let i = 0; i < body.length; i++) {
            if (body[i].length > colCount) colCount = body[i].length;
        }
        if (colCount === 0) return null;

        const MAX = 100;
        const shown = body.slice(0, MAX);
        let html = '<table class="py-data-table py-range-table"><thead><tr>';
        for (let c = 0; c < colCount; c++) {
            html += '<th>' + escapeHtml(headers[c] != null ? String(headers[c]) : '') + '</th>';
        }
        html += '</tr></thead><tbody>';
        shown.forEach(function (r) {
            html += '<tr>';
            for (let c = 0; c < colCount; c++) {
                const v = r[c];
                html += '<td>' + escapeHtml(v == null ? '' : String(v)) + '</td>';
            }
            html += '</tr>';
        });
        html += '</tbody></table>';
        if (body.length > MAX) {
            html += '<div style="font-size:12px;color:#888;margin-top:4px;">共 ' + body.length + ' 行，仅显示前 ' + MAX + ' 行（下载/插入Excel 为全量）</div>';
        }
        // 全量数据岛（含被截断的行）：下载/复制/插入Excel 用它拿完整数据
        html += buildFullDataIsland(headers, body.map(function (r) { return (r || []).slice(); }));
        return html;
    }

    App.ChatV2.renderers.json = {
        renderJsonHtml,
        renderCollapsibleJson,
        tryConvertToTable,
        renderDescriptiveStatsTable,
        renderAggregationTable,
        renderCorrelationMatrix,
        renderAnomalyTable,
        renderDistributionTable,
        renderTrendTable,
        renderArrayTable,
        tryRowsMatrixTable
    };
    ChatLog.info?.('CHAT-V2', '✅ renderers/json 已加载');
})();