/*
 * chat-v2/services/tool-registry.util.js - Pure helper submodule of tool-registry
 * Extracted from tool-registry.js. Pure functions + const tables only (DOM-free'/no closure).
 * Delegates back via App.ChatV2.toolRegistryUtil.*
 */

(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    const CHART_TYPE_CN = {
        'xlPie': '饼图', 'xlDoughnut': '圆环图',
        'xlColumnClustered': '柱状图', 'xlColumnStacked': '堆叠柱状图',
        'xlColumnStacked100': '百分比柱状图',
        'xlBarClustered': '条形图', 'xlBarStacked': '堆叠条形图',
        'xlBarStacked100': '百分比条形图',
        'xlLine': '折线图', 'xlLineMarkers': '折线带点图',
        'xlArea': '面积图', 'xlXYScatter': '散点图',
        'xlRadar': '雷达图', 'xlRadarFilled': '填充雷达图',
        'xlBubble': '气泡图', 'xlSunburst': '旭日图',
        'xlTreemap': '树状图', 'xlHistogram': '直方图',
        'xlBoxwhisker': '箱线图', 'xlFunnel': '漏斗图',
        'xlWaterfall': '瀑布图'
    };

    const CHART_TYPE_ALIASES = {
        'pie': 'xlPie', 'doughnut': 'xlDoughnut', 'donut': 'xlDoughnut',
        'column': 'xlColumnClustered', 'columnclustered': 'xlColumnClustered',
        'columnstacked': 'xlColumnStacked', 'columnstacked100': 'xlColumnStacked100',
        'bar': 'xlBarClustered', 'barclustered': 'xlBarClustered',
        'barstacked': 'xlBarStacked', 'barstacked100': 'xlBarStacked100',
        'line': 'xlLine', 'linemarkers': 'xlLineMarkers',
        'area': 'xlArea', 'scatter': 'xlXYScatter', 'xyscatter': 'xlXYScatter',
        'radar': 'xlRadar', 'radarfilled': 'xlRadarFilled',
        'bubble': 'xlBubble', 'sunburst': 'xlSunburst',
        'treemap': 'xlTreemap', 'histogram': 'xlHistogram',
        'boxwhisker': 'xlBoxwhisker', 'boxplot': 'xlBoxwhisker',
        'funnel': 'xlFunnel', 'waterfall': 'xlWaterfall'
    };

    function resolveChartType(args) {
        const raw = args?.chartType || args?.chart_type || '';
        if (!raw) return 'xlColumnClustered';
        const cleaned = String(raw).trim().replace(/\s+/g, '');
        // 1. 直接枚举匹配（忽略大小写）
        const lower = cleaned.toLowerCase();
        if (CHART_TYPE_CN[raw]) return raw;
        for (const key of Object.keys(CHART_TYPE_CN)) {
            if (key.toLowerCase() === lower) return key;
        }
        // 2. 无前缀别名匹配
        if (CHART_TYPE_ALIASES[lower]) return CHART_TYPE_ALIASES[lower];
        // 3. 补 xl 前缀再匹配
        const prefixed = 'xl' + lower;
        for (const key of Object.keys(CHART_TYPE_CN)) {
            if (key.toLowerCase() === prefixed) return key;
        }
        // 4. 中文直接映射
        if (/饼/.test(raw)) return 'xlPie';
        if (/柱/.test(raw)) return 'xlColumnClustered';
        if (/条形/.test(raw)) return 'xlBarClustered';
        if (/折线/.test(raw)) return 'xlLine';
        if (/面积/.test(raw)) return 'xlArea';
        if (/散点/.test(raw)) return 'xlXYScatter';
        if (/雷达/.test(raw)) return 'xlRadar';
        if (/气泡/.test(raw)) return 'xlBubble';
        if (/漏斗/.test(raw)) return 'xlFunnel';
        if (/瀑布/.test(raw)) return 'xlWaterfall';
        if (/树状/.test(raw)) return 'xlTreemap';
        if (/旭日/.test(raw)) return 'xlSunburst';
        if (/直方/.test(raw)) return 'xlHistogram';
        if (/箱线/.test(raw)) return 'xlBoxwhisker';
        if (/圆环|环形/.test(raw)) return 'xlDoughnut';
        return 'xlColumnClustered';
    }

    function _presentSucceeded(result) {
        return !!result && (result.ok === true || result.status === 'completed');
    }

    function _backendReported(result, keyword) {
        if (!Array.isArray(result?.content)) return false;
        for (const c of result.content) {
            const t = (c && typeof c === 'object' && typeof c.text === 'string') ? c.text
                : (typeof c === 'string' ? c : '');
            if (typeof t === 'string' && keyword.test(t)) return true;
        }
        return false;
    }

    function normalizeName(rawName) {
        if (!rawName || typeof rawName !== 'string') return 'unnamed_tool';
        // 1. 替换非法字符为下划线
        let normalized = rawName.replace(/[^a-zA-Z0-9_-]/g, '_');
        // 2. 合并连续下划线
        normalized = normalized.replace(/_+/g, '_');
        // 3. 去除首尾下划线
        normalized = normalized.replace(/^_+|_+$/g, '');
        // 4. 如果为空，使用默认名
        if (!normalized) normalized = 'unnamed_tool';
        // 5. 如果不以字母开头，加 'fn_' 前缀
        if (!/^[a-zA-Z]/.test(normalized)) {
            normalized = 'fn_' + normalized;
        }
        return normalized;
    }

    function normalizeSheetArg(args) {
        if (!args || typeof args !== 'object') return;
        const s = args.sheet;
        if (s && typeof s === 'object') {
            const name = (typeof s.sheet === 'string' && s.sheet.trim()) ? s.sheet
                : (typeof s.name === 'string' && s.name.trim()) ? s.name
                : (typeof s.title === 'string' && s.title.trim()) ? s.title
                : '';
            args.sheet = name;
        }
    }

    function inferDimColumnFromMessage(ctx) {
        const msg = (ctx && typeof ctx.userMessage === 'string') ? ctx.userMessage : '';
        if (!msg) return null;
        // 匹配：按XX汇总、按XX分组、按XX统计、按XX求和、按XX平均、对XX汇总、以XX维度
        const m = msg.match(/(?:按|对|以|根据|按照|基于|依)[\s]?([^\s，。,\.!！?？、；;：:（）()（）()【】\[\]\{\}\d]+)(?:汇总|分组|统计|求和|平均|合计|分析|维度|分类|归类)/);
        if (m && m[1]) {
            const v = m[1].trim();
            if (v && v.length <= 20) return v;
        }
        // 宽松匹配：只要"按XX"后跟"汇总/统计/分组"
        const m2 = msg.match(/按[\s]?([^\s，。,\.!！?？、；;：:（）()（）()【】\[\]\{\}\d]{1,10})(?:汇总|统计|分组|分类|合计|分析)/);
        if (m2 && m2[1]) return m2[1].trim();
        return null;
    }

    function inferTimeGroupFromMessage(ctx) {
        const msg = (ctx && typeof ctx.userMessage === 'string') ? ctx.userMessage : '';
        if (!msg) return null;
        const lower = msg.toLowerCase();
        // 优先匹配显式时间粒度词
        if (/按\s*年|每年|年度|year/.test(lower)) return 'year';
        if (/按\s*季度|每季度|季度|quarter/.test(lower)) return 'quarter';
        if (/按\s*月|每月|月度|month/.test(lower)) return 'month';
        if (/按\s*周|每周|周度|week/.test(lower)) return 'week';
        if (/按\s*日|每日|日度|day/.test(lower)) return 'day';
        return null;
    }

    function _existingSheetNames(ctx) {
        const names = [];
        const turn = ctx && ctx.context ? ctx : (ctx || {});
        const src = (turn.context && turn.context.sheets) || turn.sheets;
        if (Array.isArray(src)) {
            for (const s of src) {
                const n = (typeof s === 'string') ? s
                    : (s && typeof s === 'object' && (s.name || s.title || s.Name));
                if (n) names.push(String(n).trim());
            }
        }
        const active = (turn.context && turn.context.activeSheet && (turn.context.activeSheet.sheet || turn.context.activeSheet.Name)) || '';
        if (active) names.push(String(active).trim());
        return names.filter(Boolean);
    }

    function extractSheetNameFromMessage(ctx) {
        const msg = (ctx && typeof ctx.userMessage === 'string') ? ctx.userMessage : '';
        if (!msg) return null;
        // 统一清洗：去包裹引号；过滤"一个/张/新/工作表"等泛词，避免把量词当表名
        const _clean = (v) => {
            if (!v) return null;
            const t = String(v)
                .replace(/^[“"'‘’「『【\s]+/, '')
                .replace(/[”"'’」』】\s]+$/, '')
                .trim();
            if (!t) return null;
            if (/^(一个|一张|一份|个|张|份|下|一下|新的|新|的|工作表|工作簿|sheet|页签|标签页)$/i.test(t)) return null;
            return t;
        };
        // 0) 工作表管理场景：与"工作表/sheet/页签"紧邻的引号名（前后两种语序）
        //    例：删除工作表"异常场景测试表"  /  删除"异常场景测试表"工作表
        //    动机：此前只认"名为/叫"标记，导致 `删除工作表"X"` 抽不出表名 →
        //    excel_delete_sheet 缺 name → 被"缺参数"护栏拦下，用户既没被执行也没被反问。
        let m = msg.match(/(?:工作表|sheet|页签|标签页)\s*[“"「『【]\s*([^”"」』】]{1,31})\s*[”"」』】]/i);
        if (m) { const v = _clean(m[1]); if (v) return v; }
        m = msg.match(/[“"「『【]\s*([^”"」』】]{1,31})\s*[”"」』】]\s*(?:工作表|sheet|页签|标签页)/i);
        if (m) { const v = _clean(m[1]); if (v) return v; }
        // 1) 显式命名（如 名为"销售数据"）
        m = msg.match(/(?:名为|命名为|名字(?:叫|为)|名称(?:叫|为)|叫|叫做)[“"'‘’「『]?([^"'“”‘’「』」》\s，。、；;()（）]+)[“"'”」』》]?/);
        if (m) { const v = _clean(m[1]); if (v) return v; }
        // 2) 明确的工作表管理动词场景下的引号名兜底（避免"列/字段"等非工作表引号名被误取）
        if (/(?:删除|删掉|移除|去掉|清掉|新建|创建|重命名|改名|复制|拷贝|克隆)[^，。；;!！？?\n]{0,8}?(?:工作表|sheet|页签|标签页)/i.test(msg)) {
            const mm = msg.match(/[“"「『【]\s*([^”"」』】]{1,31})\s*[”"」』】]/);
            if (mm) { const v = _clean(mm[1]); if (v) return v; }
        }
        return null;
    }

    function defaultNewSheetName(ctx) {
        // 用户消息里已显式给了目标表名 → 尊重用户意图，而非兜底生成"新建工作表N"
        const fromMsg = extractSheetNameFromMessage(ctx);
        if (fromMsg) return fromMsg;
        const existing = new Set(_existingSheetNames(ctx));
        let i = 1;
        while (existing.has('新建工作表' + i)) i++;
        return '新建工作表' + i;
    }

    function _hasValue(v) {
        return v !== undefined && v !== null && v !== '';
    }

    function _requiredParamKeys(tool) {
        const schema = tool && (tool.inputSchema || tool.schema);
        if (schema && typeof schema === 'object' && !Array.isArray(schema)) {
            return Array.isArray(schema.required) ? schema.required : [];
        }
        const arr = (schema && Array.isArray(schema)) ? schema
            : (Array.isArray(tool && tool.parameters) ? tool.parameters : []);
        return arr.filter(p => p && p.required).map(p => p.name).filter(Boolean);
    }

    function _paramLabel(tool, key) {
        const schema = tool && (tool.inputSchema || tool.schema);
        let desc = '';
        if (schema && typeof schema === 'object' && !Array.isArray(schema) && schema.properties) {
            desc = (schema.properties[key] && schema.properties[key].description) || '';
        } else {
            const arr = (schema && Array.isArray(schema)) ? schema
                : (Array.isArray(tool && tool.parameters) ? tool.parameters : []);
            const p = arr.find(x => x && x.name === key);
            desc = (p && p.description) || '';
        }
        const norm = String(desc || key);
        // 去掉过于代码化的描述尾巴，转为口语化标签
        const cleaned = norm.split(/[（(]/)[0].trim();
        return cleaned || key;
    }

    function resolveClearIntentFromMessage(turn) {
        const msg = (turn && typeof turn.userMessage === 'string') ? turn.userMessage : '';
        if (!msg) return null;
        const hasClearVerb = /清空|清除|去掉|删除|移除|取消|清理|清掉|清除干净/.test(msg);
        const hasFormat = /格式/.test(msg);
        const hasContent = /内容|数据|数值|值|文字/.test(msg);
        // 入口放宽松：只要消息里带清/删/去/移等动作词即进入判定（该函数仅在"清空类工具且规划者未传参"时被调用，
        // 放宽是为了让"清格式保留数据"这类省略标准动词的表述也能命中数据保护分支）。
        if (!/清|删|去|移/.test(msg)) return null;
        // 明确"保留内容/数据/值"但涉及格式 → 绝不清内容，只清格式（数据保护优先）
        const keepContent = /保留.{0,4}(数据|内容|数值|值)|不动.{0,3}(数据|内容)|不清.{0,3}内容/.test(msg);
        const keepFormat = /保留格式|不动格式|不清格式/.test(msg);
        if (keepContent && hasFormat) return { clearContent: false, clearFormat: true };
        if (keepFormat && hasContent) return { clearContent: true, clearFormat: false };
        // 只谈"格式"，未提内容/数据 → 仅清格式（防止"清空格式"误删单元格内容）
        if (hasFormat && !hasContent) return { clearContent: false, clearFormat: true };
        // 只谈"内容/数据/数值"，未提格式 → 仅清内容（保留格式样式）
        if (hasContent && !hasFormat && /清空内容|清空数据|清空数值|删除内容|清除数据/.test(msg)) {
            return { clearContent: true, clearFormat: false };
        }
        return null;
    }

    const INTENT_PARAM_KEYS = {
        'excel_clear_range': ['clearContent', 'clearFormat'],
        'excel_clear_sheet': ['clearContent', 'clearFormat'],
        'excel_close_workbook': ['saveChanges'],
        'excel_autofit': ['type'],
        'excel_apply_conditional_format': ['formatType'],
        'excel_merge_cells': ['mergeType'],
        'excel_split_by_rows': ['outputMode'],
        'excel_unmerge_cells': ['fillAfterUnmerge'],
        'excel_sort_range': ['sortBy', 'hasHeader'],
        'excel_date_group': ['groupBy', 'weekStart', 'fiscalYearStart', 'fiscalQuarter']
    };

    /**
     * 从「按销售额从高到低排序」这类说法里抽出表头列名。
     * 列字母 / 第 N 列 / 颜色·图标·方向等排序选项不算列名，返回 null。
     * 后端 excel_sort_range 会拿这个名字去对表头（精确匹配）。
     */
    function extractExplicitSortColumnName(msg) {
        if (!msg || typeof msg !== 'string') return null;
        const re = /(?:按照|依照|根据|按|把|将)\s*[「『"“']?\s*([^「」『』"“”'\s，,。；;、]{1,20}?)\s*[」』"”']?\s*(?:这一?列|列|栏|字段)?\s*(?:从高到低|从低到高|从大到小|从小到大|由高到低|由低到高|由大到小|由小到大|降序|升序|倒序|正序|排序|排列)/gu;
        const reject = /^(?:降序|升序|倒序|正序|从|由|行|列|单元格|字体|文字|字|图标|条件格式|笔划|笔画|拼音|字母|大小写|红|绿|蓝|黄|橙|紫|青|黑|白|灰|红色|绿色|蓝色|黄色|橙色|紫色|青色|黑色|白色|灰色)$/;
        let m;
        while ((m = re.exec(msg)) !== null) {
            let name = String(m[1] || '').trim()
                .replace(/(?:这一?列|列|栏|字段)$/u, '')
                .replace(/(?:从|由)?(?:高到低|低到高|大到小|小到大)$/u, '')
                .replace(/(?:降序|升序|倒序|正序)$/u, '')
                .trim();
            if (!name || reject.test(name)) continue;
            if (/颜色|色块|底色|背景|图标/.test(name)) continue;
            // 字母 / 序号交给原有「第N列 / A列」解析，避免把 A 当成表头名
            if (/^[A-Za-z]{1,3}$/.test(name) || /^\d+$/.test(name)) continue;
            if (/^第?(?:\d+|[一二三四五六七八九十]+)$/.test(name)) continue;
            return name;
        }
        return null;
    }

    /**
     * 用户点了列名（如「按销售额从高到低」）时，把 sortBy.column 改成该列名。
     * 弱模型常照抄工具示例，把 column 写成 A，升降序却是对的——结果整表按第一列排。
     * 只在列名与当前不一致时改写；次序仅在用户明确说了升/降序时才覆盖。
     * @returns {boolean} 是否改写了 args.sortBy
     */
    function applyExplicitSortColumn(args, turn) {
        if (!args) return false;
        const msg = (turn && typeof turn.userMessage === 'string') ? turn.userMessage
            : (typeof turn === 'string' ? turn : '');
        const named = extractExplicitSortColumnName(msg);
        if (!named) return false;
        const orderDesc = /降序|从大到小|大到小|高到低|由高到低|由大到小|descending|\bdesc\b/iu.test(msg);
        const orderAsc = /升序|从小到大|小到大|低到高|由低到高|由小到大|ascending|\basc\b/iu.test(msg);
        const order = orderDesc ? 'descending' : (orderAsc ? 'ascending' : null);

        let parsed = null;
        const raw = args.sortBy;
        if (typeof raw === 'string' && raw.trim()) {
            try { parsed = JSON.parse(raw); } catch (_) { parsed = null; }
        } else if (Array.isArray(raw)) {
            parsed = raw.map(x => (x && typeof x === 'object') ? { ...x } : x);
        }
        if (!Array.isArray(parsed) || !parsed.length || !parsed[0] || typeof parsed[0] !== 'object') {
            parsed = [{ column: named, order: order || 'ascending' }];
        } else {
            const cur = String(parsed[0].column || '').trim();
            const sameCol = cur.toLowerCase() === named.toLowerCase();
            const sameOrder = !order || String(parsed[0].order || '').toLowerCase() === order;
            if (sameCol && sameOrder) return false;
            parsed[0].column = named;
            if (order) parsed[0].order = order;
        }
        args.sortBy = JSON.stringify(parsed);
        return true;
    }

    function resolveIntentFlagsFromMessage(turn, name) {
        const msg = (turn && typeof turn.userMessage === 'string') ? turn.userMessage : '';
        if (!msg) return null;
        switch (name) {
            case 'excel_clear_range':
            case 'excel_clear_sheet':
                return resolveClearIntentFromMessage(turn);
            case 'excel_close_workbook': {
                // 默认后端 saveChanges=false 会静默丢弃未保存改动；保守保护数据。
                const noSave = /不保存|不想保存|不要保存|放弃修改|不存盘|不存储/.test(msg);
                if (noSave) return { saveChanges: false };
                if (/保存|存储|存盘/.test(msg) || /关闭/.test(msg)) return { saveChanges: true };
                return null;
            }
            case 'excel_autofit': {
                const wantsRow = /行高|按行|行.*适应|调整行/.test(msg) && !/列/.test(msg);
                const wantsCol = /列宽|按列|列.*适应|调整列/.test(msg) && !/行/.test(msg);
                if (wantsRow) return { type: 'row' };
                if (wantsCol) return { type: 'column' };
                return null;
            }
            case 'excel_apply_conditional_format': {
                if (/数据条|数据栏|\bbar\b/iu.test(msg)) return { formatType: 'dataBar' };
                if (/色阶|三色|渐变色|颜色渐变|colorScale/iu.test(msg)) return { formatType: 'colorScale' };
                if (/图标集|iconSet/iu.test(msg)) return { formatType: 'iconSet' };
                if (/突出|高亮|highlight|大于|小于|介于|等于|文本包含|重复/iu.test(msg)) return { formatType: 'highlight' };
                return null;
            }
            case 'excel_merge_cells': {
                if (/合并后居中|合并居中|居中合并|合并并居中|合并.*居中/.test(msg)) return { mergeType: 'mergeAndCenter' };
                return null;
            }
            case 'excel_split_by_rows': {
                if (/拆分成文件|拆分.*文件|另存.*多个文件|保存为多个文件|每个.*文件|拆到文件|输出.*文件|存成文件/iu.test(msg)) return { outputMode: 'files' };
                if (/拆分.*工作表|拆到工作表|新的工作表|新表|sheet|输出.*sheet|拆成.*sheet/iu.test(msg)) return { outputMode: 'sheets' };
                return null;
            }
            case 'excel_unmerge_cells': {
                if (/取消合并.*填充|取消合并.*铺满|取消合并.*展开|取消合并.*复制|拆开.*填充|拆分.*填充|展开.*值/iu.test(msg)) return { fillAfterUnmerge: true };
                return null;
            }
            case 'excel_sort_range': {
                // 中文数字→列号：一..九、十、十三、二十六 等，支持 1..26（A..Z）
                const CH_U = { 一:1, 二:2, 三:3, 四:4, 五:5, 六:6, 七:7, 八:8, 九:9 };
                const cnToColNumber = (cn) => {
                    if (cn.startsWith('二十')) return 20 + ((cn.length > 2 && CH_U[cn[cn.length - 1]]) || 0);
                    if (cn.includes('十')) {
                        const i = cn.indexOf('十');
                        const tens = (i === 0 ? 1 : CH_U[cn[0]]) * 10;
                        const unit = (i < cn.length - 1 && CH_U[cn[cn.length - 1]]) || 0;
                        return tens + unit;
                    }
                    return CH_U[cn] || 0;
                };
                const out = {};
                const orderDesc = /降序|从大到小|大到小|高到低|descending|\bdesc\b/iu.test(msg);
                const orderAsc = /升序|从小到大|小到大|低到高|ascending|\basc\b/iu.test(msg);
                const order = orderDesc ? 'descending' : (orderAsc ? 'ascending' : null);
                const colMatch = /(?:第\s*|按\s*|排序列?\s*(?:为|是)?\s*)([一二三四五六七八九十]+|[A-Za-z]+|\d+)/iu.exec(msg);
                let column = null;
                if (colMatch) {
                    const g = colMatch[1];
                    let n = 0;
                    if (/^\d+$/.test(g)) n = parseInt(g, 10);
                    else if (/^[A-Za-z]+$/.test(g)) { column = g.toUpperCase(); }
                    else if (/^[一二三四五六七八九十]+$/.test(g)) n = cnToColNumber(g);
                    if (n >= 1 && n <= 26) column = String.fromCharCode(64 + n);
                }
                // 「按销售额从高到低」：销售额不是列字母。默认成 A 会把整表按第一列排
                // （实测：用户要按销售额降序，卡片上却是 column=A）。表头名交给后端对表头解析。
                if (!column) {
                    const namedCol = extractExplicitSortColumnName(msg);
                    if (namedCol) column = namedCol;
                }
                // ── 排序依据（还原原生对话框）：单元格颜色 / 字体颜色 / 条件格式图标 ──
                // 例："把红色的行排前面" "按单元格颜色排序，绿色在前" "按字体颜色排"
                const CN_COLOR_HEX = { 红:'#FF0000', 绿:'#008000', 蓝:'#0000FF', 黄:'#FFFF00', 橙:'#FFA500', 紫:'#800080', 青:'#008080', 黑:'#000000', 白:'#FFFFFF', 灰:'#808080' };
                const hexMatch = /#[0-9a-fA-F]{6}|\b\d{1,3}\s*,\s*\d{1,3}\s*,\s*\d{1,3}\b/iu.exec(msg);
                const cnColorMatch = /(红|绿|蓝|黄|橙|紫|青|黑|白|灰)色/iu.exec(msg);
                const colorValue = hexMatch ? hexMatch[0].replace(/\s+/g, '') : (cnColorMatch ? CN_COLOR_HEX[cnColorMatch[1]] : null);
                const mentionsFontColor = /字体颜色|文字颜色|字颜色/.test(msg);
                const mentionsCellColor = /单元格颜色|填充颜色?|背景色?|底色|色块/.test(msg);
                // "把红色的排前面"这类：有颜色词 + 排序语境，也算按单元格颜色
                const mentionsColorSort = mentionsCellColor || mentionsFontColor ||
                    (!!cnColorMatch && /排|排序|置顶|靠前|在前|上面/.test(msg)) || !!hexMatch;
                let sortOn = null;
                if (mentionsFontColor) sortOn = 'fontColor';
                else if (mentionsColorSort) sortOn = 'cellColor';
                else if (/图标|条件格式/.test(msg) && /排|排序/.test(msg)) sortOn = 'icon';

                const key = { column: column || 'A', order: order || 'ascending' };
                if (sortOn) key.sortOn = sortOn;
                if ((sortOn === 'cellColor' || sortOn === 'fontColor') && colorValue) key.color = colorValue;
                // 颜色/图标排序必须带参数才有意义；仅当确实解析到依据时才覆盖 sortBy
                if (order || column || sortOn) {
                    out.sortBy = JSON.stringify([key]);
                }
                // ── 原生"选项"对话框 ──
                if (/区分大小写/.test(msg)) out.matchCase = true;
                if (/按行排序|行方向|横向排序|按行排/.test(msg)) out.orientation = 'rows';
                else if (/按列排序|列方向|纵向排序|按列排/.test(msg)) out.orientation = 'columns';
                if (/笔划|笔画/.test(msg)) out.sortMethod = 'stroke';
                else if (/拼音|字母顺序/.test(msg)) out.sortMethod = 'pinYin';
                if (/无表头|没有表头|不含表头|无标题|没有标题|标题行不算/.test(msg)) out.hasHeader = false;
                else if (/有表头|保留表头|含表头|有标题|包含标题|含标题行/.test(msg)) out.hasHeader = true;
                return Object.keys(out).length ? out : null;
            }
            case 'excel_date_group': {
                const out = {};
                const lowerMsg = msg.toLowerCase();
                // 分组粒度
                if (/按\s*年|每年|年度|year/.test(lowerMsg)) out.groupBy = 'year';
                else if (/按\s*季度|每季度|季度|quarter/.test(lowerMsg)) out.groupBy = 'quarter';
                else if (/按\s*月|每月|月度|month/.test(lowerMsg)) out.groupBy = 'month';
                else if (/按\s*周|每周|周度|week/.test(lowerMsg)) out.groupBy = 'week';
                else if (/按\s*日|每日|日度|day/.test(lowerMsg)) out.groupBy = 'day';
                // 周起始日
                if (out.groupBy === 'week') {
                    if (/周日|星期天|sunday/.test(lowerMsg)) out.weekStart = 'sunday';
                    else out.weekStart = 'monday';
                }
                // 财年
                const fyMatch = /财年.*从\s*(\d{1,2})\s*月|财年起始\s*(\d{1,2})|fiscal.*start\s*(\d{1,2})/i.exec(msg);
                if (fyMatch) {
                    const m = parseInt(fyMatch[1] || fyMatch[2] || fyMatch[3], 10);
                    if (m >= 1 && m <= 12) {
                        out.fiscalYearStart = m;
                        out.fiscalQuarter = true;
                    }
                }
                if (/自然年|自然季度|calendar/.test(lowerMsg)) out.fiscalQuarter = false;
                return Object.keys(out).length ? out : null;
            }
            default:
                return null;
        }
    }

    function pickArgsKeys(obj, keys) {
        for (const k of keys) {
            if (obj && obj[k] !== undefined && obj[k] !== null) return obj[k];
        }
        return undefined;
    }

    function toBool(v) {
        if (v === true || v === false) return v;
        return v === 'true' || v === 1 || v === '1' || v === true;
    }

    function parseRangeAddr(addr) {
        if (typeof addr !== 'string') return null;
        const m = addr.match(/(?:(.+?)!)?\$?([A-Za-z]+)\$?(\d+):\$?([A-Za-z]+)\$?(\d+)/);
        if (!m) return null;
        const colNum = (c) => {
            let n = 0;
            for (const ch of c.toUpperCase()) { n = n * 26 + (ch.charCodeAt(0) - 64); }
            return n;
        };
        const c1 = colNum(m[2]), c2 = colNum(m[4]);
        return { sheet: m[1], col1: Math.min(c1, c2), col2: Math.max(c1, c2), row1: Math.min(parseInt(m[3], 10), parseInt(m[5], 10)), row2: Math.max(parseInt(m[3], 10), parseInt(m[5], 10)) };
    }

    function combineSourceRanges(categorySource, dataSource) {
        const cat = parseRangeAddr(categorySource);
        const data = parseRangeAddr(dataSource);
        if (!cat || !data) return dataSource || categorySource || '';
        // 要求同一张表（都无前缀=当前表）
        if ((cat.sheet || '') !== (data.sheet || '')) return dataSource || '';
        // 仅当标签列与数值列相邻（连续矩形）时才合并，
        // 否则中间多出的列会被 Excel 误当作额外数据系列（如饼图多扇区）。
        const gap = Math.abs(cat.col2 - data.col1) <= 1 || Math.abs(data.col2 - cat.col1) <= 1;
        if (!gap) return dataSource || '';
        const col1 = Math.min(cat.col1, data.col1);
        const col2 = Math.max(cat.col2, data.col2);
        const row1 = Math.min(cat.row1, data.row1);
        const row2 = Math.max(cat.row2, data.row2);
        const colLetter = (n) => {
            let s = '';
            while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
            return s;
        };
        const sheetPrefix = cat.sheet ? `${cat.sheet}!` : '';
        return `${sheetPrefix}$${colLetter(col1)}$${row1}:$${colLetter(col2)}$${row2}`;
    }

    /** @type {Object.<string, string[]>|null} name(lower) → intentKeywords */
    let _excelToolsKeywordIndex = null;

    /**
     * 从 Config/excel_tools.json（及运行时 window._mcpTools）构建工具名→intentKeywords 索引。
     * Node 测试靠 process.cwd()/Config；浏览器靠已加载的 MCP 元数据。
     */
    function _ingestToolsIntoKeywordIndex(idx, tools) {
        if (!Array.isArray(tools)) return;
        for (const t of tools) {
            if (!t || !t.name) continue;
            const n = String(t.name).replace(/^[a-z]+-local__/i, '').toLowerCase();
            const kws = t.intentKeywords || (t.meta && t.meta.intentKeywords);
            if (n && Array.isArray(kws) && kws.length) idx[n] = kws.slice();
        }
    }

    function ensureExcelToolsKeywordIndex() {
        if (_excelToolsKeywordIndex) return _excelToolsKeywordIndex;
        const idx = Object.create(null);
        try {
            const bundled = (typeof globalThis !== 'undefined' && globalThis.__SMARTPILOT_EXCEL_TOOLS__)
                || (typeof window !== 'undefined' && window.__SMARTPILOT_EXCEL_TOOLS__);
            if (bundled && Array.isArray(bundled.tools)) _ingestToolsIntoKeywordIndex(idx, bundled.tools);
        } catch (_) { /* ignore */ }
        try {
            const mcp = (typeof window !== 'undefined' && Array.isArray(window._mcpTools)) ? window._mcpTools : [];
            _ingestToolsIntoKeywordIndex(idx, mcp);
        } catch (_) { /* ignore */ }
        _excelToolsKeywordIndex = idx;
        return idx;
    }

    function resolveToolIntentKeywords(tool, kwIndex) {
        if (Array.isArray(tool && tool.intentKeywords) && tool.intentKeywords.length) {
            return tool.intentKeywords;
        }
        const name = String((tool && (tool.name || tool.originalName)) || '')
            .replace(/^[a-z]+-local__/i, '').toLowerCase();
        const idx = kwIndex || ensureExcelToolsKeywordIndex();
        return (name && idx[name]) ? idx[name] : [];
    }

    /**
     * 【工具注入瘦身】按用户消息意图预筛 MCP 工具，只保留"核心基本盘 + 意图命中的工具族"。
     *
     * 背景：全量 MCP 工具（实测 123 个）的 schema 注入 prompt 会占满 token 预算
     * （日志实测 tool tokens 7391 > message budget 7355），小模型（如 qwen2.5:7b）因此
     * 看不到完整上下文 → 不调用工具 / 乱调工具 → 被迫走强匹配兜底（fakeExecution）。
     * LLM 可见工具越少，选型越准；同时把 prompt token 预算让给对话上下文。
     *
     * 边界（只影响 LLM 可见层，执行层不受影响）：
     * - toolRegistry 注册表全量保留，invokeTracked / 参数补全仍按全量工具执行；
     * - 本函数只裁剪"进入 prompt / 原生 tools 下发"的列表；
     * - 消息为空或过短时保守保留全量（无法判断意图时不冒丢失关键工具的风险）。
     * - 意图词唯一权威源：各工具 intentKeywords（excel_tools.json）；复合任务保留少量族扩展。
     *
     * @param {Array}  tools   - 全量（或已剔除图片类）MCP 工具列表
     * @param {string} message - 用户消息
     * @param {Object} [opts]  - { keepAllIfEmpty: boolean, preferTools: string[] }
     *                           keepAllIfEmpty：消息过短时是否保留全量（默认 true）
     *                           preferTools：无条件保留的工具名（如上一轮真实执行过的工具），
     *                                        用于跟进型消息（"不够美观"/"继续"）不丢当前工具体系
     * @returns {Array} 瘦身后的工具列表（保持原顺序；极端情况回退全量）
     */
    function filterToolsForMessage(tools, message, opts) {
        if (!Array.isArray(tools) || tools.length === 0) return tools || [];

        const msg = String(message || '').trim().toLowerCase();
        const keepAllIfEmpty = !opts || opts.keepAllIfEmpty !== false;
        // 消息为空/过短：无法判断意图，保守保留全量
        if (keepAllIfEmpty && msg.length < 2) return tools;

        // 核心基本盘：分析/读写/汇总类请求的基础工具，始终保留。
        // 选择原则：泛请求"读取/分析/总结"的最小可用集；不含大而全的"所有常用工具"，
        // 其余工具由下方意图规则按消息动态追加。
        const CORE_TOOLS = [
            'excel_read_range',
            'excel_get_sheets',
            'excel_get_workbook_info',
            'excel_get_used_range',
            'excel_search',
            'excel_write_range',
            'excel_aggregate',
            'excel_execute_vba',
            'excel_clean_data',
            'excel_sort_range',
            'excel_apply_filter',
            'excel_add_formula'
        ];

        // 【词表归一】意图匹配唯一权威源 = 各工具的 intentKeywords（excel_tools.json）。
        // 测试里常只传 { name }：此时从 Config/excel_tools.json 目录索引补齐关键词。
        // 复合任务（看板）仍需少量族扩展：单工具关键词不足以解锁整条工具链。
        const COMPOSITE_FAMILIES = [
            { words: ['看板', '仪表盘', '驾驶舱', 'dashboard', '大屏', '数据面板'],
              toolKeys: ['generate_dashboard', 'create_chart', 'add_shape', 'get_layout_snapshot',
                         'create_sheet', 'aggregate', 'date_group', 'write_range', 'fill_data',
                         'format_range', 'conditional_format', 'pivot_table', 'sort_range',
                         'set_cell_value', 'merge_cells', 'autofit', 'save_workbook',
                         'clear_sheet', 'clear_range', 'create_slicer'] },
            // 美化诉求通常要 format + autofit + 条件格式一起可见（单工具关键词不够）
            { words: ['美化', '美观', '好看', '不好看', '漂亮', '太丑', 'beautify'],
              toolKeys: ['format_range', 'conditional_format', 'autofit', 'clear_format'] },
            { words: ['ppt', '幻灯片', '演示文稿', '做成ppt', '做ppt', '生成ppt', '制作ppt',
                      '做一份ppt', '汇报材料', '方案ppt', '做成幻灯片', '转成ppt', '演示'],
              toolKeys: ['ppt_'] }
        ];

        // 【上一轮工具粘性】跟进型消息（"不够美观"/"继续"/"再来一次"/"太丑了"）往往一个意图词都不含，
        // 若只按当前这一句裁剪，用户正在使用的工具体系（如格式化族）会被整体剔除，
        // 模型拿不到工具，只能回一段"接下来我将使用 excel_format_range 工具进行美化处理"的文字计划
        // （实测失败：UI 显示"任务完成 · 未调用工具"，Excel 毫无变化）。
        // 因此把最近几轮真实执行过的工具无条件保留——用户正在操作什么，就让他继续能操作。
        const preferSet = new Set(
            ((opts && Array.isArray(opts.preferTools)) ? opts.preferTools : [])
                .map(n => String(n || '').toLowerCase())
                .filter(Boolean));

        // 「并设置成中国红 / 统一为科技绿」不在 excel_format_range 的 intentKeywords 里
        // （词表只有「颜色/底色/格式化」）。瘦身会把格式化工具裁掉，模型只能口头承诺下一步。
        // 点名外观，或点了商务蓝/深蓝这种预设名，都要让格式化工具留在清单里。
        let needFormatTools = false;
        try {
            const styleFn = (typeof App !== 'undefined' && App.ChatV2 && App.ChatV2.turnGoal
                && App.ChatV2.turnGoal.semanticStyleTarget);
            if (typeof styleFn === 'function' && styleFn(message)) needFormatTools = true;
        } catch (_) { /* turn-goal 未加载时走下面的词 */ }
        if (!needFormatTools) {
            needFormatTools = /(?:统一(?:为|成|到|用)|改成|改为|换成|换为|调成|调为|设置成|设为|设成|变成|涂成|刷成|转为|转成)|中国红|商务蓝|深蓝|科技绿|底色|背景色|填充色/.test(String(message || ''));
        }

        const kwIndex = ensureExcelToolsKeywordIndex();

        const keptNames = new Set();
        for (const t of tools) {
            const name = String(t.name || t.originalName || '').toLowerCase();
            if (CORE_TOOLS.includes(name)) { keptNames.add(name); continue; }
            if (preferSet.has(name)) { keptNames.add(name); continue; }
            const kws = resolveToolIntentKeywords(t, kwIndex);
            if (kws.some(w => w && msg.includes(String(w).toLowerCase()))) {
                keptNames.add(name);
                continue;
            }
            for (const rule of COMPOSITE_FAMILIES) {
                if (!rule.words.some(w => msg.includes(w))) continue;
                if (rule.toolKeys.some(k => name.includes(k))) { keptNames.add(name); break; }
            }
            if (needFormatTools && /format_range|conditional_format|autofit|clear_format/.test(name)) {
                keptNames.add(name);
            }
        }

        const kept = tools.filter(t => {
            const name = String(t.name || t.originalName || '').toLowerCase();
            return keptNames.has(name);
        });
        // 极端安全：瘦身后不应为空（核心集保证非空）；若为空则回退全量
        return kept.length > 0 ? kept : tools;
    }

    function normalizeChartTypeValueForTool(val) {
        if (!val) return val;
        const v = String(val).trim().replace(/\s+/g, '');
        const lower = v.toLowerCase();
        if (/^xl/.test(lower)) return v; // 已是枚举，原样透传
        const map = {
            'pie': 'xlPie', '饼图': 'xlPie', '扇形图': 'xlPie', '占比图': 'xlPie', '占比': 'xlPie',
            'doughnut': 'xlDoughnut', 'donut': 'xlDoughnut', '圆环图': 'xlDoughnut', '环形图': 'xlDoughnut',
            'column': 'xlColumnClustered', 'columnclustered': 'xlColumnClustered', '柱状图': 'xlColumnClustered',
            '柱形图': 'xlColumnClustered', '分组柱': 'xlColumnClustered', '对比图': 'xlColumnClustered',
            'columnstacked': 'xlColumnStacked', 'stackedcolumn': 'xlColumnStacked', '堆叠柱': 'xlColumnStacked',
            'bar': 'xlBarClustered', 'barclustered': 'xlBarClustered', '条形图': 'xlBarClustered',
            '横向柱状图': 'xlBarClustered', '横向柱': 'xlBarClustered',
            'line': 'xlLine', '折线图': 'xlLine', 'linemarkers': 'xlLineMarkers',
            '趋势图': 'xlLineMarkers', 'area': 'xlArea', '面积图': 'xlArea',
            'scatter': 'xlXYScatter', '散点图': 'xlXYScatter', 'radar': 'xlRadar', '雷达图': 'xlRadar',
            'bubble': 'xlBubble', '气泡图': 'xlBubble', 'histogram': 'xlHistogram', '直方图': 'xlHistogram',
            'funnel': 'xlFunnel', '漏斗图': 'xlFunnel', 'waterfall': 'xlWaterfall', '瀑布图': 'xlWaterfall',
            'treemap': 'xlTreemap', '树状图': 'xlTreemap', 'sunburst': 'xlSunburst', '旭日图': 'xlSunburst',
            // 【实测】弱模型会说"combo/组合图"（柱+线组合），后端不支持该类型会直接报
            // "不支持的图表类型：combo"（2026-10-06 日志）。后端无真正的组合图能力，
            // 降级映射为簇状柱形图，避免整轮失败重试。
            'combo': 'xlColumnClustered', 'combination': 'xlColumnClustered', '组合图': 'xlColumnClustered'
        };
        return map[lower] || v;
    }

    function normalizeChartArgs(raw) {
        const a = { ...(raw || {}) };
        // 图表类型
        const ct = pickArgsKeys(a, ['chartType', 'chart_type']);
        if (ct !== undefined) {
            a.chartType = normalizeChartTypeValueForTool(ct);
        }
        delete a.chart_type;
        // 数据范围：优先合并 category_source + data_source
        // 兼容"域空壳 / 技能形态"传入的 selection（对象含 address/range，或选区地址字符串）→ dataRange：
        // skill-pipeline 走域空壳时 args 是 {prompt,text,message,excelData,selection}，不映射的话
        // 归一后的 excel_create_chart 会因缺 dataRange 被必填校验拦下。
        const _sel = a.selection;
        const _selAddr = _sel && (typeof _sel === 'object' ? (_sel.address || _sel.range) : String(_sel));
        // dataAddress/data_address：prompt 引导词曾误用该名（已改为 dataRange），此处保留别名兜底
        const ds = pickArgsKeys(a, ['dataRange', 'data_range', 'dataAddress', 'data_address', 'data_source', 'source', 'range']) || _selAddr;
        const cs = pickArgsKeys(a, ['category_source', 'categoryRange', 'category_range', 'labels']);
        if (cs && ds && ds !== cs) {
            a.dataRange = combineSourceRanges(cs, ds);
        } else if (ds) {
            a.dataRange = ds;
        }
        delete a.data_source; delete a.data_range; delete a.dataAddress; delete a.data_address; delete a.source; delete a.range;
        delete a.category_source; delete a.categoryRange; delete a.category_range; delete a.labels;
        // 技能形态遗留字段：删除，避免污染后端参数校验
        delete a.selection; delete a.prompt; delete a.text; delete a.message; delete a.excelData;
        delete a.referenceRegions; delete a.attachments;
        // 图例开关（后端用 hasLegend）
        const legend = pickArgsKeys(a, ['hasLegend', 'showLegend', 'include_legend', 'includeLegend']);
        if (legend !== undefined) a.hasLegend = toBool(legend);
        delete a.include_legend; delete a.includeLegend;
        return a;
    }

    /**
     * excel_add_shape 参数归一化：LLM 可能输出 shape_type/fill_color/border/radius/send_to_back/lines，
     * 后端只认 shapeType/fill/line/cornerRadius/sendToBack。看板渲染器直传后端形态，此处仅兜底转换。
     */
    function normalizeShapeArgs(raw) {
        const a = { ...(raw || {}) };
        const st = pickArgsKeys(a, ['shapeType', 'shape_type']);
        if (st !== undefined) a.shapeType = st;
        delete a.shape_type;
        const fill = pickArgsKeys(a, ['fill', 'fill_color', 'fillColor', 'bgColor', 'backgroundColor']);
        if (fill !== undefined) a.fill = fill;
        delete a.fill_color; delete a.fillColor; delete a.bgColor; delete a.backgroundColor;
        if (a.shapeType === 'roundedRectangle' || a.shapeType === 'rounded_rectangle') a.shapeType = 'roundedRect';
        const cr = pickArgsKeys(a, ['cornerRadius', 'corner_radius', 'radius']);
        if (cr !== undefined) a.cornerRadius = cr;
        delete a.corner_radius; delete a.radius;
        const stb = pickArgsKeys(a, ['sendToBack', 'send_to_back']);
        if (stb !== undefined) a.sendToBack = toBool(stb);
        delete a.send_to_back;
        if (a.lines && typeof a.lines !== 'string') a.lines = JSON.stringify(a.lines);
        return a;
    }

    /**
     * excel_export_csv 参数归一化：LLM 常输出 file_path/savePath/save_path/path，
     * 后端只认 filePath。归一化后必填参数校验才能命中，避免"缺少必要参数"拦截。
     */
    function normalizeExportCsvArgs(raw) {
        const a = { ...(raw || {}) };
        const fp = pickArgsKeys(a, ['filePath', 'file_path', 'savePath', 'save_path', 'path']);
        if (fp !== undefined) a.filePath = fp;
        delete a.file_path; delete a.savePath; delete a.save_path; delete a.path;
        return a;
    }

    /**
     * excel_export_csv.filePath 安全默认：用户说"保存到桌面/下载"或未指定时，
     * 自动落盘到桌面（默认）/下载目录，文件名 = 工作表名_时间戳.csv。
     * 无桥接（如纯 Web 预览）时返回 null，交由"需用户补充"兜底，绝不编造路径。
     */
    async function inferExportFilePath(ctxTurn) {
        const bridge = window.BridgeAccessor && window.BridgeAccessor.getBridge() || (App && App.excelBridge);
        const turnMsg = (ctxTurn && (ctxTurn.userMessage || (ctxTurn.context && ctxTurn.context.userMessage))) || '';
        const userText = String(turnMsg || '');
        let dir = '';
        try {
            if (bridge) {
                if (/下载|downloads?/i.test(userText) && typeof bridge.GetDownloadsPath === 'function') {
                    dir = dirFromBridgeResult(await bridge.GetDownloadsPath());
                } else if (typeof bridge.GetDesktopPath === 'function') {
                    // 【修复】必须过 dirFromBridgeResult：COM 往返后返回的是 `C:/Users/...`，
                    // 且宿主可能返回 `{"success":false,...}` 错误 JSON（PPT/Word 未 override 时）。
                    // 直接赋给 dir 会产出 `C://Users//...` 非法路径（实测 2026-10-05 23:45）。
                    dir = dirFromBridgeResult(await bridge.GetDesktopPath());
                }
            }
        } catch (_e) { dir = ''; }
        if (!dir) return null;
        const ctx = (ctxTurn && ctxTurn.context) || {};
        const sheet = (ctx.activeSheet && ctx.activeSheet.sheet)
            || (ctx.selection && ctx.selection.sheetName)
            || 'export';
        const safeSheet = String(sheet).replace(/[\\/:*?"<>|]/g, '_');
        const d = new Date();
        const pad = (n) => String(n).padStart(2, '0');
        const ts = `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
        // 【修复】不再按 dir 里有没有 `\` 猜分隔符——COM 往返后一定是 `/`，猜错会产出
        // `C://Users//...` 这种非法路径（实测 2026-10-05 23:45）。统一用反斜杠拼。
        return joinWinPath(dir, `${safeSheet}_${ts}.csv`);
    }

    /**
     * 从"已知属于当前用户的路径"反推桌面目录：`X:\Users\<当前用户>\...` → `X:\Users\<当前用户>\Desktop`。
     * 动机（实测 2026-10-05 23:20）：PptBridge / WordBridge **没有 override GetDesktopPath**，
     * 调到的基类实现返回 `{"success":false,"error":"operation not supported on this host"}`，
     * 于是"直用当前电脑用户桌面"这条路在 PPT/Word 宿主里拿不到目录。
     * 而这些宿主一定能把"当前打开文档的完整路径"交出来（PPT: GetPresentationInfo.fullName，
     * Word: GetDocumentInfo.fullName），其前 3 段就是当前登录用户的 profile —— 纯动态，不写死用户名。
     * 文档在网络盘/非用户目录时不匹配，返回空串由上层放弃。
     * @param {string} p 任意已知路径
     * @returns {string} 桌面绝对路径；无法反推时返回 ''
     */
    function desktopFromUserProfile(p) {
        if (typeof p !== 'string' || !p) return '';
        // 【关键】先归一化分隔符：COM 往返后 `C:\Users\...` 已变成 `C:/Users/...`，
        // 不归一化则下面这个要求 `\` 的正则永远匹配不上（实测 2026-10-05 23:45）。
        const m = normalizePathSeparators(p).match(/^([A-Za-z]:\\Users\\[^\\/?*:]+)\\/i);
        return m ? (m[1] + '\\Desktop') : '';
    }

    /**
     * 桥方法返回值归一化：Excel 主桥返回裸路径，HostBridgeBase 基类/新实现可能返回
     * `{"success":true,"path":"..."}` 或 `{"success":false,"error":"..."}`。
     * 这里统一取出可用的目录字符串；错误 JSON 一律视为"拿不到"。
     */
    function dirFromBridgeResult(v) {
        if (v === null || v === undefined) return '';
        const s = String(v).trim();
        if (!s) return '';
        if (s.charAt(0) === '{') {
            try {
                const j = JSON.parse(s);
                if (!j || j.success === false) return '';
                const cand = j.path || j.desktop || j.dir || j.result || '';
                return typeof cand === 'string' ? normalizePathSeparators(cand) : '';
            } catch (_e) { return ''; }
        }
        // 非 JSON 且含分隔符才可能是路径（错误文案不含路径分隔符）
        return /[\\/]/.test(s) ? normalizePathSeparators(s) : '';
    }

    /**
     * 宿主导出（PPT/Word → PDF/图片）的保存路径兜底生成器。
     * 动机（实测 2026-10-05 23:00）：用户只说"将演示文稿导出为 PDF"，模型不会传 filePath。
     * 原先只靠 tool-registry 的 BACKEND_INFERRED_PARAMS 留空交后端兜底，但宿主 DLL
     * （Hosts 下的 PptAIVSTOAddIn.dll / WordAIVSTOAddIn.dll）不一定随源码重编译——旧版
     * Bridge.ExportPdf 仍有 `filePath 不能为空` 硬校验，工具直接失败并把原始校验语甩给用户。
     * 这里在前端就把真实路径算出来，新旧后端都能落盘；目录拿不到时返回 null，
     * 交回原逻辑（后端兜底 / 反问用户），不做瞎猜。
     * 目录解析链：用户消息里的盘符路径 → Downloads → GetDesktopPath() → 由当前文档路径反推 profile 桌面。
     * @param {Object} ctxTurn 当前 turn（含 userMessage / context）
     * @param {{ext?:string, base?:string}} [opt] ext=扩展名；base=取不到文档名时的兜底文件名
     * @returns {Promise<string|null>} 绝对路径；无法确定目录时返回 null
     */
    async function inferHostExportFilePath(ctxTurn, opt) {
        const o = opt || {};
        const ext = String(o.ext || 'pdf').replace(/^\./, '').toLowerCase().replace(/[^a-z0-9]/g, '') || 'pdf';
        const bridge = (window.BridgeAccessor && window.BridgeAccessor.getBridge()) || (App && App.excelBridge);
        const userText = String((ctxTurn && (ctxTurn.userMessage || (ctxTurn.context && ctxTurn.context.userMessage))) || '');

        // 1) 用户明确说了保存目录（"保存到 D:\季度报告\报告.pdf"）→ 尊重用户
        let dir = '';
        if (/保存|存到|导出到|输出到|另存|放在/.test(userText)) {
            // 【修复】分隔符同时接受 `\` 与 `/`：宿主/COM 往返后用户消息里的 Windows 路径
            // 可能已变成 `D:/季度报告/最终稿.pdf`，只认`\` 会整个匹配失败 →
            // 用户明明指定了目录却被静默忽略、落到桌面（实测 2026-10-05 23:50）。
            const m = userText.match(/[A-Za-z]:[\\/][^\s，。；;、"”']{1,160}/);
            if (m && m[0]) {
                const raw = normalizePathSeparators(m[0]);
                dir = /\.[A-Za-z0-9]{1,5}$/.test(raw) ? raw.slice(0, Math.max(raw.lastIndexOf('\\'), 0)) : raw;
                dir = dir.replace(/\\+$/, '');
            }
        }
        // 2) "下载" → Downloads（宿主未暴露 GetDownloadsPath 时退回桌面）
        if (!dir && /下载|downloads?/i.test(userText) && bridge && typeof bridge.GetDownloadsPath === 'function') {
            try { dir = dirFromBridgeResult(await bridge.GetDownloadsPath()); } catch (_e) { dir = ''; }
        }
        // 3) 桌面：宿主若 override 了 GetDesktopPath 就直接用（Excel 主桥返回裸路径，PPT/Word 宿主
        //    未 override 时返回基类的 error JSON，这里归一化为空串继续往下走）
        if (!dir && bridge && typeof bridge.GetDesktopPath === 'function') {
            try { dir = dirFromBridgeResult(await bridge.GetDesktopPath()); } catch (_e) { dir = ''; }
        }
        // 4) 兜底：由当前打开文档的完整路径反推当前登录用户的桌面（PPT/Word 宿主唯一可靠的动态来源）
        //    顺带记下文档名：turn context 里没带 activeSheet 时（非首轮/思考轮）也能拿到好文件名
        let docNameFromBridge = '';
        if (!dir && bridge) {
            for (const fn of ['GetPresentationInfo', 'GetDocumentInfo', 'GetUIPath']) {
                if (typeof bridge[fn] !== 'function') continue;
                try {
                    const raw = await bridge[fn]();
                    const s = String(raw || '').trim();
                    // 既支持裸路径，也支持 {"success":true,"fullName":"..."} 这类 JSON
                    let cand = s;
                    if (s.charAt(0) === '{') {
                        const j = JSON.parse(s);
                        if (j && j.success !== false && j.name && !docNameFromBridge) docNameFromBridge = String(j.name);
                        cand = (j && (j.fullName || j.path || j.fileName)) || '';
                    }
                    dir = desktopFromUserProfile(cand);
                    if (dir) break;
                } catch (_e) { /* 试下一个 */ }
            }
        }
        dir = normalizePathSeparators(dir);
        if (!dir || !/[\\/]/.test(dir)) return null;

        // 文件名：**不直接用用户在消息里写的文件名**（如"最终稿.pdf"）——同名会被反复覆盖。
        // 目录尊重用户，文件名走当前文档名（chat-entry 放进 context.activeSheet.workbook）
        // → 桥给的文档名 → 工具类兜底名，最后统一加时间戳。
        const ctx = (ctxTurn && ctxTurn.context) || {};
        let base = String((ctx.activeSheet && ctx.activeSheet.workbook) || docNameFromBridge || o.base || 'export');
        base = base.replace(/\.[A-Za-z0-9]{1,5}$/, '')   // 去扩展名（"季报.pptx" → "季报"）
            .replace(/[\\/:*?"<>|]/g, '_')
            .replace(/\s+/g, ' ')
            .trim().slice(0, 60);
        if (!base) base = String(o.base || 'export');

        const d = new Date();
        const pad = (n) => String(n).padStart(2, '0');
        const ts = `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
        // 【修复】同inferExportFilePath：一律用反斜杠拼，不猜分隔符。
        return joinWinPath(dir, `${base}_${ts}.${ext}`);
    }

    // 占位符路径的"整段名"判定：token 必须独占一个路径段（前后是分隔符或串首尾），
    // 避免把 `Desktop\username_report.csv` 这类正常文件名误判成占位符。
    const PLACEHOLDER_PATH_SEG_RE = /(^|[\\/])(your[\s_-]?user\w*|yourname|your[\s_-]?name|some[\s_-]?user|test[\s_-]?user|user[\s_-]?name|placeholder|dummy|xxx+|用户名|用户目录|示例用户)([\\/]|$)/i;
    // 占位符路径的"任意位置"判定（尖/花括号模板、未展开变量、示例域名、省略号等）
    const PLACEHOLDER_PATH_ANY_RE = /(<[^<>/]{1,24}>|\{[^{}/]{1,24}\}|%username%|your[\s_-]?path|path[\\/]+to[\\/]|example\.(com|org)|\.\.\.|…)/i;

    /**
     * 判断某个参数值是否是"模型编造出来的占位符路径"。
     * 动机（实测）：用户说"将当前工作表导出为 CSV 并保存到桌面"，模型把用户名当占位符，
     * 传 `C:\Users\Yourusername\Desktop\output.csv`；该值**非空**，于是绕过了
     * inferExportFilePath 的默认路径推断，后端又无条件 CreateDirectory，
     * 磁盘上真被建出 `C:\Users\Yourusername\Desktop\`。
     * 因此"已有参数是否可信"的判定必须把这类值当成"空"，交给默认生成器/后端兜底。
     * 语义与后端 Utils.PathHelper.PlaceholderPathRegex 对齐（双层防护）。
     * @param {*} p
     * @returns {boolean}
     */
    /**
     * 【根因修复·分隔符归一化】把桥返回的目录/路径里的分隔符统一成 Windows 反斜杠。
     *
     * 为什么必须做（实测 2026-10-05 23:45）：WebView2 的 COM hostObjects 代理把
     * .NET 字符串经过 VARIANT 往返后，**反斜杠会被吃掉**——后端
     * `C:\Users\Administrator\Desktop` 到前端已变成 `C:/Users/Administrator/Desktop`。
     * 于是所有"按 `\` 分割"的既有逻辑全部失效：
     *   - `dir.includes('\\')` 判定为假 → 用 `/` 拼接 → 产出 `C://Users//...` 非法路径；
     *   - `desktopFromUserProfile` 的 `^([A-Za-z]:\\Users\\...)` 正则匹配失败 → 目录解析返回空；
     * 这就是"以前能用、现在导出全挂"的真实原因（后端兜底是新增的，前端却先崩了）。
     *
     * 统一在此收口：所有从 bridge 拿到的路径先过这里，后续代码只需处理一种分隔符。
     * 同时保留原有分隔符风格（网络盘/UNC 的 `\\` 前缀不能被误改）。
     * @param {string} p 任意路径
     * @returns {string} 分隔符统一为 `\` 的路径；非字符串返回空串
     */
    function normalizePathSeparators(p) {
        if (typeof p !== 'string') return '';
        let s = p.trim();
        if (!s) return '';
        const isUnc = /^[\\/]{2}/.test(s);
        // 先压掉所有连续分隔符（COM 往返会把 `\\` 变成 `//`），再确保结尾无分隔符。
        // UNC（\\server\share）压完是 `\server\share`，去掉这个前导单斜杠再补回双反斜杠。
        s = s.replace(/[\\/]+/g, '\\').replace(/\\+$/, '').replace(/^\\/, '');
        if (isUnc) s = '\\\\' + s;
        return s;
    }

    /**
     * 拼接目录与文件名，**永远用反斜杠**（Windows 绝对路径的正规写法）。
     * 不再依据 dir 里有没有 `\` 来猜分隔符——COM 往返后这个判断本就不可靠。
     */
    function joinWinPath(dir, name) {
        const d = normalizePathSeparators(dir).replace(/\\+$/, '');
        const n = String(name || '').replace(/^[\\/]+/, '');
        if (!d) return n;
        return n ? d + '\\' + n : d;
    }

    function isPlaceholderPath(p) {
        if (typeof p !== 'string') return false;
        const s = p.trim();
        if (!s) return false;
        return PLACEHOLDER_PATH_SEG_RE.test(s) || PLACEHOLDER_PATH_ANY_RE.test(s);
    }

    // 收窄"占位符路径"判定的适用范围：仅路径型参数名，或值本身像绝对路径。
    const PATHISH_PARAM_RE = /(path|file|folder|dir|目录|文件名)/i;
    const ABS_PATHISH_VALUE_RE = /^[a-zA-Z]:[\\/]|^\\\\|^[\\/](?![\\/])|^~\//;

    /**
     * 是否应对「参数 pk 的值 val」启用占位符路径检测。
     *
     * 为什么必须收窄：completeRequiredArgs 会对**每一个必填参数**调用占位符判定，
     * 而 isPlaceholderPath 的花括号/省略号规则是为【路径】设计的。代码/文本类参数
     * （典型：excel_run_python_script.script）天然含 {...}（字典字面量）、... 、… ，
     * 一律套用会被误判成"模型编造的占位符" → 该值被清空并反问用户，
     * 脚本根本执行不了（实测 2026-10-09：python-pipeline 自定义脚本被判
     * "缺少必要参数：请提供script"，尽管 script 已传入）。
     * 因此此处只对「路径型参数名」或「值本身像绝对路径」的参数启用检测。
     */
    function isPlaceholderPathForParam(pk, val) {
        if (typeof val !== 'string' || !val.trim()) return false;
        const nameHit = PATHISH_PARAM_RE.test(String(pk || ''));
        const valueLooksLikePath = ABS_PATHISH_VALUE_RE.test(val.trim());
        if (!nameHit && !valueLooksLikePath) return false;
        return isPlaceholderPath(val);
    }

    function inferChartTitleFromMessage(msg, chartType) {
        if (!msg) return null;
        const m = msg.match(/(?:按|对|以|根据|按照|基于|依)[\s]?([^\s，。,\.!！?？、；;：:（）()（）()【】\[\]\{\}\d]{1,12})(?:汇总|统计|分组|求和|平均|合计|分析|维度|分类|归类)/);
        if (m && m[1]) {
            const dim = m[1].trim();
            const vm = msg.match(/(销售额|销售|利润|成本|金额|收入|数量|业绩|单价|费用|毛利|净利|客户数|人数|占比)/);
            const val = vm ? vm[1] : '';
            if (dim && val) return `按${dim}${val}汇总`;
            if (dim) return `按${dim}汇总`;
        }
        const typeNames = {
            xlColumnClustered: '柱状图', xlColumnStacked: '堆叠柱状图', xlColumnStacked100: '100%堆叠柱状图',
            xlPie: '饼图', xlLine: '折线图', xlLineMarkers: '折线图', xlBarClustered: '条形图',
            xlDoughnut: '圆环图', xlArea: '面积图', xlFunnel: '漏斗图', xlWaterfall: '瀑布图',
            xlXYScatter: '散点图', xlRadar: '雷达图'
        };
        if (chartType && typeNames[chartType]) return typeNames[chartType];
        return null;
    }

    function inferValueColumnFromMessage(msg) {
        if (!msg) return null;
        const m = msg.match(/(销售额|销售|利润|成本|金额|收入|数量|业绩|单价|费用|毛利|净利|客户数|人数|占比)/);
        return m ? m[1] : null;
    }

    function inferChartDefaults(ctxTurn, args) {
        const msg = (ctxTurn && typeof ctxTurn.userMessage === 'string') ? ctxTurn.userMessage : '';
        const out = {};
        if (!_hasValue(args.title)) {
            const t = inferChartTitleFromMessage(msg, args.chartType);
            if (t) out.title = t;
        }
        if (args.hasLegend === undefined) out.hasLegend = true;
        if (args.showDataLabels === undefined) {
            const ct = String(args.chartType || '');
            out.showDataLabels = /xlPie|xlDoughnut|xlFunnel|xlWaterfall/.test(ct);
        }
        if (args.chartStyle === undefined) out.chartStyle = 201;
        if (!_hasValue(args.xAxisTitle)) {
            const dim = inferDimColumnFromMessage(ctxTurn);
            if (dim) out.xAxisTitle = dim;
        }
        if (!_hasValue(args.yAxisTitle)) {
            const val = inferValueColumnFromMessage(msg);
            if (val) out.yAxisTitle = val;
        }
        return Object.keys(out).length ? out : null;
    }

    function classifyError(rawError, toolName) {
        if (!rawError) return null;
        const errStr = typeof rawError === 'string' ? rawError : JSON.stringify(rawError);
        const toolLabel = toolName || '工具';

        // 1. 参数缺失/为空类
        if (/不能为空|缺少参数|missing.*param|required.*param|参数.*缺失|参数.*不能为空|field.*required|ArgumentNullException/i.test(errStr)
            || /源工作表.*空|目标工作表.*空|工作表.*空|sheet.*empty|range.*empty/i.test(errStr)
            || /请指定清洗操作|请指定.*operations|operations.*为空/i.test(errStr)) {
            return {
                type: '参数缺失',
                reason: '必填参数为空或未提供，后端拒绝执行。',
                suggestion: `请检查 ${toolLabel} 的必填参数（如工作表名、单元格范围、文件名等），确认选区已选中且参数值有效。`
            };
        }

        // 2. 工作表/范围不存在
        if (/不存在|not.*found|not.*exist|DISP_E_BADINDEX|无效.*工作表|无效.*名称|worksheet.*not.*found|sheet.*not.*found/i.test(errStr)) {
            return {
                type: '引用无效',
                reason: '指定的工作表、单元格范围或名称不存在。',
                suggestion: `确认工作簿中存在对应的工作表，检查单元格地址格式是否正确（如 A1:B10），或先切换到正确的工作表再重试。`
            };
        }

        // 3. Excel 未连接/桥接失败
        if (/excel.*not.*connect|excel.*未连接|桥接.*失败|bridge.*fail|COM.*fail|VSTO.*not.*load|excelBridge.*不可用/i.test(errStr)
            || /excelBridge|WebView2.*not.*found|hostObjects.*not.*available/i.test(errStr)) {
            return {
                type: 'Excel 连接异常',
                reason: 'Excel 与插件的通信桥接未建立或已断开。',
                suggestion: '请确认已打开 Excel 且插件已加载，尝试重启 Excel 或重新打开任务面板。'
            };
        }

        // 4. MCP Server 连接问题
        if (/mcp.*server.*not|mcp.*not.*connect|mcp.*未连接|server.*not.*start|connection.*refused|ECONNREFUSED|connect.*fail/i.test(errStr)) {
            return {
                type: 'MCP 服务未连接',
                reason: 'MCP Server 未启动或连接已断开。',
                suggestion: '在 MCP 面板点击"启动"按钮启动 MCP Server，或检查端口是否被占用。'
            };
        }

        // 5. 权限/保护
        if (/protected|保护|locked|read.?only|只读|无权限|permission|access.*denied/i.test(errStr)) {
            return {
                type: '权限/保护限制',
                reason: '目标工作表/工作簿受保护或为只读状态。',
                suggestion: '在 Excel 中取消工作表保护（审阅 → 撤销工作表保护），或另存为新文件后再操作。'
            };
        }

        // 6. 数据为空/无结果
        // 注意：不能用裸 /empty/（会误匹配操作名 fillEmpty/removeEmptyRows），
        // 只匹配"空数据"语境（数据/范围/选区为空、无数据、no data/result）。
        if (/(数据为空|无数据|没有数据|无可用数据|数据.*为空|no data|no result|selection.*empty|选区.*空|empty\s*(data|range|selection|area|rows?|cells))/i.test(errStr)) {
            return {
                type: '数据为空',
                reason: '当前选区或数据源为空，没有可处理的内容。',
                suggestion: '先在 Excel 中选择包含数据的单元格区域，或在调用工具时明确指定有效的范围地址。'
            };
        }

        // 7. 工具未注册/不存在
        if (/tool.*not.*found|Tool not found|未注册|not.*registered|unknown.*tool/i.test(errStr)) {
            const isSkill = /skill/i.test(toolLabel || '');
            if (isSkill) {
                return {
                    type: 'Skill 未加载',
                    reason: '该 Skill 未在系统中注册或已被移除。',
                    suggestion: '请在 Skill 管理面板确认该 Skill 已启用，或重启插件刷新 Skill 列表。Skill 由系统根据用户意图自动路由，通常不需要手动调用。'
                };
            }
            return {
                type: '工具未注册',
                reason: `工具 ${toolLabel} 未在工具注册表中找到。`,
                suggestion: '请刷新工具列表（重新加载 MCP / Skill），或使用其他可用工具。若为 Skill 调用，请注意 Skill 由系统自动路由，无需也不允许直接调用。'
            };
        }

        // 8. Skill 执行器不可用
        if (/executor.*not.*available|ExecuteSkill.*not.*available|skillExecutor.*not|Skill executor/i.test(errStr)) {
            return {
                type: 'Skill 执行器不可用',
                reason: 'Skill 的执行后端（Excel 桥接 / Skill 执行器）未就绪。',
                suggestion: '请确认 Excel 插件已完全加载，尝试重启 Excel；若仍不恢复，检查 Skill 配置是否完整。'
            };
        }

        // 9. 参数类型/格式错误
        if (/invalid.*param|invalid.*argument|参数.*无效|参数.*格式|type.*mismatch|format.*error/i.test(errStr)) {
            return {
                type: '参数格式错误',
                reason: '传入参数的类型或格式不符合工具要求。',
                suggestion: `检查 ${toolLabel} 各参数的数据类型（字符串/数字/数组）和格式要求，必要时重新描述您的意图让系统重新生成正确参数。`
            };
        }

        // 10. 文件路径/文件不存在
        if (/file.*not.*found|文件.*不存在|路径.*无效|path.*not.*exist|file.*not.*exist/i.test(errStr)) {
            return {
                type: '文件未找到',
                reason: '指定的文件路径不存在或无法访问。',
                suggestion: '确认文件路径正确、文件存在且可访问，必要时重新选择文件。'
            };
        }

        // 11. 网络/超时
        if (/timeout|timed.?out|超时|network.*error|网络.*错误|ETIMEDOUT/i.test(errStr)) {
            return {
                type: '请求超时',
                reason: '工具执行超时，可能是网络问题或后端处理耗时过长。',
                suggestion: '检查网络连接，稍后重试；若问题持续，考虑拆分任务或减小数据规模。'
            };
        }

        // 12. 聚合工具特有
        if (/aggregate|聚合/i.test(toolLabel) && /not|fail|error/i.test(errStr)) {
            return {
                type: '聚合分析失败',
                reason: '聚合工具执行过程中发生错误，可能是数据格式不匹配或选区为空。',
                suggestion: '确保已选中包含完整数据的区域（含表头），数据列类型符合聚合需求（数值列用于求和/平均、文本列用于分组）。'
            };
        }

        // 13. OCR 相关
        if (/ocr|OCR|recognize|识别/i.test(toolLabel)) {
            if (/image|图片|文件|path/i.test(errStr)) {
                return {
                    type: 'OCR 参数错误',
                    reason: 'OCR 工具需要有效的图片文件路径或图片数据。',
                    suggestion: 'OCR 仅用于图片识别。如要读取 Excel 单元格数据，请改用 excel_read_range 工具。'
                };
            }
        }

        return null;
    }

    function normalizeMcpResult(result) {
        if (result == null) {
            return { ok: true, result: null };
        }

        // 切片器成功载荷：即使外层误标 success=false / 把汇总塞进 error，只要建成了切片或透视图就视为成功
        function slicerPayloadOk(text) {
            if (typeof text !== 'string' || !text) return false;
            if (!/excel_create_slicer|slicersCreated|ptChartOkCount/i.test(text)) return false;
            const sc = /"slicersCreated"\s*:\s*(\d+)/.exec(text);
            const pc = /"ptChartOkCount"\s*:\s*(\d+)/.exec(text);
            if ((sc && Number(sc[1]) > 0) || (pc && Number(pc[1]) > 0)) return true;
            if (/Slicer:\s*[1-9]\d*\s*\/\s*[1-9]/i.test(text) || /PivotChart:\s*[1-9]\d*\s*\/\s*[1-9]/i.test(text)) return true;
            return false;
        }

        // 1. 明确的 error 字段
        if (result && typeof result === 'object') {
            if (result.error && typeof result.error === 'string' && result.error.trim()) {
                if (slicerPayloadOk(result.error) || slicerPayloadOk(result.result)) {
                    return { ok: true, result: result.result != null ? result.result : result.error };
                }
                return { ok: false, error: result.error.trim(), result };
            }
            // 2. success/ok 为 false
            if (result.success === false || result.ok === false || result.Success === false) {
                const errMsg = result.message || result.Message || result.msg || result.Msg
                    || result.error || result.Error || '工具执行失败';
                if (slicerPayloadOk(String(errMsg)) || slicerPayloadOk(result.result)) {
                    return { ok: true, result: result.result != null ? result.result : errMsg };
                }
                return { ok: false, error: String(errMsg), result };
            }
            // 3. MCP 标准格式：content 数组中 isError=true 的项
            if (Array.isArray(result.content)) {
                const errorItem = result.content.find(c => c && (c.isError === true || c.type === 'error'));
                if (errorItem) {
                    const errText = errorItem.text || errorItem.message || JSON.stringify(errorItem);
                    return { ok: false, error: String(errText), result };
                }
            }
            // 4. 有 result 字段——提取出来
            if (result.result !== undefined) {
                // 递归检查 result 字段内部是否为错误
                const inner = normalizeMcpResult(result.result);
                if (!inner.ok) return { ok: false, error: inner.error, result };
                return { ok: true, result: result.result };
            }
        }

        // 5. 纯字符串——检查是否为错误消息
        if (typeof result === 'string') {
            const s = result.trim();
            if (/^(错误|失败|异常|出错|Error|Fail|Exception)[：:]?/i.test(s)
                || /(不能为空|不能为空|不存在|执行失败|执行错误|发生错误|发生异常)$/.test(s)) {
                return { ok: false, error: s, result: s };
            }
            return { ok: true, result: s };
        }

        return { ok: true, result };
    }

    function computeNextRound(turn) {
        let max = 0;
        if (Array.isArray(turn?.toolCalls)) {
            for (const tc of turn.toolCalls) {
                const r = Number(tc?.round);
                if (Number.isFinite(r) && r > max) max = r;
            }
        }
        return max + 1;
    }

    function validateOutput(schema, value) {
        if (!schema) return { valid: true, errors: [] };
        const errors = [];

        function walk(node, v, path) {
            if (!node) return;
            // enum / const
            if (Array.isArray(node.enum) && !node.enum.some(e => JSON.stringify(e) === JSON.stringify(v))) {
                errors.push(`${path}: 值不在 enum 允许范围内`);
            }
            if (node.const !== undefined && JSON.stringify(node.const) !== JSON.stringify(v)) {
                errors.push(`${path}: 值必须等于 const ${JSON.stringify(node.const)}`);
            }
            // oneOf：恰好命中一个分支
            if (Array.isArray(node.oneOf)) {
                const hits = node.oneOf.filter(branch => {
                    const r = validateOutput(branch, v);
                    return r.valid;
                });
                if (hits.length !== 1) {
                    errors.push(`${path}: oneOf 要求恰好命中一个分支，实际命中 ${hits.length} 个`);
                    return;
                }
            }
            // type
            const t = node.type;
            if (t) {
                const okType =
                    (t === 'string' && typeof v === 'string') ||
                    (t === 'number' && typeof v === 'number') ||
                    (t === 'integer' && typeof v === 'number' && Number.isInteger(v)) ||
                    (t === 'boolean' && typeof v === 'boolean') ||
                    (t === 'null' && v === null) ||
                    (t === 'array' && Array.isArray(v)) ||
                    (t === 'object' && v !== null && typeof v === 'object' && !Array.isArray(v)) ||
                    (t === 'json' && true); // json = 任意合法 JSON
                if (!okType) {
                    errors.push(`${path}: 类型应为 ${t}，实际为 ${v === null ? 'null' : Array.isArray(v) ? 'array' : typeof v}`);
                    return;
                }
            }
            // 递归子节点
            if (v !== null && typeof v === 'object') {
                if (Array.isArray(v)) {
                    if (node.items) {
                        for (let i = 0; i < v.length; i++) walk(node.items, v[i], `${path}[${i}]`);
                    }
                } else if (node.properties) {
                    for (const key of Object.keys(node.properties)) {
                        if (v[key] !== undefined) walk(node.properties[key], v[key], `${path}.${key}`);
                    }
                    if (Array.isArray(node.required)) {
                        for (const req of node.required) {
                            if (v[req] === undefined) errors.push(`${path}: 缺少必填字段 "${req}"`);
                        }
                    }
                    if (node.additionalProperties === false) {
                        for (const key of Object.keys(v)) {
                            if (!node.properties || !(key in node.properties)) {
                                errors.push(`${path}: 不允许的额外字段 "${key}"`);
                            }
                        }
                    }
                }
            }
        }

        walk(schema, value, '$');
        return { valid: errors.length === 0, errors };
    }

    function finalizeToolContent(tool, args, result) {
        if (!tool || typeof tool.finalizeContent !== 'function') return result;
        try {
            const exec = { name: tool.name, args };
            const replacement = tool.finalizeContent(exec, { ...result });
            if (replacement !== undefined && replacement !== null) {
                return { ...result, content: replacement, _finalized: true };
            }
        } catch (e) {
            ChatLog.warn?.('TOOLS', `[tool-registry] ${tool.name} finalizeContent 异常（已忽略）:`, e);
        }
        return result;
    }

    function presentToolCall(tool, args) {
        if (!tool || typeof tool.presentCall !== 'function') return undefined;
        try {
            return tool.presentCall(args) || undefined;
        } catch (e) {
            ChatLog.warn?.('TOOLS', `[tool-registry] ${tool.name} presentCall 异常（已忽略）:`, e);
            return undefined;
        }
    }

    function presentToolResult(tool, args, result) {
        if (!tool || typeof tool.presentResult !== 'function') return undefined;
        try {
            return tool.presentResult(args, {
                ok: result.ok,
                content: result.content !== undefined ? result.content : result.result,
                error: result.error || null,
                meta: result._meta
            }) || undefined;
        } catch (e) {
            ChatLog.warn?.('TOOLS', `[tool-registry] ${tool.name} presentResult 异常（已忽略）:`, e);
            return undefined;
        }
    }

    function buildSkillAction(skill, args, mode) {
        const skillName = skill.name || skill.Name;
        const templateCode = skill.templateCode || skill.TemplateCode || '';

        let actionType = 'code_generation';
        if (mode === 'MultimodalLLM') actionType = 'image_recognition';
        else if (mode === 'LocalOCR') actionType = 'image_recognition';
        else if (mode === 'VBACode') actionType = 'code_generation';
        else if (mode === 'PythonCode') actionType = 'code_generation';

        return {
            actionType,
            parameters: {
                skillName,
                description: args?.description || args?.prompt || skill.description || '',
                templateCode,
                processingMode: mode,
                data: args,
                ...(args || {})
            }
        };
    }

    // ── 展示文本统一出口 ─────────────────────────────────────────────
    // 把任意工具/管道返回结果安全转成可靠的可读文本：字符串原样返回；对象优先取最易读的核心
    // 字段（result/resultText/message/msg/text/summary），否则 JSON.stringify，绝不 String() 成
    // "[object Object]" 或 "function () { [native code] }"。所有展示层都应走这里保证状态/结果对应。
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

    // ── JSON 回声替换守卫（弱模型/直通管道"把工具结果当总结"的兜底）─────────────
    // 背景（2026-10-09 用户反馈）：excel_merge_sheets 成功后，气泡正文就是
    // 整坨 {"success":true,"sourceSheets":[…],"mergedRows":…,"data":[[…]]}。
    // 泄漏点：skill-pipeline.formatSkillResult 对字符串结果原样返回，而后端工具的
    // content 本身就是 payload JSON 字符串 → 被当作"技能输出"写进正文。
    // 规则（只替换确定性可识别的结构，绝不改写自然语言）：
    //   ① 输入 trim 后必须整体是一段 JSON（{ 或 [ 开头），带任何说明文字的不动；
    //   ② 解析后命中"已知工具结果结构"才替换：
    //      - 合并结果（mergedRows/dataRows/sourceSheets/writtenTo 任一，含 result/data 信封层）
    //        → 交给 orchestrator.excelFriendlyText 出摘要；
    //      - 带字符串 summary 的产物对象 → 「✅ summary」+ files_created 逐行；
    //   ③ 其余 JSON（用户点名要的 JSON 输出、纯数据矩阵）一律不动，避免误伤。
    // 宽容解析：先按原样解析；失败且带转义引号（被嵌进字符串里的 JSON）时逐级去转义再试。
    // 关键（2026-10-09 实测踩坑）：当 payload 作为"字符串"被再包进一层 JSON 时，
    // JSON.stringify 会把原本的 \" 翻倍成 \\"，单层替换 /\\"/ 只去掉一个反斜杠仍无效 → 整坨漏过。
    // 因此先 /\\\\/g（双反斜杠→单）→ 再 /\\"/g（转义引号→引号），可覆盖 1~N 层嵌套。
    function _tryParseJson(s) {
        try { return JSON.parse(s); } catch (e) { /* 落到转义态重试 */ }
        let cur = s;
        for (let i = 0; i < 4; i++) {
            const un = cur.replace(/\\\\/g, '\\').replace(/\\"/g, '"');
            if (un === cur) break;
            cur = un;
            try { return JSON.parse(cur); } catch (e) { /* 继续去转义 */ }
        }
        return null;
    }

    // 【形状无关兜底】从任意文本里挖出"嵌套的合并结果 JSON 块"。
    // 为什么需要文本级扫描（2026-10-09 实测教训）：合并 payload 在运行时可能以
    // 字符串 / 双层 JSON 编码 / 未知包装键（payload/output/response…）落到正文或工具卡，
    // 结构化解包（result→data 逐层剥）只要漏一层就识别不到 → 整坨 JSON 摊进界面。
    // 这里用"在每个 { 处尝试递增长度的子串解析"取代手写括号配对——
    // 手写配对在"字符串里嵌转义 JSON"（如 payload 作为字符串再被 JSON.stringify 一层后
    // 引号翻倍成 \\"）时会把内层 } 误判为字符串内容，导致切片多截一个 } 而解析失败；
    // 子串试解析不依赖括号配对，遇到第一个能成功解析且命中合并字段的候选即返回。
    function findMergePayloadInText(text) {
        if (typeof text !== 'string' || !text) return null;
        if (text.length > 200000) return null;   // 超长文本不扫，避免卡顿
        // 整段已是"被引号包住的 JSON 字符串"（双层编码）→ 先剥外层引号再扫
        const t0 = text.trim();
        if (t0.charAt(0) === '"') {
            try {
                const inner = JSON.parse(t0);
                if (typeof inner === 'string') {
                    const r = findMergePayloadInText(inner);
                    if (r) return r;
                }
            } catch (e) { /* 不是合法编码串，继续按子串扫描 */ }
        }
        const MAX_LEN = 4000; // 单块候选上限，避免大文本上试解析过多
        const isMergeObj = (o) => o && typeof o === 'object' && !Array.isArray(o) &&
            (o.writtenTo !== undefined
                || (o.mergedRows !== undefined && Array.isArray(o.sourceSheets))
                || (o.dataRows !== undefined && Array.isArray(o.sourceSheets)));
        for (let st = 0; st < text.length; st++) {
            if (text.charAt(st) !== '{') continue;
            const limit = Math.min(text.length, st + MAX_LEN);
            for (let e = st + 1; e <= limit; e++) {
                const ch = text.charAt(e);
                // 只在看起来像 JSON 结尾（} 或 ]）的边界尝试解析，省去大量无效 parse
                if (ch !== '}' && ch !== ']') continue;
                const cand = text.slice(st, e + 1);
                const obj = _tryParseJson(cand);
                if (isMergeObj(obj)) return { obj: obj, start: st, end: e + 1 };
            }
        }
        return null;
    }

    // 把"正文里嵌的合并结果 JSON"整块替换成可读摘要（只替换命中块，其余文本原样保留）。
    function replaceMergeJsonInText(content) {
        if (typeof content !== 'string' || content.indexOf('{') < 0) return null;
        let out = content;
        let changed = false;
        for (let guard = 0; guard < 5; guard++) {
            const hit = findMergePayloadInText(out);
            if (!hit) break;
            const friendly = friendlyMergeResultText(hit.obj);
            if (!friendly) break;
            out = out.slice(0, hit.start) + friendly + out.slice(hit.end);
            changed = true;
        }
        return changed ? out : null;
    }

    // 合并结果 → 友好摘要（优先复用 orchestrator 的专属分支，保证两处文案一致）
    function friendlyMergeResultText(obj) {
        const _ef = App.ChatV2 && App.ChatV2.orchestrator && App.ChatV2.orchestrator.excelFriendlyText;
        if (typeof _ef === 'function') {
            try {
                const t = _ef('excel_merge_sheets', obj);
                if (t) return t;
            } catch (e) { /* 忽略，走 null */ }
        }
        return null;
    }

    // 正则抽取合并结果关键字段——完全脱离 JSON.parse，能处理"非标准 JSON"：
    // 运行时遇到的真实形态是"无逗号 + 换行"的 pprint 风格（Python/host 美化输出），
    // JSON.parse 必然失败；这里直接从文本抓取 mergedRows/sourceSheets/writtenTo 等专属字段，
    // 只要命中任一合并特征就把文本渲染成摘要，不再被裸 JSON 倾倒（2026-10-09 用户 v202610092157 反馈）。
    function _extractMergeFields(text) {
        if (typeof text !== 'string' || !text) return null;
        const num = (re) => { const m = text.match(re); return m ? Number(m[1]) : undefined; };
        const str = (re) => { const m = text.match(re); return m ? m[1] : undefined; };
        const bool = (re) => { const m = text.match(re); return m ? m[1].trim().toLowerCase() === 'true' : undefined; };
        const mergedRows = num(/"mergedRows"\s*:\s*(\d+)/);
        const dataRows = num(/"dataRows"\s*:\s*(\d+)/);
        const cleanedEmptySheets = num(/"cleanedEmptySheets"\s*:\s*(\d+)/);
        const headerPreserved = bool(/"headerPreserved"\s*:\s*(true|false)/);
        const writtenTo = str(/"writtenTo"\s*:\s*"([^"]*)"/);
        const targetSheet = str(/"targetSheet"\s*:\s*"([^"]*)"/);
        let sourceSheets = [];
        const sm = text.match(/"sourceSheets"\s*:\s*\[([\s\S]*?)\]/);
        if (sm) {
            sourceSheets = (sm[1].match(/"([^"]*)"/g) || []).map(s => s.slice(1, -1).trim()).filter(Boolean);
        }
        const hasMerge = mergedRows !== undefined || dataRows !== undefined ||
            sourceSheets.length > 0 || writtenTo !== undefined;
        if (!hasMerge) return null;
        return { mergedRows, dataRows, cleanedEmptySheets, headerPreserved, writtenTo, targetSheet, sourceSheets };
    }

    function replaceJsonEchoWithFriendly(content) {
        if (typeof content !== 'string') return null;
        let obj = null;
        // ① 正文整体就是一段 JSON（纯回声）→ 先走标准/转义解析（最快路径）
        const s0 = content.trim();
        if (s0 && (s0.charAt(0) === '{' || s0.charAt(0) === '[')) {
            try { obj = JSON.parse(s0); } catch (e) { obj = null; }
        }
        // ② 非标准 JSON（无逗号 pprint 风格 / 任意变体）→ 直接正则抽取合并字段。
        // 仅当整段就是 JSON 块（{ 或 [ 开头）时才整段替换，避免误删用户带说明前缀的文本。
        if (obj == null || typeof obj !== 'object') {
            if (s0.charAt(0) === '{' || s0.charAt(0) === '[') {
                const fields = _extractMergeFields(s0);
                if (fields) {
                    const _ef = App.ChatV2 && App.ChatV2.orchestrator && App.ChatV2.orchestrator.excelFriendlyText;
                    if (typeof _ef === 'function') {
                        const t = _ef('excel_merge_sheets', fields);
                        if (t) return t;
                    }
                }
            }
        }
        // ③ 正文里嵌了 JSON 块（模型带说明文字的回声 / 双层编码 / 非标准块）
        if ((!obj || typeof obj !== 'object') && content.indexOf('{') >= 0) {
            const hit = findMergePayloadInText(content) || _extractMergeFields(content);
            if (hit) {
                const t = replaceMergeJsonInText(content);
                if (t) return t;
            }
        }
        if (!obj || typeof obj !== 'object') return null;

        const _isMerge = (o) => o && typeof o === 'object' && !Array.isArray(o) &&
            (o.mergedRows !== undefined || o.dataRows !== undefined ||
                Array.isArray(o.sourceSheets) || o.writtenTo !== undefined);
        // 合并结果可能被包在 result/data 信封层里（字符串/对象两种形态），逐层下钻
        let cur = obj;
        for (let i = 0; i < 4 && cur && typeof cur === 'object' && !Array.isArray(cur); i++) {
            if (_isMerge(cur)) {
                const _ef = App.ChatV2 && App.ChatV2.orchestrator && App.ChatV2.orchestrator.excelFriendlyText;
                return (typeof _ef === 'function') ? _ef('excel_merge_sheets', cur) : null;
            }
            const nxt = (cur.result != null) ? cur.result
                : ((cur.data != null && !Array.isArray(cur.data)) ? cur.data : null);
            if (nxt == null) break;
            if (typeof nxt === 'string') {
                try { cur = JSON.parse(nxt); } catch (e) { break; }
            } else {
                cur = nxt;
            }
        }

        if (!Array.isArray(obj) && typeof obj.summary === 'string' && obj.summary.trim()) {
            const files = Array.isArray(obj.files_created)
                ? obj.files_created.filter(f => typeof f === 'string' && f.trim()) : [];
            return '✅ ' + obj.summary.trim() +
                (files.length ? '<br>' + files.map(f => '📄 ' + String(f).replace(/\\/g, '/')).join('<br>') : '');
        }
        return null;
    }

    /**
     * 编辑重发后，模型常把 excel_get_sheets / excel_read_range 的返回原样当成收尾。
     * 这种正文以 [ 或 { 开头，并带 usedRange 或一行业务数据。
     */
    function replaceToolDumpEcho(content, turn) {
        const s = String(content || '').trim();
        if (!s || (s.charAt(0) !== '[' && s.charAt(0) !== '{')) return null;
        const sheetDump = /"usedRange"/.test(s) && /"name"/.test(s);
        const rowDump = /^\s*\[/.test(s) && /"日期"|"销售额"|"部门"/.test(s);
        if (!sheetDump && !rowDump) return null;
        const calls = (turn && Array.isArray(turn._executedToolCalls)) ? turn._executedToolCalls : [];
        const lines = [];
        const merge = calls.slice().reverse().find(function (c) {
            return c && /merge_sheets/i.test(String(c.name || ''));
        });
        if (merge && merge.args) {
            let src = merge.args.sourceSheets;
            if (typeof src === 'string') {
                try { src = JSON.parse(src); } catch (_) { src = String(src).split(/[,，、]/); }
            }
            src = Array.isArray(src) ? src.map(function (x) { return String(x || '').trim(); }).filter(Boolean) : [];
            const target = String(merge.args.targetSheet || '合并结果').trim() || '合并结果';
            lines.push(src.length
                ? ('已将' + src.join('、') + '合并到「' + target + '」。')
                : ('已合并到「' + target + '」。'));
        }
        const painted = [];
        calls.forEach(function (c) {
            if (!c || !/format_range/i.test(String(c.name || ''))) return;
            const sh = c.args && String(c.args.sheet || '').trim();
            if (sh && painted.indexOf(sh) < 0) painted.push(sh);
        });
        if (painted.length) lines.push('已设置表头样式：' + painted.join('、') + '。');
        if (!lines.length && sheetDump) {
            const names = [];
            const re = /"name"\s*:\s*"([^"]+)"/g;
            let m;
            while ((m = re.exec(s)) && names.length < 8) names.push(m[1]);
            lines.push(names.length ? ('已读取工作表清单：' + names.join('、') + '。') : '已读取工作表清单。');
        }
        if (!lines.length) return '已完成操作。工具返回的原始数据不再贴出。';
        return lines.join('\n');
    }

    // ── 错误文本统一出口 ─────────────────────────────────────────────
    // 把任意错误值转成可读文本：函数/Promise/对象不再被 String() 成 "function () { [native code] }"。
    // 统一替代 agent-loop.safeErrText 与 orchestrator.safeExecError 的重复实现（见 #6 收敛）。
    // 默认文案取更通用的"未知原因"（错误不一定来自后端）。
    function safeErrorText(v) {
        if (v == null) return '未知原因';
        if (typeof v === 'string') return v || '未知原因';
        if (typeof v === 'function') return '后端返回了函数对象（可能未正确 await 或桥接异常）';
        if (v instanceof Error) return v.message || String(v);
        if (typeof v === 'object') {
            if (typeof v.message === 'string') return v.message;
            try {
                const s = JSON.stringify(v);
                return s && s.length <= 500 ? s : (s ? s.slice(0, 500) + '…' : '未知原因');
            } catch (e) { return String(v); }
        }
        return String(v);
    }

    // ── 空参 format_range 判定 ─────────────────────────────────────
    // 实测 2026-10-22 22:21：覆盖度误引导后模型回空参 format_range（{}），后端"当前选区"兜底
    // 把用户选区涂预设色（1ms 假成功 ❌）。合法调用必带 format/alignment/markLiteral/
    // conditionalFormat/autoRange 之一；全都没有即无效（纯 {range}/{sheet} 不构成格式化操作）。
    // 注意：dispatcher 批量链传 format='{}'（字符串，truthy）→ 有 payload，不会被误拦。
    function isMeaninglessFormatRangeArgs(args) {
        if (!args || typeof args !== 'object') return true;
        return !(args.format || args.alignment || args.markLiteral
            || args.conditionalFormat || args.autoRange === true);
    }

    function buildSkillInputs(args, ctx = {}) {
        if (!args || typeof args !== 'object') return {};
        const inputs = {};

        // 任务/请求内容归一化：LLM 在调用 Skill 时常把描述放在不同键下（text/prompt/message/contents/request/content/description/userPrompt）。
        // 取值优先级：显式 text > prompt > description > message > contents > request > content > userPrompt
        const rawText = args.text ?? args.prompt ?? args.description ?? args.message
            ?? args.contents ?? args.request ?? args.content ?? args.userPrompt;
        if (args.prompt) inputs.prompt = args.prompt;
        if (rawText != null && String(rawText).trim() !== '') inputs.text = String(rawText);
        else if (args.text) inputs.text = args.text;
        if (args.message) inputs.message = args.message;
        if (args.description) inputs.description = args.description;

        // 【兜底】LLM 调用 Skill 时若未提供任何请求描述（不同模型常漏传 text/prompt），
        // 用当前 turn 的用户消息补上，保证后端 Skill 执行器能拿到输入，避免多模态"contents 不能为空"。
        if (!inputs.text && !inputs.prompt && !inputs.message && ctx?.turn?.userMessage) {
            inputs.text = String(ctx.turn.userMessage);
        }
        // 数据上下文
        if (args.selection) inputs.selection = args.selection;
        if (args.excelData) inputs.excelData = args.excelData;
        if (Array.isArray(args.referenceRegions) && args.referenceRegions.length > 0) {
            inputs.referenceRegions = args.referenceRegions;
        }
        // 附件
        if (Array.isArray(args.attachments) && args.attachments.length > 0) {
            inputs.attachments = args.attachments;
            const first = args.attachments[0];
            if (first?.type === 'image' && first?.base64) {
                inputs.imageData = first.base64;
                inputs.mimeType = first.mime || 'image/png';
            }
        }
        return inputs;
    }

    /**
     * 美化 / 只点了表头时，底色不要铺到表身。
     * 用户点名表身、整表、隔行、斑马纹，或这是看板色块时，保持整段填充。
     * 已显式带 bgScope 的调用不改。
     */
    function wantsBodyFill(msg) {
        return /表身|数据区|正文底|斑马|隔行|条纹|整表.{0,8}(底色|填充|背景|涂)|全部.{0,6}(底色|填充|涂)|所有(行|单元格).{0,8}(底色|填充|背景)/.test(String(msg || ''));
    }

    function turnFillText(turn) {
        if (!turn) return '';
        const goal = turn.goal && typeof turn.goal === 'object' ? (turn.goal.text || '') : '';
        return [turn.userMessage, turn.message, goal].filter(Boolean).join('\n');
    }

    /** 用户点名要给这一段上底色（不是只改表头）。 */
    function userAskedToFillSelection(msg) {
        const s = String(msg || '');
        if (wantsBodyFill(s)) return true;
        if (/表头|首行|标题行/.test(s) && !/表身|数据区|整表|全部|所有单元格|所有行/.test(s)) return false;
        return /底色|填充色|背景色|涂成|填成|上色|铺色|高亮/.test(s);
    }

    function headerRowAddress(range) {
        const src = String(range || '').trim().replace(/：/g, ':').replace(/\s+/g, '');
        const m = src.match(/^(.*!)?\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/);
        if (!m) return '';
        const r1 = Number(m[3]);
        const r2 = m[5] ? Number(m[5]) : r1;
        if (!(r2 > r1)) return '';
        return (m[1] || '') + m[2] + r1 + ':' + (m[4] || m[2]) + r1;
    }

    function asFormatObject(args) {
        let format = args.format;
        if (typeof format === 'string') {
            const text = format.trim();
            if (!text) return null;
            try { format = JSON.parse(text); } catch (_) { return null; }
            if (!format || typeof format !== 'object' || Array.isArray(format)) return null;
            args.format = format;
        }
        return format;
    }

    /**
     * 多行区域的底色默认只留在首行。
     * 用户点名表身、整表、隔行，或直接说给这一段上底色时，整段保留。
     * 看板色块和带 bgScope=all|body 的调用不拆。
     * 返回值是只涂表头那一行的参数；原 args.format.bgColor 会被删掉。
     * 单行区域返回 null，底色留在原调用里。
     */
    function confineFormatFillToHeader(args, turn) {
        if (!args || typeof args !== 'object') return null;
        const format = asFormatObject(args);
        if (!format || typeof format !== 'object' || Array.isArray(format) || !format.bgColor) return null;
        // bgScope=header 已经声明只涂表头。不能再把 bgColor 从原调用里删掉：
        // 删掉之后完成条件看不到中国红，五张表明明涂成功了，收尾仍追加「没有执行」。
        if (format.bgScope === 'all' || format.bgScope === 'body' || format.bgScope === 'header') return null;
        const msg = turnFillText(turn);
        if (/看板|大屏|仪表盘/.test(msg)) return null;
        if (turn && turn.context && (turn.context._forceDashboard || turn.context._oneClickDashboard)) return null;
        if (userAskedToFillSelection(msg)) return null;
        const addr = String(args.range || args.address || '').replace(/：/g, ':').replace(/\s+/g, '');
        const parsed = addr.match(/^(.*!)?\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/);
        if (parsed) {
            const r1 = Number(parsed[3]);
            const r2 = parsed[5] ? Number(parsed[5]) : r1;
            if (!(r2 > r1)) return null;
        }
        const bgColor = format.bgColor;
        delete format.bgColor;
        delete format.bgScope;
        const headerRange = headerRowAddress(addr);
        if (!headerRange) return null;
        return { range: headerRange, sheet: args.sheet || '', format: { bgColor: bgColor, bgScope: 'all' } };
    }

    function resolveGuardMode(turnMode, thinkingMode) {
        if (turnMode) return turnMode;
        const map = { auto: 'standard', quick: 'quick', standard: 'standard', deep: 'deep', expert: 'expert' };
        return map[String(thinkingMode || 'auto').trim().toLowerCase()] || 'standard';
    }

    App.ChatV2.toolRegistryUtil = {
        CHART_TYPE_CN,
        resolveChartType,
        _presentSucceeded,
        _backendReported,
        normalizeName,
        normalizeSheetArg,
        inferDimColumnFromMessage,
        inferTimeGroupFromMessage,
        _existingSheetNames,
        extractSheetNameFromMessage,
        defaultNewSheetName,
        _hasValue,
        _requiredParamKeys,
        _paramLabel,
        resolveClearIntentFromMessage,
        extractExplicitSortColumnName,
        applyExplicitSortColumn,
        resolveIntentFlagsFromMessage,
        confineFormatFillToHeader,
        INTENT_PARAM_KEYS,
        pickArgsKeys,
        toBool,
        parseRangeAddr,
        combineSourceRanges,
        filterToolsForMessage,
        normalizeChartTypeValueForTool,
        normalizeChartArgs,
        normalizeShapeArgs,
        normalizeExportCsvArgs,
        inferExportFilePath,
        inferHostExportFilePath,
        normalizePathSeparators,
        joinWinPath,
        isPlaceholderPath,
        isPlaceholderPathForParam,
        inferChartTitleFromMessage,
        inferValueColumnFromMessage,
        inferChartDefaults,
        classifyError,
        normalizeMcpResult,
        computeNextRound,
        validateOutput,
        finalizeToolContent,
        presentToolCall,
        presentToolResult,
        buildSkillAction,
        buildSkillInputs,
        safeDisplayText,
        safeErrorText,
        replaceJsonEchoWithFriendly,
        replaceToolDumpEcho,
        findMergePayloadInText,
        replaceMergeJsonInText,
        friendlyMergeResultText,
        resolveGuardMode,
        isMeaninglessFormatRangeArgs,
    };
    ChatLog.info?.('TOOLS', 'tool-registry.util loaded (39 fns)');
})();