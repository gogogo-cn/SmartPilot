/**
 * ============================================================
 * chat-v2/services/intent-config.js — 意图配置加载器（单一事实源）
 * ============================================================
 * 职责：
 *   - 唯一事实源为 Config/intent_config.json（后端 PathHelper.ConfigDirectory 同源）
 *   - 前端加载链：VSTO excelBridge.GetIntentConfig() → HTTP /api/config/intent → 内嵌兜底
 *   - 将配置中的 keyword 数组编译为正则/定义，供 intent-detector / TaskPlanner / StepExecutor 消费
 *   - 配置热更新：远程配置到达后重建 defs，消费方在每次调用时经 defs() 取最新值
 *
 * 内嵌兜底与 Config/intent_config.json 保持一致（由 Tools/check-intent-consistency.ps1 校验）。
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    // ============================================================
    // 内嵌兜底（与 Config/intent_config.json 镜像，仅作离线/加载失败回退）
    // ============================================================
    const DEFAULT_CONFIG = {
        version: 1,
        chartTypes: [
            { keywords: ['饼图', '扇形图', '占比图', '占比', '饼'], label: '饼图', enum: 'xlPie' },
            { keywords: ['圆环图', '环形图', '环形'], label: '圆环图', enum: 'xlDoughnut' },
            { keywords: ['柱状图', '柱形图', '分组柱', '对比图', '柱状'], label: '柱状图', enum: 'xlColumnClustered' },
            { keywords: ['堆叠柱', '堆积柱', 'stacked column', 'stackedcolumn'], label: '堆叠柱', enum: 'xlColumnStacked', flags: 'i' },
            { keywords: ['条形图', '横向柱状图', '横向柱', '条形'], label: '条形图', enum: 'xlBarClustered' },
            { keywords: ['折线图', '趋势图', '折线'], label: '折线图', enum: 'xlLine' },
            { keywords: ['面积图', '面积'], label: '面积图', enum: 'xlArea' },
            { keywords: ['散点图', '散点'], label: '散点图', enum: 'xlXYScatter' },
            { keywords: ['雷达图', '雷达'], label: '雷达图', enum: 'xlRadar' },
            { keywords: ['漏斗图', '漏斗'], label: '漏斗图', enum: 'xlFunnel' },
            { keywords: ['瀑布图', '瀑布'], label: '瀑布图', enum: 'xlWaterfall' }
        ],
        chartWords: [
            '饼图', '柱状图', '柱形图', '条形图', '横向柱状图', '折线图', '散点图', '面积图', '雷达图',
            '瀑布图', '漏斗图', '圆环图', '柱状', '柱形', '条形', '折线', '散点', '面积', '雷达', '圆环', '饼'
        ],
        chartSplitWords: [
            '表格', '图表', '饼图', '柱状图', '条形图', '折线图', '散点图', '面积图', '雷达图',
            '图形', '视图', '数据', '列表', '明细'
        ],
        chartSpec: {
            zhMap: [
                { keywords: ['柱状', '柱形', '柱状图'], enum: 'xlColumnClustered' },
                { keywords: ['条形', '条形图'], enum: 'xlBarClustered' },
                { keywords: ['折线', '折线图', '趋势'], enum: 'xlLineMarkers' },
                { keywords: ['面积', '面积图'], enum: 'xlArea' },
                { keywords: ['饼', '饼图'], enum: 'xlPie' },
                { keywords: ['环形', '圆环'], enum: 'xlDoughnut' },
                { keywords: ['散点', '散点图'], enum: 'xlXYScatter' },
                { keywords: ['雷达', '雷达图'], enum: 'xlRadar' },
                { keywords: ['漏斗', '漏斗图'], enum: 'xlFunnel' },
                { keywords: ['瀑布', '瀑布图'], enum: 'xlWaterfall' },
                { keywords: ['箱线', '箱线图'], enum: 'xlBoxwhisker' },
                { keywords: ['树状', '树状图'], enum: 'xlTreemap' },
                { keywords: ['旭日', '旭日图'], enum: 'xlSunburst' },
                { keywords: ['直方', '直方图', '分布'], enum: 'xlHistogram' },
                { keywords: ['气泡', '气泡图'], enum: 'xlBubble' }
            ],
            semanticMap: [
                { re: '占比|构成|比例|份额|pie', enum: 'xlPie', flags: 'i' },
                { re: '趋势|走势|变化|时序|line|trend', enum: 'xlLineMarkers', flags: 'i' },
                { re: '对比|比较|排名|bar|column', enum: 'xlColumnClustered', flags: 'i' },
                { re: '分布|频率|histogram', enum: 'xlHistogram', flags: 'i' }
            ],
            displayMap: [
                { enum: 'xlPie', name: '饼图', icon: '🥧', category: '占比构成', keywords: ['pie', '饼'] },
                { enum: 'xlDoughnut', name: '圆环图', icon: '🍩', category: '占比构成', keywords: ['doughnut', 'donut', '圆环', '环形'] },
                { enum: 'xlColumnClustered', name: '分组柱状图', icon: '📊', category: '对比分析', keywords: ['column', 'clustered', '柱状', '柱形'] },
                { enum: 'xlColumnStacked', name: '堆叠柱状图', icon: '📊', category: '累积构成', keywords: ['stacked', '堆叠'] },
                { enum: 'xlColumnStacked100', name: '100%堆叠柱', icon: '📊', category: '占比构成', keywords: ['stacked100'] },
                { enum: 'xlBarClustered', name: '分组条形图', icon: '📊', category: '对比分析', keywords: ['bar', '条形'] },
                { enum: 'xlBarStacked', name: '堆叠条形图', icon: '📊', category: '累积构成', keywords: ['barstacked', '堆叠条'] },
                { enum: 'xlBarStacked100', name: '100%堆叠条', icon: '📊', category: '占比构成', keywords: ['barstacked100'] },
                { enum: 'xlLine', name: '折线图', icon: '📈', category: '趋势时序', keywords: ['line', '折线'] },
                { enum: 'xlLineMarkers', name: '折线带点图', icon: '📈', category: '趋势时序', keywords: ['marker', '折线', '趋势'] },
                { enum: 'xlArea', name: '面积图', icon: '📈', category: '趋势时序', keywords: ['area', '面积'] },
                { enum: 'xlXYScatter', name: '散点图', icon: '✳️', category: '相关性', keywords: ['scatter', '散点'] },
                { enum: 'xlRadar', name: '雷达图', icon: '🕸️', category: '多维评估', keywords: ['radar', '雷达'] },
                { enum: 'xlRadarFilled', name: '填充雷达图', icon: '🕸️', category: '多维评估', keywords: ['radarfilled'] },
                { enum: 'xlBubble', name: '气泡图', icon: '🫧', category: '相关性', keywords: ['bubble', '气泡'] },
                { enum: 'xlSunburst', name: '旭日图', icon: '🌞', category: '层级构成', keywords: ['sunburst', '旭日'] },
                { enum: 'xlTreemap', name: '树状图', icon: '🌳', category: '层级构成', keywords: ['treemap', '树状'] },
                { enum: 'xlHistogram', name: '直方图', icon: '📊', category: '分布分析', keywords: ['histogram', '直方', '分布'] },
                { enum: 'xlBoxwhisker', name: '箱线图', icon: '📦', category: '分布分析', keywords: ['box', '箱线'] },
                { enum: 'xlFunnel', name: '漏斗图', icon: '🔻', category: '转化分析', keywords: ['funnel', '漏斗'] },
                { enum: 'xlWaterfall', name: '瀑布图', icon: '💧', category: '增减分析', keywords: ['waterfall', '瀑布'] }
            ]
        },
        browserDisplayKeywords: ['浏览器', '网页', 'html', 'HTML', '在浏览器中', '用浏览器', '浏览器展示', '网页展示', 'html展示', '用.*展示', '展示.*数据', '打开.*查看', '网页表格', '表格.*查看'],
        aggregateIntent: {
            groupPrefixes: ['按', '依照', '根据', '每个', '每', '各'],
            aggWords: ['汇总', '求和', '总计', '合计', '累计', '相加'],
            avgWords: ['平均', '均值', '平均数', '平均值'],
            minMaxWords: ['最大', '最小', '最多', '最少', '最高', '最低'],
            countWords: ['计数', '统计', '数量', '个数', '多少', '几条'],
            groupWords: ['分组', '聚合', 'group', 'aggregate'],
            displayWords: ['展示', '显示', '查看', '分析', '对比', '排名', '绘制', '画', '生成', '做'],
            valueWords: ['利润', '销售(?:额)?', '成本', '数量', '客户(?:数)?', '金额', '产值', '产量', '人数', '人次', '订单', '销售额']
        },
        advancedPatterns: [
            { re: '前\\s*(?:[一二三四五六七八九十百]+|\\d+)\\s*[名条个行项记录页]|TOP\\s*\\d+|top\\s*\\d+|排名前|排行前|前几名|前几位', type: 'topn', label: 'Top N' },
            { re: '同比|环比|年同比|月环比|同期|去年同期|对比.*去年|对比.*上期', type: 'compare', label: '同比环比' },
            { re: '按月|按季|按年|按月份|按季度|按年份|按月汇总|按周汇总|按日期', type: 'date_group', label: '日期分组' },
            { re: '筛选|过滤|满足|条件|其中|只统计|只汇总|仅统计|仅汇总', type: 'filtered', label: '筛选汇总' }
        ],
        aggFuncs: [
            { keywords: ['求和', '总计', '合计', '汇总', '累计', '相加', '加总'], func: 'sum', label: '求和' },
            { keywords: ['平均', '均值', '平均数', '平均值', 'avg', 'average'], func: 'avg', label: '平均', flags: 'i' },
            { keywords: ['最大', '最多', '最高', 'max'], func: 'max', label: '最大', flags: 'i' },
            { keywords: ['最小', '最少', '最低', 'min'], func: 'min', label: '最小', flags: 'i' },
            { keywords: ['计数', '统计数量', '个数', '多少条', 'count'], func: 'count', label: '计数', flags: 'i' }
        ],
        dimColumns: [
            { keywords: ['部门', '科室', '组', '小组', '团队'], label: '部门' },
            { keywords: ['地区', '区域', '城市', '省份', '省', '市', '县'], label: '地区' },
            { keywords: ['产品', '商品', 'SKU', '品类', '类别', '分类'], label: '产品' },
            { keywords: ['日期', '时间', '月份', '季度', '年份', '月', '年', '周'], label: '日期' },
            { keywords: ['人员', '员工', '职工', '姓名', '业务员'], label: '人员' },
            { keywords: ['渠道', '方式', '途径', '来源'], label: '渠道' },
            { keywords: ['类型', '种类', '性质', '状态'], label: '类型' }
        ],
        valueColumns: [
            { keywords: ['利润', '盈利', '收益', '毛利'], label: '利润' },
            { keywords: ['销售', '营收', '收入', '营业额', '销售额'], label: '销售额' },
            { keywords: ['成本', '费用', '支出', '花费'], label: '成本' },
            { keywords: ['数量', '件数', '箱数', '吨数', '产量', '销量'], label: '数量' },
            { keywords: ['客户', '用户', '顾客', '人数', '人次', '会员'], label: '客户数' },
            { keywords: ['金额', '总额', '总值', '总计额'], label: '金额' },
            { keywords: ['平均', '均值', '单价', '价格', '费用率', '利率'], label: '平均值' }
        ],
        quickPass: ['你好', '嗨', 'hi', 'hello', 'hey', '你好呀', '您好', '早上好', '下午好', '晚上好', '再见', '拜拜', '谢谢', '感谢', '好的', 'ok', '嗯', '在吗', '你在吗', '今天', '现在几点', '你叫什么', '你是谁', '你能做什么'],
        targetSheet: {
            verbs: ['放到', '写入', '写到', '存到', '保存到', '放入', '写进', '另存为', '落到', '放在', '放进'],
            nouns: ['工作表', 'sheet', '表格', '工作簿', '页(?:面)?'],
            exclude: ['公式', '函数', '数据', '图表', '图片', '内容', '结果', '输出', '信息', '文字', '文本', '数字', '数值', '标题', '备注', '说明', '描述', '注释', '摘要', '总结', '结论', '意见', '建议', '方案', '配置', '设置', '参数', '条件', '规则', '格式', '样式', '模板', '代码', '脚本', '宏', '程序', '文件', '文档', '附件', '链接', '地址', '路径', '名称', '类型', '大小', '颜色', '字体', '边框', '背景', '前景', '位置', '方向', '长度', '宽度', '高度', '面积', '体积', '容量', '数量', '质量', '重量', '速度', '时间', '日期', '年份', '月份', '星期', '小时', '分钟', '秒', '毫秒', '桌面', '下载', '文件夹', '目录', '临时', '本地', '当前目录', '磁盘', 'U盘', '网盘', '云盘', '服务器', '邮箱', '剪贴板', '内存']
        },
        displayMode: {
            bothRe: '表格|数据(?:和|及|与|并|同时|都|两边|两个)(?:图|图表|视图)|既.*表.*又.*图|图和表|表格和图表|都展示|同时展示',
            chartRe: '饼图|柱状图|条形图|折线图|图表|图形|图\\b|视图|可视化|柱形|画图|画图展示|用图',
            tableRe: '表格|表盘|数据表|列表|明细|原始数据|作为表格|表格形式'
        },
        readOpKeywords: ['读取', '读一下', '读入', '读数据', '查看', '查找', '搜索', '分析', '看看', '浏览', '研究', '评估', '对比', '检查', '判断', '找出', '识别', '了解', '查看一下', '读出来', '分析一下', '看一下'],
        toolKeywords: {
            read: ['读取', '读一下', '查看', '查找', '搜索', '看看', '浏览', '研究', '评估', '对比', '检查', '找出', '判断', '识别', '了解', '查看一下', '看一下'],
            write: ['写入', '填写', '填入', '改', '修改', '设置', '写', '保存到', '存到', '放到', '复制到', '写到', '放进', '写入到', '另存为'],
            create: ['创建', '新建', '生成', '新增', '添加'],
            format: ['美化', '格式化', '设置格式', '调整', '对齐'],
            delete: ['删除', '清空', '移除', '去掉'],
            analyze: ['分析', '统计', '汇总', '计算', '求和', '平均', '最大', '最小', '看看', '研究', '评估', '对比', '找出', '判断', '识别', '了解', '分析一下'],
            chart: ['图表', '画图', '折线图', '柱状图', '饼图'],
            merge: ['合并', '关联', '匹配'],
            split: ['拆分', '分割', '分开'],
            send: ['发送', '邮件', '通知', '消息']
        },
        singleStepKeywords: [
            '^(读取|读一下|查看|打开)',
            '^(写入|填写|填入|设置为)',
            '^(删除|清空|移除)',
            '^(创建|新建|生成)(一个|一张|一份)?(新)?(工作簿|工作表|表格)',
            '^(美化|格式化)(一下)?(这个)?(表格|选区)',
            '^(求和|平均|统计|计算)(一下)?'
        ],
        toolMapping: [
            { keywords: ['图表', '柱状图', '折线图', '饼图', '面积图', '散点图', '雷达图', '环形图', '画图', 'chart'], tool: 'excel_create_chart', flags: 'i' },
            { keywords: ['按年', '按季度', '按月', '按周', '按日', '日期分组', '时间分组', '年度', '季度', '月度', '周度', '财年', 'year', 'quarter', 'month', 'week', 'day'], tool: 'excel_date_group', flags: 'i' },
            { keywords: ['汇总', '聚合', '统计', '求和', '分组', '分类', '平均', '计数', '合计', '分析', '最高', '最低', 'count', 'sum', 'aggregate', 'analyze'], tool: 'excel_aggregate', flags: 'i' },
            { keywords: ['格式化', '美化', '设置格式', 'format'], tool: 'excel_format_range', flags: 'i' },
            { keywords: ['读取', '查看', '读一下', 'read'], tool: 'excel_read_range', flags: 'i' },
            { keywords: ['复制', '粘贴', 'copy', 'paste'], tool: 'excel_copy_range', flags: 'i' },
            { keywords: ['新建工作表', '创建工作表', '新建表', '创建表', 'create_sheet'], tool: 'excel_create_sheet', flags: 'i' },
            { keywords: ['重命名', '改名', 'rename'], tool: 'excel_rename_sheet', flags: 'i' },
            { keywords: ['写入', '填写', '填入', 'write'], tool: 'excel_write_range', flags: 'i' },
            { keywords: ['删除', '清空', '清除', 'delete', 'clear'], tool: 'excel_clear_range', flags: 'i' }
        ],
        columnHints: {
            textHints: ['部门', '姓名', '名字', '名称', '地址', '电话', '邮箱', '性别', '备注', '描述', '说明', '类别', '分类', 'name', 'department', 'address', 'phone', 'email', 'note', 'desc', 'category', 'group'],
            categoryHints: ['部门', '地区', '区域', '大区', '分类', '类别', '类型', '产品', '商品', '品类', '渠道', '来源', '城市', '省份', '客户', '供应商', '业务', '项目', 'department', 'region', 'city', 'province', 'channel', 'product', 'category', 'group'],
            nonCategoryHints: ['日期', '时间', '年份', '月份', '数量', '金额', '价格', '销售额', '利润', '成本', '收入', '费用', '合计', '小计', '总额', '总计', '占比', '客户数', '人数', '次数', '个数', 'date', 'time', 'year', 'month', 'amount', 'price', 'quantity', 'total', 'revenue', 'profit', 'cost', 'sales'],
            numericHints: ['金额', '价格', '数量', '销售额', '利润', '总额', '总计', '数值', '数字', '收入', '成本', '费用', '合计', '小计', 'sum', 'amount', 'price', 'quantity', 'total', 'revenue', 'profit', 'value', 'sales', 'count', 'avg', 'average'],
            nonTextHints: ['部门', '姓名', '名字', '名称', '日期', '月份', '年份', '时间', '地址', '电话', '邮箱', '性别', '备注', '描述', '说明', '类别', '分类']
        },
        backend: {
            moneyHints: ['销售额', '销售', '利润', '成本', '收入', '金额', '价格', '单价', '费用', '毛利', '净利', '合计', '总额', '总计', '产值', '业绩', 'revenue', 'sales', 'profit', 'cost', 'amount', 'price', 'income'],
            nonNumericColumns: ['编号', '姓名', '名称', '日期', '时间', '备注'],
            dimensionFallback: ['部门', '地区', '区域', '大区', '分类', '类别', '类型', '产品', '商品', '品类', '渠道', '来源', '城市', '省份'],
            dimensionExclude: ['日期', '时间', '编号', '销售额', '成本', '利润', '金额', '数量', '客户数', '单价', '价格', '收入', '备注'],
            synonyms: {
                '销售额': ['月度业绩', '业绩', '销售收入', '营业额', '收入', '销售额'],
                '业绩': ['月度业绩', '业绩', '销售收入', '营业额', '收入', '销售额'],
                '提成': ['提成', '基础提成', '最终实发提成', '实发提成'],
                '部门': ['部门', '科室', '组', '小组', '团队', '所属部门', '部门名称'],
                '员工': ['员工', '员工姓名', '姓名', '人员', '职工', '业务员'],
                '数量': ['数量', '人数', '个数'],
                '金额': ['金额', '数额', '数值'],
                '地区': ['地区', '区域', '大区', '地域', '省市', '省份', '城市'],
                '日期': ['日期', '时间', '月份', '季度', '年份', '年度', '年月', '月', '年', '周'],
                '产品': ['产品', '商品', 'sku', '品类', '类别', '分类', '产品分类', '产品类别', '商品分类', '商品类别'],
                '渠道': ['渠道', '方式', '途径', '来源']
            }
        },
        rag: {
            semanticFallback: { enabled: true, threshold: 0.62, topK: 3 }
        }
    };

    // ============================================================
    // 状态
    // ============================================================
    let _config = DEFAULT_CONFIG;
    let _defs = null;
    let _loadPromise = null;
    let _subscribers = [];

    // ============================================================
    // 正则/定义构建器
    // ============================================================
    function escapeRe(s) {
        // 关键词若含正则元字符（如 "用.*展示"），保持原样；纯字面词才转义
        return s;
    }

    function buildRegex(keywords, flags) {
        const list = Array.isArray(keywords) ? keywords : [];
        if (list.length === 0) return null;
        return new RegExp(list.map(escapeRe).join('|'), flags || '');
    }

    function buildDefs(cfg) {
        const c = cfg || DEFAULT_CONFIG;
        const defs = {};

        // 图表类型定义
        defs.chartTypeDefs = (c.chartTypes || []).map(t => ({
            re: buildRegex(t.keywords, t.flags),
            label: t.label,
            enum: t.enum
        }));

        // 图表词正则（检测 + 计数）
        defs.chartWordRe = buildRegex(c.chartWords);
        defs.chartWordCountRe = buildRegex(c.chartWords, 'g');
        defs.chartSplitWords = c.chartSplitWords || [];

        // skill-spec-engine 图表 spec 规范化映射（zhMap/semanticMap/displayMap）
        const cs = c.chartSpec || {};
        defs.chartSpec = {
            zhMap: (cs.zhMap || []).map(m => ({ re: buildRegex(m.keywords), enum: m.enum })),
            semanticMap: (cs.semanticMap || []).map(m => ({ re: new RegExp(m.re, m.flags || ''), enum: m.enum })),
            displayMap: cs.displayMap || []
        };

        // 浏览器展示
        defs.browserDisplayRe = buildRegex(c.browserDisplayKeywords);

        // 目标工作表提取
        const ts = c.targetSheet || {};
        defs.targetSheet = {
            verbs: ts.verbs || [],
            nouns: ts.nouns || [],
            exclude: ts.exclude || []
        };

        // 展示模式
        const dm = c.displayMode || {};
        defs.displayMode = {
            bothRe: dm.bothRe ? new RegExp(dm.bothRe) : null,
            chartRe: dm.chartRe ? new RegExp(dm.chartRe, 'i') : null,
            tableRe: dm.tableRe ? new RegExp(dm.tableRe) : null
        };

        // 聚合意图正则（组合结构）
        const ai = c.aggregateIntent || {};
        defs._aggregateIntent = ai; // 原始配置，供 intent-detector 语法兜底拼装
        const join = arr => (arr || []).join('|');
        const aggIntentSource =
            '(?:' + join(ai.groupPrefixes) + '.*?(?:' +
            join(ai.aggWords) + '|' + join(ai.avgWords) + '|' + join(ai.minMaxWords) + '|' +
            join(ai.countWords) + '|' + join(ai.groupWords) + ')' +
            '|' + join(ai.groupPrefixes) + '.*?(?:' + join(ai.displayWords) + ')(?:的?\\s*(?:' + join(ai.valueWords) + '))?' +
            '|' + join(c.chartWords) + '(?:.*?)(?:' + join(ai.displayWords) + '))';
        defs.aggregateIntentRe = new RegExp(aggIntentSource);

        // 高级场景
        defs.advancedPatterns = (c.advancedPatterns || []).map(p => ({
            re: new RegExp(p.re),
            type: p.type,
            label: p.label
        }));

        // 聚合函数
        defs.aggFuncDefs = (c.aggFuncs || []).map(f => ({
            re: buildRegex(f.keywords, f.flags),
            func: f.func,
            label: f.label
        }));

        // 维度/数值列
        defs.dimColPatterns = (c.dimColumns || []).map(d => ({ re: buildRegex(d.keywords), label: d.label }));
        defs.valueColPatterns = (c.valueColumns || []).map(v => ({ re: buildRegex(v.keywords), label: v.label }));

        // TaskPlanner
        defs.quickPassRe = new RegExp('^(' + join(c.quickPass) + ')', 'i');
        defs.toolKeywords = {};
        for (const [k, words] of Object.entries(c.toolKeywords || {})) {
            defs.toolKeywords[k] = buildRegex(words);
        }
        // 路由守卫前置操作词（读取/分析类）：text-pipeline 图表路径守卫据此判定
        // 多步骤请求是否含"图表路径无法执行"的前置操作（读取→分析→写入等）。
        defs.readOpRe = buildRegex(c.readOpKeywords);
        defs.singleStepKeywords = (c.singleStepKeywords || []).map(s => new RegExp(s));
        defs.toolMapping = (c.toolMapping || []).map(m => ({ re: buildRegex(m.keywords, m.flags), tool: m.tool }));

        // StepExecutor 列提示
        const ch = c.columnHints || {};
        defs.columnHints = {
            textHints: buildRegex(ch.textHints, 'i'),
            categoryHints: buildRegex(ch.categoryHints, 'i'),
            nonCategoryHints: buildRegex(ch.nonCategoryHints, 'i'),
            numericHints: buildRegex(ch.numericHints, 'i'),
            nonTextHints: buildRegex(ch.nonTextHints, 'i')
        };

        return defs;
    }

    function rebuild() {
        _defs = buildDefs(_config);
        const subs = _subscribers.slice();
        _subscribers = [];
        for (const cb of subs) {
            try { cb(_config, _defs); } catch (e) { ChatLog.warn?.('INTENT-CONFIG', '订阅回调异常:', e?.message); }
        }
    }

    // ============================================================
    // 加载链：VSTO 桥接 → HTTP → 内嵌兜底
    // ============================================================
    function normalizeConfig(raw) {
        if (!raw || typeof raw !== 'object') return null;
        // 校验关键结构，避免坏配置导致运行时崩溃
        if (!Array.isArray(raw.chartTypes) && !Array.isArray(raw.chartWords) && !Array.isArray(raw.dimColumns)) {
            return null;
        }
        return raw;
    }

    async function load() {
        if (_loadPromise) return _loadPromise;
        _loadPromise = (async () => {
            let cfg = null;

            // 1. VSTO 桥接（走 BridgeAccessor 包装桥，避免裸 hostObjects 漏 0x80020006）
            try {
                const bridge = (window.BridgeAccessor && typeof BridgeAccessor.getBridge === 'function' && BridgeAccessor.getBridge())
                    || (window.App && App.excelBridge)
                    || ((window.chrome && window.chrome.webview && window.chrome.webview.hostObjects)
                        ? window.chrome.webview.hostObjects.excelBridge : null);
                if (bridge && typeof bridge.GetIntentConfig === 'function') {
                    const raw = await bridge.GetIntentConfig();
                    const parsed = typeof raw === 'string' ? JSON.parse(raw) : raw;
                    if (parsed && parsed.error === 'bridge-call-failed') {
                        /* 宿主未实现：走内嵌兜底 */
                    } else if (parsed && parsed.success && parsed.config) cfg = normalizeConfig(parsed.config);
                    else if (parsed && parsed.config) cfg = normalizeConfig(parsed.config);
                    else if (parsed && parsed.success !== false) cfg = normalizeConfig(parsed);
                    if (cfg) ChatLog.info?.('INTENT-CONFIG', '✅ 已从 VSTO 桥接加载意图配置 v' + (cfg.version || '?'));
                }
            } catch (e) {
                const msg = e?.message || String(e || '');
                if (!/e is not a function|0x80020006|未知名称|bridge-call-failed/i.test(msg)) {
                    ChatLog.warn?.('INTENT-CONFIG', 'VSTO 桥接加载意图配置失败:', msg);
                }
            }

            // 2. HTTP（Web 版 / 独立服务）
            // 【守卫修正·B】VSTO(WebView2) 走 COM 桥，本地 9527 服务不存在；
            // 即使桥未提供该方法也应回退内嵌兜底，而非去探 9527 刷屏。故 VSTO 跳过 HTTP。
            if (!cfg && !(window.chrome && window.chrome.webview && window.chrome.webview.hostObjects)) {
                try {
                    const base = (App.Endpoints ? App.Endpoints.excelBridge.baseUrl : (window.ExcelBridgeConfig && window.ExcelBridgeConfig.baseUrl)) || 'http://localhost:9527';
                    const res = await fetch(base + '/api/config/intent', { cache: 'no-store' });
                    if (res.ok) {
                        const data = await res.json();
                        if (data && data.success && data.config) cfg = normalizeConfig(data.config);
                        else if (data && data.config) cfg = normalizeConfig(data.config);
                        else cfg = normalizeConfig(data);
                        if (cfg) ChatLog.info?.('INTENT-CONFIG', '✅ 已从 HTTP 加载意图配置 v' + (cfg.version || '?'));
                    }
                } catch (e) {
                    ChatLog.warn?.('INTENT-CONFIG', 'HTTP 加载意图配置失败:', e?.message);
                }
            }

            // 3. 兜底：内嵌默认
            if (!cfg) {
                ChatLog.info?.('INTENT-CONFIG', '使用内嵌意图配置兜底 v' + (DEFAULT_CONFIG.version || '?'));
            } else {
                _config = cfg;
            }
            rebuild();
            return _config;
        })();
        return _loadPromise;
    }

    // ============================================================
    // 导出
    // ============================================================
    App.ChatV2.intentConfig = {
        get: () => _config,
        defs: () => {
            if (!_defs) rebuild();
            return _defs;
        },
        load,
        buildRegex,
        buildDefs,
        onUpdate: (cb) => { _subscribers.push(cb); },
        _defaults: DEFAULT_CONFIG
    };

    // 自动触发加载（不阻塞模块初始化；消费方经 defs() 取最新值）
    if (typeof window !== 'undefined') {
        setTimeout(() => { load().catch(() => {}); }, 0);
    }

    ChatLog.info?.('CHAT-V2', '✅ intentConfig loaded (embedded v' + (DEFAULT_CONFIG.version || '?') + ')');
})();