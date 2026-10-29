/**
 * ============================================================
 * chat-v2/services/vba-template-engine.js
 * ============================================================
 *
 * 职责：把"用户想写什么 VBA"转成"可直接运行的 VBA 代码"
 *
 * 架构原则（解决"补丁堆 prompt"的根本问题）：
 *   1. LLM 不再负责生成完整 VBA —— 它只输出结构化 IntentSpec JSON
 *      （如 {scenario:"ppt_from_data", dataSheet:"Sheet1", dataRange:"A1:G121", ...}）
 *   2. 真正的代码由人工审核过的模板 + 参数化渲染 生成
 *   3. 每个模板都是**确定性**的：相同 IntentSpec 永远产出相同代码
 *   4. 模板库覆盖高频场景，未来增加场景只需新增一个模板
 *
 * 为什么这样能解决问题：
 *   - 旧方式：LLM 一次性写完整 30+ 行 VBA 还要保证 API 调用正确
 *     → 任何一个细节记错就崩溃（Slides.Add 索引、Layout 常量、Dim 位置...）
 *   - 新方式：LLM 只填一个 JSON，模板负责"Sub 怎么写、循环怎么写、API 怎么调"
 *     → 模板是人工写好且测过的，LLM 失误不影响 API 调用正确性
 *
 * 模板格式：
 *   每个模板导出：
 *     - id: 唯一标识
 *     - title: 用户可见标题
 *     - match(intent): 根据 IntentSpec 判断是否匹配
 *     - render(intent, context): 返回完整 VBA 字符串
 *
 * 模板内部约定：
 *   - 顶层固定输出 `Sub <id>(...) ... End Sub` 结构
 *   - 所有变量声明在 Sub 顶部
 *   - 集合追加统一用 `collection.Count + 1` 作为 Index
 *   - 跨应用对象（PowerPoint）必须 Set 之后判空再使用
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.services) App.ChatV2.services = {};

    const ChatLog = window.ChatLog || console;

    // ============================================================
    // 工具：VBA 字符串安全转义
    // ============================================================
    function vbStr(v) {
        if (v == null) return '""';
        const s = String(v).replace(/"/g, '""');
        return `"${s}"`;
    }

    function vbComment(text) {
        if (!text) return '';
        return `' ${String(text).replace(/\n/g, ' ')}`;
    }

    /**
     * 构造"块注释"风格的多行注释
     */
    function vbBlockComment(lines) {
        if (!Array.isArray(lines) || lines.length === 0) return '';
        return lines.map(l => `' ${l}`).join('\n');
    }

    /**
     * 要点列表 → 可编译的 VBA 字符串表达式。
     * 旧写法把 `" & ChrW(8226)` 直接拼进 .Text，粘贴进编辑器就是语法错误。
     */
    function vbBulletText(items) {
        const list = (Array.isArray(items) ? items : [])
            .map(b => String(b == null ? '' : b).trim())
            .filter(Boolean);
        if (list.length === 0) return '""';
        return list.map(b => `ChrW(8226) & " " & ${vbStr(b)}`).join(' & vbCrLf & ');
    }

    // ============================================================
    // 配色主题定义（内置兜底）
    // 单一数据源：Config/ppt_templates.json（后端 GetPptTemplates 下发）。
    // 引擎启动时异步加载 JSON 主题并缓存；加载失败或未命中时回退到本内置兜底。
    // ============================================================
    const COLOR_THEMES_BACKUP = {
        'business_blue': {
            id: 'business_blue',
            name: '商务蓝',
            primary: [0, 90, 156],
            secondary: [0, 120, 200],
            accent: [255, 165, 0],
            bg: [245, 248, 252],
            text: [51, 51, 51],
            light: [230, 240, 250]
        },
        'tech_green': {
            id: 'tech_green',
            name: '科技绿',
            primary: [0, 128, 128],
            secondary: [0, 160, 160],
            accent: [255, 105, 97],
            bg: [240, 248, 248],
            text: [51, 51, 51],
            light: [220, 240, 240]
        },
        'vibrant_orange': {
            id: 'vibrant_orange',
            name: '活力橙',
            primary: [230, 100, 50],
            secondary: [240, 140, 80],
            accent: [50, 100, 180],
            bg: [255, 248, 240],
            text: [51, 51, 51],
            light: [255, 235, 220]
        },
        'elegant_purple': {
            id: 'elegant_purple',
            name: '优雅紫',
            primary: [100, 60, 150],
            secondary: [130, 90, 180],
            accent: [220, 180, 50],
            bg: [248, 245, 252],
            text: [51, 51, 51],
            light: [235, 225, 245]
        },
        'academic_navy': {
            id: 'academic_navy',
            name: '学术深蓝',
            primary: [15, 45, 90],
            secondary: [180, 150, 70],
            accent: [200, 80, 60],
            bg: [247, 248, 250],
            text: [40, 50, 65],
            light: [225, 232, 242]
        },
        'academic_crimson': {
            id: 'academic_crimson',
            name: '学术红',
            primary: [128, 30, 45],
            secondary: [150, 110, 60],
            accent: [60, 90, 140],
            bg: [250, 247, 243],
            text: [55, 45, 45],
            light: [240, 225, 222]
        },
        'light_fresh': {
            id: 'light_fresh',
            name: '浅色清新',
            primary: [55, 160, 140],
            secondary: [90, 170, 220],
            accent: [240, 160, 90],
            bg: [255, 255, 255],
            text: [60, 70, 70],
            light: [230, 244, 240]
        },
        'light_minimal': {
            id: 'light_minimal',
            name: '浅色极简',
            primary: [60, 65, 80],
            secondary: [120, 128, 145],
            accent: [60, 120, 200],
            bg: [255, 255, 255],
            text: [45, 50, 60],
            light: [240, 242, 246]
        }
    };

    // 运行时主题缓存：JSON 加载成功后覆盖内置兜底
    let _loadedThemes = null;

    /**
     * 版式画像归一化：兼容后端 camelCase（新）与 PascalCase（旧文件）两种键名，
     * 统一输出 camelCase 结构（title/body/table/shapes/capacity）。
     */
    function normalizeLayout(l) {
        if (!l) return null;
        const pick = (o, keys) => {
            if (!o) return null;
            for (const k of keys) if (o[k] != null) return o[k];
            return null;
        };
        const rect = o => (o ? {
            left: pick(o, ['left', 'Left']),
            top: pick(o, ['top', 'Top']),
            width: pick(o, ['width', 'Width']),
            height: pick(o, ['height', 'Height'])
        } : null);
        const tableRaw = pick(l, ['table', 'Table']);
        const shapesRaw = pick(l, ['shapes', 'Shapes']);
        const capRaw = pick(l, ['capacity', 'Capacity']);
        return {
            title: rect(pick(l, ['title', 'Title'])),
            body: rect(pick(l, ['body', 'Body'])),
            table: tableRaw ? {
                left: pick(tableRaw, ['left', 'Left']),
                top: pick(tableRaw, ['top', 'Top']),
                width: pick(tableRaw, ['width', 'Width']),
                height: pick(tableRaw, ['height', 'Height']),
                columns: pick(tableRaw, ['columns', 'Columns']),
                headerFill: pick(tableRaw, ['headerFill', 'HeaderFill'])
            } : null,
            shapes: Array.isArray(shapesRaw) ? shapesRaw.map(s => ({
                kind: pick(s, ['kind', 'Kind']),
                left: pick(s, ['left', 'Left']),
                top: pick(s, ['top', 'Top']),
                width: pick(s, ['width', 'Width']),
                height: pick(s, ['height', 'Height']),
                fill: pick(s, ['fill', 'Fill'])
            })) : null,
            capacity: capRaw ? {
                maxBullets: pick(capRaw, ['maxBullets', 'MaxBullets']),
                maxTableRows: pick(capRaw, ['maxTableRows', 'MaxTableRows']),
                maxTableCols: pick(capRaw, ['maxTableCols', 'MaxTableCols']),
                maxChartSeries: pick(capRaw, ['maxChartSeries', 'MaxChartSeries']),
                maxChartCategories: pick(capRaw, ['maxChartCategories', 'MaxChartCategories'])
            } : null
        };
    }

    /**
     * 把 ppt_templates.json 的主题条目归一化为引擎主题结构（colors → 扁平 RGB 数组）
     */
    function normalizeTheme(t) {
        const c = (t && t.colors) || {};
        const rgb = key => (Array.isArray(c[key]) && c[key].length >= 3) ? c[key].slice(0, 3) : null;
        return {
            id: t.id || 'business_blue',
            name: t.name || t.id || '商务蓝',
            primary: rgb('primary') || COLOR_THEMES_BACKUP.business_blue.primary,
            secondary: rgb('secondary') || COLOR_THEMES_BACKUP.business_blue.secondary,
            accent: rgb('accent') || COLOR_THEMES_BACKUP.business_blue.accent,
            bg: rgb('bg') || COLOR_THEMES_BACKUP.business_blue.bg,
            text: rgb('text') || COLOR_THEMES_BACKUP.business_blue.text,
            light: rgb('light') || COLOR_THEMES_BACKUP.business_blue.light,
            layout: normalizeLayout(t && t.layout)
        };
    }

    /**
     * 裸 WebView2 宿主对象（未被 BridgeAccessor 包装层覆盖的那一层）。
     * 包装层调用部分 COM 方法会抛 'e is not a function'，裸桥则一直正常。
     */
    function _rawBridge() {
        try {
            return (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) || null;
        } catch (_) { return null; }
    }

    /**
     * 本地 MCP 工具调用（只调一次、不重试）：宿主 DLL 未暴露对应 COM 方法时的回落。
     * 实测返回 {"success":true,"result":"<业务 JSON 字符串>"}，result 是嵌套字符串，需二次解析。
     */
    async function _callLocalToolOnce(toolName, argsObj) {
        try {
            const b = _rawBridge();
            if (!b || typeof b.CallLocalTool !== 'function') return null;
            const raw = await b.CallLocalTool(toolName, JSON.stringify(argsObj || {}));
            let out = raw;
            if (typeof out === 'string') { try { out = JSON.parse(out); } catch (_) { return null; } }
            if (out && typeof out.result === 'string') {
                try { out = JSON.parse(out.result); } catch (_) { return null; }
            } else if (out && Array.isArray(out.content) && out.content[0] && typeof out.content[0].text === 'string') {
                try { out = JSON.parse(out.content[0].text); } catch (_) { return null; }
            }
            return out || null;
        } catch (_) { return null; }
    }

    /**
     * 从后端 excelBridge.GetPptTemplates() 加载主题配置并缓存。
     * 宿主 DLL 未暴露该方法时（实测报 'e is not a function'）回落本地 MCP 工具 ppt_get_templates
     * （同一份 GetMergedTemplatesJson）。全部失败时保持内置兜底，不抛异常（渲染路径永远可用）。
     */
    async function loadPptTemplates() {
        try {
            let data = null;
            const bridge = window.BridgeAccessor && window.BridgeAccessor.getBridge() || (window.chrome?.webview?.hostObjects?.excelBridge || App.excelBridge);
            if (bridge && typeof bridge.GetPptTemplates === 'function') {
                try {
                    const json = await bridge.GetPptTemplates();
                    data = JSON.parse(json || '{}');
                } catch (e) { data = null; }
            }
            if (!data || !Array.isArray(data.themes) || data.themes.length === 0) {
                data = await _callLocalToolOnce('ppt_get_templates', {});
            }
            if (!data || !Array.isArray(data.themes) || data.themes.length === 0) {
                // Word 等非 PPT 宿主无主题文件：静默用内置，不刷 warn
                const _host = (window.App && App.HostContext && App.HostContext.host)
                    || (typeof URLSearchParams !== 'undefined' && new URLSearchParams(location.search).get('host'))
                    || '';
                if (String(_host).toLowerCase() === 'ppt') {
                    ChatLog.warn?.('TPL', 'GetPptTemplates 返回无有效主题，使用内置主题');
                }
                return null;
            }
            const map = {};
            for (const t of data.themes) {
                if (t && t.id) map[t.id] = normalizeTheme(t);
            }
            if (Object.keys(map).length === 0) return null;
            _loadedThemes = map;
            ChatLog.info?.('TPL', `✅ 已从后端加载 ${Object.keys(map).length} 个 PPT 主题`);
            return map;
        } catch (e) {
            ChatLog.error?.('TPL', '加载 PPT 主题失败，使用内置主题:', e);
            return null;
        }
    }

    /**
     * 解析主题：优先运行时 JSON 缓存，其次内置兜底，最后默认商务蓝。
     */
    function resolveTheme(themeId) {
        const id = themeId || 'business_blue';
        if (_loadedThemes && _loadedThemes[id]) return _loadedThemes[id];
        if (COLOR_THEMES_BACKUP[id]) return COLOR_THEMES_BACKUP[id];
        return COLOR_THEMES_BACKUP.business_blue;
    }

    /**
     * 主题列表（供 UI 配置卡渲染色板）：id + 名称 + 主色 RGB。
     * 优先运行时 JSON 缓存，其次内置兜底。
     */
    function getThemeList() {
        const map = _loadedThemes || COLOR_THEMES_BACKUP;
        return Object.keys(map).map(id => {
            const t = map[id] || {};
            return {
                id,
                name: t.name || id,
                primary: Array.isArray(t.primary) && t.primary.length >= 3 ? t.primary.slice(0, 3) : [0, 90, 156]
            };
        });
    }

    /**
     * 主题是否带版式画像（来自公司模板导入：标题/正文/表格位置 + 装饰形状）。
     */
    function hasThemeLayout(themeId) {
        const t = resolveTheme(themeId);
        return !!(t && t.layout);
    }

    /**
     * 生成主题版式画像的人类可读摘要（供 LLM 规划参考）。
     * 无版式画像时返回 null。
     */
    function getThemeLayoutSummary(themeId) {
        const t = resolveTheme(themeId);
        const l = t && t.layout;
        if (!l) return null;
        const parts = [];
        if (l.title) parts.push(`标题区(${l.title.left},${l.title.top},${l.title.width}x${l.title.height})`);
        if (l.body) parts.push(`正文区(${l.body.left},${l.body.top},${l.body.width}x${l.body.height})`);
        if (l.table) {
            const cols = Array.isArray(l.table.columns) ? l.table.columns.length : 0;
            parts.push(`表格区(${l.table.left},${l.table.top},${l.table.width}x${l.table.height},${cols}列)`);
        }
        if (Array.isArray(l.shapes) && l.shapes.length > 0) parts.push(`装饰形状${l.shapes.length}个`);
        if (l.capacity) {
            const c = l.capacity;
            parts.push(`容量上限(每页要点≤${c.maxBullets ?? 8}/表格≤${c.maxTableRows ?? 25}行${c.maxTableCols ?? 12}列/图表系列≤${c.maxChartSeries ?? 5}/分类≤${c.maxChartCategories ?? 8})`);
        }
        return parts.length > 0 ? parts.join('，') : null;
    }

    // 启动即异步预热主题缓存（不阻塞渲染路径）
    loadPptTemplates();

    // ============================================================
    // 模板库
    // ============================================================
    const TEMPLATES = [];

    /**
     * 模板 1：从 JSON 描述生成 PPT 报告（带设计功能）
     * 架构：LLM 输出 JSON 描述 PPT 结构 → 模板引擎渲染为可执行 VBA
     * 优势：VBA 代码由人工审核过的模板生成，保证正确性；LLM 只负责理解内容和组织页面
     * 设计功能：配色主题、背景色、顶部色带、分隔线、页脚、图表支持、双栏布局
     */
    TEMPLATES.push({
        id: 'ppt_from_data',
        title: '从 JSON 描述生成 PPT 报告',
        keywords: ['ppt', 'powerpoint', '幻灯片', '演示文稿', '汇报'],
        match(intent) {
            // 匹配条件：包含 PPT 关键词且 slides 数组存在
            if (!/ppt|powerpoint|幻灯片|演示文稿|汇报/i.test(intent.raw || '')) return false;
            // 如果 intent 中有 slides 数组，优先匹配此模板
            if (Array.isArray(intent.slides) && intent.slides.length > 0) return true;
            // 如果有 reportTitle 也匹配
            if (intent.reportTitle) return true;
            return false;
        },
        render(intent, ctx) {
            const reportTitle = intent.reportTitle || '数据分析报告';
            const subtitle = intent.subtitle || '';
            const dataSheet = intent.dataSheet || ctx.activeSheet || 'Sheet1';
            const dataRange = intent.dataRange || '';
            const slides = Array.isArray(intent.slides) ? intent.slides : [];
            const theme = resolveTheme(intent.theme);

            if (slides.length === 0) {
                slides.push({ type: 'title', title: reportTitle, subtitle: subtitle });
                slides.push({ type: 'text', title: '数据概览', bullets: ['数据加载中...'] });
            }

            // 生成配色主题初始化代码
            const themeInitCode = `    ' 配色主题：${theme.name}
    Dim primaryColor As Long, secondaryColor As Long, accentColor As Long
    Dim bgColor As Long, textColor As Long, lightColor As Long
    primaryColor = RGB(${theme.primary[0]}, ${theme.primary[1]}, ${theme.primary[2]})
    secondaryColor = RGB(${theme.secondary[0]}, ${theme.secondary[1]}, ${theme.secondary[2]})
    accentColor = RGB(${theme.accent[0]}, ${theme.accent[1]}, ${theme.accent[2]})
    bgColor = RGB(${theme.bg[0]}, ${theme.bg[1]}, ${theme.bg[2]})
    textColor = RGB(${theme.text[0]}, ${theme.text[1]}, ${theme.text[2]})
    lightColor = RGB(${theme.light[0]}, ${theme.light[1]}, ${theme.light[2]})`;

            // 生成每张幻灯片的 VBA 代码块
            const slideCodeBlocks = slides.map((slide, idx) => {
                const slideIndex = idx + 1;
                switch (slide.type) {
                    case 'title':
                        return `    ' --- 幻灯片 ${slideIndex}：标题页 ---
    Set pptSlide = pptPres.Slides.Add(pptPres.Slides.Count + 1, 12)

    ' 背景色
    pptSlide.Background.Fill.ForeColor.RGB = bgColor

    ' 主标题
    Set tb = pptSlide.Shapes.AddTextbox(1, 50, 120, 860, 100)
    With tb.TextFrame.TextRange
        .Text = ${vbStr(slide.title || reportTitle)}
        .Font.Size = 44
        .Font.Bold = -1
        .Font.Color.RGB = primaryColor
        .Font.Name = "微软雅黑"
        .ParagraphFormat.Alignment = 2
    End With

    ' 副标题
    Set tb = pptSlide.Shapes.AddTextbox(1, 150, 240, 660, 50)
    With tb.TextFrame.TextRange
        .Text = ${vbStr(slide.subtitle || '')}
        .Font.Size = 18
        .Font.Color.RGB = secondaryColor
        .Font.Name = "微软雅黑"
        .ParagraphFormat.Alignment = 2
    End With

    ' 装饰线
    Set tb = pptSlide.Shapes.AddShape(1, 380, 310, 200, 4)
    tb.Fill.ForeColor.RGB = accentColor
    tb.Line.Visible = 0

    ' 页脚
    Set tb = pptSlide.Shapes.AddTextbox(1, 50, 500, 860, 30)
    With tb.TextFrame.TextRange
        .Text = "由 Excel AI 自动生成 · " & Format(Date, "yyyy年mm月dd日")
        .Font.Size = 12
        .Font.Color.RGB = RGB(150, 150, 150)
        .Font.Name = "微软雅黑"
        .ParagraphFormat.Alignment = 2
    End With`;

                    case 'table': {
                        const cols = Array.isArray(slide.columns) ? slide.columns : [];
                        const rowCount = Math.min(Number(slide.rowCount) || 10, 20);
                        const startRow = Number(slide.startRow) || 2;
                        const tableCols = cols.length || 5;
                        const tableRows = rowCount + 1;

                        const colNamesVBA = cols.length > 0
                            ? cols.map(c => vbStr(c)).join(', ')
                            : Array.from({ length: tableCols }, (_, i) => vbStr(`列${i + 1}`)).join(', ');

                        return `    ' --- 幻灯片 ${slideIndex}：数据表格 ---
    Set pptSlide = pptPres.Slides.Add(pptPres.Slides.Count + 1, 12)
    pptSlide.Background.Fill.ForeColor.RGB = bgColor

    ' 标题
    Set tb = pptSlide.Shapes.AddTextbox(1, 50, 20, 860, 50)
    With tb.TextFrame.TextRange
        .Text = ${vbStr(slide.title || '')}
        .Font.Size = 32
        .Font.Bold = -1
        .Font.Color.RGB = primaryColor
        .Font.Name = "微软雅黑"
    End With

    ' 顶部色带
    Set tb = pptSlide.Shapes.AddShape(1, 0, 0, 960, 8)
    tb.Fill.ForeColor.RGB = primaryColor
    tb.Line.Visible = 0

    ' 分隔线
    Set tb = pptSlide.Shapes.AddShape(1, 50, 75, 860, 2)
    tb.Fill.ForeColor.RGB = secondaryColor
    tb.Line.Visible = 0

    ' 数据表格
    Set tblShape = pptSlide.Shapes.AddTable(${tableRows}, ${tableCols}, 50, 90, 860, 400)
    Set tbl = tblShape.Table
    colNames = Array(${colNamesVBA})
    For ci = 0 To UBound(colNames)
        tbl.Cell(1, ci + 1).Shape.Fill.ForeColor.RGB = primaryColor
        tbl.Cell(1, ci + 1).Shape.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
        tbl.Cell(1, ci + 1).Shape.TextFrame.TextRange.Text = CStr(colNames(ci))
        tbl.Cell(1, ci + 1).Shape.TextFrame.TextRange.Font.Name = "微软雅黑"
        tbl.Cell(1, ci + 1).Shape.TextFrame.TextRange.Font.Size = 12
        tbl.Cell(1, ci + 1).Shape.TextFrame.TextRange.Font.Bold = -1
    Next ci
    Set ws = ThisWorkbook.Worksheets(${vbStr(dataSheet)})
    If ${vbStr(dataRange)} <> "" Then
        Set rng = ws.Range(${vbStr(dataRange)})
    Else
        Set rng = ws.UsedRange
    End If
    data = rng.Value
    If Not IsArray(data) Then
        nRows = 1
        nCols = 1
    Else
        nRows = UBound(data, 1)
        nCols = 1
        On Error Resume Next
        nCols = UBound(data, 2)
        If Err.Number <> 0 Then nCols = 1: Err.Clear
        On Error GoTo ErrHandler
    End If
    destRow = 2
    For ri = ${startRow} To nRows
        If destRow > ${tableRows} Then Exit For
        For ci = 1 To nCols
            If ci <= ${tableCols} Then
                If IsArray(data) Then
                    tbl.Cell(destRow, ci).Shape.TextFrame.TextRange.Text = CStr(data(ri, ci))
                Else
                    tbl.Cell(destRow, ci).Shape.TextFrame.TextRange.Text = CStr(data)
                End If
                tbl.Cell(destRow, ci).Shape.TextFrame.TextRange.Font.Name = "微软雅黑"
                tbl.Cell(destRow, ci).Shape.TextFrame.TextRange.Font.Size = 11
                If destRow Mod 2 = 0 Then
                    tbl.Cell(destRow, ci).Shape.Fill.ForeColor.RGB = lightColor
                End If
            End If
        Next ci
        destRow = destRow + 1
    Next ri

    ' 页脚
    Set tb = pptSlide.Shapes.AddTextbox(1, 50, 520, 860, 20)
    With tb.TextFrame.TextRange
        .Text = "第 " & pptPres.Slides.Count & " 页"
        .Font.Size = 10
        .Font.Color.RGB = RGB(150, 150, 150)
        .Font.Name = "微软雅黑"
        .ParagraphFormat.Alignment = 2
    End With`;
                    }

                    case 'text': {
                        const bulletVBA = vbBulletText(slide.bullets);
                        return `    ' --- 幻灯片 ${slideIndex}：文本要点 ---
    Set pptSlide = pptPres.Slides.Add(pptPres.Slides.Count + 1, 12)
    pptSlide.Background.Fill.ForeColor.RGB = bgColor

    ' 标题
    Set tb = pptSlide.Shapes.AddTextbox(1, 50, 20, 860, 50)
    With tb.TextFrame.TextRange
        .Text = ${vbStr(slide.title || '')}
        .Font.Size = 32
        .Font.Bold = -1
        .Font.Color.RGB = primaryColor
        .Font.Name = "微软雅黑"
    End With

    ' 顶部色带
    Set tb = pptSlide.Shapes.AddShape(1, 0, 0, 960, 8)
    tb.Fill.ForeColor.RGB = primaryColor
    tb.Line.Visible = 0

    ' 分隔线
    Set tb = pptSlide.Shapes.AddShape(1, 50, 75, 860, 2)
    tb.Fill.ForeColor.RGB = secondaryColor
    tb.Line.Visible = 0

    ' 内容文本框
    Set tb = pptSlide.Shapes.AddTextbox(1, 50, 90, 860, 400)
    With tb.TextFrame.TextRange
        .Text = ${bulletVBA}
        .Font.Size = 16
        .Font.Color.RGB = textColor
        .Font.Name = "微软雅黑"
        .ParagraphFormat.SpaceAfter = 12
    End With

    ' 页脚
    Set tb = pptSlide.Shapes.AddTextbox(1, 50, 520, 860, 20)
    With tb.TextFrame.TextRange
        .Text = "第 " & pptPres.Slides.Count & " 页"
        .Font.Size = 10
        .Font.Color.RGB = RGB(150, 150, 150)
        .Font.Name = "微软雅黑"
        .ParagraphFormat.Alignment = 2
    End With`;
                    }

                    case 'chart': {
                        const chartType = slide.chartType || 'bar';
                        const chartTypes = {
                            'bar': 51,
                            'pie': 5,
                            'line': 4
                        };
                        const pptChartType = chartTypes[chartType] || 51;
                        return `    ' --- 幻灯片 ${slideIndex}：${chartType === 'bar' ? '柱状图' : chartType === 'pie' ? '饼图' : '折线图'} ---
    Set pptSlide = pptPres.Slides.Add(pptPres.Slides.Count + 1, 12)
    pptSlide.Background.Fill.ForeColor.RGB = bgColor

    ' 标题
    Set tb = pptSlide.Shapes.AddTextbox(1, 50, 20, 860, 50)
    With tb.TextFrame.TextRange
        .Text = ${vbStr(slide.title || '')}
        .Font.Size = 32
        .Font.Bold = -1
        .Font.Color.RGB = primaryColor
        .Font.Name = "微软雅黑"
    End With

    ' 顶部色带
    Set tb = pptSlide.Shapes.AddShape(1, 0, 0, 960, 8)
    tb.Fill.ForeColor.RGB = primaryColor
    tb.Line.Visible = 0

    ' 分隔线
    Set tb = pptSlide.Shapes.AddShape(1, 50, 75, 860, 2)
    tb.Fill.ForeColor.RGB = secondaryColor
    tb.Line.Visible = 0

    ' AddChart2(Style, Type, Left, Top, Width, Height)。Style=-1 用默认样式，第二参才是图表类型
    Set chartShape = pptSlide.Shapes.AddChart2(-1, ${pptChartType}, 50, 90, 860, 400)
    Set chartData = chartShape.Chart.ChartData
    chartData.Activate
    Set chartWorkbook = chartData.Workbook
    Set ws = ThisWorkbook.Worksheets(${vbStr(dataSheet)})
    If ${vbStr(dataRange)} <> "" Then
        Set rng = ws.Range(${vbStr(dataRange)})
    Else
        Set rng = ws.UsedRange
    End If
    rng.Copy
    chartWorkbook.Worksheets(1).Range("A1").PasteSpecial -4163
    Application.CutCopyMode = False
    chartWorkbook.Close True
    chartShape.Chart.HasTitle = True
    chartShape.Chart.ChartTitle.Text = ${vbStr(slide.title || '')}
    chartShape.Chart.ChartTitle.Font.Size = 18
    chartShape.Chart.ChartTitle.Font.Color.RGB = primaryColor`;
                    }

                    case 'two_column': {
                        const leftBulletVBA = vbBulletText(slide.leftBullets);
                        const rightBulletVBA = vbBulletText(slide.rightBullets);
                        return `    ' --- 幻灯片 ${slideIndex}：双栏布局 ---
    Set pptSlide = pptPres.Slides.Add(pptPres.Slides.Count + 1, 12)
    pptSlide.Background.Fill.ForeColor.RGB = bgColor

    ' 标题
    Set tb = pptSlide.Shapes.AddTextbox(1, 50, 20, 860, 50)
    With tb.TextFrame.TextRange
        .Text = ${vbStr(slide.title || '')}
        .Font.Size = 32
        .Font.Bold = -1
        .Font.Color.RGB = primaryColor
        .Font.Name = "微软雅黑"
    End With

    ' 顶部色带
    Set tb = pptSlide.Shapes.AddShape(1, 0, 0, 960, 8)
    tb.Fill.ForeColor.RGB = primaryColor
    tb.Line.Visible = 0

    ' 分隔线
    Set tb = pptSlide.Shapes.AddShape(1, 50, 75, 860, 2)
    tb.Fill.ForeColor.RGB = secondaryColor
    tb.Line.Visible = 0

    ' 左栏
    Set tb = pptSlide.Shapes.AddTextbox(1, 50, 90, 420, 400)
    With tb.TextFrame.TextRange
        .Text = ${leftBulletVBA}
        .Font.Size = 14
        .Font.Color.RGB = textColor
        .Font.Name = "微软雅黑"
    End With

    ' 中间分隔线
    Set tb = pptSlide.Shapes.AddShape(1, 475, 90, 2, 400)
    tb.Fill.ForeColor.RGB = lightColor
    tb.Line.Visible = 0

    ' 右栏
    Set tb = pptSlide.Shapes.AddTextbox(1, 490, 90, 420, 400)
    With tb.TextFrame.TextRange
        .Text = ${rightBulletVBA}
        .Font.Size = 14
        .Font.Color.RGB = textColor
        .Font.Name = "微软雅黑"
    End With

    ' 页脚
    Set tb = pptSlide.Shapes.AddTextbox(1, 50, 520, 860, 20)
    With tb.TextFrame.TextRange
        .Text = "第 " & pptPres.Slides.Count & " 页"
        .Font.Size = 10
        .Font.Color.RGB = RGB(150, 150, 150)
        .Font.Name = "微软雅黑"
        .ParagraphFormat.Alignment = 2
    End With`;
                    }

                    default:
                        return `    ' --- 幻灯片 ${slideIndex}：未知类型 ${slide.type} ---`;
                }
            }).join('\n\n');

            // 生成安全的文件名
            const safeName = reportTitle.replace(/[^\w\u4e00-\u9fa5]/g, '_');

            return `' ============================================================
' PPT 报告生成器 — 由模板引擎渲染（带设计功能）
' 报告标题：${reportTitle}
' 数据源：${dataSheet}${dataRange ? '!' + dataRange : ''}
' 配色主题：${theme.name}
' 幻灯片数量：${slides.length}
' ============================================================
Sub GeneratePPTReport()
    Dim pptApp As Object, pptPres As Object, pptSlide As Object
    Dim tb As Object, tbl As Object, tblShape As Object
    Dim ws As Object, rng As Object, data As Variant
    Dim colNames As Variant
    Dim ri As Long, ci As Long, destRow As Long
    Dim nRows As Long, nCols As Long
    Dim chartShape As Object, chartData As Object, chartWorkbook As Object
    Dim startedPpt As Boolean
    Dim savePath As String

    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    ' 已打开的 PowerPoint 只借用，不要 Quit，否则会关掉用户正在看的文稿
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    If pptApp Is Nothing Then
        Set pptApp = CreateObject("PowerPoint.Application")
        startedPpt = True
    Else
        startedPpt = False
    End If
    Err.Clear
    On Error GoTo ErrHandler
    If pptApp Is Nothing Then
        MsgBox "无法创建 PowerPoint 对象，请确认已安装 PowerPoint", vbCritical
        GoTo CleanExit
    End If
    pptApp.Visible = True
    Set pptPres = pptApp.Presentations.Add
    If pptPres Is Nothing Then
        MsgBox "无法创建演示文稿", vbCritical
        GoTo CleanExit
    End If

${themeInitCode}

${slideCodeBlocks}

    ' 未保存的工作簿没有 Path，改存到桌面，避免生成 "\\报表.pptx" 这种非法路径
    If Len(ThisWorkbook.Path) = 0 Then
        savePath = Environ("USERPROFILE") & "\\Desktop\\${safeName}_" & Format(Date, "yyyymmdd") & ".pptx"
    Else
        savePath = ThisWorkbook.Path & "\\${safeName}_" & Format(Date, "yyyymmdd") & ".pptx"
    End If
    On Error Resume Next
    pptPres.SaveAs savePath
    If Err.Number <> 0 Then
        MsgBox "PPT 已打开，但保存失败：" & Err.Description, vbExclamation
        Err.Clear
    Else
        MsgBox "PPT 已生成：" & savePath, vbInformation, "完成"
    End If
    On Error GoTo ErrHandler

CleanExit:
    On Error Resume Next
    If Not pptPres Is Nothing Then pptPres.Close
    If startedPpt And Not pptApp Is Nothing Then pptApp.Quit
    Set pptSlide = Nothing: Set pptPres = Nothing: Set pptApp = Nothing
    Set tbl = Nothing: Set tblShape = Nothing
    Set tb = Nothing: Set ws = Nothing: Set rng = Nothing
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    MsgBox "错误 " & Err.Number & "：" & Err.Description, vbCritical
    Resume CleanExit
End Sub
`;
        }
    });

    /**
     * 模板 2：批量重命名工作表
     */
    TEMPLATES.push({
        id: 'rename_sheets_batch',
        title: '批量重命名工作表',
        keywords: ['批量', '重命名', '工作表', 'sheet'],
        match(intent) {
            return /批量.*重命名|重命名.*工作表|sheet.*重命名|rename.*sheet/i.test(intent.raw || '');
        },
        render(intent, ctx) {
            const pattern = intent.pattern || 'Sheet{i}';
            return `' ============================================================
' 批量重命名工作表
' 命名规则：${pattern}
' 运行模式：手动交互（结束后弹出汇总）
' ============================================================
Sub RenameSheetsBatch()
    Dim ws As Worksheet
    Dim i As Long
    Dim renamedCount As Long
    Dim runLog As String
    Dim newName As String

    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    i = 0
    renamedCount = 0
    runLog = ""

    ' 关键：用 For Each 而非 For i，避免删除工作表后索引错位
    For Each ws In ThisWorkbook.Worksheets
        i = i + 1
        newName = Replace(${vbStr(pattern)}, "{i}", CStr(i))
        newName = Trim(newName)
        newName = Replace(newName, ":", "")
        newName = Replace(newName, "\\", "")
        newName = Replace(newName, "/", "")
        newName = Replace(newName, "?", "")
        newName = Replace(newName, "*", "")
        newName = Replace(newName, "[", "")
        newName = Replace(newName, "]", "")
        If Len(newName) = 0 Then newName = "Sheet" & CStr(i)
        If Len(newName) > 31 Then newName = Left(newName, 31)
        On Error Resume Next
        ws.Name = newName
        If Err.Number = 0 Then
            renamedCount = renamedCount + 1
        Else
            runLog = runLog & "第" & i & "个工作表重命名失败：" & Err.Description & vbCrLf
            Err.Clear
        End If
        On Error GoTo ErrHandler
    Next ws

    MsgBox "已重命名 " & renamedCount & " 个工作表" & vbCrLf & runLog, vbInformation, "完成"

CleanExit:
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "运行时错误：" & Err.Description, vbCritical
End Sub
`;
        }
    });

    /**
     * 模板 3：按列值拆分工作表
     */
    TEMPLATES.push({
        id: 'split_sheets_by_column',
        title: '按列值拆分工作表',
        keywords: ['拆分', '按列', '分组', 'split'],
        match(intent) {
            return /按.*拆分|按.*分组|按.*列.*分|按.*拆.*表|split.*sheet/i.test(intent.raw || '');
        },
        render(intent, ctx) {
            const dataSheet = intent.dataSheet || 'Sheet1';
            const headerRow = intent.headerRow != null ? Number(intent.headerRow) : 1;
            const columnIndex = intent.columnIndex || 1;
            return `' ============================================================
' 按第 ${columnIndex} 列的值拆分到不同工作表
' 运行模式：后台自动（不弹窗，结束时输出汇总）
' ============================================================
Sub SplitSheetsByColumn()
    Dim srcWs As Worksheet
    Dim newWs As Worksheet
    Dim srcRange As Range
    Dim data As Variant
    Dim headers As Variant
    Dim uniqueValues As Object
    Dim key As Variant
    Dim i As Long, j As Long
    Dim lastRow As Long
    Dim runLog As String
    Dim colIdx As Long
    Dim col As Long
    Dim created As Long
    Dim sheetName As String
    Dim destRow As Long
    Dim headerRowNum As Long
    Dim firstCol As Long
    Dim lastCol As Long
    Dim lastDataRow As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo ErrHandler

    colIdx = ${Number(columnIndex) > 0 ? Number(columnIndex) : 1}

    ' --- 1. 按工作表行号取表头和数据（不用 UsedRange 的相对行，避免表头不在第 1 行时取错） ---
    Set srcWs = ThisWorkbook.Worksheets(${vbStr(dataSheet)})
    Set srcRange = srcWs.UsedRange
    If srcRange Is Nothing Then
        MsgBox "没有可拆分的数据", vbExclamation
        GoTo CleanExit
    End If
    headerRowNum = ${Number.isFinite(Number(headerRow)) ? Number(headerRow) : 1}
    If headerRowNum < srcRange.Row Then headerRowNum = srcRange.Row
    firstCol = srcRange.Column
    lastCol = srcRange.Column + srcRange.Columns.Count - 1
    lastDataRow = srcRange.Row + srcRange.Rows.Count - 1
    If lastDataRow <= headerRowNum Then
        MsgBox "数据不足", vbExclamation
        GoTo CleanExit
    End If

    headers = srcWs.Range(srcWs.Cells(headerRowNum, firstCol), srcWs.Cells(headerRowNum, lastCol)).Value
    data = srcWs.Range(srcWs.Cells(headerRowNum + 1, firstCol), srcWs.Cells(lastDataRow, lastCol)).Value
    If Not IsArray(data) Then
        MsgBox "数据不足", vbExclamation
        GoTo CleanExit
    End If
    lastRow = UBound(data, 1)
    If colIdx > UBound(data, 2) Then
        MsgBox "拆分列超出数据范围", vbExclamation
        GoTo CleanExit
    End If

    ' --- 2. 收集唯一值 ---
    Set uniqueValues = CreateObject("Scripting.Dictionary")
    For i = 1 To lastRow
        key = CStr(data(i, colIdx))
        If Not uniqueValues.Exists(key) Then uniqueValues.Add key, 0
    Next i

    ' --- 3. 为每个唯一值创建新工作表 ---
    runLog = ""
    created = 0
    For Each key In uniqueValues.Keys
        If IsError(key) Then GoTo NextKey
        sheetName = Trim(CStr(key))
        sheetName = Replace(sheetName, ":", "")
        sheetName = Replace(sheetName, "\\", "")
        sheetName = Replace(sheetName, "/", "")
        sheetName = Replace(sheetName, "?", "")
        sheetName = Replace(sheetName, "*", "")
        sheetName = Replace(sheetName, "[", "")
        sheetName = Replace(sheetName, "]", "")
        If Len(sheetName) = 0 Then sheetName = "分组" & CStr(created + 1)
        If Len(sheetName) > 31 Then sheetName = Left(sheetName, 31)
        On Error Resume Next
        Set newWs = ThisWorkbook.Worksheets(sheetName)
        On Error GoTo ErrHandler
        If newWs Is Nothing Then
            Set newWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            newWs.Name = sheetName
            created = created + 1
        End If

        If IsArray(headers) Then
            For col = 1 To UBound(headers, 2)
                newWs.Cells(1, col).Value = headers(1, col)
            Next col
        Else
            newWs.Cells(1, 1).Value = headers
        End If

        destRow = 2
        For i = 1 To lastRow
            If CStr(data(i, colIdx)) = CStr(key) Then
                For j = 1 To UBound(data, 2)
                    newWs.Cells(destRow, j).Value = data(i, j)
                Next j
                destRow = destRow + 1
            End If
        Next i
        Set newWs = Nothing
NextKey:
    Next key

    MsgBox "已创建 " & created & " 个工作表", vbInformation, "完成"

CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "运行时错误 " & Err.Number & "：" & Err.Description, vbCritical
End Sub
`;
        }
    });

    /**
     * 模板 4：遍历文件夹合并 CSV
     */
    TEMPLATES.push({
        id: 'merge_csv_files',
        title: '批量合并 CSV 文件',
        keywords: ['合并', 'csv', '文件夹', '遍历', 'merge'],
        match(intent) {
            return /合并.*csv|csv.*合并|文件夹.*csv|遍历.*文件|merge.*csv/i.test(intent.raw || '');
        },
        render(intent, ctx) {
            const folderPath = intent.folderPath || 'C:\\YourFolder';
            return `' ============================================================
' 遍历文件夹，合并所有 CSV 到当前工作簿
' 运行模式：后台自动（不弹窗，写日志到 Sheet1）
' ============================================================
Sub MergeCsvFiles()
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim ts As Object
    Dim ws As Worksheet
    Dim line As String
    Dim fields() As String
    Dim destRow As Long
    Dim fileCount As Long
    Dim runLog As String
    Dim k As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrHandler

    ' --- 1. 准备输出工作表 ---
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("合并结果")
    On Error GoTo ErrHandler
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "合并结果"
    End If
    ws.Cells.Clear
    destRow = 1

    ' --- 2. 遍历文件夹 ---
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(${vbStr(folderPath)})
    If folder Is Nothing Then
        MsgBox "文件夹不存在：" & ${vbStr(folderPath)}, vbCritical
        GoTo CleanExit
    End If

    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "csv" Then
            Set ts = fso.OpenTextFile(file.Path, 1, False, -2)
            Do While Not ts.AtEndOfStream
                line = ts.ReadLine
                If Len(Trim(line)) > 0 Then
                    fields = Split(line, ",")
                    For k = LBound(fields) To UBound(fields)
                        ws.Cells(destRow, k + 1).Value = fields(k)
                    Next k
                    destRow = destRow + 1
                End If
            Loop
            ts.Close
            Set ts = Nothing
            fileCount = fileCount + 1
        End If
    Next file

    MsgBox "已合并 " & fileCount & " 个 CSV 文件，共 " & destRow - 1 & " 行数据", vbInformation, "完成"

CleanExit:
    Set fso = Nothing
    Set folder = Nothing
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "运行时错误 " & Err.Number & "：" & Err.Description, vbCritical
End Sub
`;
        }
    });

    /**
     * 模板 5：数据格式批量调整
     */
    TEMPLATES.push({
        id: 'batch_format',
        title: '批量格式化选中区域',
        keywords: ['格式化', '批量格式', '批量设置', 'format'],
        match(intent) {
            return /格式化|批量.*格式|批量.*设置.*格式/i.test(intent.raw || '');
        },
        render(intent, ctx) {
            return `' ============================================================
' 批量格式化选中区域
' 运行模式：手动交互（需先选中区域）
' ============================================================
Sub BatchFormat()
    Dim selRng As Range
    Dim scrUpd As Boolean
    Dim calcMode As Long

    scrUpd = Application.ScreenUpdating
    calcMode = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrHandler

    Set selRng = Selection
    If Not TypeOf selRng Is Range Then
        MsgBox "请先选中需要格式化的区域", vbExclamation
        GoTo CleanExit
    End If

    With selRng
        .Font.Name = "微软雅黑"
        .Font.Size = 11
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(180, 180, 180)
        .Columns.AutoFit
    End With

CleanExit:
    Application.ScreenUpdating = scrUpd
    Application.Calculation = calcMode
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = scrUpd
    Application.Calculation = calcMode
    MsgBox "运行时错误：" & Err.Description, vbCritical
End Sub
`;
        }
    });

    // ============================================================
    // 模板匹配 + 渲染
    // ============================================================

    /**
     * 从 LLM 输出 / 用户描述中提取 IntentSpec
     * IntentSpec 是结构化 JSON，LLM 只负责填充以下字段：
     *   - scenario: 场景标识（可选，由模板自身识别）
     *   - dataSheet / dataRange / column / columnIndex / fields: 业务参数
     *   - raw: 原始用户描述，用于关键词匹配
     */
    function buildIntentFromLLM(llmOutput, userMessage, ctx) {
        let intent = null;
        if (typeof llmOutput === 'string') {
            // 尝试解析 ```json ... ``` 块
            const m = llmOutput.match(/```(?:json)?\s*([\s\S]*?)```/);
            const text = m ? m[1].trim() : llmOutput.trim();
            try {
                intent = JSON.parse(text);
            } catch (e) {
                // 不是合法 JSON，用 fallback
            }
        } else if (llmOutput && typeof llmOutput === 'object') {
            intent = llmOutput;
        }

        if (!intent || typeof intent !== 'object') intent = {};
        intent.raw = userMessage || intent.raw || '';
        // 兜底：从用户消息中识别 sheet 名
        if (!intent.dataSheet && ctx?.activeSheet) {
            intent.dataSheet = ctx.activeSheet;
        }
        return intent;
    }

    /**
     * 匹配最佳模板
     * 优先级：scenario 精确匹配 > keywords 命中数 > fallback
     */
    function pickTemplate(intent) {
        // 1. scenario 精确匹配
        if (intent.scenario) {
            const exact = TEMPLATES.find(t => t.id === intent.scenario);
            if (exact) return exact;
        }

        // 带 slides / reportTitle 的 JSON 就是 PPT 模板的输入，不必再靠关键词碰运气
        if ((Array.isArray(intent.slides) && intent.slides.length > 0) || intent.reportTitle) {
            const ppt = TEMPLATES.find(t => t.id === 'ppt_from_data');
            if (ppt) return ppt;
        }

        // 2. 只用各模板自己的 match()。
        //    不能按单个关键词计数：几乎每句都有“工作表”，会把“把 A1 加粗”渲染成重命名宏。
        const matched = [];
        for (const t of TEMPLATES) {
            try {
                if (typeof t.match === 'function' && t.match(intent)) matched.push(t);
            } catch (_) { /* 单条 match 失败不影响其它模板 */ }
        }
        if (matched.length === 0) return null;
        if (matched.length === 1) return matched[0];

        let best = matched[0];
        let bestHits = -1;
        for (const t of matched) {
            const hits = (t.keywords || []).filter(kw => {
                try { return new RegExp(kw, 'i').test(intent.raw || ''); }
                catch (_) { return false; }
            }).length;
            if (hits > bestHits) {
                bestHits = hits;
                best = t;
            }
        }
        return best;
    }

    /**
     * 把模型或模板产出整理成可以原样粘进 VBA 编辑器的标准模块。
     * 只做不改变业务含义的整理：去围栏和说明文字、弯引号、过程外的 Attribute、
     * 重复 Dim、缺失的 End Sub / 被引用却不存在的错误标签。
     */
    function toRunnableVba(raw) {
        let code = extractVbaBody(raw);
        if (!code) return '';
        code = normalizeVbaPunctuation(code);
        code = code.replace(/^[ \t]*Attribute\s+VB_\w+\s*=.*$/gim, '');
        code = hoistOptionExplicit(code);
        code = dedupeProcedureDims(code);
        code = ensureProcedureWrapper(code);
        if (!code) return '';
        code = ensureReferencedLabels(code);
        return code.replace(/\n{3,}/g, '\n\n').trim() + '\n';
    }

    function extractVbaBody(raw) {
        if (raw == null) return '';
        let s = String(raw).replace(/^\uFEFF/, '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');
        const fenced = [...s.matchAll(/```(?:vba|vb|visualbasic)?[ \t]*\n?([\s\S]*?)```/gi)].map(m => m[1]);
        const fencedProc = fenced.find(block => /\b(?:Sub|Function)\s+\S+/i.test(block));
        if (fencedProc) s = fencedProc;
        else if (fenced.length === 1) s = fenced[0];

        const lines = s.split('\n');
        const isCodeStart = (line) => /^(?:'|Rem\b|Option\s+Explicit\b|Attribute\s+VB_|((?:Public|Private)\s+)?(?:Sub|Function|Dim|Const|Enum|Type)\b)/i.test(line.trim());
        let start = 0;
        while (start < lines.length && lines[start].trim() === '') start++;
        if (start < lines.length && !isCodeStart(lines[start])) {
            const idx = lines.findIndex(l => /^(?:Option\s+Explicit\b|(?:(?:Public|Private)\s+)?(?:Sub|Function)\s+)/i.test(l.trim()));
            if (idx >= 0) start = idx;
        }
        let body = lines.slice(start).join('\n').trim();
        const endRe = /\bEnd\s+(?:Sub|Function|Property)\b/gi;
        let last = -1;
        let m;
        while ((m = endRe.exec(body))) last = m.index + m[0].length;
        if (last > 0) {
            const tail = body.slice(last);
            const tailMeaningful = tail.split('\n').some(l => {
                const t = l.trim();
                return t && !t.startsWith("'") && !/^Rem\b/i.test(t);
            });
            if (tailMeaningful && !/\b(?:Sub|Function)\s+/i.test(tail)) body = body.slice(0, last);
        }
        return body.trim();
    }

    function normalizeVbaPunctuation(code) {
        const out = [];
        let inString = false;
        for (let i = 0; i < code.length; i++) {
            const ch = code[i];
            if (inString) {
                out.push(ch === '“' || ch === '”' ? '"' : ch);
                if (ch === '"') {
                    if (code[i + 1] === '"') { out.push('"'); i++; }
                    else inString = false;
                }
                continue;
            }
            if (ch === '"') { inString = true; out.push(ch); continue; }
            if (ch === '“' || ch === '”') { out.push('"'); continue; }
            if (ch === '‘' || ch === '’') { out.push("'"); continue; }
            if (ch === '（') { out.push('('); continue; }
            if (ch === '）') { out.push(')'); continue; }
            if (ch === '，') { out.push(','); continue; }
            if (ch === '：') { out.push(':'); continue; }
            if (ch === '；') { out.push('\n'); continue; }
            out.push(ch);
        }
        return out.join('');
    }

    function hoistOptionExplicit(code) {
        const lines = code.split('\n').filter(l => !/^\s*Option\s+Explicit\b/i.test(l));
        if (!/^\s*Option\s+Explicit\b/im.test(code)) return lines.join('\n').trim();
        return 'Option Explicit\n\n' + lines.join('\n').replace(/^\n+/, '');
    }

    function splitProcedures(code) {
        const re = /^([ \t]*(?:(?:Public|Private)\s+)?(?:Sub|Function)\s+[^\n]+)$/gim;
        const starts = [];
        let m;
        while ((m = re.exec(code))) starts.push({ index: m.index, header: m[1] });
        if (starts.length === 0) return [{ header: '', body: code, index: 0 }];
        const parts = [];
        if (starts[0].index > 0) parts.push({ header: '', body: code.slice(0, starts[0].index), index: 0 });
        for (let i = 0; i < starts.length; i++) {
            const from = starts[i].index;
            const to = i + 1 < starts.length ? starts[i + 1].index : code.length;
            parts.push({ header: starts[i].header, body: code.slice(from, to), index: from });
        }
        return parts;
    }

    function dimNames(line) {
        const m = line.match(/^(\s*)((?:Private\s+|Public\s+)?)Dim\s+(.+)$/i);
        if (!m) return null;
        const decl = m[3].replace(/'.*$/, '').trim();
        if (!decl) return null;
        const items = decl.split(',').map(p => p.trim()).filter(Boolean);
        const names = items.map(p => (p.split(/\s+/)[0] || '').replace(/[()]/g, '')).filter(Boolean);
        return { indent: m[1], scope: m[2] || '', items, names };
    }

    function dedupeProcedureDims(code) {
        return splitProcedures(code).map(part => {
            if (!part.header) return part.body;
            const seen = new Set();
            const lines = part.body.split('\n').map(line => {
                const parsed = dimNames(line);
                if (!parsed) return line;
                const fresh = [];
                parsed.items.forEach((item, idx) => {
                    const name = parsed.names[idx];
                    const key = name.toLowerCase();
                    if (!name || seen.has(key)) return;
                    seen.add(key);
                    fresh.push(item);
                });
                if (fresh.length === 0) return `${parsed.indent}' 已去掉重复声明 ${parsed.names.join(', ')}`;
                if (fresh.length === parsed.items.length) return line;
                return `${parsed.indent}${parsed.scope}Dim ${fresh.join(', ')}`;
            });
            return lines.join('\n');
        }).join('');
    }

    function looksLikeVbaStatements(code) {
        return /^\s*(?:Dim|Set|If|For|Do|While|With|Select|On\s+Error|Application\.|ActiveSheet|ThisWorkbook|Worksheets|Sheets|Range|Cells|MsgBox|Debug\.|Const|Call|GoTo|Let)\b/im.test(code || '');
    }

    function ensureProcedureWrapper(code) {
        if (/\b(?:Sub|Function)\s+\S+/i.test(code)) return code;
        // 说明文字不能包进 Sub，否则看起来像模块，粘进去却无法编译
        if (!looksLikeVbaStatements(code)) return '';
        return `Sub GeneratedMacro()\n${code}\nEnd Sub`;
    }

    function stripVbaCommentsAndStrings(code) {
        return code
            .replace(/"[^"]*(?:""[^"]*)*"/g, '""')
            .replace(/'[^\n]*/g, '');
    }

    function ensureReferencedLabels(code) {
        const plain = stripVbaCommentsAndStrings(code);
        const needs = [];
        if (/\bGoTo\s+ErrHandler\b/i.test(plain) && !/^\s*ErrHandler\s*:/im.test(plain)) needs.push('err');
        if (/\bGoTo\s+CleanExit\b/i.test(plain) && !/^\s*CleanExit\s*:/im.test(plain)) needs.push('clean');
        // 必须锚定行首。\s 会跨行，把 "Exit Sub" 下一行误认成另一个过程，从而多补一个 End Sub。
        const subCount = (plain.match(/^[ \t]*(?:(?:Public|Private)[ \t]+)?Sub[ \t]+\S+/gim) || []).length;
        const endSubCount = (plain.match(/^[ \t]*End[ \t]+Sub\b/gim) || []).length;
        const missingEnd = Math.max(0, subCount - endSubCount);
        if (needs.length === 0 && missingEnd === 0) return code;

        let block = '';
        if (needs.includes('clean')) block += '\nCleanExit:\n    Exit Sub\n';
        if (needs.includes('err')) block += '\nErrHandler:\n    MsgBox "运行时错误 " & Err.Number & "：" & Err.Description, vbCritical\n';
        if (!block && missingEnd === 0) return code;

        const lines = code.replace(/\s+$/, '').split('\n');
        let insertAt = lines.length;
        for (let i = lines.length - 1; i >= 0; i--) {
            if (/^\s*End\s+Sub\b/i.test(lines[i])) { insertAt = i; break; }
        }
        if (block) lines.splice(insertAt, 0, block.replace(/^\n/, ''));
        let joined = lines.join('\n');
        for (let i = 0; i < missingEnd; i++) joined += '\nEnd Sub';
        return joined;
    }

    /**
     * 主入口：从 LLM 输出的 IntentSpec 渲染最终 VBA 代码
     */
    function render(llmOutput, userMessage, ctx = {}) {
        const intent = buildIntentFromLLM(llmOutput, userMessage, ctx);
        const tpl = pickTemplate(intent);
        if (tpl) {
            try {
                const code = toRunnableVba(tpl.render(intent, ctx));
                return { template: tpl.id, code, matched: true };
            } catch (e) {
                ChatLog.error?.('TPL', `[vba-template-engine] 模板 ${tpl.id} 渲染失败:`, e);
                return { template: null, code: '', matched: false, error: String(e?.message || e) };
            }
        }
        return { template: null, code: '', matched: false };
    }

    /**
     * 列出所有可用模板（供 UI 提示用户）
     */
    function listTemplates() {
        return TEMPLATES.map(t => ({ id: t.id, title: t.title, keywords: t.keywords }));
    }

    App.ChatV2.services.vbaTemplateEngine = {
        render,
        toRunnableVba,
        listTemplates,
        loadPptTemplates,
        resolveTheme,
        getThemeList,
        hasThemeLayout,
        getThemeLayoutSummary,
        TEMPLATES
    };
    ChatLog.info?.('TPL', '✅ vba-template-engine loaded: ' + TEMPLATES.length + ' templates');
})();
